# -*- coding: utf-8 -*-
"""
MercadoLibre seller-panel download flows, driven from the left-tab menu.

Two methods, both starting from the seller panel:

  download_billing_reports(driver, ...)
      Facturación -> latest N months -> Ir al detalle -> Reportes
      -> Seleccionar todos los reportes -> Descargar

      The number of files is NOT fixed: MercadoLibre only offers the report
      types that have data for that period, so select-all yields a different
      count per month. Observed for BOCINA_SM on 2026-08-28:
          Agosto (EN CURSO)  4  BILL_ML, FULL, PAYMENT, NC_ML
          Julio  (cerrado)   4  BILL_ML, FULL, PAYMENT, NC_ML
          Junio  (cerrado)   2  BILL_ML, PAYMENT      (only 2 types offered)
      Callers should not hard-code an expected count.

  download_sales_excel(driver, ...)
      Ventas -> remove the "Envíos de hoy" filter -> Descargar Excel de ventas
      -> wait for the notification widget -> Descargar on the newest entry.
      Yields exactly 1 file.

DOM notes that matter:
  * The Reportes panel is a micro-frontend rendered inside a SHADOW ROOT
    (section.remote-module.fbi-billing-fe-reporting-reporting). Ordinary
    querySelector / By.XPATH cannot see into it - every lookup here goes
    through the deep helpers below.
  * The select-all checkbox has a React-generated id (e.g. "«r2»") that changes
    between renders. It must be found by its label text, never by id.
  * The per-report ids ARE stable: BILL_ML, FULL, PAYMENT, NC_ML.
"""
import datetime
import os
import calendar
import time

BILLING_BASE = "https://vendedores.mercadolibre.com.mx"
BILLING_RESUME = BILLING_BASE + "/billing/resume"
VENTAS_LISTADO = "https://vendedores.mercadolibre.com.mx/ventas/omni/listado"
SPACE_MANAGEMENT = ("https://vendedores.mercadolibre.com.mx"
                    "/publicaciones/listado/space_management")

# Spanish month names, as they appear in the datepicker's aria-labels.
ES_MONTHS = [None, "enero", "febrero", "marzo", "abril", "mayo", "junio",
             "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

# ---------------------------------------------------------------- shadow DOM

# Every helper below recurses through open shadow roots. Returning the element
# itself lets Selenium wrap it as a WebElement, so we can issue a real click
# (React handles synthetic events from real clicks more reliably than from
# dispatchEvent).

JS_DEEP_QUERY = """
const sel = arguments[0];
const scan = (root) => {
  const hit = root.querySelector(sel);
  if (hit) return hit;
  for (const el of root.querySelectorAll('*')) {
    if (el.shadowRoot) { const r = scan(el.shadowRoot); if (r) return r; }
  }
  return null;
};
return scan(document);
"""

JS_DEEP_CHECKBOX_BY_LABEL = """
const want = arguments[0].toLowerCase();
const scan = (root) => {
  for (const el of root.querySelectorAll('input[type=checkbox]')) {
    if ((el.id||'').startsWith('nav-')) continue;
    const holder = el.closest('label,div,li,tr');
    if (((holder && holder.innerText) || '').toLowerCase().includes(want)) return el;
  }
  for (const el of root.querySelectorAll('*')) {
    if (el.shadowRoot) { const r = scan(el.shadowRoot); if (r) return r; }
  }
  return null;
};
return scan(document);
"""

JS_DEEP_CHECKBOX_STATE = """
const scan = (root, out) => {
  root.querySelectorAll('input[type=checkbox]').forEach(e => {
    if ((e.id||'').startsWith('nav-')) return;
    out.push({id: e.id||'', checked: e.checked,
              label: (e.closest('label,div')?.innerText||'').trim()
                     .replace(/\\s+/g,' ').slice(0,50)});
  });
  root.querySelectorAll('*').forEach(el => { if (el.shadowRoot) scan(el.shadowRoot, out); });
  return out;
};
return scan(document, []);
"""

JS_WIDGET_ROWS = """
return [...document.querySelectorAll('.process-notification-process')].map(e => ({
  title: (e.querySelector('.process-notification-process__content-title')?.innerText||'').trim(),
  sub:   (e.querySelector('.process-notification-process__content-subtitle')?.innerText||'').trim(),
  date:  (e.querySelector('.process-notification-process__content-date')?.innerText||'').trim(),
  hasDl: !!e.querySelector('.process-notification-link')
}));
"""


def deep_find(driver, css):
    return driver.execute_script(JS_DEEP_QUERY, css)


def deep_checkbox(driver, label_text):
    return driver.execute_script(JS_DEEP_CHECKBOX_BY_LABEL, label_text)


def click(driver, el):
    """Scroll into view, then real-click, falling back to a JS click."""
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    time.sleep(0.6)
    try:
        el.click()
    except Exception:
        driver.execute_script("arguments[0].click();", el)


# ---------------------------------------------------------------- downloads

def snapshot_dir(path):
    if not os.path.isdir(path):
        return {}
    return {f: os.path.getsize(os.path.join(path, f)) for f in os.listdir(path)}


def set_download_dir(driver, path):
    """Route downloads to `path` for this session via CDP.

    Without this, files land in the store's own downloadPath, which is shared
    across runs and makes counting new files unreliable.
    """
    if not os.path.isdir(path):
        os.makedirs(path)
    try:
        driver.execute_cdp_cmd("Browser.setDownloadBehavior", {
            "behavior": "allow", "downloadPath": path})
        return True
    except Exception as e:
        print("  [警告] 无法重定向下载目录：%s" % str(e)[:90])
        return False


def wait_for_downloads(path, before, min_files=1, timeout=240, settle=12, label=""):
    """Block until downloads have *settled*, and return the new filenames.

    "Select all reports" fires several downloads a few seconds apart, so
    returning at the first completed file loses the rest. There is no signal
    from the site saying "that was the last one", so completion is inferred
    from the download folder going quiet:

      - got everything we expected (`min_files`) and quiet for `settle`  -> return
      - fewer than expected, but quiet for `settle * 3`                  -> give up

    The second branch matters: `min_files` is a prediction (the number of report
    checkboxes ticked), not a promise. Without the escape hatch a wrong
    prediction would burn the whole `timeout` on every affected period.
    """
    deadline = time.time() + timeout
    seen = set()
    last_change = time.time()
    while time.time() < deadline:
        time.sleep(2)
        new = set(snapshot_dir(path)) - set(before)
        pending = [k for k in new if k.endswith((".crdownload", ".tmp"))]
        done = {k for k in new if not k.endswith((".crdownload", ".tmp"))}
        if done != seen:
            seen = done
            last_change = time.time()
        quiet = time.time() - last_change
        if seen and not pending and (
                (len(seen) >= min_files and quiet >= settle)
                or quiet >= settle * 3):
            return sorted(seen)
    if len(seen) < min_files:
        print("  [警告] %s：预期至少 %d 个文件，等了 %d 秒实际只有 %d 个"
              % (label or "download", min_files, timeout, len(seen)))
    return sorted(seen)


# ---------------------------------------------------- 0. 语言关卡

SELLER_HOME = BILLING_BASE + "/resumen"
WANT_LOCALE = "es_MX"

# 整个下载层靠可见的西语文字定位元素（"Ir al detalle"、"descargar"、
# "almacenamiento"、/ultimos?\s*(\d+)\s*mes/ …），导出文件的列名也是西语，
# 而**导出文件的语言跟着账号的界面语言走**。有人在浏览器里把语言改成 English，
# 后果分两种，后一种才可怕：
#
#   * 元素找不到 -> 抛错，这家店整个失败，看得见
#   * 账期页的月份名核对失败 -> **那一期被静默跳过**，报表少算费用而毫无提示。
#     只下目标月而漏掉后续账期，实测让 EWTTO_SM 少算佣金 973.23、代扣税率
#     从 9.172% 虚高到 9.221%。
#
# 与其把几十处选择器翻译成多语言（还要为每种语言维护一套列名表），不如在跑
# 任何路线之前先确认语言，不对就改回来。
#
# 关键点：判据用的是 `data-selected-locale="es_MX"` 这个 **locale 代码**，
# 不是可见文字。所以这道检查本身不会因为语言变了而失效 —— 用文字判断语言，
# 语言一变判断自己就先瞎了。
JS_GET_LOCALE = """
const el = document.querySelector('.nav-user-menu-language-switcher');
if (el && el.getAttribute('data-selected-locale'))
  return el.getAttribute('data-selected-locale');
const s = document.querySelector(
  '.nav-user-menu-language-switcher__option[aria-selected="true"]');
return s ? s.getAttribute('data-value') : null;
"""

# 切换器藏在用户菜单里（折叠状态），但 JS 点击不受可见性限制：先点触发器把
# 组件唤醒，再点目标选项。已经是目标语言时直接返回 already，不做无谓点击。
JS_SET_LOCALE = """
const want = arguments[0];
const root = document.querySelector('.nav-user-menu-language-switcher');
if (!root) return 'no-switcher';
if (root.getAttribute('data-selected-locale') === want) return 'already';
const t = document.getElementById('nav-user-menu-language-switcher-trigger');
if (t) t.click();
const o = root.querySelector('[data-value="' + want + '"]');
if (!o) return 'no-option';
o.click();
return 'clicked';
"""


def ensure_language(driver, want=WANT_LOCALE, timeout=30):
    """进店后、跑任何路线之前：确认卖家后台是西语，不是就切回去。

    返回 (进来时的 locale 或 None, 是否动过)。

    读不到切换器时**不拦路**，只发警告：那多半是页面改版，而为了一个读不到
    的判据把 12 家店全挡住，比它要防的问题更糟。真的改不回来才抛错 —— 那种
    情况下继续跑必然产出错语言的导出。
    """
    if not (driver.current_url or "").startswith(BILLING_BASE):
        driver.get(SELLER_HOME)
        time.sleep(6)
    cur = driver.execute_script(JS_GET_LOCALE)
    if cur is None:
        print("  [警告] 顶栏找不到语言切换器，无法确认界面语言。"
              "若平台改版，下载层的西语选择器可能已经失效。")
        return None, False
    if cur == want:
        print("  界面语言：%s ✓" % cur)
        return cur, False

    print("  [警告] 界面语言是 %s，不是 %s —— 导出文件的列名会跟着变，"
          "正在切回去。" % (cur, want))
    r = driver.execute_script(JS_SET_LOCALE, want)
    if r not in ("clicked", "already"):
        raise RuntimeError("切换界面语言失败（%s）：当前 %s，需要 %s。"
                           "请在浏览器里手动改回西班牙语。" % (r, cur, want))
    time.sleep(4)
    driver.get(SELLER_HOME)
    time.sleep(6)
    now = driver.execute_script(JS_GET_LOCALE)
    if now != want:
        raise RuntimeError("界面语言切换后仍是 %s（期望 %s）。"
                           "请在浏览器里手动改回西班牙语再跑。" % (now, want))
    print("  界面语言已从 %s 改回 %s ✓" % (cur, want))
    return cur, True


# ---------------------------------------------------- 1. Facturación reports

def _open_reports_tab(driver, timeout=90):
    """Select the Reportes tab and wait for its shadow content to render."""
    tab = deep_find(driver, "button.billing-detail_tab-reports")
    if tab is None:
        raise RuntimeError("Reportes tab not found")
    if "andes-tab--selected" not in (tab.get_attribute("class") or ""):
        click(driver, tab)

    deadline = time.time() + timeout
    while time.time() < deadline:
        boxes = driver.execute_script(JS_DEEP_CHECKBOX_STATE)
        if boxes:
            return boxes
        time.sleep(2)
    raise RuntimeError("Reportes panel never rendered")


JS_CARD_INFO = """
let e = arguments[0];
for (let i = 0; i < 8 && e; i++) {
  e = e.parentElement; if (!e) break;
  const txt = e.innerText || '';
  const m = txt.match(
    /(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Octubre|Noviembre|Diciembre)/i);
  if (m) return {month: m[1], enCurso: /EN CURSO/i.test(txt)};
}
return {month: '?', enCurso: false};
"""


def billing_periods(driver, months=2, skip_current=False):
    """Return the indices of the 'Ir al detalle' buttons to visit, newest first.

    `skip_current` drops the EN CURSO period - its charges are still accruing,
    so its reports are provisional.
    """
    driver.get(BILLING_RESUME)
    time.sleep(8)
    buttons = driver.find_elements("xpath", "//button[normalize-space()='Ir al detalle']")
    picked = []
    for i, b in enumerate(buttons):
        info = driver.execute_script(JS_CARD_INFO, b)
        tag = "EN CURSO" if info["enCurso"] else "cerrado"
        if skip_current and info["enCurso"]:
            print("  跳过 %s（%s）" % (info["month"], tag))
            continue
        picked.append((i, info["month"]))
        if len(picked) >= months:
            break
    return picked


def download_billing_reports(driver, download_dir, months=2, timeout=240,
                             skip_current=False):
    """Facturación -> newest `months` periods -> all reports -> Descargar.

    Returns {period: [filenames]}.
    """
    results = {}
    picked = billing_periods(driver, months=months, skip_current=skip_current)
    print("  本次下载的账期：%s" % ", ".join(m for _, m in picked))

    for idx, _month in picked:
        driver.get(BILLING_RESUME)
        time.sleep(8)

        buttons = driver.find_elements(
            "xpath", "//button[normalize-space()='Ir al detalle']")
        if idx >= len(buttons):
            print("  [警告] 只有 %d 个账期可选" % len(buttons))
            break

        month = driver.execute_script("""
            let e = arguments[0];
            for (let i = 0; i < 8 && e; i++) {
              e = e.parentElement; if (!e) break;
              const m = (e.innerText||'').match(
                /(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Octubre|Noviembre|Diciembre)/i);
              if (m) return m[1];
            }
            return '?';
        """, buttons[idx])

        print("\n--- Facturación [%d] %s：进入明细（Ir al detalle）---" % (idx, month))
        click(driver, buttons[idx])
        time.sleep(9)

        period = driver.current_url.rstrip("/").split("/")[-1].split("?")[0]
        print("    明细页：%s" % driver.current_url)

        files = _download_detail_reports(driver, download_dir, month, timeout)
        results["%s (%s)" % (month, period)] = files

    return results


SPANISH_MONTH_NAMES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
                       "Julio", "Agosto", "Septiembre", "Octubre",
                       "Noviembre", "Diciembre"]


def billing_detail_url(year, month):
    """某个会计月的账单明细页地址。

    明细页按**账期结束日**寻址，而结束日就是当月最后一天 —— 实测
    Agosto → 20260831、Septiembre → 20260930、Julio → 20260731 都成立。

    这条路径的意义：resume 页只摆最近 3 期卡片，更早的账期在页面上根本
    点不到，但明细页本身还在，直接访问就能拿到。做历史月份全靠它。
    """
    last = calendar.monthrange(year, month)[1]
    return "%s/billing/detail/%04d%02d%02d" % (BILLING_BASE, year, month, last)


def download_billing_for_months(driver, download_dir, year, month,
                                extra=2, timeout=240):
    """下载目标会计月**以及随后 extra 期**的账单。

    为什么不能只下目标月：订单级费用按**计费日**入账，月末几天的订单，佣金常常
    落在下一期账单上。实测 EWTTO_SM 的 7 月订单有 56 行费用记在 8/9 月账单里
    ——只下 7 月的话纯佣金少算 973.23，而代扣税是用「Ventas 合并值 − 纯佣金」
    倒推的，于是等额虚高，税率从 9.172% 变成 9.221%。

    超出当月的期数会被平台的页面校验挡掉（月份名对不上就跳过），所以这里
    不必自己算上限。
    """
    out = {}
    y, mo = year, month
    for i in range(max(1, extra + 1)):
        got = download_billing_for_period(driver, download_dir, y, mo, timeout)
        if not got and i > 0:
            break            # 后续期拿不到很正常（还没到/已关闭），不算失败
        out.update(got)
        mo += 1
        if mo > 12:
            y, mo = y + 1, 1
    return out


def download_billing_for_period(driver, download_dir, year, month, timeout=240):
    """下载**指定会计月**的全部账单报表，不管它还在不在 resume 页的卡片里。

    返回 {期间: [文件名]}；该账期不存在或页面打不开时返回 {}（不抛异常 ——
    历史月份取不到是常态，不该让整个店铺的下载失败）。
    """
    url = billing_detail_url(year, month)
    label = "%s %d" % (SPANISH_MONTH_NAMES[month - 1], year)
    print("\n--- Facturación %s：直接访问明细页 ---" % label)
    print("    %s" % url)
    driver.get(url)
    time.sleep(9)

    # 页面存在≠是我们要的那一期。用页面自己写的月份名核对，避免平台把
    # 无效日期悄悄重定向到最近一期、结果把 9 月的账单当成 7 月的存下来。
    shown = driver.execute_script("""
        const re = /(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Octubre|Noviembre|Diciembre)/i;
        const m = (document.body.innerText || '').match(re);
        return m ? m[1] : null;
    """)
    want = SPANISH_MONTH_NAMES[month - 1]
    if shown and shown.lower() != want.lower():
        print("    [警告] 页面显示的是 %s，不是 %s —— 该账期可能已不可访问，跳过"
              % (shown, want))
        return {}
    if not shown:
        print("    [警告] 页面没有可识别的月份，跳过")
        return {}

    period = url.rstrip("/").split("/")[-1]
    try:
        files = _download_detail_reports(driver, download_dir, want, timeout)
    except RuntimeError as e:
        print("    [警告] %s：%s" % (label, e))
        return {}
    return {"%s (%s)" % (want, period): files}


def _download_detail_reports(driver, download_dir, month, timeout=240):
    """已经站在某个账期的明细页上：勾选全部报表并下载。返回文件名列表。

    从 download_billing_reports 里抽出来，好让"按指定账期下载"能走同一段代码 ——
    两条路径的差别只在怎么到达明细页，页面上的操作完全一样。
    """
    _open_reports_tab(driver)

    cb = deep_checkbox(driver, "seleccionar todos los reportes")
    if cb is None:
        raise RuntimeError("select-all checkbox not found for %s" % month)
    # The select-all id is React-generated ("«r2»") and changes between
    # renders, so remember this element's own id rather than pattern
    # matching it out of the list below.
    select_all_id = cb.get_attribute("id")
    if not cb.is_selected():
        click(driver, cb)
        time.sleep(2)

    state = driver.execute_script(JS_DEEP_CHECKBOX_STATE)
    checked = [c["id"] for c in state
               if c["checked"] and c["id"] and c["id"] != select_all_id]
    expected = len(checked)
    print("    已勾选报表（%d）：%s" % (expected, checked or "（无）"))

    btn = None
    deadline = time.time() + 30
    while time.time() < deadline:
        btn = deep_find(driver, "button.billing-reporting-download__button")
        if btn is not None and btn.is_enabled():
            break
        time.sleep(1)
    if btn is None or not btn.is_enabled():
        raise RuntimeError("bulk Descargar never enabled for %s" % month)

    before = snapshot_dir(download_dir)
    click(driver, btn)
    print("    已点击 Descargar，等待文件落盘…")
    files = wait_for_downloads(download_dir, before,
                               min_files=max(expected, 1),
                               timeout=timeout, label="billing %s" % month)
    for f in files:
        print("      + %s" % f)
    if expected and len(files) != expected:
        print("    [警告] %s：勾选了 %d 张报表，实际得到 %d 个文件"
              % (month, expected, len(files)))
    return files


# --------------------------------------------------------- 2. Ventas Excel

# Accent-insensitive text match. The UI writes "Últimos 6 meses"; a plain
# lowercase compare fails on the accented U, so both sides are normalised.
JS_NORM = """
const norm = (s) => (s||'').normalize('NFD').replace(/[\\u0300-\\u036f]/g,'')
                    .toLowerCase().replace(/\\s+/g,' ').trim();
"""

JS_FIND_PERIOD_TRIGGER = """
// Verified live: the control is a combobox button inside .sc-dropdown-date-range
return document.querySelector('.sc-dropdown-date-range button.andes-dropdown__trigger')
    || document.querySelector('button.andes-dropdown__trigger');
"""

# NEVER match an option by its full text. The menu renders the label followed by
# the concrete date range it resolves to today:
#     "Últimos 2 meses  29 jun. al 29 ago."
#     "Últimos 6 meses  28 feb. al 29 ago."
# Those dates move every single day, so an exact compare works once and then
# silently stops finding the option. Match the label only.
JS_FIND_PERIOD_OPTION = JS_NORM + """
const re = new RegExp('ultimos?\\\\s*' + arguments[0] + '\\\\s*mes');
for (const el of document.querySelectorAll(
        'li[role=option], [role=option], .andes-list__item')) {
  if (re.test(norm(el.innerText))) return el;
}
return null;
"""

JS_LIST_PERIOD_OPTIONS = """
const out = [];
for (const el of document.querySelectorAll('li[role=option], [role=option]')) {
  const t = (el.innerText||'').trim().replace(/\\s+/g,' ');
  if (t && t.length < 60) out.push(t);
}
return [...new Set(out)];
"""

# The trigger's innerText repeats the label (a visually-hidden copy sits beside
# the visible one, giving "Últimos 2 meses Últimos 2 meses"), so read the
# display span instead of the button.
JS_CURRENT_PERIOD = JS_NORM + """
const el = document.querySelector(
      '.sc-dropdown-date-range .andes-dropdown__display-values')
   || document.querySelector('.andes-dropdown__display-values');
if (!el) return null;
const m = norm(el.innerText).match(/ultimos?\\s*(\\d+)\\s*mes/);
return m ? parseInt(m[1], 10) : null;
"""


def set_sales_period(driver, months=6, timeout=45):
    """Ventas -> click the period dropdown -> pick "Últimos N meses".

    The human flow is: click "Últimos 2 meses", a menu appears, click
    "Últimos 6 meses". There is no URL parameter for this that survives a
    reload, so it has to be driven through the menu.

    Returns True if the period reads N months afterwards.
    """
    want_label = "Últimos %d meses" % months

    current = driver.execute_script(JS_CURRENT_PERIOD)
    if current == months:
        print("    时间范围已是 %s" % want_label)
        return True
    print("    当前时间范围为 %s 个月，切换为 %d 个月" % (current or "?", months))

    trigger = driver.execute_script(JS_FIND_PERIOD_TRIGGER)
    if trigger is None:
        print("    [警告] 未找到时间范围下拉框，保持原值不变")
        return False
    click(driver, trigger)
    time.sleep(2)

    option = None
    deadline = time.time() + 15
    while time.time() < deadline:
        option = driver.execute_script(JS_FIND_PERIOD_OPTION, months)
        if option is not None:
            break
        time.sleep(1)
    if option is None:
        print("    [警告] 菜单中没有 %r。可选项为：%s"
              % (want_label, driver.execute_script(JS_LIST_PERIOD_OPTIONS)))
        return False
    print("    选择：%s" % option.text.replace("\n", " ").strip()[:52])

    click(driver, option)

    # the list reloads; wait for the trigger to actually read the new period
    deadline = time.time() + timeout
    while time.time() < deadline:
        time.sleep(2)
        if driver.execute_script(JS_CURRENT_PERIOD) == months:
            print("    时间范围已设为 %s" % want_label)
            time.sleep(6)          # let the order list finish reloading
            return True
    print("    [警告] 时间范围未能切换到 %s" % want_label)
    return False


# 最近一次 request_sales_excel 实际生效的时间范围。用模块级变量而不是改返回值，
# 是因为返回值的形状被 run_downloads 的 pending 机制约束着，改它要牵动收取逻辑。
SALES_PERIOD_ACTUAL = {"months": None}


def request_sales_excel(driver, period_months=6):
    """PHASE 1 - ask MercadoLibre to build the sales Excel, do not wait.

    Generation is server-side (~25-30s) and the file appears in the corner
    notification widget when done, so the wait can overlap with other routes.
    Returns a pending dict, or None if the button never became available.
    """
    print("\n--- Ventas：申请 Excel de ventas ---")
    driver.get(VENTAS_LISTADO)
    time.sleep(12)

    # remove the "Envíos de hoy" applied filter; without this the list is
    # today's shipments only (often 0 sales) and the Excel button stays disabled
    removed = driver.execute_script("""
        for (const t of document.querySelectorAll('[data-testid="filters--applied"] .andes-tag')) {
          const lbl = (t.querySelector('.andes-tag__label')?.innerText||'').trim();
          if (/env[ií]os de hoy/i.test(lbl)) {
            const b = t.querySelector('button.andes-tag__close');
            if (b) { b.click(); return lbl; }
          }
        }
        return null;
    """)
    print("    已移除筛选：%s" % (removed or "（无已应用筛选）"))
    time.sleep(10)

    # widen the period AFTER clearing the filter - the dropdown re-renders when
    # filters change, and a handle grabbed before that goes stale
    # 读回**实际**生效的窗口，而不是我们请求的那个。菜单里没有的档位会被
    # set_sales_period 拒绝并保持原值（实测要 3 拿到的是 2），而这一步失败
    # 之后导出照样能跑，只是少了一大段数据 —— 不读回来就完全看不出。
    if period_months:
        set_sales_period(driver, months=period_months)
    actual = driver.execute_script(JS_CURRENT_PERIOD)
    if period_months and actual != period_months:
        print("    [警告] 请求 %s 个月，实际生效 %s 个月 —— 导出会少一段数据"
              % (period_months, actual))
    SALES_PERIOD_ACTUAL["months"] = actual

    btn = None
    deadline = time.time() + 90
    while time.time() < deadline:
        btn = driver.execute_script("""
            return [...document.querySelectorAll('button.report-link')]
              .find(e => /Descargar Excel de ventas/i.test(
                  e.getAttribute('aria-label')||e.innerText||'')) || null;
        """)
        if btn is not None and btn.is_enabled():
            break
        time.sleep(2)
    if btn is None or not btn.is_enabled():
        print("    [警告] 'Descargar Excel de ventas' 始终不可点击")
        return None

    rows_before = len(driver.execute_script(JS_WIDGET_ROWS))
    click(driver, btn)
    print("    已提交申请，文件在服务端继续生成")
    return {"kind": "ventas", "rows_before": rows_before,
            "requested_at": time.time()}


def collect_sales_excel(driver, pending, download_dir, timeout=600, poll=15):
    """PHASE 3 - return to Ventas and download the generated Excel.

    Polls up to `timeout` seconds printing status. Not-ready is reported, not
    raised: the file stays in the widget and can be collected on a later run.
    """
    waited = time.time() - pending.get("requested_at", time.time())
    print("\n--- Ventas：取回 Excel de ventas ---")
    print("    距提交申请已过去 %.0f 秒" % waited)

    if "ventas/omni" not in driver.current_url:
        driver.get(VENTAS_LISTADO)
        time.sleep(12)

    before = snapshot_dir(download_dir)
    rows_before = pending.get("rows_before", 0)

    ready = False
    deadline = time.time() + timeout
    while time.time() < deadline:
        rows = driver.execute_script(JS_WIDGET_ROWS)
        left = int(deadline - time.time())
        if len(rows) > rows_before and rows[0]["hasDl"] and "xito" in rows[0]["sub"]:
            print("    [剩余 %4d 秒] 就绪：%s | %s"
                  % (left, rows[0]["date"], rows[0]["sub"]))
            ready = True
            break
        top = rows[0]["sub"] if rows else "(no entries)"
        print("    [剩余 %4d 秒] %s" % (left, top))
        time.sleep(poll)

    if not ready:
        print("    [警告] %d 秒内没有新的成功记录，改用最新一条"
              % timeout)

    clicked = driver.execute_script("""
        const row = document.querySelector('.process-notification-process');
        if (!row) return null;
        const link = row.querySelector('.process-notification-link');
        if (!link) return null;
        link.scrollIntoView({block:'center'});
        link.click();
        return (row.querySelector('.process-notification-process__content-date')?.innerText||'').trim();
    """)
    if clicked is None:
        print("    [警告] 最新一条通知里没有 Descargar 链接")
        return []
    print("    已在第 %s 条记录上点击 Descargar" % clicked)

    files = wait_for_downloads(download_dir, before, min_files=1,
                               timeout=180, settle=8, label="ventas excel")
    for f in files:
        print("      + %s" % f)
    return files


def download_sales_excel(driver, download_dir, timeout=300, period_months=6):
    """Request and collect in one call - for running the Ventas route alone.

    The split version above is what a full run uses, so the generation wait
    overlaps with the other routes.
    """
    pending = request_sales_excel(driver, period_months=period_months)
    if not pending:
        return []
    return collect_sales_excel(driver, pending, download_dir, timeout=timeout)



# ------------------------------------------------- 3. Stock reports (Full)

# Publicaciones -> Control de stock -> "Descargar reportes de stock" -> 4 reports.
#
# Verified live 2026-08-29. Each behaves differently:
#
#   a. Reporte de costos por el servicio de almacenamiento
#         a "Periodo" dropdown of billing periods; pick the newest, then Descargar
#   b. Reporte general de stock
#         NO period control - it is a point-in-time snapshot ("Actualizado el
#         29 de agosto"), so a date range would be meaningless. Downloads on click.
#   c. Reporte consolidado de movimientos
#         a range datepicker, then Aplicar, then Descargar
#   d. Reporte de devoluciones a la bodega
#         downloads on click
#
# Three mechanics that are easy to get wrong:
#
#   * The "Descargar reportes de stock" dropdown opens ONLY via a JS click. A
#     native Selenium click and an ActionChains click both leave
#     aria-expanded="false". (The billing Reportes tab is the opposite - it
#     needs a real click. Do not "harmonise" these.)
#   * Each menu <li> is inert; the real target is an empty overlay
#     button.andes-list__item-actionable inside it.
#   * In the range picker, pick the EARLIER date first. Once a start is
#     selected every earlier day becomes --disabled, so doing it backwards
#     silently selects only one end of the range.

STOCK_REPORTS = [
    ("costos",       "reporte de costos por el servicio de almacenamiento"),
    ("general",      "reporte general de stock"),
    ("consolidado",  "reporte consolidado de movimientos"),
    ("devoluciones", "reporte de devoluciones a la bodega"),
]

JS_OPEN_STOCK_MENU = JS_NORM + """
for (const el of document.querySelectorAll('button')) {
  if (/descargar reportes de stock/.test(norm(el.innerText))) {
    el.click();               // JS click only - see the note above
    return true;
  }
}
return false;
"""

# MercadoLibre changed this menu's markup: the entries used to be
# li.andes-button-dropdown__menu-item and are now plain li.andes-list__item.
# Accept either, or the whole stock route silently returns "menu has: []" while
# the menu is in fact open (aria-expanded="true") right in front of it.
JS_STOCK_MENU_SEL = ("li.andes-button-dropdown__menu-item, "
                     "li.andes-list__item")

JS_STOCK_ITEMS = """
return [...document.querySelectorAll('%s')]
       .map(e => (e.innerText||'').trim())
       .filter(t => /reporte/i.test(t));
""" % JS_STOCK_MENU_SEL

JS_CLICK_STOCK_ITEM = JS_NORM + """
const want = norm(arguments[0]);
for (const li of document.querySelectorAll('%s')) {
  if (norm(li.innerText).indexOf(want) === 0) {
    (li.querySelector('button.andes-list__item-actionable') || li).click();
    return true;
  }
}
return false;
""" % JS_STOCK_MENU_SEL

# Find the modal by its heading text, so we never grab the page-level
# "Descargar reportes de stock" button sitting behind the overlay.
JS_MODAL_BUTTON = """
const key = arguments[0].toLowerCase();
for (const el of document.querySelectorAll('[role=dialog], .andes-modal')) {
  if (!(el.innerText||'').toLowerCase().includes(key)) continue;
  for (const b of el.querySelectorAll('button')) {
    if (/descargar/i.test(b.innerText||'')) return {found: true, disabled: b.disabled};
  }
}
return {found: false, disabled: true};
"""

JS_CLICK_MODAL_BUTTON = """
const key = arguments[0].toLowerCase();
for (const el of document.querySelectorAll('[role=dialog], .andes-modal')) {
  if (!(el.innerText||'').toLowerCase().includes(key)) continue;
  for (const b of el.querySelectorAll('button')) {
    if (/descargar/i.test(b.innerText||'')) { b.click(); return true; }
  }
}
return false;
"""

JS_OPEN_PERIOD_DROPDOWN = """
for (const el of document.querySelectorAll('[role=dialog],.andes-modal')) {
  if (!/almacenamiento/i.test(el.innerText||'')) continue;
  const t = el.querySelector('button.andes-dropdown__trigger');
  if (t) { t.click(); return true; }
}
return false;
"""

# Billing-period options read like "Del 1 de julio al 31 de julio del 2026".
# The Espanol / English / Chinese entries in the DOM belong to the nav language
# switcher, so match the date shape rather than taking the first list item.
# Two shapes appear in this dropdown, and which ones are present varies:
#   "Período actual - ..."                     the open, still-accruing period
#   "Del 1 de julio al 31 de julio del 2026"   a closed period
# The list is already newest-first, so DOM order is preserved and the caller
# just takes the first N. On 2026-08-31 only the two closed periods were
# offered - no "Período actual" - so nothing may assume it exists.
JS_LIST_PERIODS = """
const re = /^(per[ií]odo\\s+actual|del\\s+\\d+\\s+de\\s+\\w+)/i;
const out = [];
for (const el of document.querySelectorAll('li[role=option], li.andes-list__item')) {
  const t = (el.innerText||'').trim().replace(/\\s+/g, ' ');
  if (re.test(t)) out.push(t);
}
return [...new Set(out)];
"""

# Pick by text rather than by index: the option list is rebuilt each time the
# modal opens, and an index would silently drift if the newest period rolls
# over between iterations.
#
# Whitespace MUST be collapsed on both sides. JS_LIST_PERIODS normalises, and
# the "Período actual" entry is multi-line ("... del 2026\\nAcumulas $ 415.81
# hasta el momento"), so comparing against raw innerText never matches it - the
# current period is silently skipped while the closed ones still work.
JS_PICK_PERIOD = """
const flat = (s) => (s||'').trim().replace(/\\s+/g, ' ');
const want = flat(arguments[0]);
for (const el of document.querySelectorAll('li[role=option], li.andes-list__item')) {
  if (flat(el.innerText) === want) {
    (el.querySelector('button.andes-list__item-actionable') || el).click();
    return true;
  }
}
return false;
"""

JS_CAL_CAPTION = """
const e = document.querySelector('.andes-datepicker__caption-label');
return e ? e.innerText.trim() : null;
"""

JS_CAL_NAV = """
const b = document.querySelector(arguments[0]);
if (!b) return false;
b.click();
return true;
"""

# aria-label is "sabado 29 de agosto de 2026 " - endsWith on the trimmed label
# is exact and needs no regex escaping.
JS_CAL_PICK = """
const want = arguments[0] + ' de ' + arguments[1] + ' de ' + arguments[2];
for (const b of document.querySelectorAll('button.andes-datepicker__day')) {
  if ((b.getAttribute('aria-label')||'').trim().endsWith(want)) {
    const cell = b.closest('.andes-datepicker__cell');
    if (cell && /--disabled/.test(cell.className)) return 'disabled';
    b.click();
    return 'ok';
  }
}
return 'not-found';
"""

JS_CAL_FIRST_ENABLED = """
for (const b of document.querySelectorAll('button.andes-datepicker__day')) {
  const cell = b.closest('.andes-datepicker__cell');
  if (cell && /--disabled/.test(cell.className)) continue;
  return parseInt((b.innerText||'').trim(), 10);
}
return null;
"""

JS_OPEN_DATEPICKER = """
const t = document.querySelector('#report-date-picker__trigger');
if (!t) return false;
t.click();
return true;
"""

JS_CLICK_APLICAR = """
for (const b of document.querySelectorAll('button')) {
  if (b.getBoundingClientRect().width > 4 &&
      /^aplicar$/i.test((b.innerText||'').trim())) { b.click(); return true; }
}
return false;
"""


def _cal_goto(driver, year, month, tries=18):
    """Walk the single-month calendar to year/month."""
    want = "%s %d" % (ES_MONTHS[month], year)
    for _ in range(tries):
        cap = (driver.execute_script(JS_CAL_CAPTION) or "").lower()
        if cap == want:
            return True
        cur_m = next((i for i, n in enumerate(ES_MONTHS) if n and cap.startswith(n)), None)
        parts = cap.split()
        cur_y = int(parts[-1]) if parts and parts[-1].isdigit() else year
        back = (cur_y, cur_m or 0) > (year, month)
        driver.execute_script(
            JS_CAL_NAV,
            ".andes-datepicker__button--previous" if back
            else ".andes-datepicker__button--next")
        time.sleep(1.2)
    return False


def _select_date_range(driver, months_back=2):
    """Select (today - months_back) .. today in the range picker.

    The picker enforces its own floor - on 2026-08-29 every day before 1 July
    was disabled - so a requested start that is out of range is clamped forward
    to the earliest day the calendar actually offers, rather than failing or
    silently selecting only one end of the range.
    """
    today = datetime.date.today()
    m, y = today.month - months_back, today.year
    while m <= 0:
        m += 12
        y -= 1
    start_day = min(today.day, 28)

    if not driver.execute_script(JS_OPEN_DATEPICKER):
        print("      [警告] 未找到日期选择器")
        return False
    time.sleep(3)

    picked_start = None
    for bump in range(months_back + 1):
        mm, yy = m + bump, y
        while mm > 12:
            mm -= 12
            yy += 1
        if not _cal_goto(driver, yy, mm):
            continue

        # Only the originally requested month keeps the requested day. Once we
        # have had to skip forward, the target day is arbitrary - taking it
        # would throw away the front of the month (June disabled + day 28 gave
        # 28 Jul instead of 1 Jul, halving the range), so take the earliest day
        # the calendar allows.
        if bump == 0:
            if driver.execute_script(JS_CAL_PICK, start_day,
                                     ES_MONTHS[mm], yy) == "ok":
                picked_start = "%04d-%02d-%02d" % (yy, mm, start_day)
                break
        first = driver.execute_script(JS_CAL_FIRST_ENABLED)
        if first and driver.execute_script(
                JS_CAL_PICK, first, ES_MONTHS[mm], yy) == "ok":
            picked_start = "%04d-%02d-%02d" % (yy, mm, first)
            print("      起始日期收敛为 %s（更早的日期不可选）" % picked_start)
            break
    if not picked_start:
        print("      [警告] 没有可选的起始日期")
        return False
    time.sleep(1)

    _cal_goto(driver, today.year, today.month)
    res = driver.execute_script(JS_CAL_PICK, today.day,
                                ES_MONTHS[today.month], today.year)
    if res != "ok":
        print("      [警告] 结束日期 %s -> %s" % (today, res))
    print("      区间：%s .. %s" % (picked_start, today))
    time.sleep(1)

    driver.execute_script(JS_CLICK_APLICAR)
    time.sleep(3)
    return True



def _open_stock_menu(driver, timeout=60):
    """Open the stock-reports dropdown, waiting for the page to be ready.

    A fixed sleep is not enough: this page finishes rendering late, and under
    load a 12s wait left the button absent, so the whole route returned
    "menu has: []" and downloaded nothing. Poll for the menu ITEMS, not just
    the trigger - the trigger can exist while the list is still empty.
    """
    deadline = time.time() + timeout
    while time.time() < deadline:
        if driver.execute_script(JS_OPEN_STOCK_MENU):
            time.sleep(3)
            if driver.execute_script(JS_STOCK_ITEMS):
                return True
        time.sleep(3)
    return False


def _download_storage_costs(driver, download_dir, label, periods=2, timeout=240):
    """Report (a), once per billing period, newest first.

    The modal only holds one period at a time, so this reopens it per period
    rather than trying to multi-select.
    """
    out = []
    known = None
    for i in range(periods):
        driver.get(SPACE_MANAGEMENT)
        time.sleep(12)
        before = snapshot_dir(download_dir)

        if not _open_stock_menu(driver):
            print("      [警告] 库存报表菜单始终为空")
            break
        if not driver.execute_script(JS_CLICK_STOCK_ITEM, label):
            print("      [警告] 未找到该菜单项")
            break
        time.sleep(5)

        driver.execute_script(JS_OPEN_PERIOD_DROPDOWN)
        time.sleep(3)
        available = driver.execute_script(JS_LIST_PERIODS)
        if known is None:
            known = available
            print("      可选账期：%d 个 -> %s" % (len(known), known))
        if i >= len(available):
            print("      仅有 %d 个账期可选，停止" % len(available))
            break

        want = available[i]
        if not driver.execute_script(JS_PICK_PERIOD, want):
            print("      [警告] 无法选中 %r" % want)
            continue
        print("      [%d/%d] %s" % (i + 1, min(periods, len(available)), want))
        time.sleep(2)

        st = driver.execute_script(JS_MODAL_BUTTON, "almacenamiento")
        if not st["found"] or st["disabled"]:
            print("      [警告] 该账期的 Descargar 不可用")
            continue
        driver.execute_script(JS_CLICK_MODAL_BUTTON, "almacenamiento")

        files = wait_for_downloads(download_dir, before, min_files=1,
                                   timeout=timeout, settle=6,
                                   label="stock/costos[%d]" % i)
        for f in files:
            print("        + %s" % f)
        out.extend(files)
    return out


def download_stock_reports(driver, download_dir, months_back=2, timeout=240,
                           only=None, storage_periods=2):
    """Download the four stock reports. Returns {key: [filenames]}.

    `storage_periods` - how many billing periods to pull for report (a),
    newest first. 2 gives the newest and the one below it.
    """
    print("\n--- Publicaciones：库存报表（reportes de stock）---")
    results = {}

    for key, label in STOCK_REPORTS:
        if only and key not in only:
            continue

        if key == "costos":
            print("    [costos] %s" % label)
            results[key] = _download_storage_costs(
                driver, download_dir, label,
                periods=storage_periods, timeout=timeout)
            continue

        driver.get(SPACE_MANAGEMENT)
        time.sleep(12)
        before = snapshot_dir(download_dir)

        if not _open_stock_menu(driver):
            print("    [警告] 库存报表菜单始终为空，跳过 %s" % key)
            results[key] = []
            continue

        if not driver.execute_script(JS_CLICK_STOCK_ITEM, label):
            print("    [警告] 未找到菜单项 %r，当前菜单为：%s"
                  % (label, driver.execute_script(JS_STOCK_ITEMS)))
            results[key] = []
            continue
        print("    [%s] %s" % (key, label))
        time.sleep(5)

        if key == "consolidado":
            _select_date_range(driver, months_back=months_back)
            # only (c) still needs its modal's Descargar; (b) and (d) download
            # on click, and (a) is handled by _download_storage_costs above
            st = driver.execute_script(JS_MODAL_BUTTON, "consolidado")
            if not st["found"]:
                print("      [警告] 弹窗里没有 Descargar 按钮")
            elif st["disabled"]:
                print("      [警告] Descargar 仍不可点击，账期未被接受")
            else:
                driver.execute_script(JS_CLICK_MODAL_BUTTON, "consolidado")

        files = wait_for_downloads(download_dir, before, min_files=1,
                                   timeout=timeout, settle=6,
                                   label="stock/%s" % key)
        for f in files:
            print("      + %s" % f)
        results[key] = files

    return results


def probe_sales_excel(driver, pending, settle=10):
    """Cheap readiness check for the Ventas Excel. Leaves the driver on the page."""
    if "ventas/omni" not in driver.current_url:
        driver.get(VENTAS_LISTADO)
        time.sleep(settle)
    else:
        driver.refresh()
        time.sleep(settle)
    rows = driver.execute_script(JS_WIDGET_ROWS)
    return bool(rows
                and len(rows) > pending.get("rows_before", 0)
                and rows[0]["hasDl"]
                and "xito" in rows[0]["sub"])


def fetch_sales_excel(driver, pending, download_dir):
    """Download the generated Excel. Assumes probe_sales_excel() just passed."""
    before = snapshot_dir(download_dir)
    clicked = driver.execute_script("""
        const row = document.querySelector('.process-notification-process');
        if (!row) return null;
        const link = row.querySelector('.process-notification-link');
        if (!link) return null;
        link.scrollIntoView({block:'center'});
        link.click();
        return (row.querySelector('.process-notification-process__content-date')?.innerText||'').trim();
    """)
    if clicked is None:
        print("      [警告] 最新一条通知里没有 Descargar 链接")
        return []
    files = wait_for_downloads(download_dir, before, min_files=1,
                               timeout=180, settle=8, label="ventas excel")
    for f in files:
        print("      + %s" % f)
    return files
