# -*- coding: utf-8 -*-
"""
【历史数据补采·独立脚本】尽可能往回拉一家店的全部数据。

    python download_history.py --store EWTTO_SM
    python download_history.py --store EWTTO_SM --probe-only   # 只探查上限，不下载
    python download_history.py --store EWTTO_SM --since 2025-01-01

与日常跑批的区别
----------------
`run_downloads.py` 是**按月例行**：销售取最近 6 个月、账单取最近 2 期、MercadoPago
取最近 60 天。本脚本是**一次性补历史**：每一项都往平台允许的最远处拉。

    项目            日常          本脚本
    销售 Ventas     Últimos 6 meses   Último año（或 --since 指定的自定义日期）
    账单            最近 2 期      平台提供的**全部**期（实测 EWTTO_SM 只有 3 期）
    MercadoPago     60 天          365 天（日历会自动收敛到平台允许的最早日期）
    库存            当前快照       同左 —— 库存是时点数，没有"历史"可补

**不修改任何现有代码。** 全部通过 import 复用 `meli_forms` / `mercadopago` /
`ziniao_client` 的现成函数与 JS 常量。唯一自己实现的是"按标签选时间范围"——因为
`meli_forms.set_sales_period()` 只会拼 `"Últimos N meses"`，够不到 `Último año`
和 `Fecha personalizada`，而那个函数正在日常跑批里用，不能动。

为什么要尽快跑
--------------
账单是**滚动**的：实测 EWTTO_SM 的账单页只保留 3 期（Septiembre / Agosto /
Julio）。到 10 月初 Julio 就会掉出去，**永久拿不回来**。销售虽然能拉一年，同样
是滚动窗口，越早存下来越好。
"""
import argparse
import console  # noqa: F401  中文输出编码保护
import datetime
import os
import sys
import time

from ziniao_client import ZiniaoClient, ZiniaoError, StoreNotAllowed
from store_config import StoreConfigError
import store_config
import meli_forms as mf
import mercadopago as mp
import run_downloads          # 复用 run_folder / collect_pending

ROOT = os.path.dirname(os.path.abspath(__file__))


# ════════════════════════════════════════════════════════════════════
# 时间范围：按**标签**选，而不是按"N 个月"拼
# ════════════════════════════════════════════════════════════════════

def list_period_options(driver):
    """打开时间范围下拉，返回全部可选项文本。菜单不展开就读不到。"""
    trig = driver.execute_script(mf.JS_FIND_PERIOD_TRIGGER)
    if trig is None:
        return [], None
    mf.click(driver, trig)
    time.sleep(3)
    opts = driver.execute_script(mf.JS_LIST_PERIOD_OPTIONS) or []
    return [str(o).replace("\n", " ").strip() for o in opts], trig


def pick_period_by_label(driver, want, timeout=20):
    """在时间范围菜单里点选包含 `want` 的那一项（不区分重音与大小写）。

    `meli_forms.set_sales_period()` 只认 "Últimos N meses"，拼不出 "Último año"。
    那个函数在日常跑批里用着，不能改，所以这里自己按标签匹配 —— 复用它的
    JS_FIND_PERIOD_TRIGGER 定位下拉，只是匹配规则换成子串。
    """
    import unicodedata
    norm = lambda s: "".join(
        c for c in unicodedata.normalize("NFD", str(s).lower())
        if unicodedata.category(c) != "Mn")
    target = norm(want)

    opts, trig = list_period_options(driver)
    if trig is None:
        print("    [警告] 未找到时间范围下拉框")
        return False
    hit = [o for o in opts if target in norm(o)]
    if not hit:
        print("    [警告] 菜单里没有 %r。可选：%s" % (want, "、".join(opts)))
        return False
    print("    选择：%s" % hit[0][:64])

    el = driver.execute_script("""
        const want = arguments[0];
        const norm = s => s.normalize('NFD').replace(/[\\u0300-\\u036f]/g,'').toLowerCase();
        for (const e of document.querySelectorAll('li[role=option], li.andes-list__item, li, button')) {
          const t = (e.innerText||'').trim().replace(/\\s+/g,' ');
          if (t && norm(t).indexOf(norm(want)) >= 0
              && e.getBoundingClientRect().width > 4) return e;
        }
        return null;
    """, hit[0][:40])
    if el is None:
        print("    [警告] 找到了选项文本却拿不到可点击元素")
        return False
    mf.click(driver, el)

    # 列表要重新加载；等到触发器文案真的变了为止
    deadline = time.time() + timeout
    while time.time() < deadline:
        time.sleep(2)
        cur = driver.execute_script("""
            const b = document.querySelector('[data-testid=date-filter] button, .andes-dropdown__trigger');
            return b ? (b.innerText||'').trim().replace(/\\s+/g,' ') : '';
        """)
        if cur and target in norm(cur):
            print("    时间范围已设为 %s" % hit[0][:40])
            time.sleep(6)
            return True
    print("    [提示] 未能确认触发器文案已更新，继续尝试导出")
    time.sleep(4)
    return True


def request_sales_excel_history(driver, period_label):
    """申请销售 Excel，但时间范围用 `period_label`（如 "Último año"）。

    流程完全照搬 meli_forms.request_sales_excel：先清掉 "Envíos de hoy" 筛选
    （不清的话列表只剩今天的发货、导出按钮是灰的），**再**改时间范围（筛选一变
    下拉会重渲染，先拿的句柄会失效），最后点导出。
    """
    print("\n--- Ventas：申请 Excel de ventas（%s）---" % period_label)
    driver.get(mf.VENTAS_LISTADO)
    time.sleep(12)

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

    if not pick_period_by_label(driver, period_label):
        print("    [警告] 时间范围没设成 %s，导出的仍是默认范围" % period_label)

    btn, deadline = None, time.time() + 90
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

    rows_before = len(driver.execute_script(mf.JS_WIDGET_ROWS))
    mf.click(driver, btn)
    print("    已提交申请，文件在服务端继续生成")
    return {"kind": "ventas", "rows_before": rows_before,
            "requested_at": time.time()}


# ════════════════════════════════════════════════════════════════════
# 探查：平台侧到底能给多少
# ════════════════════════════════════════════════════════════════════

def probe_limits(driver):
    """账单有几期、销售能选到多久。只读，不下载。"""
    out = {}
    driver.get(mf.BILLING_RESUME)
    time.sleep(16)
    btns = driver.find_elements("xpath", "//button[normalize-space()='Ir al detalle']")
    months = []
    for b in btns:
        t = driver.execute_script("""
            let e=arguments[0];
            for(let i=0;i<8&&e;i++){e=e.parentElement;if(!e)break;
              const m=(e.innerText||'').replace(/\\s+/g,' ').match(
                /(Enero|Febrero|Marzo|Abril|Mayo|Junio|Julio|Agosto|Septiembre|Octubre|Noviembre|Diciembre)\\s*\\d*/i);
              if(m)return m[0].trim();}
            return '?';""", b)
        months.append(t)
    out["billing_periods"] = months
    print("【账单】平台提供 %d 期：%s" % (len(months), "、".join(months) or "(无)"))

    driver.get(mf.VENTAS_LISTADO)
    time.sleep(16)
    opts, _ = list_period_options(driver)
    # 语言切换项也在同一个菜单里，滤掉
    opts = [o for o in opts if o not in ("Español", "English", "中文")]
    out["sales_periods"] = opts
    print("【销售】时间范围可选 %d 项：" % len(opts))
    for o in opts:
        print("     %s" % o[:70])
    return out


# ════════════════════════════════════════════════════════════════════

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--store", default="EWTTO_SM")
    ap.add_argument("--out", default=None, help="输出目录；默认 downloads/<店>/<时间戳>_history")
    ap.add_argument("--probe-only", action="store_true", help="只探查上限，不下载")
    ap.add_argument("--sales-period", default="Último año",
                    help='销售时间范围标签，默认 "Último año"')
    ap.add_argument("--billing-months", type=int, default=12,
                    help="尝试取多少期账单；平台给不了那么多时自动截止（默认 12）")
    ap.add_argument("--mp-days", type=int, default=365,
                    help="MercadoPago 回溯天数，日历会自动收敛到允许的最早日期")
    ap.add_argument("--collect-timeout", type=int, default=1800)
    ap.add_argument("--collect-poll", type=int, default=20)
    ap.add_argument("--skip-mp", action="store_true", help="跳过 MercadoPago")
    ap.add_argument("--skip-stock", action="store_true", help="跳过库存")
    ap.add_argument("--no-restart", action="store_true")
    args = ap.parse_args()

    store = store_config.require_allowed(args.store)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = os.path.abspath(args.out or os.path.join(
        ROOT, "downloads", store, stamp + "_history"))

    print("=" * 66)
    print("历史数据补采 —— %s" % store)
    print("=" * 66)

    client = ZiniaoClient()
    client.start(restart=not args.no_restart)
    st = client.find_store(store)

    result = {"store": store, "out_dir": out_dir, "billing": {}, "sales": [],
              "stock": {}, "mp": [], "pending": [], "empty": [], "errors": []}

    with client.open_store(st, close_when_done=True) as session:
        d = session.driver
        d.get(session.launcher_page)

        if args.probe_only:
            probe_limits(d)
            d.get(session.launcher_page)
            print("\n--probe-only：未下载任何文件。")
            return 0

        if not os.path.isdir(out_dir):
            os.makedirs(out_dir)
        print("输出目录：%s" % out_dir)
        mf.set_download_dir(d, out_dir)

        limits = probe_limits(d)
        n_bill = min(args.billing_months, len(limits["billing_periods"])) or 1

        # ---- 阶段一：先把要排队生成的都提交出去 ----
        pending = []
        p = run_downloads._guard(result, "ventas/request",
                                 request_sales_excel_history, d, args.sales_period)
        if p and p is not run_downloads.FAILED:
            pending.append(p)
        if not args.skip_mp:
            for kind in mp.REPORTS:
                try:
                    q = run_downloads._guard(
                        result, "mercadopago/%s/request" % kind,
                        mp.request_report, d, report=kind, days=args.mp_days)
                    if q and q is not run_downloads.FAILED:
                        pending.append(q)
                except mp.NoReportAccess as e:
                    print("  [错误] mercadopago：%s" % e)
                    result["errors"].append("mercadopago: %s" % e)
                    break

        # ---- 阶段二：能立刻下载的 ----
        print("\n>>> 账单：尝试取 %d 期（平台提供 %d 期）"
              % (n_bill, len(limits["billing_periods"])))
        got = run_downloads._guard(result, "billing",
                                   mf.download_billing_reports, d, out_dir,
                                   months=n_bill)
        result["billing"] = got if isinstance(got, dict) else {}
        if not args.skip_stock:
            got = run_downloads._guard(result, "stock",
                                       mf.download_stock_reports, d, out_dir,
                                       months_back=2)
            result["stock"] = got if isinstance(got, dict) else {}

        # ---- 阶段三：收取 ----
        if pending:
            print("\n>>> 收取 %d 张待取报表，预算 %d 秒"
                  % (len(pending), args.collect_timeout))
            remaining = run_downloads.collect_pending(
                d, pending, out_dir, result, args.collect_timeout,
                poll=args.collect_poll)
            for x in remaining:
                result["pending"].append(run_downloads.label_of(x))
        d.get(session.launcher_page)

    # ---- 汇总 ----
    print("\n" + "=" * 66)
    print("补采结果 —— %s" % store)
    print("=" * 66)
    n = 0
    for period, files in result["billing"].items():
        print("  账单 %-22s %d 个文件" % (period, len(files)))
        n += len(files)
    print("  销售 %-22s %d 个文件" % ("", len(result["sales"])))
    n += len(result["sales"])
    for k, files in result["stock"].items():
        print("  库存 %-22s %d 个文件" % (k, len(files)))
        n += len(files)
    print("  MercadoPago %-15s %d 个文件" % ("", len(result["mp"])))
    n += len(result["mp"])
    if result["pending"]:
        print("\n  [待取] %s" % "、".join(result["pending"]))
    if result["empty"]:
        print("  [无数据] %s" % "、".join(result["empty"]))
    if result["errors"]:
        print("\n  [错误]")
        for e in result["errors"]:
            print("    %s" % e)
    on_disk = [f for f in os.listdir(out_dir)
               if f.lower().endswith((".xlsx", ".xls", ".csv"))]
    print("\n  合计下载 %d，目录中 %d 个文件" % (n, len(on_disk)))
    print("  %s" % out_dir)
    return 1 if result["errors"] else 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except StoreConfigError as e:
        print("\n配置错误：%s" % e)
        sys.exit(4)
    except StoreNotAllowed as e:
        print("\n店铺不在白名单：%s" % e)
        sys.exit(3)
    except ZiniaoError as e:
        print("\n紫鸟客户端错误：%s" % e)
        sys.exit(2)
