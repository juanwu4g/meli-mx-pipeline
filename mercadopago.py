# -*- coding: utf-8 -*-
"""
MercadoPago "Todas las transacciones" (settlement) report.

Generation takes minutes on MercadoPago's side, so this route is split in two:

    request_transactions_report(driver)      run FIRST, kicks off generation
    collect_transactions_report(driver, ...) run LAST, downloads the result

Between them the other routes run, so the wait costs almost nothing.
Measured 2026-08-31: a 60-day csv took ~6.5 minutes.

Two cases, per the agreed rule:

  * Day 1-5 of the month - prefer the automatic report MercadoPago generates on
    the 1st. If none exists, fall back to requesting a manual one. That fallback
    matters: on this account the Automáticos filter currently returns
    "No encontramos reportes con los filtros que aplicaste" - every existing
    report is typed Manual, including the ones that look scheduled.
  * Any other day - request a manual report for the last `days` days.

Everything is requested in csv.

Page mechanics, all verified live:
  * The report list is a plain <table>; the FORMAT column doubles as the status
    column. While generating it reads "En preparación"; when ready that cell
    becomes a download button (button[class*=statement-button-download]).
    That transition is the completion signal - there is nothing else to poll.
  * A period can appear twice with different formats (.xlsx and .csv rows for
    the same range), so a row must be matched on range AND format.
  * Datepicker aria-labels read
    "lunes 27 de julio de 2026, Día no seleccionado, Inicio de rango" - the
    date is embedded, not at the end, so matching uses a leading space to keep
    "2 de julio" from matching inside "22 de julio".
"""
import datetime
import re
import time

_BASE = "https://www.mercadopago.com.mx/balance/reports/"

# Every report type under /balance/reports shares one UI: the same list table,
# the same "Crear reporte" split button, the same datepicker and format radios,
# and the same "En preparación" -> download-button transition. Adding a type is
# a row here, nothing else. Verified identical for both on 2026-08-31.
# Two UI flavours live under /balance/reports and they are NOT interchangeable:
#
#   "typed"   settlement_v2, release
#             "Crear reporte" is a split button with data-testid
#             menu item "Manual" (also offers "Programado")
#             format chosen with #idCsv / #idXlsx radios
#             list has a "Tipo de reporte" filter (Manuales / Automáticos)
#             one row per format
#
#   "accordion"  account_statement_generic
#             NO results table at all - statements are accordions, one per
#             period. A toolbar button labelled exactly "Generar" opens a modal
#             with a Período dropdown (discrete months, not a range picker) and
#             a Formato dropdown that DEFAULTS TO .pdf. Inside each expanded
#             accordion every format is a row whose link reads "Generar" when
#             that format does not exist and "Abrir" when it does - "Abrir" is
#             both the readiness signal and the download.
#
#   "simple"  collection
#             "Crear reporte" is a plain button, found by its text
#             menu item "Crear" (also offers "Ajustes")
#             NO format choice - the report is produced in both, and the row
#             carries two download buttons
#             third column is "Estado del cobro", so no Automáticos filter
#
# Both share the same andes-ui-datepicker and the same
# "En preparación" -> download-button transition, which is why one collect
# routine serves both.
REPORTS = {
    "settlement": {"url": _BASE + "settlement_v2",
                   "label": "Todas las transacciones",
                   "flavor": "typed", "create_item": "manual"},
    "release":    {"url": _BASE + "release",
                   "label": "Liberaciones",
                   "flavor": "typed", "create_item": "manual"},
    # Cobros is asked for a 30-day window rather than the 60-day default the
    # other types use. Its modal allows up to a year, so this is a choice, not
    # a limit.
    "collection": {"url": _BASE + "collection",
                   "label": "Cobros",
                   "flavor": "simple", "create_item": "crear", "days": 30},
    # Retiros is Cobros plus one optional filter. That filter only renders
    # AFTER a period is picked, and MercadoPago labels it "Estado del cobro"
    # even though this page's own column reads "Estado del retiro".
    "withdraw":   {"url": _BASE + "withdraw",
                   "label": "Retiros",
                   "flavor": "simple", "create_item": "crear", "days": 30,
                   "state_filter": "Todos los estados"},
    # Poscobro is a two-step wizard - see the "wizard flavour" block below.
    "after_collection": {"url": _BASE + "after_collection",
                         "label": "Poscobro",
                         "flavor": "wizard", "create_item": "crear",
                         "days": 60},
    "account_statement": {"url": _BASE + "account_statement_generic",
                          "label": "Estados de saldos y movimientos",
                          "flavor": "accordion"},
}

# Kept for callers that predate the multi-type support.
REPORTS_URL = REPORTS["settlement"]["url"]

ES_MONTHS = [None, "enero", "febrero", "marzo", "abril", "mayo", "junio",
             "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]

# Spanish short month names as the table renders them: "2/jul/2026"
ES_ABBR = [None, "ene", "feb", "mar", "abr", "may", "jun",
           "jul", "ago", "sep", "oct", "nov", "dic"]

AUTO_WINDOW_DAYS = 5          # day-of-month <= this prefers the automatic report


JS_ROWS = """
return [...document.querySelectorAll('tr')].map(tr => {
  const cells = [...tr.querySelectorAll('td')]
      .map(c => (c.innerText||'').trim().replace(/\\s+/g,' '));
  const btn = tr.querySelector('button[class*=statement-button-download]');
  return {cells: cells, ready: !!btn,
          fmt: btn ? (btn.innerText||'').trim() : (cells[3] || '')};
}).filter(r => r.cells.length);
"""

# The typed pages expose a data-testid; the simple page does not, so fall back
# to matching the button's text.
JS_OPEN_CREATE_MENU = """
const b = document.querySelector('[data-testid=dropdown-button-create-report]');
if (b) { b.click(); return true; }
for (const e of document.querySelectorAll('button')) {
  if (/^crear reporte$/i.test((e.innerText||'').trim())
      && e.getBoundingClientRect().width > 4) { e.click(); return true; }
}
return false;
"""

# "Manual" on the typed pages, "Crear" on the simple one.
JS_CLICK_CREATE_ITEM = """
const want = arguments[0].toLowerCase();
for (const e of document.querySelectorAll('li,[role=option],[role=menuitem],button,a')) {
  if ((e.innerText||'').trim().toLowerCase() === want
      && e.getBoundingClientRect().width > 4) {
    (e.querySelector('button') || e).click();
    return true;
  }
}
return false;
"""
JS_CLICK_MANUAL = JS_CLICK_CREATE_ITEM     # kept for older callers

JS_OPEN_DATEPICKER = """
const b = document.querySelector('.andes-ui-datepicker__button');
if (!b) return false;
b.click();
return true;
"""

# A leading space disambiguates "2 de julio" from "22 de julio".
JS_PICK_DAY = """
const want = ' ' + arguments[0] + ' de ' + arguments[1] + ' de ' + arguments[2];
for (const b of document.querySelectorAll('button.andes-ui-datepicker__day')) {
  if ((b.getAttribute('aria-label') || '').indexOf(want) >= 0) {
    const cell = b.closest('.andes-ui-datepicker__cell');
    if (cell && /--disabled/.test(cell.className)) return 'disabled';
    b.click();
    return 'ok';
  }
}
return 'not-found';
"""

JS_CAL_NAV = """
const nav = document.querySelector('.andes-ui-datepicker__nav-buttons')
         || document.querySelector('.andes-ui-datepicker__nav');
if (!nav) return false;
const bs = nav.querySelectorAll('button');
if (!bs.length) return false;
bs[arguments[0] === 'back' ? 0 : bs.length - 1].click();
return true;
"""

JS_DATEPICKER_VALUE = """
const e = document.querySelector('.andes-ui-datepicker__selection')
       || document.querySelector('.andes-ui-datepicker__button');
return e ? (e.innerText||'').trim().replace(/\\s+/g,' ') : null;
"""

JS_CONFIRM_PICKER = """
const c = document.querySelector('.andes-ui-datepicker__buttons-container');
if (!c) return false;
for (const b of c.querySelectorAll('button')) {
  if (/aplicar|listo|confirmar|aceptar/i.test(b.innerText||'')) { b.click(); return true; }
}
return false;
"""

# The optional filters on the simple pages appear only once a period has been
# chosen - before that the modal has no dropdown at all, which is why this is
# always called after _select_range.
JS_OPEN_MODAL_DROPDOWN = """
const bs = [...document.querySelectorAll('button.andes-ui-dropdown__button')]
             .filter(b => b.getBoundingClientRect().width > 3);
const i = arguments[0] || 0;
if (i >= bs.length) return null;
bs[i].click();
return (bs[i].innerText || '').trim();
"""

JS_PICK_MODAL_OPTION = """
const want = arguments[0];
for (const e of document.querySelectorAll('[role=option]')) {
  if ((e.innerText||'').trim().replace(/\\s+/g,' ') === want) {
    (e.querySelector('button') || e).click();
    return true;
  }
}
return false;
"""

JS_PICK_FORMAT = """
const r = document.querySelector(arguments[0] === 'csv' ? '#idCsv' : '#idXlsx');
if (!r) return 'missing';
(r.closest('label') || r).click();
r.click();
return r.checked;
"""

JS_GENERAR_STATE = """
const m = document.querySelector('.andes-ui-modal');
if (!m) return null;
for (const b of m.querySelectorAll('button')) {
  if (/^generar$/i.test((b.innerText||'').trim())) return {disabled: b.disabled};
}
return null;
"""

JS_CLICK_GENERAR = """
const m = document.querySelector('.andes-ui-modal');
if (!m) return false;
for (const b of m.querySelectorAll('button')) {
  if (/^generar$/i.test((b.innerText||'').trim())) { b.click(); return true; }
}
return false;
"""

JS_FILTER_TYPE = """
for (const b of document.querySelectorAll('button.andes-ui-dropdown-standalone__trigger')) {
  if ((b.innerText||'').trim().toLowerCase().startsWith('tipo de reporte')) {
    b.click();
    return true;
  }
}
return false;
"""

JS_PICK_TYPE_OPTION = """
const want = arguments[0].toLowerCase();
for (const e of document.querySelectorAll('li[role=option],[role=option],li.andes-list__item')) {
  if ((e.innerText||'').trim().toLowerCase().startsWith(want)) {
    (e.querySelector('button') || e).click();
    return true;
  }
}
return false;
"""

# A row may carry SEVERAL download buttons - the Cobros rows expose .csv and
# .xlsx side by side - so scan them all rather than taking the first. On the
# typed pages the same period instead appears as two separate rows, one per
# format, which the same range+format test handles.
# Scan EVERY row whose period matches, then answer with the best of them.
#
# This used to answer from the first matching row and return immediately, which
# hung the collect loop in two ways:
#
#   * a row with the same period but no download buttons ended the search, so a
#     LATER row with that period which was ready never got looked at. Duplicate
#     periods are the norm, not the exception - every run asks for the same
#     rolling window, so a second run the same day creates a second row for it.
#   * a finished report offering only .xlsx when we asked for .csv was skipped
#     with `continue`, the loop fell off the end returning null, and probe read
#     that as "still generating" - forever, because that report is done and will
#     never grow a .csv button.
#
# Both looked identical from the terminal: "not ready yet" against a page that
# plainly showed the report ready.
JS_FIND_ROW = """
const wantRange = arguments[0], wantFmt = ('.' + arguments[1]).toLowerCase();
let best = null;
for (const tr of document.querySelectorAll('tr')) {
  const cells = [...tr.querySelectorAll('td')]
      .map(c => (c.innerText||'').trim().replace(/\\s+/g,' '));
  if (!cells.length || cells[1] !== wantRange) continue;
  // The row carries its own state, e.g. "statements-table__row--delayed".
  // Ignore the long variant (…--settlement_v2-delayed-manual--not-seen) and
  // take only the bare one-word suffix.
  let state = '';
  for (const cls of tr.className.split(/\\s+/)) {
    const m = cls.match(/^statements-table__row--([a-z]+)$/);
    if (m) { state = m[1]; break; }
  }
  const btns = [...tr.querySelectorAll('button[class*=statement-button-download]')];
  const fmts = btns.map(b => (b.innerText||'').trim().toLowerCase());
  const row = {cells: cells, state: state, formats: fmts,
               status: cells[cells.length-1] || ''};

  if (fmts.indexOf(wantFmt) >= 0) {         // exactly what we asked for: done
    row.ready = true;
    return row;
  }
  // Otherwise remember the most useful thing seen so far and keep looking.
  //   ready-in-another-format  beats  a terminal state  beats  still-generating
  const rank = fmts.length ? 3 : (state === 'empty' || state === 'delayed' ? 2 : 1);
  if (!best || rank > best._rank) { row.ready = false; row._rank = rank; best = row; }
}
return best;
"""

JS_CLICK_ROW = """
const wantRange = arguments[0], wantFmt = ('.' + arguments[1]).toLowerCase();
for (const tr of document.querySelectorAll('tr')) {
  const cells = [...tr.querySelectorAll('td')]
      .map(c => (c.innerText||'').trim().replace(/\\s+/g,' '));
  if (!cells.length || cells[1] !== wantRange) continue;
  for (const b of tr.querySelectorAll('button[class*=statement-button-download]')) {
    if ((b.innerText||'').trim().toLowerCase() === wantFmt) { b.click(); return cells; }
  }
}
return null;
"""



# ---------------------------------------------------- accordion flavour ------
# The page carries THREE things labelled "Generar" - the toolbar button that
# opens the modal, the modal's own submit (andes-ui-button--large), and the
# per-format links inside each accordion - plus a "Generar nuevo estado"
# button. Every selector below is scoped so they cannot be confused.

# Two different buttons open this modal and which one exists depends on state:
#   "Generar"             on the promo card, shown only while the newest period
#                         has no statement yet - it disappears once one exists
#   "Generar nuevo estado" the toolbar button, always present
# Both lead to the same modal. Prefer the exact "Generar", fall back to the
# toolbar one. Zero-width matches are the per-format textlinks inside collapsed
# accordions - there can be 40+ of them, so the width test is essential.
JS_ACCT_OPEN_MODAL = """
const shown = (e) => e.getBoundingClientRect().width > 4
                  && !/textlink/.test(e.className||'');

// "Generar" alone is ambiguous - the modal's own submit carries the same text -
// so that one also has to exclude --large. "Generar nuevo estado" is unique by
// text, and it IS --large itself, so it must NOT be filtered on that.
for (const e of document.querySelectorAll('button')) {
  if (/^generar$/i.test((e.innerText||'').trim())
      && shown(e) && !/--large/.test(e.className||'')) {
    e.click();
    return 'generar';
  }
}
for (const e of document.querySelectorAll('button')) {
  if (/^generar nuevo estado$/i.test((e.innerText||'').trim()) && shown(e)) {
    e.click();
    return 'generar nuevo estado';
  }
}
return null;
"""

# The modal's own dropdowns are .andes-ui-dropdown__button; the page's filters
# are .andes-ui-dropdown-standalone__trigger, so these cannot collide.
JS_ACCT_OPEN_DD = """
const bs = [...document.querySelectorAll('button.andes-ui-dropdown__button')]
             .filter(b => b.getBoundingClientRect().width > 3);
if (arguments[0] >= bs.length) return null;
bs[arguments[0]].click();
return (bs[arguments[0]].innerText || '').trim();
"""

# role=option strictly: the nav sidebar is full of plain <li>s.
JS_ACCT_OPTIONS = """
return [...document.querySelectorAll('[role=option]')]
  .filter(e => e.getBoundingClientRect().width > 3)
  .map(e => (e.innerText||'').trim().replace(/\\s+/g,' '));
"""

JS_ACCT_PICK_OPTION = """
const want = arguments[0];
for (const e of document.querySelectorAll('[role=option]')) {
  if ((e.innerText||'').trim().replace(/\\s+/g,' ') === want) {
    (e.querySelector('button') || e).click();
    return true;
  }
}
return false;
"""

JS_ACCT_MODAL_GENERAR = """
for (const b of document.querySelectorAll('button')) {
  if (/^generar$/i.test((b.innerText||'').trim())
      && /--large/.test(b.className||'')
      && b.getBoundingClientRect().width > 4) {
    if (arguments[0]) { b.click(); return true; }
    return {disabled: b.disabled};
  }
}
return null;
"""

# Find the accordion for a month, expand it, and report what the wanted format
# offers: "abrir" (ready), "generar" (that format not produced), or missing.
JS_ACCT_FORMAT_STATE = """
const monthKey = arguments[0].toLowerCase(), wantFmt = arguments[1];
for (const h of document.querySelectorAll('.andes-ui-accordion-header')) {
  const head = (h.innerText||'').trim().replace(/\\s+/g,' ');
  if (head.toLowerCase().indexOf(monthKey) < 0) continue;
  if (h.getAttribute('aria-expanded') !== 'true') { h.click(); return {head: head, action: 'expanding'}; }
  const p = document.getElementById(h.getAttribute('aria-controls'));
  if (!p) return {head: head, action: 'no-panel'};
  for (const row of p.querySelectorAll('.card-list__item')) {
    if ((row.innerText||'').replace(/\\s+/g,' ').indexOf(wantFmt) < 0) continue;
    const b = row.querySelector('button.andes-ui-textlink');
    return {head: head, action: b ? (b.innerText||'').trim().toLowerCase() : 'no-link'};
  }
  return {head: head, action: 'no-format-row'};
}
return null;
"""

# The label lives on .card-list__item; closest('div') stops at an inner wrapper
# that contains only the link text, which is why the row selector is explicit.
JS_ACCT_CLICK_FORMAT = """
const monthKey = arguments[0].toLowerCase(), wantFmt = arguments[1];
for (const h of document.querySelectorAll('.andes-ui-accordion-header')) {
  if ((h.innerText||'').toLowerCase().indexOf(monthKey) < 0) continue;
  const p = document.getElementById(h.getAttribute('aria-controls'));
  if (!p) return 'no-panel';
  for (const row of p.querySelectorAll('.card-list__item')) {
    if ((row.innerText||'').replace(/\\s+/g,' ').indexOf(wantFmt) < 0) continue;
    const b = row.querySelector('button.andes-ui-textlink');
    if (!b) return 'no-link';
    b.click();
    return 'clicked:' + (b.innerText||'').trim();
  }
}
return 'not-found';
"""


def _month_key(period_label):
    """'Agosto de 2026' -> 'agosto 2026', which is what the accordion header
    ('1 agosto 2026 - 31 agosto 2026') contains."""
    parts = (period_label or "").lower().replace(" de ", " ").split()
    return " ".join(parts[-2:]) if len(parts) >= 2 else (period_label or "").lower()


def _request_account_statement(driver, spec, fmt="csv"):
    """Toolbar Generar -> newest Período -> chosen Formato -> Generar."""
    # This page finishes rendering later than the table-based ones, so poll for
    # the toolbar button rather than trusting a fixed sleep.
    opened = None
    deadline = time.time() + 45
    while time.time() < deadline:
        opened = driver.execute_script(JS_ACCT_OPEN_MODAL)
        if opened:
            break
        time.sleep(3)
    if not opened:
        print("    [警告] 'Generar' 和 'Generar nuevo estado' 都没找到")
        return None
    print("      通过 %r 打开" % opened)
    time.sleep(8)

    # Período: the newest real month. "Periodo personalizado" is a range option,
    # not a month, so it is excluded.
    driver.execute_script(JS_ACCT_OPEN_DD, 0)
    time.sleep(3)
    months = [o for o in driver.execute_script(JS_ACCT_OPTIONS)
              if "personalizado" not in o.lower()]
    if not months:
        print("    [警告] 没有可选的 período")
        return None
    period = months[0]
    print("      período：%s（共 %d 个可选）" % (period, len(months)))
    driver.execute_script(JS_ACCT_PICK_OPTION, period)
    time.sleep(3)

    # Formato defaults to .pdf - it MUST be set explicitly.
    driver.execute_script(JS_ACCT_OPEN_DD, 1)
    time.sleep(3)
    want = "." + fmt
    if not driver.execute_script(JS_ACCT_PICK_OPTION, want):
        print("    [警告] 不提供 %s 格式，可选：%s"
              % (want, driver.execute_script(JS_ACCT_OPTIONS)))
        return None
    print("      formato：%s" % want)
    time.sleep(3)

    st = driver.execute_script(JS_ACCT_MODAL_GENERAR, False)
    if not st or st.get("disabled"):
        print("    [警告] 弹窗里的 Generar 不可点击")
        return None
    driver.execute_script(JS_ACCT_MODAL_GENERAR, True)
    print("    已点击 Generar，MercadoPago 端继续生成")
    time.sleep(8)

    return {"kind": "mercadopago", "report": "account_statement",
            "url": spec["url"], "label": spec["label"], "flavor": "accordion",
            "range": period, "month_key": _month_key(period), "fmt": fmt,
            "requested_at": time.time(), "already_ready": False}


def _collect_account_statement(driver, pending, download_dir, timeout, poll):
    import meli_forms as mf

    month_key = pending.get("month_key") or _month_key(pending.get("range"))
    want = "." + pending.get("fmt", "csv")
    driver.get(pending["url"])
    time.sleep(16)
    before = mf.snapshot_dir(download_dir)

    deadline = time.time() + timeout
    ready = False
    while time.time() < deadline:
        st = driver.execute_script(JS_ACCT_FORMAT_STATE, month_key, want)
        left = int(deadline - time.time())
        if st is None:
            print("    [剩余 %4d 秒] 尚未出现 %r 的折叠面板" % (left, month_key))
        elif st["action"] == "expanding":
            time.sleep(4)
            continue                      # re-check now that it is open
        elif st["action"] == "abrir":
            print("    [剩余 %4d 秒] 就绪 —— %s | %s Abrir" % (left, st["head"], want))
            ready = True
            break
        else:
            print("    [剩余 %4d 秒] %s | %s -> %s"
                  % (left, st["head"], want, st["action"]))
        time.sleep(poll)
        driver.refresh()
        time.sleep(10)

    if not ready:
        print("    [警告] %d 秒后仍未就绪，留到后续运行再取"
              % timeout)
        return []

    print("    %s" % driver.execute_script(JS_ACCT_CLICK_FORMAT, month_key, want))
    files = mf.wait_for_downloads(download_dir, before, min_files=1,
                                  timeout=150, settle=5,
                                  label="mercadopago/account_statement")
    for f in files:
        print("      + %s" % f)
    return files



# ------------------------------------------------- wizard flavour (Poscobro) --
#
# after_collection is a TWO-STEP wizard and shares almost nothing with the
# other pages beyond the datepicker:
#
#   step 1  a range datepicker on the left, a multi-select "Tipo de operacion"
#           on the right (Reclamos / Contracargos / Devoluciones). That
#           dropdown has its own Aplicar. "Siguiente" stays disabled until
#           BOTH the period and at least one operation type are set.
#   step 2  five dropdowns (estado del reclamo / contracargo / devolucion,
#           canal, herramienta de cobro) which already default to their
#           "Todos ..." values, then Generar.
#
# Three traps:
#   * The results table uses plain .csv / .xlsx buttons, NOT
#     button[class*=statement-button-download], so the shared row helpers see
#     nothing here.
#   * The row's period text is written unlike any other page
#     ("4 julio 2026 - 2 septiembre 2026", full month names) AND its end date
#     is recorded a day later than the one selected. Rather than reconstruct
#     that, the request reads the top row's period back after Generar and uses
#     it verbatim as the match key.
#   * Ticking an operation row updates its checkbox asynchronously - reading
#     `checked` immediately after the click still returns false.

JS_AC_ROWS = """
return [...document.querySelectorAll('tr')].map(tr => {
  const cells = [...tr.querySelectorAll('td')]
      .map(c => (c.innerText||'').trim().replace(/\\s+/g,' '));
  const btns = [...tr.querySelectorAll('button')]
      .filter(b => /^\\.(csv|xlsx)$/i.test((b.innerText||'').trim()))
      .map(b => (b.innerText||'').trim().toLowerCase());
  return {cells: cells, formats: btns};
}).filter(r => r.cells.length);
"""

JS_AC_FIND_ROW = """
const wantPeriod = arguments[0], wantFmt = ('.' + arguments[1]).toLowerCase();
for (const tr of document.querySelectorAll('tr')) {
  const cells = [...tr.querySelectorAll('td')]
      .map(c => (c.innerText||'').trim().replace(/\\s+/g,' '));
  if (cells.length < 2 || cells[1] !== wantPeriod) continue;
  let state = '';
  for (const cls of tr.className.split(/\\s+/)) {
    const m = cls.match(/^statements-table__row--([a-z]+)$/);
    if (m) { state = m[1]; break; }
  }
  if (state === 'empty' || state === 'delayed') {
    return {cells: cells, ready: false, state: state, formats: []};
  }
  // Same as JS_FIND_ROW: report every format offered, so a caller asked for
  // .csv can still take a report that only came out as .xlsx.
  const fmts = [...tr.querySelectorAll('button')]
      .map(b => (b.innerText||'').trim().toLowerCase())
      .filter(t => t === '.csv' || t === '.xlsx');
  return {cells: cells, ready: fmts.indexOf(wantFmt) >= 0, formats: fmts,
          status: cells[cells.length-1] || ''};
}
return null;
"""

JS_AC_CLICK_ROW = """
const wantPeriod = arguments[0], wantFmt = ('.' + arguments[1]).toLowerCase();
for (const tr of document.querySelectorAll('tr')) {
  const cells = [...tr.querySelectorAll('td')]
      .map(c => (c.innerText||'').trim().replace(/\\s+/g,' '));
  if (cells.length < 2 || cells[1] !== wantPeriod) continue;
  const hit = [...tr.querySelectorAll('button')]
      .find(b => (b.innerText||'').trim().toLowerCase() === wantFmt);
  if (!hit) return null;
  hit.click();
  return cells;
}
return null;
"""

JS_AC_TOP_PERIOD = """
for (const tr of document.querySelectorAll('tr')) {
  const cells = [...tr.querySelectorAll('td')]
      .map(c => (c.innerText||'').trim().replace(/\\s+/g,' '));
  if (cells.length >= 2) return cells[1];      // first DATA row
}
return null;
"""

# Tick exactly ONE option per call. Clicking re-renders the list, so a loop
# over a NodeList captured up front holds stale nodes after the first click -
# that is why a single pass ticked only Devoluciones and the report came back
# scoped to one operation type instead of three.
JS_AC_CHECK_ONE = """
const want = arguments[0];
for (const li of document.querySelectorAll('li.andes-ui-list__item')) {
  if ((li.innerText||'').trim() !== want) continue;
  const cb = li.querySelector('input[type=checkbox]');
  if (!cb) return 'no-checkbox';
  if (cb.checked) return 'already';
  li.click();
  return 'clicked';
}
return 'not-found';
"""

JS_AC_OP_STATES = """
return [...document.querySelectorAll('li.andes-ui-list__item')]
  .filter(e => e.getBoundingClientRect().width > 3
            && /Reclamos|Contracargos|Devoluciones/.test(e.innerText||''))
  .map(e => ({t: (e.innerText||'').trim(),
              chk: !!(e.querySelector('input[type=checkbox]')||{}).checked}));
"""

JS_AC_DD_TEXTS = """
return [...document.querySelectorAll('button.andes-ui-dropdown__button')]
  .filter(e => e.getBoundingClientRect().width > 3)
  .map(e => (e.innerText||'').trim());
"""

JS_AC_PICK_TODO = """
for (const e of document.querySelectorAll('[role=option]')) {
  const t = (e.innerText||'').trim();
  if (/^tod[oa]s?\\b/i.test(t)) { e.click(); return t; }
}
return null;
"""

JS_AC_CLICK_TEXT = """
const want = arguments[0].toLowerCase();
for (const b of document.querySelectorAll('button')) {
  if (b.getBoundingClientRect().width > 4
      && (b.innerText||'').trim().toLowerCase() === want) { b.click(); return true; }
}
return false;
"""

AC_OPERATIONS = ["Reclamos", "Contracargos", "Devoluciones"]


def _request_after_collection(driver, spec, days=60, fmt="csv"):
    """Poscobro: Crear -> range -> operation types -> Aplicar -> Siguiente
    -> Todos everywhere -> Generar."""
    if not driver.execute_script(JS_OPEN_CREATE_MENU):
        print("    [警告] 未找到 'Crear reporte'")
        return None
    time.sleep(4)
    if not driver.execute_script(JS_CLICK_CREATE_ITEM,
                                 spec.get("create_item", "crear")):
        print("    [警告] 新建菜单里没有 'Crear'")
        return None
    time.sleep(9)

    today = datetime.date.today()
    chosen = _select_range(driver, today - datetime.timedelta(days=days), today)
    if not chosen:
        return None

    # --- step 1: operation types, then that dropdown's own Aplicar ---
    driver.execute_script(JS_OPEN_MODAL_DROPDOWN, 0)
    time.sleep(3)
    for op in AC_OPERATIONS:
        res = driver.execute_script(JS_AC_CHECK_ONE, op)
        time.sleep(1.5)                 # let React re-render before the next
        if res not in ("clicked", "already"):
            print("      [警告] %s -> %s" % (op, res))
    states = driver.execute_script(JS_AC_OP_STATES)
    print("      operaciones：%s"
          % ", ".join("%s=%s" % (s["t"], "on" if s["chk"] else "OFF")
                      for s in states))
    if states and not all(s["chk"] for s in states):
        print("      [警告] 并非所有操作类型都已勾选")
    driver.execute_script(JS_AC_CLICK_TEXT, "aplicar")
    time.sleep(4)

    if not driver.execute_script(JS_AC_CLICK_TEXT, "siguiente"):
        print("    [警告] 'Siguiente' 不可点击，第一步未完成")
        return None
    time.sleep(9)

    # --- step 2: every filter on its "Todos ..." value (already the default) --
    texts = driver.execute_script(JS_AC_DD_TEXTS)
    for i, t in enumerate(texts):
        if not t.lower().startswith(("todos", "todas")):
            print("      filtro[%d] %r -> Todos" % (i, t))
            driver.execute_script(JS_OPEN_MODAL_DROPDOWN, i)
            time.sleep(3)
            driver.execute_script(JS_AC_PICK_TODO)
            time.sleep(2)
    print("      filtros：%s" % ", ".join(driver.execute_script(JS_AC_DD_TEXTS)))

    before_top = driver.execute_script(JS_AC_TOP_PERIOD)
    if not driver.execute_script(JS_AC_CLICK_TEXT, "generar"):
        print("    [警告] 'Generar' 不可点击")
        return None
    print("    已点击 Generar，MercadoPago 端继续生成")
    time.sleep(10)

    period = driver.execute_script(JS_AC_TOP_PERIOD)
    if not period or period == before_top:
        print("      [警告] 没有出现新行，将按 %r 匹配" % period)
    print("      该行期间：%s" % period)

    return {"kind": "mercadopago", "report": "after_collection",
            "url": spec["url"], "label": spec["label"], "flavor": "wizard",
            "range": period, "fmt": fmt,
            "requested_at": time.time(), "already_ready": False}

def _confirm_row_period(driver, want_range, tries=4, pause=4):
    """Match on the period MercadoPago actually recorded, not the one we picked.

    The two are usually identical, but not always - Poscobro records an end date
    a day later than the one selected, and the pages do not agree with each
    other about "today": on 2026-09-03 Liberaciones recorded
    "5/jul/2026 a 2/sep/2026" while Retiros, same store, same day, recorded
    "4/ago/2026 a 3/sep/2026".

    When the recorded period differs from ours, nothing ever matches the row and
    the collect phase waits out its whole budget on a report that is sitting
    there finished. So: give the row a few seconds to appear under the period we
    chose, and if it does not, adopt whatever the newest row says.

    The wizard flavour has always done this (it has no choice - its period text
    is unlike any other page). This extends it to the rest.
    """
    for _ in range(tries):
        if driver.execute_script(JS_FIND_ROW, want_range, "csv"):
            return want_range               # the site echoed what we selected
        time.sleep(pause)
    top = driver.execute_script(JS_AC_TOP_PERIOD)
    if top and top != want_range:
        print("      [提示] 站点记录的期间是 %r，不是 %r "
              "- matching on the site's" % (top, want_range))
        return top
    return want_range


def _fmt_range(start, end):
    """Render a range the way the table's 'Período de fechas' column does."""
    return "%d/%s/%d a %d/%s/%d" % (start.day, ES_ABBR[start.month], start.year,
                                    end.day, ES_ABBR[end.month], end.year)


# The site's "today" is NOT this machine's today. This box runs on CST (UTC+8)
# while MercadoPago Mexico is UTC-6, so for much of the day the local clock is a
# day ahead and an end date built from datetime.date.today() does not exist yet
# on the site - the cell comes back disabled or not-found and the whole range
# selection fails. Read the calendar's own --today cell instead.
JS_CAL_TODAY = """
for (const b of document.querySelectorAll('button.andes-ui-datepicker__day')) {
  const cell = b.closest('[class*=cell]');
  if (cell && /--today/.test(cell.className)) return b.getAttribute('aria-label');
}
return null;
"""


def _parse_cal_label(label):
    """'Hoy, martes 1 de septiembre de 2026, ...' -> date(2026, 9, 1)."""
    m = re.search(r"(\d{1,2})\s+de\s+([A-Za-zÁ-úá-ú]+)\s+de\s+(\d{4})", label or "")
    if not m:
        return None
    day, month_name, year = m.groups()
    month_name = month_name.lower()
    if month_name not in ES_MONTHS:
        return None
    return datetime.date(int(year), ES_MONTHS.index(month_name), int(day))


def _select_range(driver, start, end):
    """Drive the range datepicker. Returns the (start, end) actually selected.

    `end` is clamped to the site's own today when the local clock is ahead of
    it - see JS_CAL_TODAY. Callers must use the returned dates to build the
    range string they will later match the results row on, or they will look
    for a row that was never created.
    """
    if not driver.execute_script(JS_OPEN_DATEPICKER):
        print("      [警告] 未找到期间日期选择器")
        return None
    time.sleep(4)

    site_today = _parse_cal_label(driver.execute_script(JS_CAL_TODAY))
    if site_today and end > site_today:
        print("      结束日期 %s 晚于站点当天（%s），已收敛"
              % (end, site_today))
        end = site_today

    picked = False
    for _ in range(14):
        res = driver.execute_script(JS_PICK_DAY, start.day,
                                    ES_MONTHS[start.month], start.year)
        if res == "ok":
            picked = True
            break
        if res == "disabled":
            print("      [警告] 起始日期 %s 不可选" % start)
            return None
        driver.execute_script(JS_CAL_NAV, "back")
        time.sleep(1.2)
    if not picked:
        print("      [警告] 无法定位到 %s" % start)
        return None
    time.sleep(1)

    res = driver.execute_script(JS_PICK_DAY, end.day, ES_MONTHS[end.month], end.year)
    if res != "ok":
        for _ in range(14):
            driver.execute_script(JS_CAL_NAV, "fwd")
            time.sleep(1.2)
            res = driver.execute_script(JS_PICK_DAY, end.day,
                                        ES_MONTHS[end.month], end.year)
            if res == "ok":
                break
    if res != "ok":
        print("      [警告] 结束日期 %s -> %s" % (end, res))
        return None

    time.sleep(2)
    print("      选择器显示：%s" % driver.execute_script(JS_DATEPICKER_VALUE))
    driver.execute_script(JS_CONFIRM_PICKER)
    time.sleep(2)
    return start, end


def _newest_automatic(driver):
    """Filter the list to Automáticos and return the newest row, or None."""
    if not driver.execute_script(JS_FILTER_TYPE):
        return None
    time.sleep(2)
    driver.execute_script(JS_PICK_TYPE_OPTION, "autom")
    time.sleep(8)
    rows = [r for r in driver.execute_script(JS_ROWS) if len(r["cells"]) >= 4]
    return rows[0] if rows else None


# probe() answers. WAITING is the only one worth polling again.
#
# EMPTY and DELAYED both mean "not coming during this run", but they are not the
# same and must not be reported as the same. MercadoPago encodes them in the
# row's own class - statements-table__row--<state> - alongside the tooltip text:
#
#   --empty    "No pudimos generar tu reporte porque no hay movimientos en este
#               rango de fechas."          -> permanent for that range
#   --delayed  "No pudimos generar tu reporte porque hay datos en proceso en
#               este momento. Espera unas horas a que te notifiquemos."
#                                          -> transient, but on a scale of hours
#
# Reading the class rather than the tooltip means no hover, and no dependence on
# Spanish copy that MercadoPago may reword.
READY = "ready"
WAITING = "waiting"
EMPTY = "empty"        # no movements in range: never retry this range
DELAYED = "delayed"    # data in process: retry on a later run, not this one


class NoReportAccess(Exception):
    """The logged-in MercadoPago user cannot see reports at all.

    Not a bug, and not transient. MercadoPago collaborator accounts need
    "Acceder a reportes de tus cobros y facturación" and "Acceder a reportes de
    operaciones" granted under Colaboradores. Without them every report page
    renders a permission notice and no "Crear reporte" button - which the route
    would otherwise report as `[warn] 'Crear reporte' not found`, sending you
    hunting for a markup change that never happened.

    Found on EWTTO_SM 2026-09-02: all six reports "failed" this way while
    BOCINA_SM, whose login holds the permission, ran clean.
    """


# The page renders this instead of the reports UI. Matched on the sentence,
# not the surrounding markup, since it appears on every report type's page.
PERMISSION_WALL = "no tienes permisos para ver los reportes"


def _require_access(driver, label):
    """Raise NoReportAccess if the page is a permission notice.

    Called before anything is clicked, so the failure names its own cause
    instead of surfacing as a missing selector five steps later.
    """
    txt = driver.execute_script("return document.body.innerText || ''") or ""
    if PERMISSION_WALL in txt.lower():
        raise NoReportAccess(
            "this MercadoPago login has no reports permission (page says "
            "'No tienes permisos para ver los reportes'). An account admin must "
            "grant 'Acceder a reportes de tus cobros y facturación' and "
            "'Acceder a reportes de operaciones' under Colaboradores. "
            "Blocked on: %s" % label)


# 每家店只检查一次语言。run_downloads 在每家店开跑前把它清掉 —— 和
# meli_forms.SALES_PERIOD_ACTUAL 同一个套路：模块级变量在批量跑时会串店。
LANG_CHECKED = {"done": False}


def _check_language(driver):
    """MercadoPago 页面是不是西语。只警告，不拦路。

    这边没有语言切换器（整页扫不到任何 language/idioma 元素），能用的只有
    `<html lang>`。所以发现不对也改不了，只能说出来 —— 而这已经足够：MP 的
    报表列名同样跟着语言走，出问题时至少不是无声无息的。

    语言多半跟着 MercadoLibre 那边的账号设置走（meli_forms.ensure_language
    已经把它摆正），但这一点**没有实测过**，所以这里独立检查而不是假定。
    """
    if LANG_CHECKED["done"]:
        return
    LANG_CHECKED["done"] = True
    try:
        lang = driver.execute_script("return document.documentElement.lang") or ""
    except Exception:
        return
    if not lang.lower().startswith("es"):
        print("  [警告] MercadoPago 页面语言是 %r，不是西语 —— 导出报表的列名"
              "会跟着变，后续解析可能对不上。请在账号里改回西班牙语。" % lang)
    else:
        print("  页面语言：%s ✓" % lang)


def request_report(driver, report="settlement", days=60, fmt="csv"):
    """PHASE 1 - kick off generation. Returns a pending dict, or None.

    The dict identifies the row to collect later:
        {"kind", "report", "url", "label", "range", "fmt",
         "requested_at", "already_ready"}
    """
    spec = REPORTS[report]
    url, label = spec["url"], spec["label"]
    flavor = spec.get("flavor", "typed")
    create_item = spec.get("create_item", "manual")
    # A report may pin its own window; otherwise the caller's default applies.
    days = spec.get("days", days)
    print("\n--- MercadoPago：申请 %s ---" % label)
    driver.get(url)
    time.sleep(15)
    _check_language(driver)
    _require_access(driver, label)

    if flavor == "accordion":
        return _request_account_statement(driver, spec, fmt=fmt)

    if flavor == "wizard":
        return _request_after_collection(driver, spec, days=days, fmt=fmt)

    today = datetime.date.today()

    # The simple flavour has no "Tipo de reporte" filter, so there is no
    # automatic report to prefer - it always creates one.
    if flavor == "typed" and today.day <= AUTO_WINDOW_DAYS:
        print("    当月第 %d 天，优先查找自动生成的报表" % today.day)
        auto = _newest_automatic(driver)
        if auto:
            print("    找到自动报表：%s" % auto["cells"])
            return {"kind": "mercadopago", "report": report, "url": url,
                    "label": label, "range": auto["cells"][1],
                    "fmt": fmt, "requested_at": time.time(),
                    "already_ready": auto["ready"]}
        print("    没有自动报表，改为手动申请")
        driver.get(url)
        time.sleep(12)

    start = today - datetime.timedelta(days=days)
    want_range = _fmt_range(start, today)
    print("    手动申请 %s（%d 天），格式 %s" % (want_range, days, fmt))

    if not driver.execute_script(JS_OPEN_CREATE_MENU):
        print("    [警告] 未找到 'Crear reporte'")
        return None
    time.sleep(4)
    if not driver.execute_script(JS_CLICK_CREATE_ITEM, create_item):
        print("    [警告] 新建菜单里没有 %r 选项" % create_item)
        return None
    time.sleep(8)

    chosen = _select_range(driver, start, today)
    if not chosen:
        return None
    start, end = chosen
    want_range = _fmt_range(start, end)      # rebuild: end may have been clamped

    # Optional filters render only after the period is set. Retiros wants
    # "Todos los estados"; Cobros deliberately leaves its filter alone.
    state = spec.get("state_filter")
    if state:
        opened = driver.execute_script(JS_OPEN_MODAL_DROPDOWN, 0)
        time.sleep(3)
        if opened is None:
            print("      [警告] 弹窗里没有筛选下拉框")
        else:
            ok = driver.execute_script(JS_PICK_MODAL_OPTION, state)
            print("      estado：%s -> %s" % (state, "已设置" if ok else "未找到"))
            time.sleep(2)

    if flavor == "typed":
        print("      格式 %s -> %s"
              % (fmt, driver.execute_script(JS_PICK_FORMAT, fmt)))
        time.sleep(2)
    else:
        # No radios here: Cobros is generated in both formats and the row
        # carries a button for each. Collection still asks for csv.
        print("      格式：两种都会生成，取回 .%s" % fmt)

    st = driver.execute_script(JS_GENERAR_STATE)
    if not st or st["disabled"]:
        print("    [警告] Generar 仍不可点击，期间未被接受")
        return None
    driver.execute_script(JS_CLICK_GENERAR)
    print("    已点击 Generar，MercadoPago 端继续生成")
    time.sleep(6)

    print("    实际选中的区间：%s" % want_range)
    want_range = _confirm_row_period(driver, want_range)
    return {"kind": "mercadopago", "report": report, "url": url, "label": label,
            "range": want_range, "fmt": fmt, "flavor": flavor,
            "requested_at": time.time(), "already_ready": False}


def collect_report(driver, pending, download_dir, timeout=600, poll=20):
    """PHASE 3 - return to the page, wait for the row, download it.

    Polls up to `timeout` seconds (default 10 min) printing status each time.
    Returns the downloaded filenames, or [] if still not ready - a report that
    is not ready is reported, not treated as a failure; it stays available and
    can be collected on a later run.
    """
    import meli_forms as mf

    if pending.get("flavor") == "wizard":
        # same shape as the typed pages, but its own row helpers
        waited = time.time() - pending.get("requested_at", time.time())
        print("\n--- MercadoPago：取回 %s | %s (.%s) ---"
              % (pending.get("label"), pending.get("range"),
                 pending.get("fmt", "csv")))
        print("    距提交申请已过去 %.0f 秒" % waited)
        driver.get(pending["url"])
        time.sleep(14)
        deadline = time.time() + timeout
        while time.time() < deadline:
            row = driver.execute_script(JS_AC_FIND_ROW, pending["range"],
                                        pending.get("fmt", "csv"))
            left = int(deadline - time.time())
            if row is None:
                print("    [剩余 %4d 秒] 列表中尚未出现该行" % left)
            elif row["ready"]:
                print("    [剩余 %4d 秒] 就绪（%s）" % (left, row["status"]))
                return fetch(driver, pending, download_dir)
            else:
                print("    [剩余 %4d 秒] %s" % (left, row["status"]))
            time.sleep(poll)
            driver.refresh()
            time.sleep(9)
        print("    [警告] %d 秒后仍未就绪" % timeout)
        return []

    if pending.get("flavor") == "accordion":
        waited = time.time() - pending.get("requested_at", time.time())
        print("\n--- MercadoPago：取回 %s | %s (.%s) ---"
              % (pending.get("label"), pending.get("range"),
                 pending.get("fmt", "csv")))
        print("    距提交申请已过去 %.0f 秒" % waited)
        return _collect_account_statement(driver, pending, download_dir,
                                          timeout, poll)

    want_range, fmt = pending["range"], pending.get("fmt", "csv")
    url = pending.get("url", REPORTS_URL)
    label = pending.get("label", "reporte")
    waited = time.time() - pending.get("requested_at", time.time())
    print("\n--- MercadoPago：取回 %s | %s (.%s) ---" % (label, want_range, fmt))
    print("    距提交申请已过去 %.0f 秒" % waited)

    driver.get(url)
    time.sleep(14)
    before = mf.snapshot_dir(download_dir)

    deadline = time.time() + timeout
    row = None
    while time.time() < deadline:
        row = driver.execute_script(JS_FIND_ROW, want_range, fmt)
        left = int(deadline - time.time())
        if row is None:
            print("    [剩余 %4d 秒] 列表中尚未出现该行" % left)
        elif row["ready"]:
            print("    [剩余 %4d 秒] 就绪（%s）" % (left, row["status"]))
            break
        else:
            print("    [剩余 %4d 秒] %s" % (left, row["status"] or "en preparación"))
        time.sleep(poll)
        driver.refresh()
        time.sleep(9)

    if not row or not row["ready"]:
        print("    [警告] %d 秒后仍未就绪，留到后续运行再取"
              % timeout)
        return []

    clicked = driver.execute_script(JS_CLICK_ROW, want_range, fmt)
    print("    正在下载：%s" % clicked)
    files = mf.wait_for_downloads(download_dir, before, min_files=1,
                                  timeout=120, settle=5, label="mercadopago")
    for f in files:
        print("      + %s" % f)
    return files


# --- back-compat aliases -----------------------------------------------------

def request_transactions_report(driver, days=60, fmt="csv"):
    return request_report(driver, report="settlement", days=days, fmt=fmt)


def collect_transactions_report(driver, pending, download_dir, timeout=600,
                                poll=20):
    return collect_report(driver, pending, download_dir, timeout=timeout,
                          poll=poll)


# ---------------------------------------------------- shared-wait collection --

# The reports generate CONCURRENTLY on MercadoPago's side. Time spent waiting on
# one advances all of them, so a collector must poll them in a shared loop
# rather than handing each a slice of the budget. Splitting a 600s budget across
# six pending reports gave each 100s, and nothing that takes 2.6-6.5 minutes
# could ever finish in that. probe()/fetch() exist so the caller can cycle
# cheaply: probe navigates and looks once, fetch downloads what probe found.

def probe(driver, pending, settle=12):
    """Navigate to the report's page and check state once. No waiting.

    Returns one of READY / WAITING / EMPTY / DELAYED. Leaves the driver ON that
    page so fetch() can act without navigating again.

    EMPTY and DELAYED are terminal answers for this run, not slow ones. Neither
    row looks any different to a naive check: the empty one still renders the
    text ".csv .xlsx" (so matching on cell text calls it ready) and neither has
    real download buttons (so matching on buttons calls it still generating, and
    polls it for the entire collect budget). Both have to be recognised, because
    neither report is coming while we wait.
    """
    driver.get(pending["url"])
    time.sleep(settle)

    if pending.get("flavor") == "accordion":
        month_key = pending.get("month_key") or _month_key(pending.get("range"))
        want = "." + pending.get("fmt", "csv")
        for _ in range(4):          # first hit may only expand the accordion
            st = driver.execute_script(JS_ACCT_FORMAT_STATE, month_key, want)
            if st and st.get("action") == "expanding":
                time.sleep(4)
                continue
            return READY if (st and st.get("action") == "abrir") else WAITING
        return WAITING

    js = JS_AC_FIND_ROW if pending.get("flavor") == "wizard" else JS_FIND_ROW
    row = driver.execute_script(js, pending["range"], pending.get("fmt", "csv"))
    if not row:
        return WAITING
    if row.get("ready"):
        return READY

    # Ready, just not in the format we asked for. MercadoPago decides which
    # formats a report comes out in, and a finished report never grows the
    # missing one - so waiting for it is waiting forever. Take what exists and
    # say so; the file is the point, not its extension.
    fmts = [f for f in (row.get("formats") or []) if f in (".csv", ".xlsx")]
    if fmts:
        alt = fmts[0].lstrip(".")
        print("      %s 生成的是 %s 而非 .%s，改取 %s"
              % (pending.get("label", "?"), ",".join(fmts),
                 pending.get("fmt", "csv"), fmts[0]))
        pending["fmt"] = alt
        return READY

    if row.get("state") in (EMPTY, DELAYED):
        return row["state"]
    return WAITING


def fetch(driver, pending, download_dir):
    """Download a report probe() has just reported ready. Does not navigate."""
    import meli_forms as mf

    before = mf.snapshot_dir(download_dir)
    fmt = pending.get("fmt", "csv")
    if pending.get("flavor") == "accordion":
        month_key = pending.get("month_key") or _month_key(pending.get("range"))
        res = driver.execute_script(JS_ACCT_CLICK_FORMAT, month_key, "." + fmt)
    elif pending.get("flavor") == "wizard":
        res = driver.execute_script(JS_AC_CLICK_ROW, pending["range"], fmt)
    else:
        res = driver.execute_script(JS_CLICK_ROW, pending["range"], fmt)
    if not res:
        print("      [警告] %s 没有可点击的下载入口" % pending.get("label"))
        return []

    files = mf.wait_for_downloads(download_dir, before, min_files=1,
                                  timeout=150, settle=5,
                                  label="mercadopago/%s" % pending.get("report", "?"))
    for f in files:
        print("      + %s" % f)
    return files
