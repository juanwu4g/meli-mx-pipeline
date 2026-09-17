# -*- coding: utf-8 -*-
"""
Download the MercadoLibre + MercadoPago forms for ONE store.

    python run_downloads.py                    # config.json stores.default
    python run_downloads.py --store EWTTO_SM
    python run_downloads.py --only ventas --period-months 6

For several stores in one go use `run_batch.py`. It imports run_store() from
here, so both entry points drive exactly the same code - there is no second
copy of the pipeline to keep in sync.

--months         how many billing periods (Facturación) to pull
--period-months  the Ventas date range, set through the page's period dropdown
                 ("Últimos 6 meses"). This is what bounds the sales export, so
                 the default of 6 avoids silently truncating older months.

File counts vary by period - MercadoLibre only offers report types that have
data - so nothing here asserts a fixed number.

Files land in ./downloads/<STORE>/<timestamp>/ .
"""
import argparse
import datetime
import io
import json
import os
import sys
import time

from ziniao_client import ZiniaoClient, ZiniaoError, StoreNotAllowed
from store_config import StoreConfigError
import console  # noqa: F401  中文输出编码保护
import store_config
import meli_forms
import mercadopago
import run_report

ROOT = os.path.dirname(os.path.abspath(__file__))


def add_route_args(ap):
    """Options that describe *what to download*.

    Shared with run_batch.py so a batch and a single run are configured
    identically; anything specific to one entry point stays in that file.
    """
    ap.add_argument("--month", default=None, metavar="YYYY-MM",
                    help="补做某个历史会计月（如 2026-07）。账单改为直接访问该月的"
                         "明细页，销售窗口自动放大到覆盖该月。不给则按常规下最近的。")
    ap.add_argument("--billing-after", type=int, default=2, metavar="N",
                    help="配合 --month：目标月之后再多下 N 期账单（默认 2）。"
                         "月末订单的费用常记在下一期账单上，不下就会少算佣金。")
    ap.add_argument("--months", type=int, default=2,
                    help="how many billing periods, newest first")
    ap.add_argument("--skip-current", action="store_true",
                    help="skip the EN CURSO period (charges still accruing)")
    ap.add_argument("--only",
                    choices=["billing", "ventas", "stock", "mercadopago", "all"],
                    default="all",
                    help="run one route only")
    ap.add_argument("--mp-days", type=int, default=60,
                    help="MercadoPago manual report window in days (default 60)")
    ap.add_argument("--period-months", type=int, default=6,
                    help="Ventas date range: 2 or 6 (default 6)")
    ap.add_argument("--collect-timeout", type=int, default=1800,
                    help="total budget for the collect phase, seconds "
                         "(default 1800). This is a SHARED wait - all pending "
                         "reports generate concurrently, so it is not divided "
                         "between them.")
    ap.add_argument("--session-collect", type=int, default=180,
                    help="批量模式下，店内为 Ventas 单独兜底等待的秒数（默认 "
                         "180）。其余报表在店内只扫一趟即推迟到补收阶段，不占用 "
                         "这个预算。单店运行不走该路径，改用 --collect-timeout。")
    ap.add_argument("--late-timeout", type=int, default=420,
                    help="collect budget per store in run_batch.py's late pass, "
                         "seconds (default 420). Short on purpose: by then the "
                         "reports are usually hours old.")
    ap.add_argument("--collect-poll", type=int, default=20,
                    help="pause between collect cycles, seconds (default 20)")
    ap.add_argument("--skip-ip-check", action="store_true")
    ap.add_argument("--skip-lang-check", action="store_true",
                    help="跳过进店时的界面语言检查（调试用）")
    return ap


def month_stamp(stamp, args):
    """补做历史月份的目录加 _mYYYYMM 后缀。

    不加的话，一个只为 7 月拉的目录会混在常规目录里，出 8 月报表时可能被选中
    —— 而它里面根本没有 8 月的账单。后缀让人和程序都一眼看出这是补做的。
    """
    want = month_arg(args)
    return stamp + ("_m%04d%02d" % want if want else "")


def run_folder(store_name, stamp=None, base=None):
    """Absolute output folder for one store's run.

    MUST be absolute: CDP's Browser.setDownloadBehavior resolves a relative
    downloadPath against the browser process's own working directory, not ours,
    so files vanish somewhere under the Ziniao install and every route reports
    "got 0 files" while appearing to click correctly.
    """
    stamp = stamp or datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    if base:
        return os.path.abspath(os.path.join(base, store_name, stamp))
    return os.path.abspath(os.path.join(ROOT, "downloads", store_name, stamp))


# Distinguishable from a route's own None, which means "nothing to do" rather
# than "this blew up".
FAILED = object()

# A report MercadoPago refuses to build. Terminal for this run, but not a fault.
# NO_DATA: no movements in the range - that range will never produce a report.
# DELAYED: "hay datos en proceso ... espera unas horas" - worth a later run.
NO_DATA = object()
DELAYED = object()

# pending 列表里"数据处理中"那一类条目的稳定标记。
# run_batch.py 需要把这类条目和"仍在生成"的条目区分开，早先它是直接匹配显示
# 文案里的 "datos en proceso"。显示文案一旦改动（例如改成中文），那个匹配就会
# 静默失效 —— 判断逻辑不该依赖给人看的字符串。
DELAYED_MARK = "[delayed]"


# Errors _guard must NOT absorb, because the caller can act on them. A missing
# reports permission applies to the login, not to one report, so the caller
# stops asking the other five rather than logging the same wall six times.
PASSTHROUGH = (mercadopago.NoReportAccess,)


def _guard(result, label, fn, *a, **kw):
    """Run one route, recording a failure instead of aborting the store.

    A route that breaks - a changed selector, a page that never renders - must
    not cost the other three routes, and in a batch must not cost the remaining
    stores. The error is recorded and surfaced in the summary and exit code, so
    isolating it does not mean hiding it.

    Returns FAILED on an exception, except for PASSTHROUGH errors which are
    re-raised for the caller to handle.
    """
    try:
        return fn(*a, **kw)
    except PASSTHROUGH:
        raise
    except Exception as e:
        msg = "%s: %s: %s" % (label, e.__class__.__name__, e)
        print("  [错误] %s" % msg)
        result["errors"].append(msg)
        return FAILED


def _request(result, label, pending, fn, *a, **kw):
    """Phase-1 request: treat a silent no-op as a failure too.

    The request_* helpers warn and return None when they cannot find what they
    need, rather than raising - right, because one bad report should not sink
    the other five. But that must not vanish from the summary. On 2026-09-02 all
    six MercadoPago requests returned None on EWTTO_SM (no reports permission on
    that login) and the store still reported `ok, 13 files`. A route asked to
    run and producing nothing is a failure, not a success.
    """
    p = _guard(result, label, fn, *a, **kw)
    if p is FAILED:
        return                      # already recorded
    if p is None:
        msg = "%s: requested nothing (see the [warn] above)" % label
        print("  [错误] %s" % msg)
        result["errors"].append(msg)
        return
    pending.append(p)


def label_of(p):
    return ("ventas" if p["kind"] == "ventas"
            else "mercadopago/%s" % p.get("report", "?"))


def collect_pending(d, pending, out_dir, result, budget, poll=20):
    """Cycle over outstanding reports until they are collected or time runs out.

    ONE SHARED wait, not a budget divided between items. The reports generate
    concurrently on the server, so time spent waiting on one advances all of
    them. An earlier version split the budget - with six reports and 600s that
    gave each 100s, and nothing taking 2.6-6.5 minutes could ever finish in its
    slice.

    Terminal answers (no data, data-in-process, a probe that blew up) drop out
    of the cycle immediately rather than burning the budget for the reports that
    are still genuinely generating.

    Returns the entries still outstanding when it stops.
    """

    def probe_and_fetch(p):
        """Check one item; download it if ready.

        Returns the files, [] to poll again, or a terminal sentinel.
        """
        if p["kind"] == "ventas":
            if not meli_forms.probe_sales_excel(d, p):
                return []
            got = meli_forms.fetch_sales_excel(d, p, out_dir)
            result["sales"].extend(got)
            return got
        state = mercadopago.probe(d, p)
        if state == mercadopago.EMPTY:
            return NO_DATA
        if state == mercadopago.DELAYED:
            return DELAYED
        if state != mercadopago.READY:
            return []
        got = mercadopago.fetch(d, p, out_dir)
        result["mp"].extend(got)
        return got

    deadline = time.time() + budget
    remaining = list(pending)
    cycle = 0
    # do-while：至少完整扫一轮再谈超时。budget=0 就是"只扫一遍、不等待"，
    # 这正是店内收取需要的语义 —— 探测本身的开销省不掉，省掉的是轮询空等。
    while remaining:
        cycle += 1
        print("\n>>> 收取轮次 %d —— 剩余 %d 秒，%d 张待取：%s"
              % (cycle, int(deadline - time.time()), len(remaining),
                 ", ".join(label_of(p) for p in remaining)))
        for p in list(remaining):
            # 第一轮必须完整扫完再谈超时：budget=0 的语义是"扫一遍就走"，
            # 若在这里提前跳出，这一趟等于什么都没做。
            if cycle > 1 and time.time() >= deadline:
                break
            got = _guard(result, "%s/collect" % label_of(p), probe_and_fetch, p)
            if got is FAILED:
                # the probe itself blew up; stop retrying this one
                remaining.remove(p)
            elif got is DELAYED:
                # MercadoPago has data still settling and will not build the
                # report for hours. Nothing to wait for now, but the request
                # stands - a later run collects it.
                print("    %-28s MercadoPago 数据处理中，留到后续运行再取"
                      % label_of(p))
                result["pending"].append(
                    "mercadopago/%s (%s) - %s 数据处理中"
                    % (p["report"], p["range"], DELAYED_MARK))
                # Out of THIS cycle, but not given up on: the request stands and
                # the report appears in a few hours, so it stays in the registry
                # for a later run. An EMPTY one is dropped instead - that range
                # will never produce a report, so keeping it would mean probing
                # it forever.
                result.setdefault("retry_later", []).append(p)
                remaining.remove(p)
            elif got is NO_DATA:
                # Terminal, and not a fault: MercadoPago will not build a report
                # for a range with no movements.
                print("    %-28s 该区间无流水，跳过"
                      % label_of(p))
                result["empty"].append(
                    "mercadopago/%s (%s)" % (p["report"], p["range"]))
                remaining.remove(p)
            elif got:
                remaining.remove(p)
            else:
                print("    %-28s 尚未就绪" % label_of(p))
        if not remaining or time.time() >= deadline:
            break
        time.sleep(poll)
    return remaining


# 销售报表的"最近 N 个月"菜单**只有 2 和 6 两档**。
# 别加 3 ——实测要 3 的时候 set_sales_period 找不到选项、打一行警告就返回 False，
# 而调用方不看返回值，于是照样按默认的 2 个月导出：做 7 月报表时窗口停在
# 07-13，7 月 1-12 日整段丢失，GMV 少了 36%（1,241,893 vs 1,961,431）。
# 超过 6 个月的目标月这里够不着，菜单里的 Último año / Fecha personalizada
# 要用 download_history.py 那条路径。
SALES_WINDOWS = (2, 6)


def month_arg(args):
    """--month 解析成 (年, 月)；没给返回 None。格式不对直接报错，不猜。"""
    v = getattr(args, "month", None)
    if not v:
        return None
    try:
        y, m = int(v[:4]), int(v[5:7])
        if v[4] != "-" or not 1 <= m <= 12:
            raise ValueError
    except Exception:
        raise SystemExit("--month 要写成 YYYY-MM，比如 2026-07；收到的是 %r" % v)
    return y, m


def sales_window(args, today=None):
    """销售报表要拉几个月，才能盖住目标会计月。

    平台的"最近 N 个月"是从今天往回算的，所以做 7 月报表时，窗口必须长到
    退回 7 月**月初**。取菜单里能覆盖到的最小档 —— 窗口越大，导出越慢、
    行数越多，没必要一律拉满。
    """
    want = month_arg(args)
    if not want:
        return args.period_months
    today = today or datetime.date.today()
    y, m = want
    # 目标月月初距今几个月（向上取整到整月）
    back = (today.year - y) * 12 + (today.month - m) + 1
    for w in SALES_WINDOWS:
        if w >= back:
            if w != args.period_months:
                print("  [%s] 销售窗口放大到最近 %d 个月，以覆盖 %04d-%02d"
                      % ("month", w, y, m))
            return w
    print("  [警告] %04d-%02d 距今 %d 个月，超出销售报表菜单最大的 %d 个月，"
          "该月销售数据可能拉不全" % (y, m, back, SALES_WINDOWS[-1]))
    return SALES_WINDOWS[-1]


def run_store(client, store_name, args, out_dir=None, close_when_done=True,
              defer_slow=False):
    """Open one store, run the requested routes, return what was collected.

    Raises only if the store cannot be opened at all (not on the allowlist, not
    on the account, startBrowser failed). Once the browser is up, individual
    route failures are captured in result["errors"] rather than raised, so a
    caller looping over stores keeps going.
    """
    store_config.require_allowed(store_name)
    out_dir = out_dir or run_folder(store_name)
    if not os.path.isdir(out_dir):
        os.makedirs(out_dir)
    print("输出目录：%s" % out_dir)

    result = {"store": store_name, "out_dir": out_dir,
              "billing": {}, "sales": [], "stock": {}, "mp": [],
              "pending": [], "empty": [], "deferred": [],
              "errors": [], "seconds": 0}
    began = time.time()

    store = client.find_store(store_name)
    print("店铺：%s（%s）" % (store.get("browserName"), store.get("platform_name")))

    with client.open_store(store, close_when_done=close_when_done) as session:
        print("调试端口：%s | 内核 %s"
              % (session.info.get("debuggingPort"), session.info.get("core_version")))

        if not args.skip_ip_check:
            ok = session.check_ip()
            print("IP 检测：%s" % ("通过" if ok else "失败"))
            if not ok:
                result["errors"].append("ip check failed")
        else:
            session.driver.get(session.launcher_page)

        d = session.driver
        meli_forms.set_download_dir(d, out_dir)
        want0 = month_arg(args)
        result["month"] = "%04d-%02d" % want0 if want0 else None
        # 模块级变量，批量跑时会带着上一家店的值进来 —— 这家店若跳过销售
        # 路线，就会把上一家的窗口错记到这家头上。每家店开跑前清掉。
        meli_forms.SALES_PERIOD_ACTUAL["months"] = None
        mercadopago.LANG_CHECKED["done"] = False

        # ---------------- PHASE 0: 语言关卡 ----------------------------------
        # 必须排在所有路线之前：整个下载层按西语文字定位元素，导出文件的列名
        # 也是西语，而导出语言跟着账号的界面语言走。放在后面就来不及了 ——
        # 账期页的月份名核对会静默跳过整期，费用少算而毫无提示。
        if not args.skip_lang_check:
            got = _guard(result, "language", meli_forms.ensure_language, d)
            if got is not FAILED:
                was, changed = got
                result["locale"] = was
                if changed:
                    result["errors"].append(
                        "language: 界面语言原本是 %s，已改回 es_MX。"
                        "请查清是谁改的 —— 期间导出的报表列名可能是错的。" % was)

        # ---------------- PHASE 1: request everything that generates slowly ---
        # These sit in a server-side queue while phase 2 does real work, so the
        # wait costs almost nothing. Nothing here downloads.
        pending = []
        if args.only in ("ventas", "all"):
            _request(result, "ventas/request", pending,
                     meli_forms.request_sales_excel,
                     d, period_months=sales_window(args))
        if args.only in ("mercadopago", "all"):
            # every report type under /balance/reports shares one UI
            for kind in mercadopago.REPORTS:
                try:
                    _request(result, "mercadopago/%s/request" % kind, pending,
                             mercadopago.request_report,
                             d, report=kind, days=args.mp_days)
                except mercadopago.NoReportAccess as e:
                    # The wall is on the login, not the report, so the other
                    # five would fail identically. Say it once and move on.
                    print("  [错误] mercadopago：%s" % e)
                    result["errors"].append("mercadopago: %s" % e)
                    break
        if pending:
            print("\n>>> 阶段一完成：%d 张报表已在后台生成"
                  % len(pending))

        # ---------------- PHASE 2: everything that downloads immediately ------
        if args.only in ("billing", "all"):
            # 指定了历史月份就直接访问那一期的明细页 —— resume 页只摆最近 3 期
            # 卡片，更早的账期在页面上点不到，但明细页还在。
            want = month_arg(args)
            if want:
                got = _guard(result, "billing",
                             meli_forms.download_billing_for_months,
                             d, out_dir, want[0], want[1],
                             extra=args.billing_after)
            else:
                got = _guard(result, "billing",
                             meli_forms.download_billing_reports,
                             d, out_dir, months=args.months,
                             skip_current=args.skip_current)
            result["billing"] = got if isinstance(got, dict) else {}
        if args.only in ("stock", "all"):
            got = _guard(result, "stock", meli_forms.download_stock_reports,
                         d, out_dir, months_back=2)
            result["stock"] = got if isinstance(got, dict) else {}

        # ---------------- PHASE 3: go back and collect ------------------------
        # In a batch this is deliberately SHORT. The slow MercadoPago reports
        # keep generating server-side whether or not this store's browser is
        # open, so waiting here blocks twelve other stores for nothing. Whatever
        # is not ready is written to the pending registry and collected by a
        # late pass once every store has been downloaded.
        #
        # Ventas is the exception and is always collected in-session: it is
        # ready in 30-60s, was never the straggler in any run, and it arrives
        # through a notification widget whose behaviour across a browser
        # close/reopen is unverified. No reason to risk it for no gain.
        if defer_slow:
            # 店内只扫一趟，不等待。慢报表在 MercadoPago 端照样生成，守在这里
            # 等它没有任何意义 —— 上次运行有 5 家店各空等约 280 秒，最后仍然
            # 推迟。所有轮询等待都归补收阶段。
            if pending:
                print("\n>>> 阶段三：扫描 %d 张待取报表（只扫一趟，不等待）"
                      % len(pending))
            remaining = collect_pending(d, pending, out_dir, result, budget=0,
                                        poll=args.collect_poll)

            # 唯一的例外是 Ventas：本店对账簿要用它。它在阶段一最先申请，之后
            # 阶段二还跑了几分钟，实测一直是第一趟就绪，兜底基本不会触发。
            ventas_left = [p for p in remaining if p["kind"] == "ventas"]
            if ventas_left and args.session_collect > 0:
                print("\n>>> Ventas 尚未就绪，单独再等最多 %d 秒（对账簿需要它）"
                      % args.session_collect)
                still = collect_pending(d, ventas_left, out_dir, result,
                                        budget=args.session_collect,
                                        poll=args.collect_poll)
                remaining = [p for p in remaining
                             if p["kind"] != "ventas"] + still
        else:
            budget = args.collect_timeout
            if pending:
                print("\n>>> 阶段三：收取 %d 张待取报表，预算 %d 秒"
                      % (len(pending), budget))
            remaining = collect_pending(d, pending, out_dir, result, budget,
                                        poll=args.collect_poll)

        if defer_slow and remaining:
            # Ventas is deferrable too. Verified 2026-09-09: requested in one
            # browser session, the store closed, reopened on a different
            # debugging port, and the file collected on the first probe. Its
            # readiness lives in MercadoLibre's notification widget, which is
            # server state, not anything held by the session.
            #
            # It is still normally collected in-session - it is ready in 30-60s,
            # and the store's accounting workbook needs it - so this is a safety
            # net, not the usual path. Before this, a ventas that missed the
            # budget was recorded as pending and then collected by nobody.
            print("\n    %d 张报表推迟到补收阶段"
                  % len(remaining))
            for p in remaining:
                print("      %s (%s)" % (label_of(p), p.get("range", "-")))
            result["deferred"] = list(remaining)
            remaining = []

        for p in remaining:
            result["pending"].append(
                "ventas" if p["kind"] == "ventas"
                else "mercadopago/%s (%s)" % (p["report"], p["range"]))

        # leave the browser on a normal page so a later run can reattach
        d.get(session.launcher_page)

    result["seconds"] = int(time.time() - began)

    # A written record lands beside the files themselves, so the folder still
    # explains itself once the terminal output is gone.
    result["sales_window_months"] = meli_forms.SALES_PERIOD_ACTUAL.get("months")
    write_run_meta(out_dir, result, began)
    result["report"] = run_report.write(
        result, args=args,
        started=datetime.datetime.fromtimestamp(began))
    return result


def write_run_meta(out_dir, result, began):
    """往下载目录写一份机器可读的 run_meta.json。

    只放报表端**判断数据完整性**需要的东西。目前就一项：销售报表实际生效的
    时间范围。没有它就分不清"销售数据只到 7 月 13 号"是窗口不够、还是这家店
    7 月 13 号才开张 —— 两者在数据里长得一模一样，结论却完全相反。
    """
    path = os.path.join(out_dir, "run_meta.json")
    meta = {}
    if os.path.isfile(path):          # 合并，不是覆盖
        try:
            with io.open(path, encoding="utf-8") as fh:
                meta = json.load(fh) or {}
        except Exception:
            meta = {}
    meta["downloaded_at"] = datetime.datetime.fromtimestamp(began).isoformat(timespec="seconds")
    meta["month"] = result.get("month")
    # 这一趟没跑销售路线就别把上一趟记下的窗口抹掉 —— --only billing 补下账单
    # 到同一目录时会走到这里，覆盖式写入会把 6 改成 null，校验就从"通过"
    # 变成"跳过"，白白丢掉一条能救命的检查。
    if result.get("sales_window_months") is not None:
        meta["sales_window_months"] = result["sales_window_months"]
    meta.setdefault("sales_window_months", None)
    try:
        with io.open(path, "w", encoding="utf-8") as fh:
            json.dump(meta, fh, ensure_ascii=False, indent=2)
    except Exception as e:            # 元数据写不出不该拖垮已经下好的文件
        print("  [警告] run_meta.json 写入失败：%s" % str(e)[:80])
    return meta


def count_files(result):
    """Files this run actually downloaded."""
    n = sum(len(v) for v in result["billing"].values())
    n += len(result["sales"])
    n += sum(len(v) for v in result["stock"].values())
    n += len(result["mp"])
    return n


def print_summary(result):
    print("\n" + "=" * 60)
    print("汇总 —— %s" % result["store"])
    print("=" * 60)
    for period, files in result["billing"].items():
        print("  Facturación %-18s %d 个文件" % (period, len(files)))
        for f in files:
            print("      %s" % f)
    print("  Ventas %-23s %d 个文件" % ("", len(result["sales"])))
    for f in result["sales"]:
        print("      %s" % f)
    for key, files in result["stock"].items():
        print("  Stock %-24s %d 个文件" % (key, len(files)))
        for f in files:
            print("      %s" % f)
    print("  MercadoPago %-18s %d 个文件" % ("", len(result["mp"])))
    for f in result["mp"]:
        print("      %s" % f)

    if result["pending"]:
        print("\n  [待取] 本次未能及时就绪，后续运行可取回：")
        for x in result["pending"]:
            print("      %s" % x)
    if result.get("empty"):
        print("\n  [无数据] 该区间没有可报的内容，这不是错误：")
        for x in result["empty"]:
            print("      %s" % x)
    if result["errors"]:
        print("\n  [错误] %d 条路线失败：" % len(result["errors"]))
        for x in result["errors"]:
            print("      %s" % x)

    # .csv matters: the MercadoPago reports are csv, and counting only xlsx
    # under-reported every run that included them.
    on_disk = [f for f in os.listdir(result["out_dir"])
               if f.lower().endswith((".xlsx", ".xls", ".csv"))]
    print("\n  本次下载合计 ： %d" % count_files(result))
    print("  目录中文件数 ： %d  位于 %s" % (len(on_disk), result["out_dir"]))
    print("  用时         ： %d 分 %02d 秒"
          % (result["seconds"] // 60, result["seconds"] % 60))
    if result.get("report"):
        print("  运行报告     ： %s" % result["report"])
    return on_disk


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--store", default=None,
                    help="store name; defaults to config.json stores.default")
    ap.add_argument("--expect", type=int, default=0,
                    help="expected excel count; 0 (default) disables the check, "
                         "since the count varies by period")
    ap.add_argument("--out", default=None, help="output directory")
    ap.add_argument("--keep-open", action="store_true",
                    help="leave the store browser open when finished")
    ap.add_argument("--no-restart", action="store_true",
                    help="assume the client is already in webdriver mode")
    add_route_args(ap)
    args = ap.parse_args()

    # Resolve once: the name is needed for the output path before the client
    # is even started, and --store now defaults to None.
    store_name = store_config.require_allowed(
        args.store or store_config.default_store())

    # stores.allowed is a PERMISSION list, not a work list. Adding a second
    # store there is the natural way to expect two downloads, and this script
    # will still run exactly one - so say which one, and where the other went.
    others = [n for n in store_config.allowed_stores() if n != store_name]
    if others:
        print("提示：本次只跑 %s。白名单里还有 %d 家店铺（%s）——"
              "要全部下载请用 run_batch.py。"
              % (store_name, len(others), ", ".join(others)))

    out_dir = (os.path.abspath(args.out) if args.out else
               run_folder(store_name, month_stamp(
                   datetime.datetime.now().strftime("%Y%m%d_%H%M%S"), args)))

    client = ZiniaoClient()
    client.start(restart=not args.no_restart)

    result = run_store(client, store_name, args, out_dir=out_dir,
                       close_when_done=not args.keep_open)
    on_disk = print_summary(result)

    if args.expect and len(on_disk) != args.expect:
        print("\n  [警告] 预期 %d 个文件，实际 %d 个"
              % (args.expect, len(on_disk)))
        return 1
    if result["errors"]:
        return 1
    print("\n  %d 个文件全部就位" % len(on_disk))
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except StoreConfigError as e:
        print("\n配置错误：%s" % e)
        sys.exit(4)
    except StoreNotAllowed as e:
        # Scope error, not a failure: the store simply is not on the allowlist.
        print("\n店铺不在白名单：%s" % e)
        sys.exit(3)
    except ZiniaoError as e:
        print("\n紫鸟客户端错误：%s" % e)
        sys.exit(2)
