# -*- coding: utf-8 -*-
"""
Run the download pipeline over several stores in one client session.

    python run_batch.py                          # every store in stores.batch
    python run_batch.py --stores BOCINA_SM EWTTO_SM
    python run_batch.py --only ventas --limit 2
    python run_batch.py --dry-run                # just show what would run

Which stores
------------
`--stores` if given, else `stores.batch` from config.json, else the whole
`stores.allowed` list. Every name still goes through the allowlist, so a typo
in `--stores` is rejected locally without a request reaching Ziniao.

If the allowlist is empty (= unrestricted) there is nothing to iterate and the
batch refuses to start. "Every store on the account" is 93 stores and is never
what anyone means.

Why serial
----------
One browser at a time. Whether Ziniao will hold two stores open at once is
untested, and each store has its own proxy and fingerprint - opening several in
parallel multiplies both load and the chance of tripping a verification
challenge. The client itself starts ONCE and is reused, so the per-store cost
is just startBrowser/stopBrowser, not a full client restart.

Failure isolation
-----------------
Nothing one store does can end the batch. A route that breaks is recorded by
run_store(); a store that cannot even be opened is caught here. Both show up in
the summary and in the exit code:

    0   every store ok
    1   at least one store failed or finished with route errors
    4   config.json is unreadable, or the batch list is empty
"""
import argparse
import datetime
import io
import json
import os
import sys
import time
import traceback

from ziniao_client import ZiniaoClient, ZiniaoError, StoreNotAllowed
from store_config import StoreConfigError
import console  # noqa: F401  中文输出编码保护
import store_config
import run_downloads
import transform
import pending_store
import meli_forms

ROOT = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(ROOT, "logs")
NEWLINE = chr(10)


def resolve_stores(args):
    """The store list for this batch, already checked against the allowlist."""
    if args.stores:
        names = [store_config.require_allowed(n) for n in args.stores]
    else:
        names = store_config.batch_stores()
    if not names:
        raise StoreConfigError(
            "no stores to run: config.json -> stores.allowed is empty "
            "(unrestricted), so there is no list to walk. Name the stores in "
            "stores.allowed, or pass --stores.")
    if args.limit:
        names = names[:args.limit]
    return names


def write_summary(path, stamp, rows, transform_result=None):
    """Machine-readable record of the run.

    A scheduled batch throws stdout away, so without this a silent failure is
    invisible until someone notices missing data.
    """
    if not os.path.isdir(LOG_DIR):
        os.makedirs(LOG_DIR)
    doc = {"started": stamp, "stores": rows}
    if transform_result is not None:
        doc["transform"] = transform_result
    # utf-8 explicitly: store names are Chinese on some accounts and the
    # default Windows encoding raises UnicodeEncodeError (GUIDE.md 6.4).
    with io.open(path, "w", encoding="utf-8") as f:
        f.write(json.dumps(doc, ensure_ascii=False, indent=2))


def rel_to_downloads(out_dir, name):
    """A downloaded file as a path relative to downloads/.

    That is what the cleaner's matcher wants. A bare filename also matches, but
    against every copy of that name in the tree.
    """
    return os.path.relpath(os.path.join(out_dir, os.path.basename(name)),
                           transform.PIPELINE_ROOT).replace(chr(92), "/")


def make_workbook(row, stamp, args):
    """Write this store's accounting workbook, now that its ventas is in hand.

    Per store, as soon as that store is done - not batched to the end. A store
    that finishes at 09:10 has a usable workbook at 09:10 rather than after the
    twelfth store, which matters when a run is interrupted: work already done is
    already delivered.

    `validate` reads the RAW export and pairs it with the facturacion files in
    the SAME run folder, so it needs neither the Parquet build nor any other
    store. That independence is what makes per-store generation correct rather
    than merely convenient.
    """
    if row.get("workbook") or not row.get("ventas"):
        return
    out = os.path.join(transform.PIPELINE_ROOT, "data", "reports",
                       "%s_%s.xlsx" % (row["store"], stamp))
    res = transform.validate(row["ventas"][0], out)
    if res["ok"]:
        row["workbook"] = out
    else:
        row["errors"] = list(row.get("errors", [])) + [
            "workbook: %s" % (res["output"] or "").strip()[-200:]]
        if row["status"] == "ok":
            row["status"] = "partial"


def collect_pass(client, names, args, rows):
    """Reopen each store once and collect the reports deferred during its run.

    This is where the reordering pays off. The slow MercadoPago reports keep
    generating server-side whether or not the store's browser is open, so the
    per-store run no longer waits on them - by the time this pass starts, the
    first store's reports are hours old and the last store's have had the whole
    batch to build.

    Stores are visited in the SAME ORDER they were downloaded, which is what
    gives the last store its head start: it is collected last, so it gains the
    duration of every other store's collection for free.

    The budget per store is short on purpose (--late-timeout, default 420s).
    Anything still not ready stays in the registry and a later run picks it up;
    waiting longer here just re-creates the serial stall this pass removes.
    """
    outstanding = pending_store.load()
    mine = [e for e in outstanding if e.get("store") in set(names)]
    if not mine:
        print("%s>>> 补收阶段：没有待取回的报表" % NEWLINE)
        return
    others = len(outstanding) - len(mine)

    print("%s%s" % (NEWLINE, "=" * 60))
    print("补收阶段 —— %d 张报表，涉及 %d 家店铺"
          % (len(mine), len({e["store"] for e in mine})))
    if others:
        print("  （另有 %d 张属于本批次以外的店铺，不处理）" % others)
    print("=" * 60)

    by_store = {}
    for e in mine:
        by_store.setdefault(e["store"], []).append(e)

    for name in names:                      # same order as the download loop
        entries = by_store.get(name)
        if not entries:
            continue
        row = next((r for r in rows if r["store"] == name), None)
        print("%s--- 补收：%s（%d 张报表）"
              % (NEWLINE, name, len(entries)))

        result = {"store": name, "sales": [], "mp": [], "pending": [],
                  "empty": [], "errors": [], "retry_later": []}
        late_dir = None
        keep = []
        try:
            store = client.find_store(name)
            with client.open_store(store, close_when_done=True) as session:
                d = session.driver
                d.get(session.launcher_page)
                # Entries from different runs land in different folders, and the
                # download directory is a per-SESSION CDP setting - so it has to
                # be re-applied here, once per folder. Forgetting this is how a
                # route ends up reporting "got 0 files" while clicking correctly.
                folders = {}
                for e in entries:
                    folders.setdefault(e.get("out_dir"), []).append(e)
                for out_dir, group in folders.items():
                    if not out_dir or not os.path.isdir(out_dir):
                        # the run folder is gone; put the file with this batch
                        out_dir = run_downloads.run_folder(name, args._stamp,
                                                           args.out)
                        if not os.path.isdir(out_dir):
                            os.makedirs(out_dir)
                        print("    原下载目录已不存在，改存到 %s" % out_dir)
                    meli_forms.set_download_dir(d, out_dir)
                    late_dir = out_dir
                    left = run_downloads.collect_pending(
                        d, group, out_dir, result,
                        budget=args.late_timeout, poll=args.collect_poll)
                    keep.extend(left)
                d.get(session.launcher_page)
        except Exception as e:
            # A store that will not open must not cost the other stores' reports.
            # Its entries stay in the registry, so nothing is lost.
            print("    [失败] %s: %s" % (e.__class__.__name__, e))
            result["errors"].append("late collect: %s: %s"
                                    % (e.__class__.__name__, e))
            keep = list(entries)

        # Still generating when the budget ran out. Separated from retry_later
        # because that one already has its own summary line ("datos en proceso")
        # and would otherwise be counted twice.
        waiting = list(keep)
        keep.extend(result.get("retry_later", []))
        collected = len(entries) - len(keep)
        print("    已取回 %d 张，仍欠 %d 张" % (collected, len(keep)))

        # Everything not kept is done with - downloaded, or a range that will
        # never produce a report. Removing it is what stops the registry growing
        # without bound.
        done = [e for e in entries if e not in keep]
        if done:
            pending_store.remove(done)

        if row is not None:
            if result["sales"] and late_dir:
                # ventas arrived late, so this store has no workbook yet
                row["ventas"] = list(row.get("ventas", [])) + [
                    rel_to_downloads(late_dir, f) for f in result["sales"]]
            row["files"] = row.get("files", 0) + len(result["mp"]) + len(result["sales"])
            row["late_collected"] = len(result["mp"]) + len(result["sales"])
            row["empty"] = list(row.get("empty", [])) + result["empty"]
            # Report what is still owed. The registry knowing about it is not
            # enough: a summary that says pending=[] while a report is still
            # generating is the same "silence reads as success" failure that
            # once let six failed requests pass as `ok, 13 files`.
            row["pending"] = (
                [p for p in row.get("pending", [])
                 if run_downloads.DELAYED_MARK not in p]
                + result["pending"]
                + ["mercadopago/%s (%s) - still generating"
                   % (e.get("report"), e.get("range")) for e in waiting])
            if result["errors"]:
                row["errors"] = list(row.get("errors", [])) + result["errors"]
                if row["status"] == "ok":
                    row["status"] = "partial"



def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--stores", nargs="+", default=None,
                    help="store names to run; default is config stores.batch")
    ap.add_argument("--limit", type=int, default=0,
                    help="run only the first N stores (0 = all)")
    ap.add_argument("--out", default=None,
                    help="base output dir; each store gets <base>/<STORE>/<stamp>/")
    ap.add_argument("--pace", type=int, default=5,
                    help="seconds between stores (default 5)")
    ap.add_argument("--dry-run", action="store_true",
                    help="print the store list and exit without opening anything")
    ap.add_argument("--no-restart", action="store_true",
                    help="assume the client is already in webdriver mode")
    ap.add_argument("--collect-only", action="store_true",
                    help="download nothing; just reopen the stores that still "
                         "owe reports (logs/pending.json) and collect them. This "
                         "is how you pick up what a previous run left generating.")
    ap.add_argument("--no-late-collect", action="store_true",
                    help="do not reopen stores to collect the reports their own "
                         "run deferred; they stay in logs/pending.json for a "
                         "later run")
    ap.add_argument("--workbooks", action="store_true",
                    help="after cleaning, also write one accounting workbook per "
                         "store into downloads/data/reports/. With this the run "
                         "goes download -> clean -> report with no human input.")
    ap.add_argument("--no-transform", action="store_true",
                    help="skip the cleaning pipeline in downloads/mx_sales, "
                         "which otherwise runs once after every store is done")
    ap.add_argument("--stop-on-error", action="store_true",
                    help="abort the batch on the first store that fails "
                         "(default is to carry on)")
    run_downloads.add_route_args(ap)
    args = ap.parse_args()

    names = resolve_stores(args)
    stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    stale = pending_store.prune()
    if stale:
        print("已清理 %d 张超过 %d 天的过期待取记录"
              % (len(stale), pending_store.MAX_AGE_DAYS))

    print("=" * 60)
    print("批次 %s —— %d 家店铺：%s" % (stamp, len(names), ", ".join(names)))
    print("路线：%s | 白名单：%s" % (args.only, store_config.describe()))
    print("=" * 60)
    if args.dry_run:
        for i, n in enumerate(names, 1):
            print("  %2d. %-16s -> %s"
                  % (i, n, run_downloads.run_folder(n, stamp, args.out)))
        return 0

    client = ZiniaoClient()
    client.start(restart=not args.no_restart)

    rows = []
    began = time.time()

    if args.collect_only:
        # Nothing to download: this run exists only to finish what an earlier
        # one started. Restrict to stores that actually owe something, so a
        # 12-store allowlist does not open 12 browsers to do nothing.
        owed = {e.get("store") for e in pending_store.load()}
        names = [n for n in names if n in owed]
        if not names:
            print("%s这些店铺没有待取回的报表，无需补收"
                  % NEWLINE)
            return 0
        print("%s仅补收模式：%d 家店铺有待取回的报表：%s"
              % (NEWLINE, len(names), ", ".join(names)))
        rows = [{"store": n, "status": "ok", "files": 0, "errors": [],
                 "pending": [], "empty": [], "out_dir": None, "seconds": 0}
                for n in names]

    for i, name in enumerate([] if args.collect_only else names, 1):
        print("\n" + "#" * 60)
        print("# [%d/%d] %s" % (i, len(names), name))
        print("#" * 60)
        row = {"store": name, "status": "failed", "files": 0,
               "errors": [], "pending": [], "empty": [],
               "out_dir": None, "seconds": 0}
        t0 = time.time()
        try:
            out_dir = run_downloads.run_folder(name, stamp, args.out)
            # close_when_done is not negotiable in a batch: the next store
            # needs the browser slot, and a store left open would still be
            # holding its proxy.
            result = run_downloads.run_store(client, name, args,
                                             out_dir=out_dir,
                                             close_when_done=True,
                                             defer_slow=True)
            run_downloads.print_summary(result)
            # Keep the ventas export's path RELATIVE TO downloads/, which is
            # what the cleaner's matcher wants. A bare filename would also match,
            # but against every copy in the tree.
            row["ventas"] = [rel_to_downloads(result["out_dir"], f)
                             for f in result.get("sales", [])]
            deferred = result.get("deferred", []) + result.get("retry_later", [])
            if deferred:
                entries = []
                for p in deferred:
                    e = dict(p)
                    e["store"] = name
                    e["out_dir"] = result["out_dir"]
                    e.setdefault("requested_at", time.time())
                    entries.append(e)
                # Written NOW, while the store is fresh in hand. A crash between
                # this store and the late pass would otherwise leave reports
                # generating on MercadoPago that nobody knows to collect.
                pending_store.add(entries)
            row.update(
                status="partial" if result["errors"] else "ok",
                files=run_downloads.count_files(result),
                errors=result["errors"],
                pending=result["pending"],
                empty=result.get("empty", []),
                out_dir=result["out_dir"])
        except StoreNotAllowed as e:
            row["status"] = "not-allowed"
            row["errors"] = [str(e)]
            print("  [跳过] %s" % e)
        except Exception as e:
            # Deliberately broad. The point of a batch is that store 7 dying in
            # a way nobody predicted still leaves stores 8-25 to run.
            row["errors"] = ["%s: %s" % (e.__class__.__name__, e)]
            print("  [失败] %s: %s" % (e.__class__.__name__, e))
            traceback.print_exc()
        row["seconds"] = int(time.time() - t0)
        # The workbook is written HERE, per store, not batched to the end.
        if args.workbooks:
            make_workbook(row, stamp, args)
        rows.append(row)

        if row["status"] not in ("ok", "partial") and args.stop_on_error:
            print("\n--stop-on-error：在 %s 之后中止批次" % name)
            break
        if i < len(names):
            time.sleep(args.pace)

    # ---- collect what the per-store runs deferred ------------------------
    args._stamp = stamp
    if not args.no_late_collect:
        collect_pass(client, names, args, rows)
    elif pending_store.load():
        print("%s>>> 已跳过补收阶段（--no-late-collect），仍欠：%s"
              % (NEWLINE, pending_store.describe()))

    # ---- hand the new files to the cleaning pipeline --------------------
    # Once, after every store: it walks the whole download tree, so running it
    # per store would repeat the same work. A failure here is reported but
    # cannot undo the downloads, which are already on disk.
    tr = None
    if not args.no_transform:
        tr = transform.build()
    elif transform.available():
        print("\n>>> 已跳过清洗（--no-transform），之后可手动执行：")
        print("    cd downloads && python -m mx_sales build")

    # ---- catch up a store whose ventas only arrived in the late pass -----
    # Workbooks are normally written the moment their store finishes. This is
    # only for a store whose ventas missed the in-session budget and was
    # collected late; without it that store would silently end up without one.
    if args.workbooks:
        late = [r for r in rows if r.get("ventas") and not r.get("workbook")]
        if late:
            print("%s>>> 对账簿：%d 家店铺的 Ventas 是补收回来的，现在补生成"
                  % (NEWLINE, len(late)))
            for row in late:
                make_workbook(row, stamp, args)
        for row in rows:
            if not row.get("ventas") and row["status"] in ("ok", "partial"):
                print("    %-16s 没有 Ventas 导出文件，不生成对账簿" % row["store"])

    log_path = os.path.join(LOG_DIR, "batch_%s.json" % stamp)
    write_summary(log_path, stamp, rows, transform_result=tr)

    elapsed = int(time.time() - began)
    print("\n" + "=" * 60)
    print("批次汇总 —— 用时 %d 分 %02d 秒" % (elapsed // 60, elapsed % 60))
    print("=" * 60)
    print("  %-16s %-12s %6s  %s" % ("店铺", "状态", "文件数", "备注"))
    for r in rows:
        # Every fact that applies, not just the first one found. An earlier
        # version stopped at the first match, so a store with one report still
        # generating was reported as "1 still generating" and nothing else -
        # hiding that it had produced its workbook perfectly well.
        parts = []
        if r["errors"]:
            parts.append("%d 处错误：%s" % (len(r["errors"]), r["errors"][0][:44]))
        if r["pending"]:
            parts.append("%d 张仍在生成" % len(r["pending"]))
        if r["empty"]:
            # Worth showing, but not a problem: the range simply held nothing.
            parts.append("%d 张无流水" % len(r["empty"]))
        if r.get("workbook"):
            parts.append("已出对账簿")
        print("  %-16s %-12s %6d  %s"
              % (r["store"], r["status"], r["files"], " | ".join(parts)))
    print("\n  汇总已写入 %s" % log_path)

    if tr is not None:
        if tr.get("skipped"):
            print("  清洗            已跳过（%s）" % tr["skipped"])
        else:
            print("  清洗            %s" % ("成功" if tr["ok"] else "失败"))

    bad = [r for r in rows if r["status"] != "ok"]
    if bad:
        print("  %d/%d 家店铺需要关注" % (len(bad), len(rows)))
        return 1
    if tr is not None and not tr["ok"]:
        # The downloads succeeded; only the transform did not. Worth a non-zero
        # exit so a scheduled run does not look clean, but say which half broke.
        print("  所有店铺下载成功，但清洗失败")
        return 1
    print("  全部 %d 家店铺正常" % len(rows))
    return 0


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
