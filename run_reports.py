# -*- coding: utf-8 -*-
"""
生成月度财务报表：每店一份 + 全部店铺一份汇总。

    python run_reports.py                        # 上个完整月份，白名单里全部店铺
    python run_reports.py --month 2026-08
    python run_reports.py --stores BOCINA_TA02 UNIT_TA04
    python run_reports.py --no-per-store         # 只出汇总
    python run_reports.py --dry-run              # 只列出会用哪个下载目录

店铺清单来自 config.json 的白名单，和下载流程同一个来源，不在这里另写一份。
每家店用它**最近一个文件齐全**的下载目录 —— 不是无脑取最新，因为一次
`--only ventas` 的运行会留下缺账单的目录。
"""
import argparse
import console  # noqa: F401  中文输出编码保护
import datetime
import os
import sys

import store_config
from store_config import StoreConfigError, StoreNotAllowed
import financial_report

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(ROOT, "downloads", "data", "reports", "financial")


def prev_month(today=None):
    """上一个完整月份。当月还没过完，做月报没有意义。"""
    today = today or datetime.date.today()
    first = today.replace(day=1)
    last_prev = first - datetime.timedelta(days=1)
    return "%04d-%02d" % (last_prev.year, last_prev.month)


def resolve(names, month=None):
    """把店铺名解析成 (名字, 下载目录)，并报告哪些用不了。

    传 month 是为了让每家店优先用**确实含那个月账单**的目录 —— 做历史月份时
    最新目录往往只有最近两期账单。
    """
    usable, unusable = [], []
    for n in names:
        folder, note = financial_report.latest_run(n, month=month)
        if folder:
            usable.append((n, folder))
            print("  %-22s %s" % (n, note))
        else:
            unusable.append((n, note))
            print("  %-22s [跳过] %s" % (n, note))
    return usable, unusable


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", default=None,
                    help="会计期间，如 2026-08；默认为上一个完整月份")
    ap.add_argument("--stores", nargs="+", default=None,
                    help="店铺名；默认取 config.json 白名单里的全部")
    ap.add_argument("--out-dir", default=OUT_DIR, help="报表输出目录")
    ap.add_argument("--no-per-store", action="store_true", help="不出单店报表")
    ap.add_argument("--no-combined", action="store_true", help="不出汇总报表")
    ap.add_argument("--metrics", action="store_true",
                    help="同时导出月度指标 JSON（用于累积指标库）")
    ap.add_argument("--strict", action="store_true",
                    help="任一校验未通过就不写出文件")
    ap.add_argument("--dry-run", action="store_true",
                    help="只列出会用哪个下载目录，不生成报表")
    args = ap.parse_args()

    month = args.month or prev_month()
    names = ([store_config.require_allowed(n) for n in args.stores]
             if args.stores else store_config.batch_stores())
    if not names:
        raise StoreConfigError("config.json 的白名单为空，没有店铺可出报表。")

    print("=" * 62)
    print("月度财务报表 —— %s，%d 家店铺" % (month, len(names)))
    print("=" * 62)
    usable, unusable = resolve(names, month=month)

    if not usable:
        print("\n没有一家店有可用的下载目录，无法生成报表。")
        return 1
    if args.dry_run:
        print("\n--dry-run：不生成报表。")
        return 0

    if not financial_report.has_libreoffice():
        # 公式本身照常写进文件，只是没有缓存值；Excel 打开时会自己算。
        # 说出来，免得看到"跳过重算"以为报表不完整。
        print("\n[提示] 未检测到 LibreOffice，跳过公式重算。")
        print("       公式已写入文件，Excel/WPS 打开时会自动计算；")
        print("       但本次无法自动检查公式是否报错。")

    results = []

    # ---- 单店报表：一家出问题不影响其余 ----
    if not args.no_per_store:
        for name, folder in usable:
            out = os.path.join(args.out_dir, "%s_%s.xlsx" % (name, month))
            print("\n" + "-" * 62)
            print("单店报表：%s" % name)
            print("-" * 62)
            r = financial_report.build(
                [(name, folder)], month, out, strict=args.strict,
                metrics_out=(os.path.join(args.out_dir, "%s_%s_指标.json"
                                          % (name, month)) if args.metrics else None))
            r["label"] = name
            results.append(r)

    # ---- 汇总报表：12 家店传进同一次调用 ----
    if not args.no_combined:
        out = os.path.join(args.out_dir, "MX_ML_全店汇总_%s.xlsx" % month)
        print("\n" + "-" * 62)
        print("汇总报表：%d 家店铺" % len(usable))
        print("-" * 62)
        r = financial_report.build(
            usable, month, out, strict=args.strict,
            metrics_out=(os.path.join(args.out_dir, "全店_%s_指标.json" % month)
                         if args.metrics else None))
        r["label"] = "全店汇总"
        results.append(r)

    # ---- 汇总 ----
    print("\n" + "=" * 62)
    print("报表生成结果")
    print("=" * 62)
    print("  %-22s %-10s %8s  %s" % ("报表", "结果", "用时", "文件"))
    for r in results:
        if r.get("skipped"):
            state = "跳过"
        elif r.get("stale"):
            state = "写入失败"
        elif r["ok"]:
            state = "成功"
        elif r.get("checks_failed"):
            state = "有校验未过"      # 文件已写出，但有 error 级校验没通过
        else:
            state = "失败"
        print("  %-22s %-10s %6ds  %s"
              % (r["label"], state, r["seconds"],
                 os.path.basename(r["out"]) if r.get("wrote") else "—"))
    if unusable:
        print("\n  未参与（没有可用下载目录）：")
        for n, why in unusable:
            print("    %-22s %s" % (n, why))

    print("\n  输出目录：%s" % os.path.abspath(args.out_dir))
    bad = [r for r in results if not r["ok"]]
    if bad:
        print("  %d/%d 份报表需要关注" % (len(bad), len(results)))
        return 1
    print("  全部 %d 份报表生成成功" % len(results))
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
