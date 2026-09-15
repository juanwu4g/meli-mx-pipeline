"""Command line entry point: ``python -m mx_sales <command>``.

Commands
--------
``discover``  list raw files and the report type each one matched
``build``     clean the raw downloads to Parquet and refresh the DuckDB views
``query``     run SQL against the warehouse
``validate``  apply the Ventas MX validation rules to one export

Exit codes
----------
``0``  everything processed
``1``  some files failed; the rest were written and the warehouse refreshed
``2``  nothing usable -- every file failed, or none matched a report type

1 and 2 are kept apart on purpose. ``pipeline.run`` already isolates a bad
export ("one bad export must not sink the run"), but this used to collapse any
error into a single non-zero code, so a caller had no way to tell 1 bad file out
of 55 from a run that produced nothing -- and reported both as "cleaning
failed".
"""

from __future__ import annotations

import argparse
import logging
import sys
import traceback
from pathlib import Path

import pandas as pd

from mx_sales import pipeline, warehouse
from mx_sales.accounting import runner
from mx_sales.accounting import ventas_mx as rules
from mx_sales.config import SETTINGS
from mx_sales.registry import SPECS


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="mx-sales", description=__doc__)
    parser.add_argument("-v", "--verbose", action="store_true", help="log per-file progress")
    subparsers = parser.add_subparsers(dest="command", required=True)

    discover_parser = subparsers.add_parser("discover", help="list recognized raw downloads")
    discover_parser.add_argument("--unmatched", action="store_true", help="show unrecognized files too")

    build_parser = subparsers.add_parser("build", help="clean raw downloads to Parquet")
    build_parser.add_argument("--report", help="limit to one report, e.g. ventas_mx")
    build_parser.add_argument("--file", help="process a single download (name or relative path)")
    build_parser.add_argument("--no-warehouse", action="store_true", help="skip DuckDB view refresh")

    query_parser = subparsers.add_parser("query", help="run SQL against the warehouse")
    query_parser.add_argument("sql", help="SQL to execute, e.g. 'SELECT * FROM ventas_mx LIMIT 5'")

    validate_parser = subparsers.add_parser("validate", help="apply the Ventas MX rules")
    validate_parser.add_argument("file", help="the sales export to validate (name or path)")
    validate_parser.add_argument("--excel", type=Path, help="also write an .xlsx workbook here")

    args = parser.parse_args(argv)
    logging.basicConfig(
        level=logging.INFO if args.verbose else logging.WARNING,
        format="%(levelname)s %(message)s",
    )
    pd.set_option("display.width", 200)

    if args.command == "discover":
        return _discover(args.unmatched)
    if args.command == "build":
        return _build(args.report, args.file, warehouse_refresh=not args.no_warehouse)
    if args.command == "query":
        print(warehouse.query(args.sql).to_string(index=False))
        return 0
    if args.command == "validate":
        return _validate(args.file, args.excel)
    return 2          # unreachable (subparsers are required); never borrow 1


def _discover(show_unmatched: bool) -> int:
    items = pipeline.discover()
    print(f"raw dir: {SETTINGS.raw_dir}")
    print(f"{len(items)} file(s) matched a registered report:\n")
    for item in items:
        print(f"  {item.spec.name:<14} {item.path.relative_to(SETTINGS.raw_dir)}")

    if show_unmatched:
        matched = {item.path for item in items}
        unmatched = [
            p
            for p in sorted(SETTINGS.raw_dir.rglob("*"))
            if p.is_file()
            and p.suffix.lower() in pipeline.RAW_SUFFIXES
            and p not in matched
            and not (set(p.relative_to(SETTINGS.raw_dir).parts) & {"data", "mx_sales"})
        ]
        print(f"\n{len(unmatched)} file(s) with no registered report:\n")
        for path in unmatched:
            print(f"  {'-':<14} {path.relative_to(SETTINGS.raw_dir)}")

    print(f"\nregistered reports: {', '.join(spec.name for spec in SPECS)}")
    return 0


def _build(report: str | None, only: str | None, *, warehouse_refresh: bool) -> int:
    results = pipeline.run(report, only=only)
    if not results:
        print("nothing to do: no raw files matched a registered report")
        return 2

    # `files` counts only the files that made it through -- pipeline.run skips a
    # failed one before incrementing -- so errors and files never double count.
    written = sum(result.files for result in results)
    failed = sum(len(result.errors) for result in results)

    for result in results:
        note = f", {len(result.errors)} FAILED" if result.errors else ""
        print(f"{result.report}: {result.rows} rows from {result.files} file(s)"
              f"{note} -> {result.output}")
        for error in result.errors:
            print(f"  ERROR {error}")

    if warehouse_refresh:
        views = warehouse.refresh_views()
        print(f"warehouse: {SETTINGS.warehouse_path} (views: {', '.join(views) or 'none'})")

    if not failed:
        return 0
    total = written + failed
    if written:
        print(f"{failed} of {total} file(s) failed to clean; the other {written} "
              f"were written and the warehouse was refreshed")
        return 1
    print(f"all {total} file(s) failed to clean; nothing was written")
    return 2


def _validate(file: str, excel: Path | None) -> int:
    source, frame, charges = runner.prepare(file)
    store = runner.store_of(source)
    output = runner.write(frame, store=store)
    failed = rules.exceptions(frame)

    print(f"{store or '(store unknown)'}: {len(frame)} rows -> {output}\n")
    print("control -- money in the sheet vs money in the output:")
    print(rules.control_totals(source, frame).to_string(index=False))
    print()
    print(rules.summarize(frame).to_string(index=False))
    print()
    print(rules.monthly_summary(frame, charges).to_string(index=False))

    if not failed.empty:
        columns = [
            "row_index", "order_id", "sold_at", "order_status",
            "components_sum_mxn", "total_mxn", "total_diff_mxn",
        ]
        print(f"\nrule 1 differences ({len(failed)}):")
        print(failed[columns].to_string(index=False))

    if excel is not None:
        print(f"\nworkbook: {runner.to_excel(frame, excel, source, charges, export=file, store=store)}")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except SystemExit:
        raise
    except BaseException:
        # 1 now means exactly one thing -- "some files failed, the rest were
        # written" -- and callers key off that. An unhandled crash (a bad
        # --file, an unreadable warehouse) must not borrow that code, or a total
        # failure gets reported as a partial one. Traceback still printed: this
        # widens the exit code, it does not hide the error.
        traceback.print_exc()
        sys.exit(2)
