# -*- coding: utf-8 -*-
"""
Write a human-readable record of one run into its own download folder.

Every run drops `REPORT.md` next to the files it downloaded, so the folder
explains itself months later: what was pulled, from which route, what was
skipped and why, and what is still owed.

Kept apart from run_downloads.py because it is presentation, not pipeline:
nothing here touches the browser, and a failure to write the report must never
cost the downloads it is describing.
"""
import datetime
import io
import os

REPORT_NAME = "REPORT.md"

# MercadoPago names its downloads after the report, so the group column can be
# filled from the filename. Longest prefix first: "settlement_v2-" would
# otherwise be missed by a shorter match, and account_statement_generic- shares
# no prefix with the rest only by luck.
MP_PREFIXES = [
    ("account_statement_generic-", "Estados de saldos"),
    ("after_collection-", "Poscobro"),
    ("reserve-release-", "Liberaciones"),
    ("settlement_v2-", "Todas las transacciones"),
    ("collection-", "Cobros"),
    ("withdraw-", "Retiros"),
]


def _mp_group(filename):
    low = filename.lower()
    for prefix, label in MP_PREFIXES:
        if low.startswith(prefix):
            return label
    return ""


def _size(path):
    try:
        n = os.path.getsize(path)
    except OSError:
        return "missing"
    if n < 1024:
        return "%d B" % n
    if n < 1024 * 1024:
        return "%.1f KB" % (n / 1024.0)
    return "%.1f MB" % (n / (1024.0 * 1024.0))


def _rows(result, out_dir):
    """(route, group, filename, size) for everything this run downloaded."""
    rows = []
    for period, files in sorted(result.get("billing", {}).items()):
        for f in files:
            f = os.path.basename(f)
            rows.append(("Facturación", period, f, _size(os.path.join(out_dir, f))))
    for f in result.get("sales", []):
        f = os.path.basename(f)
        rows.append(("Ventas", "Excel de ventas", f, _size(os.path.join(out_dir, f))))
    for key, files in sorted(result.get("stock", {}).items()):
        for f in files:
            f = os.path.basename(f)
            rows.append(("Stock", key, f, _size(os.path.join(out_dir, f))))
    for f in result.get("mp", []):
        f = os.path.basename(f)
        rows.append(("MercadoPago", _mp_group(f), f,
                     _size(os.path.join(out_dir, f))))
    return rows


def render(result, args=None, started=None):
    """The report body, as markdown."""
    out_dir = result["out_dir"]
    rows = _rows(result, out_dir)
    known = set(r[2] for r in rows)
    # Anything in the folder we did not attribute to a route. Usually nothing;
    # if it appears, it is a leftover from an interrupted run and the reader
    # should know rather than quietly counting it as this run's output.
    extra = sorted(f for f in os.listdir(out_dir)
                   if f.lower().endswith((".xlsx", ".xls", ".csv"))
                   and f not in known)

    started = started or datetime.datetime.now()
    secs = result.get("seconds", 0)

    L = []
    L.append("# Descargas — %s" % result["store"])
    L.append("")
    L.append("| | |")
    L.append("|---|---|")
    L.append("| Store | `%s` |" % result["store"])
    L.append("| Run started | %s |" % started.strftime("%Y-%m-%d %H:%M:%S"))
    L.append("| Duration | %d min %02d s |" % (secs // 60, secs % 60))
    L.append("| Files downloaded | **%d** |" % len(rows))
    L.append("| Folder | `%s` |" % out_dir)
    if args is not None:
        # Only the parameters that actually shaped this run. Printing the
        # billing period count on an --only ventas run invites the reader to
        # wonder where the billing files went.
        only = getattr(args, "only", "all")
        L.append("| Routes | %s |" % only)
        if only in ("billing", "all"):
            L.append("| Billing periods | %s |" % getattr(args, "months", "-"))
        if only in ("ventas", "all"):
            L.append("| Ventas range | últimos %s meses |"
                     % getattr(args, "period_months", "-"))
        if only in ("mercadopago", "all"):
            L.append("| MercadoPago window | %s days |"
                     % getattr(args, "mp_days", "-"))
    L.append("")

    if rows:
        L.append("## Files")
        L.append("")
        L.append("| Route | Group | File | Size |")
        L.append("|---|---|---|---|")
        for route, group, f, size in rows:
            L.append("| %s | %s | `%s` | %s |" % (route, group, f, size))
        L.append("")
    else:
        L.append("## Files")
        L.append("")
        L.append("Nothing was downloaded. See the sections below for why.")
        L.append("")

    empty = result.get("empty") or []
    if empty:
        L.append("## No data — not an error")
        L.append("")
        L.append("MercadoPago had no movements in the requested range, so it")
        L.append("would not build these. Re-requesting the same range will not")
        L.append("help; a different period might.")
        L.append("")
        for x in empty:
            L.append("- %s" % x)
        L.append("")

    pending = result.get("pending") or []
    if pending:
        L.append("## Still owed")
        L.append("")
        L.append("Requested but not collected in time. The request stands on the")
        L.append("provider's side — **a later run picks these up**; nothing was")
        L.append("lost. Entries marked `datos en proceso` need a few hours.")
        L.append("")
        for x in pending:
            L.append("- %s" % x)
        L.append("")

    errors = result.get("errors") or []
    if errors:
        L.append("## Errors")
        L.append("")
        L.append("These need someone to look at them.")
        L.append("")
        for x in errors:
            L.append("- %s" % x)
        L.append("")

    if extra:
        L.append("## Unattributed files in this folder")
        L.append("")
        L.append("Present on disk but not produced by this run — most likely a")
        L.append("leftover from an interrupted run.")
        L.append("")
        for f in extra:
            L.append("- `%s` (%s)" % (f, _size(os.path.join(out_dir, f))))
        L.append("")

    L.append("---")
    L.append("")
    L.append("Generated by `run_downloads.py`. Form-by-form detail: `FORMS.md`.")
    L.append("")
    return "\n".join(L)


def write(result, args=None, started=None):
    """Write REPORT.md into the run folder. Returns its path, or None.

    Never raises: a run that downloaded 18 files must not be reported as failed
    because a summary file could not be written.
    """
    try:
        path = os.path.join(result["out_dir"], REPORT_NAME)
        # utf-8: the content is Spanish, and Windows' default codepage cannot
        # encode "Facturación" (GUIDE.md 6.4).
        with io.open(path, "w", encoding="utf-8") as f:
            f.write(render(result, args=args, started=started))
        return path
    except Exception as e:
        print("  [warn] could not write %s: %s" % (REPORT_NAME, e))
        return None
