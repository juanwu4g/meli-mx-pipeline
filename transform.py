# -*- coding: utf-8 -*-
"""
Hand the freshly downloaded files to the cleaning pipeline in `downloads/`.

`downloads/mx_sales` is a separate project (its own README, its own pyproject,
its own tests) that cleans the raw exports into Parquet and DuckDB views. This
module is the seam between the two: it runs that project's CLI and reports what
happened. It deliberately knows nothing about how the pipeline works.

Two decisions worth keeping:

* **A separate process, not an import.** A pandas or duckdb failure must not be
  able to reach back into a batch that has just spent half an hour driving a
  browser. The worst a broken transform can do here is print an error.

* **After ALL stores, not after each one.** The pipeline is tree-shaped - its
  own README says "a snapshot tree holds the same report many times over", and
  `discover` walks the whole raw directory. Running it once at the end sees
  every store from the batch; running it per store would do the same work N
  times over.
"""
import os
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.abspath(__file__))

# The pipeline's project root: `python -m mx_sales` is run from here, so the
# package resolves without needing to be pip-installed.
PIPELINE_ROOT = os.path.join(ROOT, "downloads")
PIPELINE_PKG = os.path.join(PIPELINE_ROOT, "mx_sales")


def available():
    """Is the cleaning pipeline present in this checkout?"""
    return os.path.isdir(PIPELINE_PKG)


def _echo(out):
    """Print the pipeline's output, collapsing repeated lines.

    Reading the exports emits the same openpyxl "Workbook contains no default
    style" warning once per file, which buries the three lines that actually say
    what was built. Collapsed rather than filtered: a warning that fires 22 times
    is still shown, once, with its count - suppressing it outright would also
    hide the next warning nobody has seen yet.
    """
    seen = {}
    order = []
    for line in out.rstrip().splitlines():
        line = line.rstrip()
        if not line:
            continue
        if line not in seen:
            seen[line] = 0
            order.append(line)
        seen[line] += 1
    for line in order:
        n = seen[line]
        print("    %s%s" % (line, "   (x%d)" % n if n > 1 else ""))


def _invoke(subcmd, label, timeout, echo):
    """Run one `python -m mx_sales <subcmd>` and describe what happened.

    Returns a dict: ok / skipped / returncode / seconds / output.
    Never raises - the downloads are already on disk and safe either way.
    """
    if not available():
        return {"ok": True, "skipped": "downloads/ 下没有 mx_sales 包",
                "returncode": None, "seconds": 0, "output": ""}

    # sys.executable, so the transform runs in the same venv as the downloader.
    # pyarrow and duckdb are pinned in requirements.txt for exactly this.
    cmd = [sys.executable, "-m", "mx_sales"] + list(subcmd)
    env = dict(os.environ)
    # The pipeline prints Spanish report names; without this a pipe on Windows
    # raises UnicodeEncodeError and the whole step fails for a print statement.
    env["PYTHONIOENCODING"] = "utf-8"

    print("\n>>> %s：%s" % (label, " ".join(cmd[1:])))
    print("    工作目录：%s" % PIPELINE_ROOT)
    began = time.time()
    try:
        p = subprocess.run(cmd, cwd=PIPELINE_ROOT, env=env, timeout=timeout,
                           stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                           encoding="utf-8", errors="replace")
        out, rc = p.stdout or "", p.returncode
    except subprocess.TimeoutExpired:
        return {"ok": False, "skipped": None, "returncode": None,
                "seconds": int(time.time() - began),
                "output": "超时（%d 秒）" % timeout}
    except Exception as e:
        return {"ok": False, "skipped": None, "returncode": None,
                "seconds": int(time.time() - began),
                "output": "%s: %s" % (e.__class__.__name__, e)}

    secs = int(time.time() - began)
    if echo and out.strip():
        _echo(out)
    print("    %s %s，用时 %d 分 %02d 秒"
          % (label, "成功" if rc == 0 else "失败（退出码 %d）" % rc,
             secs // 60, secs % 60))
    return {"ok": rc == 0, "skipped": None, "returncode": rc,
            "seconds": secs, "output": out[-4000:]}


def build(args=(), timeout=3600, echo=True):
    """Clean every download in the tree to Parquet and refresh the DuckDB views."""
    return _invoke(["build"] + list(args), "清洗", timeout, echo)


def validate(export, excel_out, timeout=1800, echo=True):
    """Apply the accounting rules to ONE ventas export and write its workbook.

    `export` must be the path RELATIVE TO downloads/, e.g.
    "UNIT_TA04/20260903_164329/20260903_Ventas_MX_...xlsx". The pipeline's
    matcher also accepts a bare filename, but that matches any file of that name
    anywhere in the tree - fine for ventas exports, whose names carry a timestamp
    and seller id, and quietly wrong for anything else. Passing the full relative
    path removes the question.

    Note this reads the RAW export, not the Parquet, and pairs it with the
    facturacion files in the SAME run folder - so the workbook is consistent to
    one snapshot and does not depend on build() having run.
    """
    excel_out = os.path.abspath(excel_out)
    # `validate --excel` writes where it is told and does not create the folder.
    parent = os.path.dirname(excel_out)
    if parent and not os.path.isdir(parent):
        os.makedirs(parent)
    return _invoke(["validate", export, "--excel", excel_out],
                   "对账簿", timeout, echo)
