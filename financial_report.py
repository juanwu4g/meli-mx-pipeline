# -*- coding: utf-8 -*-
"""
调用 report_build.py 生成月度财务报表。

`report_build.py` 是一个独立的报表生成器（不 import 项目里任何模块），本文件是
它与下载流程之间的接缝：负责挑出每家店该用哪个下载目录、拼出命令、跑起来、把结
果讲清楚。它刻意不了解报表内部怎么算。

沿用 transform.py 的两条约定：

* **子进程，不是 import。** pandas / openpyxl 抛异常不能反过来打断调用方；报表
  里的公式重算本来也要外调 LibreOffice。最坏情况只是打印一段错误。
* **绝不抛异常。** 报表没生成不该让已经下载好的文件跟着失败。
"""
import io
import glob
import os
import re
import subprocess
import sys
import time

ROOT = os.path.dirname(os.path.abspath(__file__))
SCRIPT = os.path.join(ROOT, "report_build.py")
DOWNLOADS = os.path.join(ROOT, "downloads")

# 报表必需的输入。缺 Ventas 或账单就算不出损益表，这样的目录不能用 —— 光取"最新
# 目录"是不够的：一次只跑了 --only ventas 的运行会留下一个看着很新、其实缺账单的
# 目录（BOCINA_SM 的 20260903_110457 就是）。
REQUIRED = (
    ("*Ventas_MX*.xlsx", "Ventas"),
    ("Reporte_Facturacion_MercadoLibre_*.xlsx", "账单"),
)


# 账单文件名里的月份缩写。与 report_build.FILE_MONTH_TOKENS 同源；这里单独留一份，
# 是因为 financial_report 刻意不 import report_build（两者靠子进程相隔，见文件头）。
FILE_MONTH_TOKENS = {
    "ene": 1, "enero": 1, "feb": 2, "mar": 3, "abr": 4, "may": 5, "jun": 6,
    "jul": 7, "ago": 8, "sep": 9, "sept": 9, "oct": 10, "nov": 11, "dic": 12,
}


def billing_months(folder):
    """这个目录里有哪几个月的账单。返回 {(年, 月)}。

    账单文件名自带月份（Reporte_Facturacion_MercadoLibre_Ago2026.xlsx），
    不用打开文件就能知道覆盖了哪几期。
    """
    out = set()
    for f in glob.glob(os.path.join(folder, "Reporte_*.xlsx")):
        m = re.search(r"_([A-Za-z]{3,5})(\d{4})\.xlsx$", os.path.basename(f))
        if m and m.group(1).lower() in FILE_MONTH_TOKENS:
            out.add((int(m.group(2)), FILE_MONTH_TOKENS[m.group(1).lower()]))
    return out


def available():
    return os.path.isfile(SCRIPT)


def has_libreoffice():
    """报表里的公式需要 LibreOffice 重算出缓存值，没有就只能跳过。"""
    if os.environ.get("SOFFICE"):
        return True
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe",
              r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"):
        if os.path.isfile(p):
            return True
    for d in os.environ.get("PATH", "").split(os.pathsep):
        for exe in ("soffice.exe", "soffice"):
            if d and os.path.isfile(os.path.join(d, exe)):
                return True
    return False


def _usable(folder):
    """这个下载目录是否够算一份报表。"""
    import glob
    missing = [label for pat, label in REQUIRED
               if not glob.glob(os.path.join(folder, pat))]
    return (not missing), missing


def latest_run(store, base=None, month=None):
    """这家店最近一个**能用**的下载目录。

    返回 (目录, 说明)。没有可用目录时返回 (None, 原因)。
    按时间戳倒序找第一个文件齐全的，而不是无脑取最新 —— 见 REQUIRED 的注释。

    给了 month（"YYYY-MM"）时，**优先**挑账单里确实含那个月的目录。做历史月份
    报表时这一步是必须的：最新那个目录只有最近两期账单，拿它做 7 月报表，
    佣金全是 0，代扣税会被算成佣金的好几倍（实测 28%，正常 9%）。
    找不到含该月账单的目录时退回原来的规则，并在说明里讲清楚。
    """
    want = None
    if month:
        try:
            want = (int(month[:4]), int(month[5:7]))
        except Exception:
            want = None
    base = base or DOWNLOADS
    d = os.path.join(base, store)
    if not os.path.isdir(d):
        return None, "没有下载目录"
    runs = sorted((x for x in os.listdir(d)
                   if os.path.isdir(os.path.join(d, x)) and x[:2] == "20"),
                  reverse=True)
    if not runs:
        return None, "目录下没有任何运行记录"
    usable = []
    skipped = []
    for r in runs:
        folder = os.path.join(d, r)
        ok, missing = _usable(folder)
        if ok:
            usable.append((r, folder))
        else:
            skipped.append(r)

    if want:
        for r, folder in usable:
            if want in billing_months(folder):
                return folder, "%s（含 %04d-%02d 账单）" % ((r,) + want)

    for r, folder in usable:
        note = r if not skipped else "%s（跳过 %d 个不完整的更新目录）" % (r, len(skipped))
        if want:
            note += "  ⚠ 没有任何目录含 %04d-%02d 的账单，该月费用会缺失" % want
        return folder, note
    return None, "%d 个目录都缺文件（最新的缺：%s）" % (
        len(runs), "、".join(_usable(os.path.join(d, runs[0]))[1]))


def build(pairs, month, out, metrics_out=None, strict=False,
          no_recalc=None, timeout=3600, echo=True):
    """生成一份报表。pairs 是 [(店铺名, 下载目录), ...]。

    一份报表可以覆盖多家店 —— 汇总报表就是把 12 家一起传进来。
    返回 dict：ok / returncode / seconds / out / output。绝不抛异常。
    """
    if not available():
        return {"ok": False, "skipped": "找不到 report_build.py",
                "returncode": None, "seconds": 0, "out": out, "output": ""}
    if not pairs:
        return {"ok": False, "skipped": "没有可用的店铺", "returncode": None,
                "seconds": 0, "out": out, "output": ""}

    if no_recalc is None:
        no_recalc = not has_libreoffice()

    d = os.path.dirname(os.path.abspath(out))
    if d and not os.path.isdir(d):
        os.makedirs(d)

    cmd = [sys.executable, SCRIPT, "--month", month, "--out", out]
    for name, folder in pairs:
        cmd += ["--store", "%s=%s" % (name, folder)]
    if metrics_out:
        cmd += ["--metrics-out", metrics_out]
    if strict:
        cmd.append("--strict")
    if no_recalc:
        cmd.append("--no-recalc")

    env = dict(os.environ)
    # 报表全程中文输出，管道里必须是 UTF-8，否则一句 print 就会 UnicodeEncodeError
    env["PYTHONIOENCODING"] = "utf-8"

    # 记下调用前的修改时间：文件"存在"不等于"本次写成功了"。报表被 Excel 打开时
    # openpyxl 会 PermissionError，旧文件原封不动留在原地 —— 只判断存在就会把
    # 写失败报成成功。
    before_mtime = os.path.getmtime(out) if os.path.isfile(out) else None

    began = time.time()
    try:
        p = subprocess.run(cmd, cwd=ROOT, env=env, timeout=timeout,
                           stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                           encoding="utf-8", errors="replace")
        output, rc = p.stdout or "", p.returncode
    except subprocess.TimeoutExpired:
        return {"ok": False, "skipped": None, "returncode": None,
                "seconds": int(time.time() - began), "out": out,
                "output": "超时（%d 秒）" % timeout}
    except Exception as e:
        return {"ok": False, "skipped": None, "returncode": None,
                "seconds": int(time.time() - began), "out": out,
                "output": "%s: %s" % (e.__class__.__name__, e)}

    secs = int(time.time() - began)
    if echo and output.strip():
        for line in output.rstrip().splitlines():
            print("    %s" % line)

    # report_build 的退出码：0 = 全通过；1 = 有 error 级校验未通过（文件通常仍已
    # 写出，除非 --strict）。所以"文件在不在"和"校验过没过"是两件事，分别报。
    exists = os.path.isfile(out)
    wrote = exists and (before_mtime is None
                        or os.path.getmtime(out) > before_mtime)
    stale = exists and not wrote          # 旧文件还在，但本次没写进去

    if stale:
        print("    [错误] 未能写入 %s" % os.path.basename(out))
        print("           磁盘上是上一次的旧文件，不是本次结果。")
        if "PermissionError" in output:
            print("           原因：文件正被 Excel/WPS 打开，请关闭后重跑。")

    return {"ok": rc == 0 and wrote, "skipped": None, "returncode": rc,
            "seconds": secs, "out": out, "wrote": wrote, "stale": stale,
            "checks_failed": rc == 1 and wrote, "output": output[-6000:]}
