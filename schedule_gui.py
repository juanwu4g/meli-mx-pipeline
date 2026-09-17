# -*- coding: utf-8 -*-
"""定时任务的可视化入口：不用敲一条命令就能把每月自动跑配好。

    双击 定时设置.cmd    （推荐，它会用对解释器）
    或   .venv\\Scripts\\python schedule_gui.py

做四件事：建/改每月的计划任务、彩排一次、看状态、检查这台机器够不够格
无人值守。**业务逻辑一行都没有** —— 真正跑的是 run_monthly.cmd，这里只是
把 Windows 任务计划程序那一堆勾选项替人点好。

为什么用 XML 而不是 schtasks 的命令行参数
-----------------------------------------
有三项设置命令行根本没有开关，只能在任务计划程序界面里手点：

* 「如果错过计划的开始时间，立即启动任务」（StartWhenAvailable）—— 15 号
  机器关着就补跑，对一台没人管的老 PC 来说这条最要紧
* 「如果任务运行超过 N 小时则停止」（ExecutionTimeLimit）—— 脚本自己不做
  超时，全靠它兜底
* 电池相关的两项 —— 笔记本上默认会拦住任务

让不写代码的人去 taskschd.msc 里找这三个勾，比让他们敲命令还难。所以这里
直接生成完整的任务 XML 再 `schtasks /create /xml` 导入，一次到位。

为什么查询走 PowerShell 而不是 schtasks
---------------------------------------
`schtasks /query` 的输出是本地化的，中文机器上列名和值都是中文，解析它等于
赌系统语言。`Get-ScheduledTaskInfo` 返回的是对象，属性名与语言无关。
"""
import datetime
import io
import os
import queue
import re
import subprocess
import sys
import tempfile
import threading

import tkinter as tk
from tkinter import ttk, messagebox

from gui import ROOT, interpreter, store_list, load_settings, save_settings

TASK = "MX月度报表"
REHEARSAL = TASK + "-彩排"
RUNNER = os.path.join(ROOT, "run_monthly.cmd")
LOG_DIR = os.path.join(ROOT, "logs")

# 任务超时。一次 12 家店的完整跑批实测 2h48m–4h00m（logs/batch_*.json），
# 留到 5 小时：既能兜住卡死，又不会在正常的慢批次上误杀。
TIME_LIMIT = "PT5H"

MONTHS_XML = "".join("<%s/>" % m for m in (
    "January February March April May June July August September October "
    "November December").split())


# ────────────────────────────────────────────────────────── 系统调用

def _run(argv, timeout=120):
    """跑一个命令，返回 (退出码, 输出)。永不抛异常。"""
    try:
        p = subprocess.run(argv, capture_output=True, timeout=timeout,
                           encoding="utf-8", errors="replace")
        return p.returncode, (p.stdout or "") + (p.stderr or "")
    except Exception as e:
        return -1, "%s: %s" % (e.__class__.__name__, e)


def ps(script, timeout=120):
    """跑一段 PowerShell。

    开头强制 UTF-8：任务名里有中文，不设的话输出按 cp936 回来，Python 这边
    按 utf-8 解会变成乱码。
    """
    return _run(["powershell", "-NoProfile", "-NonInteractive",
                 "-ExecutionPolicy", "Bypass", "-Command",
                 "$OutputEncoding=[Console]::OutputEncoding="
                 "[Text.Encoding]::UTF8;" + script], timeout)


def kv(out):
    """把 `KEY=值` 形式的输出解析成字典。"""
    d = {}
    for line in (out or "").splitlines():
        if "=" in line:
            k, v = line.split("=", 1)
            d[k.strip()] = v.strip()
    return d


# ────────────────────────────────────────────────────────── 任务读写

def task_state(name):
    """任务现状。不存在返回 None。"""
    rc, out = ps(
        "$t = Get-ScheduledTask -TaskName '%s' -ErrorAction SilentlyContinue;"
        "if (-not $t) { 'EXISTS=0' } else {"
        "  'EXISTS=1';"
        "  'STATE=' + $t.State;"
        "  $i = Get-ScheduledTaskInfo -TaskName '%s' -TaskPath $t.TaskPath;"
        "  'NEXT=' + $i.NextRunTime;"
        "  'LAST=' + $i.LastRunTime;"
        "  'RESULT=' + $i.LastTaskResult;"
        "  'XML=' + ((Export-ScheduledTask -TaskName '%s') -replace '\\s+',' ')"
        "}" % (name, name, name))
    d = kv(out)
    if rc != 0 or d.get("EXISTS") != "1":
        return None
    xml = d.get("XML", "")
    day = re.search(r"<Day>(\d+)</Day>", xml)
    at = re.search(r"<StartBoundary>[^<]*T(\d\d:\d\d)", xml)
    d["DAY"] = day.group(1) if day else "?"
    d["TIME"] = at.group(1) if at else "?"
    return d


def task_xml(args=None, day=None, hour=0, minute=0):
    """生成任务 XML。

    day 为 None 时不带触发器 —— 彩排任务只靠手工触发，不该自己跑起来。
    """
    user = os.environ.get("USERNAME", "")
    domain = os.environ.get("USERDOMAIN", "")
    who = ("%s\\%s" % (domain, user)) if domain else user
    trigger = ""
    if day is not None:
        # StartBoundary 的日期部分只决定"从哪天起生效"，具体哪天跑由
        # DaysOfMonth 决定。取今天，避免写一个已经过去很久的日期。
        today = datetime.date.today()
        trigger = (
            "<CalendarTrigger>"
            "<StartBoundary>%04d-%02d-%02dT%02d:%02d:00</StartBoundary>"
            "<Enabled>true</Enabled>"
            "<ScheduleByMonth><DaysOfMonth><Day>%d</Day></DaysOfMonth>"
            "<Months>%s</Months></ScheduleByMonth>"
            "</CalendarTrigger>"
            % (today.year, today.month, today.day, hour, minute, day, MONTHS_XML))
    return (
        '<?xml version="1.0" encoding="UTF-16"?>'
        '<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">'
        "<RegistrationInfo><Description>"
        "每月自动下载上个月的数据并生成财务报表。由 schedule_gui.py 创建。"
        "</Description></RegistrationInfo>"
        "<Triggers>%s</Triggers>"
        "<Principals><Principal id=\"Author\">"
        "<UserId>%s</UserId>"
        # InteractiveToken = 「只在用户登录时运行」。紫鸟是带界面的 Electron
        # 程序，换成 Password/S4U 会被丢进 session 0，浏览器根本起不来。
        "<LogonType>InteractiveToken</LogonType>"
        "<RunLevel>LeastPrivilege</RunLevel>"
        "</Principal></Principals>"
        "<Settings>"
        # 上一轮还没跑完就到了下一个触发点：忽略新的。run_monthly 自己也有
        # 锁文件，这里是第二道。
        "<MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>"
        "<DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>"
        "<StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>"
        "<AllowHardTerminate>true</AllowHardTerminate>"
        "<StartWhenAvailable>true</StartWhenAvailable>"
        "<RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>"
        "<IdleSettings><StopOnIdleEnd>false</StopOnIdleEnd>"
        "<RestartOnIdle>false</RestartOnIdle></IdleSettings>"
        "<AllowStartOnDemand>true</AllowStartOnDemand>"
        "<Enabled>true</Enabled><Hidden>false</Hidden>"
        "<RunOnlyIfIdle>false</RunOnlyIfIdle>"
        "<WakeToRun>true</WakeToRun>"
        "<ExecutionTimeLimit>%s</ExecutionTimeLimit>"
        "<Priority>7</Priority>"
        "</Settings>"
        "<Actions Context=\"Author\"><Exec>"
        "<Command>%s</Command>%s"
        "<WorkingDirectory>%s</WorkingDirectory>"
        "</Exec></Actions></Task>"
        % (trigger, who, TIME_LIMIT, RUNNER,
           ("<Arguments>%s</Arguments>" % args) if args else "", ROOT))


def install(name, xml):
    """导入任务 XML。

    必须写成 UTF-16：schtasks /xml 读 UTF-8 文件会报 ERROR: 无效的 XML。
    """
    fd, path = tempfile.mkstemp(suffix=".xml")
    os.close(fd)
    try:
        with io.open(path, "w", encoding="utf-16") as f:
            f.write(xml)
        return _run(["schtasks", "/create", "/tn", name, "/xml", path, "/f"])
    finally:
        try:
            os.remove(path)
        except OSError:
            pass


# ────────────────────────────────────────────────────────── 机器自检

def machine_checks():
    """返回 [(通过?, 标题, 说明)]。全部只读，不改系统。"""
    out = []

    venv = os.path.join(ROOT, ".venv", "Scripts", "python.exe")
    out.append((os.path.exists(venv), "虚拟环境 .venv",
                "有" if os.path.exists(venv) else
                "缺。先跑：python -m venv .venv 再 pip install -r requirements.txt"))

    out.append((os.path.exists(RUNNER), "run_monthly.cmd",
                "有" if os.path.exists(RUNNER) else "缺，git pull 一下"))

    names = store_list()
    out.append((bool(names), "店铺白名单",
                "%d 家：%s" % (len(names), "、".join(names[:3]) + ("…" if len(names) > 3 else ""))
                if names else "读不到 config.json，或白名单为空"))

    # 睡眠：powercfg 的标签是本地化的，但这一段永远是 5 行带 0x 的值，
    # 顺序固定为 最小/最大/步进/交流/电池，所以取倒数第二行（交流）。
    rc, txt = _run(["powercfg", "/q", "SCHEME_CURRENT", "SUB_SLEEP", "STANDBYIDLE"])
    hexes = re.findall(r"0x[0-9a-fA-F]{8}", txt or "")
    if rc == 0 and len(hexes) >= 2:
        ac = int(hexes[-2], 16)
        out.append((ac == 0, "交流电源下不睡眠",
                    "已设为永不睡眠" if ac == 0 else
                    "%d 分钟后睡眠 —— 睡着了任务跑不了" % (ac // 60)))
    else:
        out.append((None, "交流电源下不睡眠", "查不到，手动确认"))

    rc, txt = ps("'V=' + (Get-ItemProperty "
                 "'HKLM:\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\Winlogon' "
                 "-Name AutoAdminLogon -ErrorAction SilentlyContinue).AutoAdminLogon")
    auto = kv(txt).get("V", "")
    out.append((auto == "1", "自动登录",
                "已开启" if auto == "1" else
                "没开。断电重启后没人登录，任务就不会跑（netplwiz 里设）"))

    rc, txt = ps("$d='HKCU:\\Control Panel\\Desktop';"
                 "'S=' + (Get-ItemProperty $d -Name ScreenSaverIsSecure "
                 "-ErrorAction SilentlyContinue).ScreenSaverIsSecure")
    sec = kv(txt).get("S", "")
    out.append((sec != "1", "屏保不锁屏",
                "没开锁屏屏保" if sec != "1" else
                "屏保会锁屏。锁屏后浏览器渲染可能被挂起，建议关掉"))
    return out


def fix_power():
    """把三项电源超时改成永不。powercfg /change 要管理员，所以走 UAC。"""
    return ps(
        "Start-Process powershell -Verb RunAs -Wait -ArgumentList "
        "'-NoProfile','-Command',"
        "'powercfg /change standby-timeout-ac 0; "
        "powercfg /change hibernate-timeout-ac 0; "
        "powercfg /change monitor-timeout-ac 0'", timeout=180)


# ────────────────────────────────────────────────────────── 界面

class SchedApp(object):
    def __init__(self, root):
        self.root = root
        self.q = queue.Queue()
        self.busy = False
        root.title("定时任务设置 —— 每月自动出报表")
        root.geometry("860x660")
        root.minsize(760, 560)

        s = load_settings()
        self.day = tk.IntVar(value=int(s.get("sched_day", 15)))
        self.hour = tk.IntVar(value=int(s.get("sched_hour", 2)))
        self.minute = tk.IntVar(value=int(s.get("sched_minute", 0)))

        self._build_status()
        self._build_form()
        self._build_actions()
        self._build_log()
        self.root.after(80, self._drain)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.refresh()

    # ---------------------------------------------------------- 布局

    def _build_status(self):
        f = ttk.LabelFrame(self.root, text=" 当前状态 ", padding=10)
        f.pack(fill="x", padx=12, pady=(10, 6))
        self.status = tk.StringVar(value="查询中…")
        self.status_lbl = ttk.Label(f, textvariable=self.status,
                                    font=("Microsoft YaHei UI", 11, "bold"))
        self.status_lbl.pack(anchor="w")
        self.detail = tk.StringVar(value="")
        ttk.Label(f, textvariable=self.detail, foreground="#555",
                  justify="left").pack(anchor="w", pady=(4, 0))
        ttk.Button(f, text="刷新", command=self.refresh, width=10).pack(
            anchor="w", pady=(8, 0))

    def _build_form(self):
        f = ttk.LabelFrame(self.root, text=" 什么时候跑 ", padding=10)
        f.pack(fill="x", padx=12, pady=6)
        row = ttk.Frame(f)
        row.pack(anchor="w")
        ttk.Label(row, text="每月").pack(side="left")
        ttk.Spinbox(row, from_=1, to=28, width=4, textvariable=self.day).pack(
            side="left", padx=4)
        ttk.Label(row, text="号").pack(side="left")
        ttk.Spinbox(row, from_=0, to=23, width=4, format="%02.0f",
                    textvariable=self.hour).pack(side="left", padx=(16, 4))
        ttk.Label(row, text="点").pack(side="left")
        ttk.Spinbox(row, from_=0, to=59, width=4, increment=5, format="%02.0f",
                    textvariable=self.minute).pack(side="left", padx=(8, 4))
        ttk.Label(row, text="分").pack(side="left")
        ttk.Label(f, foreground="#555", justify="left", text=(
            "跑的是【上个月】的报表：15 号跑，出的是上月的数 —— 那时候上月的账单\n"
            "已经闭账，数字基本稳定。最多只能选到 28 号，29～31 号在二月不会触发。\n"
            "凌晨跑最好：跑批会重启紫鸟客户端，那个点不会关掉任何人开着的窗口。")
                  ).pack(anchor="w", pady=(8, 0))

    def _build_actions(self):
        f = ttk.LabelFrame(self.root, text=" 操作 ", padding=10)
        f.pack(fill="x", padx=12, pady=6)
        rows = [
            ("创建 / 更新任务", self.do_install,
             "按上面的时间写进 Windows 任务计划程序。已存在就覆盖。"),
            ("彩排（约 1 分钟）", self.do_rehearse,
             "建一个临时任务只跑「出报表」那一步，跑完自动删。\n"
             "一分钟就能验完整条链路：能不能被计划任务拉起、中文会不会乱码、日志落没落盘。"),
            ("立即完整跑一次", self.do_run_now,
             "现在就按正式任务跑一遍，下载 + 出报表，约 3 小时。"),
            ("检查这台机器", self.do_check,
             "只读检查：睡眠、自动登录、锁屏、虚拟环境、店铺白名单。"),
            ("删除任务", self.do_delete, "把计划任务撤掉，脚本和数据都不动。"),
            ("打开日志目录", self.do_logs, "看 logs\\monthly_*.log。"),
        ]
        self.buttons = []
        for i, (label, cmd, desc) in enumerate(rows):
            b = ttk.Button(f, text=label, command=cmd, width=18)
            b.grid(row=i, column=0, sticky="w", pady=3)
            ttk.Label(f, text=desc, foreground="#555", justify="left").grid(
                row=i, column=1, sticky="w", padx=12)
            self.buttons.append(b)

    def _build_log(self):
        f = ttk.Frame(self.root)
        f.pack(fill="both", expand=True, padx=12, pady=(6, 12))
        self.txt = tk.Text(f, wrap="word", height=10,
                           font=("Consolas", 9), background="#fbfbfb")
        self.txt.pack(side="left", fill="both", expand=True)
        sb = ttk.Scrollbar(f, orient="vertical", command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set, state="disabled")
        sb.pack(side="right", fill="y")
        for tag, color in (("good", "#1a7f37"), ("warn", "#9a6700"),
                           ("bad", "#b42318"), ("dim", "#666666")):
            self.txt.tag_configure(tag, foreground=color)

    # ---------------------------------------------------------- 工具

    def log(self, line, tag=None):
        self.txt.configure(state="normal")
        self.txt.insert("end", line + "\n", tag or ())
        self.txt.see("end")
        self.txt.configure(state="disabled")

    def _lock(self, busy):
        self.busy = busy
        for b in self.buttons:
            b.configure(state="disabled" if busy else "normal")

    def _work(self, fn, done):
        """把慢活丢到子线程。Tk 控件只在主线程碰。"""
        if self.busy:
            messagebox.showwarning("正在忙", "等当前操作结束再点。")
            return
        self._lock(True)

        def body():
            try:
                self.q.put((done, fn()))
            except Exception as e:
                self.q.put((done, ("err", "%s: %s" % (e.__class__.__name__, e))))
        threading.Thread(target=body, daemon=True).start()

    def _drain(self):
        try:
            while True:
                fn, payload = self.q.get_nowait()
                fn(payload)
                self._lock(False)
        except queue.Empty:
            pass
        self.root.after(80, self._drain)

    # ---------------------------------------------------------- 动作

    def refresh(self):
        self._work(lambda: task_state(TASK), self._show_state)

    def _show_state(self, st):
        if isinstance(st, tuple):                    # 出错了
            self.status.set("查询失败")
            self.status_lbl.configure(foreground="#b42318")
            self.log(st[1], "bad")
            return
        if not st:
            self.status.set("● 还没有设置定时任务")
            self.status_lbl.configure(foreground="#9a6700")
            self.detail.set("按下面的时间点「创建 / 更新任务」即可。")
            return
        code = st.get("RESULT", "")
        verdict = {"0": ("上次跑完，全部正常", "#1a7f37"),
                   "1": ("上次跑完，有需要关注的项 —— 看日志", "#9a6700"),
                   "267009": ("正在运行", "#9a6700"),
                   "267011": ("还没跑过", "#666666")}.get(
            code, ("上次退出码 %s —— 看日志" % code, "#b42318"))
        self.status.set("● 已设置：每月 %s 号 %s" % (st.get("DAY"), st.get("TIME")))
        self.status_lbl.configure(foreground="#1a7f37")
        self.detail.set("下次运行：%s\n上次运行：%s\n%s"
                        % (st.get("NEXT") or "—", st.get("LAST") or "—",
                           verdict[0]))
        # 界面上的时间跟已登记的对齐，免得改了没保存的人看错
        try:
            self.day.set(int(st.get("DAY")))
            h, m = st.get("TIME", "02:00").split(":")
            self.hour.set(int(h))
            self.minute.set(int(m))
        except (ValueError, TypeError):
            pass

    def do_install(self):
        if not os.path.exists(RUNNER):
            messagebox.showerror("缺文件", "找不到 run_monthly.cmd：\n%s" % RUNNER)
            return
        d, h, m = self.day.get(), self.hour.get(), self.minute.get()
        if not messagebox.askokcancel(
                "创建定时任务",
                "每月 %d 号 %02d:%02d 自动跑一次：\n"
                "  ① 下载上个月的数据（全部店铺）\n"
                "  ② 生成上个月的报表\n\n"
                "任务会设成【只在用户登录时运行】—— 紫鸟是带界面的程序，\n"
                "这条不能改。所以这台机器要保持登录、不睡眠、不锁屏。\n\n"
                "继续？" % (d, h, m)):
            return
        save_settings(dict(load_settings(), sched_day=d, sched_hour=h,
                           sched_minute=m))
        self.log("")
        self.log("正在写入计划任务：每月 %d 号 %02d:%02d" % (d, h, m))
        self._work(lambda: install(TASK, task_xml(day=d, hour=h, minute=m)),
                   self._after_install)

    def _after_install(self, r):
        rc, out = r
        if rc == 0:
            self.log("✓ 已写入。错过时补跑、超 5 小时自动停止、电池限制已关，"
                     "都一并设好了。", "good")
            self.log("建议现在点一次「彩排」，一分钟验完整条链路。", "dim")
            self.refresh()
        else:
            self.log("✗ 写入失败（退出码 %s）" % rc, "bad")
            self.log(out.strip() or "(没有输出)", "bad")

    def do_rehearse(self):
        if not messagebox.askokcancel(
                "彩排",
                "会建一个临时任务，只跑「出报表」那一步（不下载、不开浏览器），\n"
                "跑完自动删掉。约 1 分钟。\n\n"
                "这一步验的是：任务能不能被计划程序拉起、工作目录对不对、\n"
                "中文会不会乱码、日志有没有落盘。\n\n开始？"):
            return
        self.log("")
        self.log("=" * 64)
        self.log(">>> 彩排  %s" % datetime.datetime.now().strftime("%H:%M:%S"))
        self._work(self._rehearse_body, self._after_rehearse)

    def _rehearse_body(self):
        rc, out = install(REHEARSAL, task_xml(args="--step report"))
        if rc != 0:
            return ("err", "建临时任务失败：%s" % (out.strip() or rc))
        rc, out = _run(["schtasks", "/run", "/tn", REHEARSAL])
        if rc != 0:
            _run(["schtasks", "/delete", "/tn", REHEARSAL, "/f"])
            return ("err", "启动临时任务失败：%s" % (out.strip() or rc))
        # 轮询到不再是 Running 为止。出 13 份报表实测 60～70 秒。
        for _ in range(90):
            st = task_state(REHEARSAL) or {}
            if st.get("STATE") != "Running":
                break
            threading.Event().wait(2)
        st = task_state(REHEARSAL) or {}
        _run(["schtasks", "/delete", "/tn", REHEARSAL, "/f"])
        return ("ok", st.get("RESULT", "?"))

    def _after_rehearse(self, r):
        kind, val = r
        if kind == "err":
            self.log("✗ %s" % val, "bad")
            return
        meaning = {"0": ("✓ 彩排通过：任务被正常拉起，报表也出来了", "good"),
                   "1": ("✓ 彩排通过：链路是通的。报表有校验未过（每月常态），"
                         "看 logs 里的日志", "good"),
                   "267009": ("⚠ 还在跑，等会儿点「刷新」", "warn")}
        msg, tag = meaning.get(
            val, ("✗ 彩排失败，退出码 %s —— 翻 logs\\monthly_*.log" % val, "bad"))
        self.log(msg, tag)
        self.log("临时任务已删除。", "dim")

    def do_run_now(self):
        if not task_state(TASK):
            messagebox.showinfo("还没有任务", "先点「创建 / 更新任务」。")
            return
        if not messagebox.askokcancel(
                "立即跑一次",
                "现在就完整跑一遍：下载 + 出报表，约 3 小时。\n\n"
                "期间这台机器会反复开关紫鸟浏览器，请不要手动去点紫鸟。\n"
                "这个窗口可以关掉，任务在后台照跑。\n\n开始？"):
            return
        self._work(lambda: _run(["schtasks", "/run", "/tn", TASK]),
                   self._after_run_now)

    def _after_run_now(self, r):
        rc, out = r
        if rc == 0:
            self.log("✓ 已启动。进度看 logs\\monthly_*.log，"
                     "或过一会儿点「刷新」。", "good")
        else:
            self.log("✗ 启动失败：%s" % (out.strip() or rc), "bad")

    def do_check(self):
        self.log("")
        self.log("=" * 64)
        self.log(">>> 机器自检")
        self._work(machine_checks, self._after_check)

    def _after_check(self, items):
        bad = 0
        for ok, title, note in items:
            mark, tag = ("  ✓", "good") if ok else (
                ("  ?", "warn") if ok is None else ("  ✗", "bad"))
            if ok is False:
                bad += 1
            self.log("%s  %-16s %s" % (mark, title, note), tag)
        if bad:
            self.log("%d 项没过。没过的项不修，定时任务可能半夜跑不起来。" % bad,
                     "warn")
            if messagebox.askyesno(
                    "修电源设置",
                    "要现在把「不睡眠 / 不休眠 / 不关屏」设好吗？\n\n"
                    "会弹一个管理员确认框。自动登录和屏保要手动设，\n"
                    "改不了的项日志里写了怎么办。"):
                self._work(fix_power, self._after_fix)
        else:
            self.log("全部通过，这台机器可以无人值守。", "good")

    def _after_fix(self, r):
        rc, out = r
        self.log("电源设置已执行，重新自检一次看看。" if rc == 0
                 else "没改成：%s" % (out.strip() or rc),
                 "good" if rc == 0 else "bad")

    def do_delete(self):
        if not messagebox.askokcancel(
                "删除定时任务",
                "只撤掉 Windows 里的计划任务，脚本、下载的数据、已生成的报表\n"
                "都不会动。以后想再开，回来点「创建 / 更新任务」即可。\n\n确定？"):
            return
        self._work(lambda: _run(["schtasks", "/delete", "/tn", TASK, "/f"]),
                   self._after_delete)

    def _after_delete(self, r):
        rc, out = r
        self.log("✓ 已删除。" if rc == 0 else "✗ 删除失败：%s" % (out.strip() or rc),
                 "good" if rc == 0 else "bad")
        self.refresh()

    def do_logs(self):
        if not os.path.isdir(LOG_DIR):
            messagebox.showinfo("还没有日志", "目录还不存在：\n%s" % LOG_DIR)
            return
        if os.name == "nt":
            os.startfile(LOG_DIR)             # noqa: S606
        else:
            subprocess.Popen(["xdg-open", LOG_DIR])

    def _on_close(self):
        save_settings(dict(load_settings(), sched_day=self.day.get(),
                           sched_hour=self.hour.get(),
                           sched_minute=self.minute.get()))
        self.root.destroy()


def main():
    if os.name != "nt":
        print("这个工具只在 Windows 上有意义（用的是 Windows 任务计划程序）。")
        return 1
    root = tk.Tk()
    try:
        root.call("tk", "scaling", 1.3)
    except tk.TclError:
        pass
    SchedApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    sys.exit(main())
