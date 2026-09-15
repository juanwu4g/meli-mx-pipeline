# -*- coding: utf-8 -*-
"""
给不写代码的人用的小窗口：选月份、选店铺、点按钮。

    双击 启动.cmd     （推荐，它会用对解释器）
    或   .venv\\Scripts\\python gui.py

这个界面只做三件事：把参数拼成命令行、把子进程的输出实时贴出来、把结果
说清楚。**所有真正的逻辑都在原来的脚本里**，这里一行业务判断都没有 ——
界面坏了不该影响出数，出数的口径也不该藏在界面里。

三个刻意的设计：

* **绝不吞掉校验结果。** 这套报表最值钱的就是那十几条硬校验。跑完只显示
  一个绿色"完成"等于把它们全扔了，所以状态栏直接写"13 份报表，12 份有
  校验未过"，并把未过的行在日志里标红。

* **边跑边显示。** 出一次报表 495 行输出，下载要跑三个小时。攒到最后再
  显示，用的人会以为卡死了。子线程读 stdout、主线程渲染，Tk 控件只在
  主线程碰。

* **跑的时候锁住按钮。** 下载脚本启动时会 taskkill 掉已运行的紫鸟客户端 ——
  同时点两次"下载"，第二次会把第一次杀掉。
"""
import datetime
import os
import queue
import subprocess
import sys
import threading

import tkinter as tk
from tkinter import ttk, messagebox

ROOT = os.path.dirname(os.path.abspath(__file__))
REPORT_DIR = os.path.join(ROOT, "downloads", "data", "reports", "financial")


def interpreter():
    """永远优先用项目自己的 .venv。

    不是洁癖：requirements.txt 锁的是 pandas 3，而机器上常见的 conda 环境
    是 pandas 2，两者对 NaN 转字符串的处理不同，报表会崩在 ⑨ 库存页。
    """
    p = os.path.join(ROOT, ".venv", "Scripts",
                     "python.exe" if os.name == "nt" else "python")
    return p if os.path.exists(p) else sys.executable


def recent_months(n=15):
    d = datetime.date.today().replace(day=1)
    out = []
    for _ in range(n):
        d -= datetime.timedelta(days=1)
        out.append("%04d-%02d" % (d.year, d.month))
        d = d.replace(day=1)
    return out


def store_list():
    """店铺白名单。读不到就返回空 —— 界面照常能开，按钮会提示配置有问题。"""
    try:
        sys.path.insert(0, ROOT)
        import store_config
        return list(store_config.batch_stores())
    except Exception:
        return []


class App(object):
    def __init__(self, root):
        self.root = root
        self.proc = None
        self.q = queue.Queue()
        root.title("meli-mx-pipeline —— 墨西哥店铺报表")
        root.geometry("980x680")
        root.minsize(820, 560)

        self._build_top()
        self._build_buttons()
        self._build_log()
        self._build_status()
        self.root.after(80, self._drain)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        if interpreter() == sys.executable and not os.path.exists(
                os.path.join(ROOT, ".venv")):
            self.log("[警告] 没找到 .venv，正在用 %s" % sys.executable, "warn")
            self.log("       依赖版本可能不对，建议先跑：python -m venv .venv"
                     " && .venv\\Scripts\\pip install -r requirements.txt", "warn")

    # ---------------------------------------------------------------- 布局

    def _build_top(self):
        f = ttk.LabelFrame(self.root, text=" 参数 ", padding=10)
        f.pack(fill="x", padx=12, pady=(12, 6))

        ttk.Label(f, text="会计月份").grid(row=0, column=0, sticky="w")
        self.month = ttk.Combobox(f, values=recent_months(), width=12,
                                  state="readonly")
        self.month.set(recent_months()[0])
        self.month.grid(row=0, column=1, sticky="w", padx=(6, 18))
        ttk.Label(f, text="默认是上一个完整月份。当月还没过完，做月报没有意义。",
                  foreground="#666").grid(row=0, column=2, columnspan=3, sticky="w")

        ttk.Label(f, text="店铺").grid(row=1, column=0, sticky="nw", pady=(10, 0))
        box = ttk.Frame(f)
        box.grid(row=1, column=1, columnspan=4, sticky="w", pady=(10, 0))
        self.all_stores = tk.BooleanVar(value=True)
        ttk.Checkbutton(box, text="全部店铺", variable=self.all_stores,
                        command=self._toggle_stores).pack(anchor="w")
        self.stores = store_list()
        self.lb = tk.Listbox(box, selectmode="extended", height=5,
                             exportselection=False, width=30)
        for s in self.stores:
            self.lb.insert("end", s)
        self.lb.pack(side="left", pady=(4, 0))
        self.lb.configure(state="disabled")
        hint = ("取消勾选“全部店铺”后，可按住 Ctrl 多选。\n"
                "白名单来自 config.json，共 %d 家。\n"
                "只选部分店铺时不会出全店汇总表（避免覆盖真正的那份）。"
                % len(self.stores))
        ttk.Label(box, text=hint, foreground="#666", justify="left").pack(
            side="left", anchor="n", padx=12, pady=(4, 0))

    def _toggle_stores(self):
        self.lb.configure(state="disabled" if self.all_stores.get() else "normal")

    def _build_buttons(self):
        f = ttk.LabelFrame(self.root, text=" 操作 ", padding=10)
        f.pack(fill="x", padx=12, pady=6)
        rows = [
            ("① 下载数据", self.do_download,
             "打开紫鸟，下载【上面所选店铺】的报表并清洗。每家约 15 分钟，\n"
             "全部 12 家约 3 小时。选了历史月份时会去那个月的账单明细页取数。"),
            ("② 出报表", self.do_reports,
             "用已下载的数据，生成【上面所选店铺】的财务报表。约 1 分钟。\n"
             "不联网、不开浏览器，可以随便重跑。"),
            ("先看用哪批数据", self.do_dry_run,
             "不生成任何文件，只列出每家店会用哪个下载目录、缺不缺该月账单。\n"
             "秒出。做历史月份前先点这个，能提前看到哪几家数据不全。"),
            ("补取欠的报表", self.do_collect,
             "MercadoPago 有些报表要生成几十分钟。这个只打开确实欠着报表的店，\n"
             "不重新下载，与上面的店铺选择无关。上一轮提示“still generating”就跑它。"),
            ("打开报表目录", self.do_open_dir,
             "在资源管理器里打开生成好的 .xlsx 所在目录。"),
        ]
        self.buttons = []
        for i, (label, cmd, desc) in enumerate(rows):
            b = ttk.Button(f, text=label, command=cmd, width=16)
            b.grid(row=i, column=0, sticky="w", pady=3)
            ttk.Label(f, text=desc, foreground="#555", justify="left").grid(
                row=i, column=1, sticky="w", padx=12)
            self.buttons.append(b)
        self.stop_btn = ttk.Button(f, text="中止", command=self.do_stop,
                                   width=16, state="disabled")
        self.stop_btn.grid(row=len(rows), column=0, sticky="w", pady=(10, 0))
        ttk.Label(f, text="强制结束正在跑的任务。下载中止后已下好的文件会保留。",
                  foreground="#555").grid(row=len(rows), column=1, sticky="w",
                                          padx=12, pady=(10, 0))

    def _build_log(self):
        f = ttk.LabelFrame(self.root, text=" 运行日志 ", padding=6)
        f.pack(fill="both", expand=True, padx=12, pady=6)
        self.txt = tk.Text(f, wrap="none", height=14, bg="#1e1e1e",
                           fg="#d4d4d4", insertbackground="#d4d4d4",
                           font=("Consolas", 9))
        sb = ttk.Scrollbar(f, orient="vertical", command=self.txt.yview)
        self.txt.configure(yscrollcommand=sb.set)
        self.txt.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.txt.tag_config("bad", foreground="#f48771")
        self.txt.tag_config("good", foreground="#89d185")
        self.txt.tag_config("warn", foreground="#e2c08d")
        self.txt.configure(state="disabled")

    def _build_status(self):
        self.status = tk.StringVar(value="就绪")
        bar = ttk.Frame(self.root)
        bar.pack(fill="x", padx=12, pady=(0, 12))
        self.status_lbl = ttk.Label(bar, textvariable=self.status,
                                    font=("", 10, "bold"))
        self.status_lbl.pack(side="left")
        ttk.Label(bar, text="解释器：%s" % interpreter(),
                  foreground="#888").pack(side="right")

    # ---------------------------------------------------------------- 日志

    def log(self, line, tag=None):
        self.txt.configure(state="normal")
        self.txt.insert("end", line + "\n", tag or ())
        self.txt.see("end")
        self.txt.configure(state="disabled")

    @staticmethod
    def _tag_for(line):
        if "✗" in line or "失败" in line or "Traceback" in line or "Error" in line:
            return "bad"
        if "[警告]" in line or "跳过" in line or "⚠" in line:
            return "warn"
        if "✓" in line or "成功" in line or "全部就位" in line:
            return "good"
        return None

    # ---------------------------------------------------------------- 执行

    def _args_stores(self):
        if self.all_stores.get():
            return []
        picked = [self.stores[i] for i in self.lb.curselection()]
        return ["--stores"] + picked if picked else []

    def do_download(self):
        a = ["run_batch.py"] + self._args_stores()
        if self.month.get() != recent_months()[0]:
            a += ["--month", self.month.get()]
            if not messagebox.askokcancel(
                    "补做历史月份",
                    "你选的是 %s，不是最近一个月。\n\n"
                    "脚本会去那个月的账单明细页单独取数，并把下载目录标上 "
                    "_m%s 后缀。\n\n继续？"
                    % (self.month.get(), self.month.get().replace("-", ""))):
                return
        self.run(a, "下载数据")

    def do_reports(self):
        self.run(["run_reports.py", "--month", self.month.get()]
                 + self._args_stores() + self._args_combined(), "出报表")

    def do_dry_run(self):
        self.run(["run_reports.py", "--month", self.month.get(), "--dry-run"]
                 + self._args_stores(), "先看用哪批数据")

    def _args_combined(self):
        """只选了部分店铺时，不要出汇总表。

        run_reports 的汇总表文件名是写死的 MX_ML_全店汇总_<月>.xlsx，不管实际
        跑了几家。所以选 2 家店跑一次，就会用一份"2 家店的汇总"**覆盖掉**真正
        12 家店的那份，而且文件名还写着"全店" —— 用的人完全看不出来。
        选了部分店铺就只出单店报表。
        """
        return [] if self.all_stores.get() else ["--no-combined"]

    def do_collect(self):
        self.run(["run_batch.py", "--collect-only"], "补取欠的报表")

    def do_open_dir(self):
        if not os.path.isdir(REPORT_DIR):
            messagebox.showinfo("还没有报表",
                                "目录还不存在：\n%s\n\n先点“出报表”。" % REPORT_DIR)
            return
        if os.name == "nt":
            os.startfile(REPORT_DIR)          # noqa: S606
        else:
            subprocess.Popen(["xdg-open", REPORT_DIR])

    def run(self, argv, title):
        if self.proc is not None:
            messagebox.showwarning("正在运行", "请等当前任务结束，或先点“中止”。")
            return
        self._lock(True)
        self.status.set("正在运行：%s …" % title)
        self.log("")
        self.log("=" * 72)
        self.log(">>> %s    %s" % (title, datetime.datetime.now().strftime("%H:%M:%S")))
        self.log("    " + " ".join([os.path.basename(interpreter())] + argv))
        self.log("=" * 72)

        env = dict(os.environ, PYTHONIOENCODING="utf-8", PYTHONUTF8="1")
        try:
            self.proc = subprocess.Popen(
                [interpreter()] + argv, cwd=ROOT, env=env,
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT, bufsize=0)
        except Exception as e:
            self.log("启动失败：%s" % e, "bad")
            self._lock(False)
            self.proc = None
            self.status.set("启动失败")
            return
        threading.Thread(target=self._pump, args=(self.proc, title),
                         daemon=True).start()

    def _pump(self, proc, title):
        """子线程：读子进程输出塞进队列。绝不碰 Tk 控件。"""
        for raw in iter(proc.stdout.readline, b""):
            self.q.put(("line", raw.decode("utf-8", "replace").rstrip("\r\n")))
        proc.stdout.close()
        self.q.put(("done", (proc.wait(), title)))

    def _drain(self):
        """主线程：把队列里的内容渲染出来。"""
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "line":
                    self.log(payload, self._tag_for(payload))
                else:
                    self._finish(*payload)
        except queue.Empty:
            pass
        self.root.after(80, self._drain)

    def _finish(self, code, title):
        self.proc = None
        self._lock(False)
        # 退出码含义见 README §9。1 不是"失败"，是"有事要看"——
        # 报表已经写出来了，只是有校验没过，这恰恰是最该让人看见的情况。
        meaning = {
            0: ("✓ %s 完成，全部正常" % title, "good"),
            1: ("⚠ %s 完成，但有需要关注的项 —— 翻上面日志里标红的行，"
                "或看报表 ⑩ 页" % title, "warn"),
            2: ("✗ %s 失败：紫鸟客户端出错" % title, "bad"),
            3: ("✗ %s 失败：店铺不在 config.json 的白名单里" % title, "bad"),
            4: ("✗ %s 失败：config.json 读不了或格式不对" % title, "bad"),
        }.get(code, ("✗ %s 结束，退出码 %s" % (title, code), "bad"))
        self.status.set(meaning[0])
        self.status_lbl.configure(
            foreground={"good": "#1a7f37", "warn": "#9a6700", "bad": "#b42318"}[meaning[1]])
        self.log("")
        self.log(meaning[0], meaning[1])

    def do_stop(self):
        if self.proc is None:
            return
        if not messagebox.askokcancel("中止", "确定要结束正在跑的任务吗？\n"
                                              "已经下好的文件会保留。"):
            return
        try:
            self.proc.terminate()
        except Exception as e:
            self.log("中止失败：%s" % e, "bad")

    def _lock(self, running):
        for b in self.buttons[:4]:            # 前四个会起子进程，最后一个只是开目录
            b.configure(state="disabled" if running else "normal")
        self.stop_btn.configure(state="normal" if running else "disabled")

    def _on_close(self):
        if self.proc is not None and not messagebox.askokcancel(
                "还在运行", "任务还在跑，关掉窗口会一起结束。确定？"):
            return
        if self.proc is not None:
            try:
                self.proc.terminate()
            except Exception:
                pass
        self.root.destroy()


def main():
    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista" if os.name == "nt" else "clam")
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
