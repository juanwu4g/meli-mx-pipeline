# -*- coding: utf-8 -*-
"""每月定时跑一次：下载**上个月**的数据，再出上个月的报表。

    python run_monthly.py                    # 上个月，下载 + 出报表
    python run_monthly.py --month 2026-07    # 指定月份（补做）
    python run_monthly.py --step report      # 只重出报表，不重新下载
    python run_monthly.py --dry-run          # 只打印会执行什么

给计划任务用的入口，人也可以手动跑。定在每月 15 号：那时候上个月的账单已经
闭账，数字基本稳定了。

为什么是子进程而不是 import
---------------------------
和 transform.py 同一个理由：下载要驱动浏览器跑半小时以上，出报表要吃 pandas。
任何一边崩了都不该把另一边连坐。子进程最坏也就是返回一个非零退出码。

例外是文件开头那句 `from run_reports import prev_month` —— 它是**故意**放在
最前面的：出报表排在四十分钟的下载之后，等跑到那一步才发现 pandas 装坏了太
晚了。这句 import 失败就立刻退出，那时候什么都还没做。

下载失败了还出不出报表
----------------------
看下载那步的退出码：

* 0（全成）、1（个别店铺失败）→ **出**。缺的店会体现在报表的
  【⑩ 数据缺口与待办】里，有总比没有强。
* 2/3/4（紫鸟错误 / 白名单 / 配置）→ **不出**。这几种情况多半一个文件都没下
  下来，这时候出报表会拿旧目录的数据生成一份看起来正常的报表，比没有报表更
  危险。

没有超时控制
------------
跑批卡住由计划任务的"如果任务运行超过 N 小时则停止"兜底，这里不自己实现 ——
边流式转发子进程输出边计时会把这个壳搞复杂，而计划任务那个开关是免费的。

退出码
------
    0  全部成功
    1  完成，但有需要关注的项（个别店铺失败，或报表有校验未过）
    2  紫鸟客户端 / 浏览器错误，下载没完成，报表未生成
    3  店铺不在白名单
    4  配置错误
    5  上一次还在跑（锁没释放）
"""
import argparse
import datetime
import io
import json
import os
import re
import subprocess
import sys

import console  # noqa: F401  中文输出编码保护
# 见上：故意的前置 import，用来在下载之前就验证报表侧能不能加载。
from run_reports import prev_month, OUT_DIR

ROOT = os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(ROOT, "logs")
LOCK = os.path.join(LOG_DIR, "monthly.lock")

# 锁超过这个时长就认为是上次被强杀留下的残骸，直接接管。计划任务的超时通常
# 设 4 小时，这里留足余量。
STALE_HOURS = 8


def now():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def config():
    """读 config.json 里的 monthly 段。读不到就当没配，不影响主流程。"""
    try:
        with io.open(os.path.join(ROOT, "config.json"), encoding="utf-8") as f:
            return json.load(f).get("monthly") or {}
    except Exception:
        return {}


# ---------------------------------------------------------------- 锁

def alive(pid):
    """这个进程还在吗。判断不了就返回 True（宁可多等，不要抢锁）。

    不能用 os.kill(pid, 0) —— Windows 上 CPython 的 os.kill 走的是
    TerminateProcess，信号 0 会把目标进程**杀掉**，而不是探测它。
    """
    if not pid:
        return True
    try:
        out = subprocess.run(["tasklist", "/FI", "PID eq %d" % pid, "/NH"],
                             capture_output=True, timeout=20,
                             encoding="utf-8", errors="replace").stdout or ""
    except Exception:
        return True
    # 没匹配到时 tasklist 打的是一句本地化的提示，里面不会有这个 PID。
    return str(pid) in out


def acquire():
    """防重入。手动跑和计划任务撞上时，两个进程同时驱动同一个紫鸟客户端，
    后启动的那个会把前一个的浏览器连同会话一起重启掉。

    锁里记了 PID，进程没了就立刻接管。只看文件年龄是不够的：跑批中途被
    强杀（关窗口、任务超时、断电）时锁会留下，而下一次定时运行往往就在几
    分钟后 —— 实测过一次 10:01 被杀、10:10 的定时任务拿到退出码 5，
    再按 STALE_HOURS 还要空等 8 小时。
    """
    if not os.path.isdir(LOG_DIR):
        os.makedirs(LOG_DIR)
    if os.path.exists(LOCK):
        age = (datetime.datetime.now()
               - datetime.datetime.fromtimestamp(os.path.getmtime(LOCK)))
        try:
            with io.open(LOCK, encoding="utf-8") as f:
                who = f.read().strip()
        except Exception:
            who = ""
        m = re.search(r"pid=(\d+)", who)
        pid = int(m.group(1)) if m else 0
        if not alive(pid):
            print("[提示] 锁是 %s 留下的，但那个进程已经不在了（上次被中断），"
                  "本次接管。" % (who or "上一次运行"))
        elif age < datetime.timedelta(hours=STALE_HOURS):
            print("上一次还在跑（%s），本次跳过。" % (who or "?"))
            print("确认没有在跑的话，删掉 %s 再试。" % LOCK)
            return False
        else:
            print("[警告] 锁已存在 %.1f 小时且进程仍在，超过 %d 小时判为异常，接管。"
                  % (age.total_seconds() / 3600.0, STALE_HOURS))
    with io.open(LOCK, "w", encoding="utf-8") as f:
        f.write("pid=%d 起于 %s" % (os.getpid(), now()))
    return True


def release():
    try:
        os.remove(LOCK)
    except OSError:
        pass


# ---------------------------------------------------------------- 执行

def child_env():
    env = os.environ.copy()
    # 店名和提示文案都是中文；计划任务里 stdout 是管道，默认会按 cp936 编码
    # 写入并抛 UnicodeEncodeError（GUIDE.md 6.4）。
    env["PYTHONUTF8"] = "1"
    env["PYTHONIOENCODING"] = "utf-8"
    return env


def run(label, cmd, log_path):
    """跑一步，输出同时进控制台和日志文件。返回退出码。"""
    head = "%s  %s\n%s\n" % (now(), label, " ".join(cmd))
    print("\n" + "=" * 70 + "\n" + head + "=" * 70)
    began = datetime.datetime.now()
    with io.open(log_path, "a", encoding="utf-8") as log:
        log.write("\n" + "=" * 70 + "\n" + head + "=" * 70 + "\n")
        log.flush()
        p = subprocess.Popen(cmd, cwd=ROOT, stdout=subprocess.PIPE,
                             stderr=subprocess.STDOUT, text=True,
                             encoding="utf-8", errors="replace",
                             env=child_env())
        for line in p.stdout:
            sys.stdout.write(line)
            # 每行都 flush：计划任务超时会强杀进程，缓冲区里的内容全丢，
            # 而那恰恰是最需要看的部分。
            log.write(line)
            log.flush()
        code = p.wait()
        secs = int((datetime.datetime.now() - began).total_seconds())
        tail = "%s  %s 结束，退出码 %d，用时 %d 秒\n" % (now(), label, code, secs)
        log.write(tail)
    print(tail.rstrip())
    return code, secs


def alert(month, lines, out_dir):
    """把失败摆到财务看得见的地方。

    计划任务会把 stdout 丢掉，没人会主动去翻 logs/。报表输出目录是财务每个月
    一定会打开的地方，所以告警文件放那儿。

    **只在真出事的时候写。** 报表侧退出码 1（有校验未过）是常态 —— 桥B
    「账单应付 vs 实际扣款」每家店每个月都不过，那是已知的口径问题。真按退
    出码 1 就告警，财务每个月都会看到一个"报表异常"，两个月后这个文件就彻底
    没人看了。会写这个文件的只有两类：**店铺没下到数据**，和**报表根本没生
    成**。报表本身有校验没过，报表的【⑩ 数据缺口与待办】里写着，不劳告警。
    """
    body = ["%s 月报生成异常" % month, "时间：%s" % now(), ""]
    body += lines
    body += ["", "完整日志：%s" % os.path.join(LOG_DIR, "monthly_%s.log"
                                               % month.replace("-", "")),
             "处理办法见 README.md 第 12 节《出问题先做这三步》。"]
    text = "\n".join(body)
    try:
        if not os.path.isdir(out_dir):
            os.makedirs(out_dir)
        p = os.path.join(out_dir, "！报表异常_%s.txt" % month.replace("-", ""))
        with io.open(p, "w", encoding="utf-8") as f:
            f.write(text)
        print("\n已写出告警文件：%s" % p)
    except Exception as e:
        print("\n[警告] 告警文件写失败：%s" % str(e)[:120])

    url = config().get("webhook")
    if not url:
        return
    try:
        import requests
        requests.post(url, json={"msgtype": "text",
                                 "text": {"content": text}}, timeout=15)
        print("已推送 webhook")
    except Exception as e:
        print("[警告] webhook 推送失败：%s" % str(e)[:120])


# ---------------------------------------------------------------- 主流程

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", default=None, metavar="YYYY-MM",
                    help="目标月份，默认上个月")
    ap.add_argument("--step", choices=["both", "download", "report"],
                    default="both", help="只跑其中一步")
    ap.add_argument("--out-dir", default=OUT_DIR, help="报表输出目录")
    ap.add_argument("--dry-run", action="store_true",
                    help="只打印会执行什么")
    args = ap.parse_args()

    month = args.month or prev_month()
    tag = month.replace("-", "")
    log_path = os.path.join(LOG_DIR, "monthly_%s.log" % tag)
    py = sys.executable

    dl = [py, "-u", os.path.join(ROOT, "run_batch.py"), "--month", month]
    # 不开放 --stores：店铺子集会让 run_reports 写出一份只含子集、却仍叫
    # MX_ML_全店汇总_<月>.xlsx 的文件，把真正的 12 店汇总覆盖掉。需要补单店
    # 的时候走界面或者直接敲 run_reports.py，那边有对应的开关。
    rp = [py, "-u", os.path.join(ROOT, "run_reports.py"), "--month", month,
          "--out-dir", args.out_dir]

    print("目标月份：%s" % month)
    print("日志　　：%s" % log_path)
    print("输出目录：%s" % os.path.abspath(args.out_dir))
    if args.dry_run:
        print("\n[dry-run] 会依次执行：")
        if args.step in ("both", "download"):
            print("  " + " ".join(dl))
        if args.step in ("both", "report"):
            print("  " + " ".join(rp))
        return 0

    if not acquire():
        return 5

    began = now()
    # bad    ：写进控制台和摘要，供人翻查
    # urgent ：还要写成告警文件推到财务眼前。判断标准见 alert() 的说明。
    steps, worst, bad, urgent = [], 0, [], []
    try:
        if not os.path.isdir(LOG_DIR):
            os.makedirs(LOG_DIR)

        if args.step in ("both", "download"):
            code, secs = run("① 下载 %s 的数据" % month, dl, log_path)
            steps.append({"step": "download", "code": code, "seconds": secs})
            worst = max(worst, code)
            if code >= 2:
                urgent.append("下载失败（退出码 %d），报表未生成。"
                              "多半是紫鸟没起来或者配置有问题。" % code)
                bad += urgent
                alert(month, urgent, args.out_dir)
                return code
            if code == 1:
                urgent.append("部分店铺下载失败，报表已按现有数据生成，"
                              "缺口见报表的【⑩ 数据缺口与待办】。")
                bad += urgent[-1:]

        if args.step in ("both", "report"):
            code, secs = run("② 生成 %s 的报表" % month, rp, log_path)
            steps.append({"step": "report", "code": code, "seconds": secs})
            worst = max(worst, code)
            if code == 1:
                # 常态，不告警。理由见 alert()。
                bad.append("有报表校验未过，文件已写出。桥B 每月都不过，"
                           "属已知口径问题；其余的看报表里的校验清单。")
            elif code >= 2:
                urgent.append("报表生成失败（退出码 %d）。" % code)
                bad.append(urgent[-1])
    finally:
        release()

    doc = {"month": month, "began": began, "ended": now(), "steps": steps,
           "exit": worst, "urgent": urgent,
           "log": log_path, "out_dir": os.path.abspath(args.out_dir)}
    with io.open(os.path.join(LOG_DIR, "monthly_%s.json" % tag),
                 "w", encoding="utf-8") as f:
        f.write(json.dumps(doc, ensure_ascii=False, indent=2))

    print("\n" + "=" * 70)
    if worst == 0:
        print("%s 月报完成，全部成功。" % month)
    else:
        print("%s 月报完成，退出码 %d —— 有需要关注的项：" % (month, worst))
        for b in bad:
            print("  · %s" % b)
        if urgent:
            alert(month, urgent, args.out_dir)
    print("=" * 70)
    return worst


if __name__ == "__main__":
    sys.exit(main())
