# -*- coding: utf-8 -*-
"""
Environment check / setup for the ziniao automation on a new machine.

    python bootstrap.py              # check everything, report what's missing
    python bootstrap.py --drivers    # also download the fallback chromedriver set (~300 MB)

Checks, in order:
  1. Python version (and whether you're inside .venv)
  2. requests + selenium importable, at the pinned versions
  3. config.json present and complete
  4. ziniao client located on disk
  5. ELECTRON_RUN_AS_NODE not poisoning the environment
  6. fallback chromedriver folder (optional)

Exits non-zero if anything required is missing.
"""
import argparse
import hashlib
import io
import json
import os
import console  # noqa: F401  中文输出编码保护
import platform
import sys

ROOT = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(ROOT, "config.json")

IS_WIN = platform.system() == "Windows"
IS_MAC = platform.system() == "Darwin"
IS_LINUX = platform.system() == "Linux"

OK, BAD, WARN = "[ 通过 ]", "[ 失败 ]", "[ 警告 ]"
problems = []


def fail(msg, fix):
    problems.append((msg, fix))
    print("{} {}".format(BAD, msg))
    print("       解决：{}".format(fix))


# --- candidate install locations, checked in order -------------------------
WIN_CLIENT_CANDIDATES = [
    r"D:\ziniao\ziniao.exe",
    r"C:\ziniao\ziniao.exe",
    os.path.expandvars(r"%LOCALAPPDATA%\Programs\ziniao\ziniao.exe"),
    os.path.expandvars(r"%PROGRAMFILES%\ziniao\ziniao.exe"),
    # V5 client is named starter.exe
    r"D:\SuperBrowser\starter.exe",
    r"C:\SuperBrowser\starter.exe",
]
MAC_CLIENT_CANDIDATES = ["/Applications/ziniao.app"]
LINUX_CLIENT_CANDIDATES = ["/opt/ziniao/ziniaobrowser"]


def find_client():
    if IS_WIN:
        cands = WIN_CLIENT_CANDIDATES
    elif IS_MAC:
        cands = MAC_CLIENT_CANDIDATES
    else:
        cands = LINUX_CLIENT_CANDIDATES
    for c in cands:
        if c and os.path.exists(c):
            return c
    return None


def load_config():
    if not os.path.exists(CONFIG_PATH):
        return None
    with io.open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


def check_python():
    v = sys.version_info
    mark = OK if v >= (3, 8) else BAD
    print("{} python {}.{}.{}".format(mark, v.major, v.minor, v.micro))
    print("       {}".format(sys.executable))
    if v < (3, 8):
        fail("需要 python 3.8 或更高版本", "安装更新的 python")
    in_venv = sys.prefix != getattr(sys, "base_prefix", sys.prefix)
    if not in_venv:
        print("{} 当前不在 .venv 环境中，依赖可能被装到全局".format(WARN))
        print("       解决：python -m venv .venv，之后使用 venv 里的 python")


def check_deps():
    for mod, pin in (("requests", "2.32.2"), ("selenium", "4.48.0")):
        try:
            m = __import__(mod)
        except ImportError:
            fail("{} 未安装".format(mod),
                 "pip install -r requirements.txt")
            continue
        ver = getattr(m, "__version__", "?")
        if ver == pin:
            print("{} {} {}".format(OK, mod, ver))
        else:
            print("{} {} {}  （要求版本：{}）".format(WARN, mod, ver, pin))


def check_config():
    try:
        cfg = load_config()
    except ValueError as e:
        fail("config.json 不是合法 JSON：{}".format(e), "修正文件语法")
        return None
    if cfg is None:
        fail("缺少 config.json",
             "把 config.example.json 复制为 config.json 并填写账号信息")
        return None

    ui = cfg.get("user_info") or {}
    missing = [k for k in ("company", "username", "password")
               if not ui.get(k) or str(ui[k]).startswith("YOUR_")]
    if missing:
        fail("config.json 的 user_info 不完整：{}".format(", ".join(missing)),
             "填写企业登录的公司名/用户名/密码")
    else:
        print("{} config.json  company={}  port={}".format(
            OK, ui.get("company"), cfg.get("socket_port")))
    return cfg


def check_client(cfg):
    path = (cfg or {}).get("client_path")
    if path and os.path.exists(path):
        print("{} 紫鸟客户端（来自配置）：{}".format(OK, path))
        return path
    if path:
        print("{} 配置里的 client_path 不存在：{}".format(WARN, path))
    found = find_client()
    if found:
        print("{} 紫鸟客户端（自动检测）：{}".format(OK, found))
        print('       如需固定路径，可在 config.json 里设置 "client_path"')
        return found
    fail("未找到紫鸟客户端",
         "安装紫鸟浏览器，或在 config.json 里设置 client_path")
    return None


def check_electron_env():
    if os.environ.get("ELECTRON_RUN_AS_NODE"):
        print("{} 当前环境设置了 ELECTRON_RUN_AS_NODE".format(WARN))
        print("       启动器会在子进程里清掉它，所以本身没问题，")
        print("       但不要用裸 os.system() 去启动客户端。")
    else:
        print("{} 未设置 ELECTRON_RUN_AS_NODE".format(OK))


def driver_dir(cfg):
    d = (cfg or {}).get("driver_folder_path") or "webdriver"
    return d if os.path.isabs(d) else os.path.join(ROOT, d)


def check_drivers(cfg, download):
    d = driver_dir(cfg)
    if os.path.isdir(d):
        present = sorted(f for f in os.listdir(d) if f.startswith("chromedriver"))
    else:
        present = []
    print("{} 备用 chromedriver：{} 个，位于 {}".format(
        OK if present else WARN, len(present), d))
    if not present:
        print("       通常无需处理：每个店铺内核自带 webdriver.exe，")
        print("       优先使用它。需要备用驱动时加 --drivers 下载。")
    if download:
        download_drivers(d)


DRIVER_CONFIG_URLS = {
    "win": "https://cdn-superbrowser-attachment.ziniao.com/webdriver/exe_32/config.json",
    "mac_x64": "https://cdn-superbrowser-attachment.ziniao.com/webdriver/mac/x64/config.json",
    "mac_arm": "https://cdn-superbrowser-attachment.ziniao.com/webdriver/mac/arm64/config.json",
}


def download_drivers(dest):
    import requests
    if IS_WIN:
        url = DRIVER_CONFIG_URLS["win"]
    elif IS_MAC:
        key = "mac_arm" if platform.machine() == "arm64" else "mac_x64"
        url = DRIVER_CONFIG_URLS[key]
    else:
        print("       Linux 使用内核自带驱动，无需下载")
        return
    print("--- 正在下载 chromedriver 到 {} ---".format(dest))
    cfg = json.loads(requests.get(url, timeout=60).text)
    if not os.path.isdir(dest):
        os.makedirs(dest)
    for item in cfg:
        name = item["name"] + (".exe" if IS_WIN else "")
        path = os.path.join(dest, name)
        if os.path.exists(path):
            with open(path, "rb") as f:
                if hashlib.new("sha1", f.read()).hexdigest() == item["sha1"]:
                    print("  已存在   {}".format(name))
                    continue
        print("  下载中   {}".format(name))
        r = requests.get(item["url"], stream=True, timeout=300)
        if r.status_code != 200:
            print("           失败，状态码 {}".format(r.status_code))
            continue
        with open(path, "wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)
        if not IS_WIN:
            os.chmod(path, 0o755)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--drivers", action="store_true",
                    help="download the fallback chromedriver set (~300 MB)")
    args = ap.parse_args()

    print("=" * 62)
    print(" 紫鸟自动化 —— 环境自检")
    print(" 系统平台：{} {}".format(platform.system(), platform.machine()))
    print("=" * 62)
    check_python()
    check_deps()
    cfg = check_config()
    check_client(cfg)
    check_electron_env()
    check_drivers(cfg, args.drivers)
    print("=" * 62)
    if problems:
        print(" 运行前必须先解决 {} 个问题：".format(len(problems)))
        for msg, fix in problems:
            print("   - {}".format(msg))
            print("       -> {}".format(fix))
        return 1
    print(" 环境检查通过")
    return 0


if __name__ == "__main__":
    sys.exit(main())
