# -*- coding: utf-8 -*-
"""
Ziniao (紫鸟) SuperBrowser client wrapper.

Handles the parts the vendor demo gets wrong or leaves implicit:
  - strips ELECTRON_RUN_AS_NODE when spawning the client (see GUIDE.md 6.1)
  - caps the updateCore retry loop instead of spinning forever (6.2)
  - always leaves the browser on a normal page, never the extension IP-check
    page, so a later run can reattach (6.3)

Typical use:

    from ziniao_client import ZiniaoClient

    with ZiniaoClient() as zn:
        zn.start()                              # launch + wait for IPC
        store = zn.find_store("BOCINA_SM")
        with zn.open_store(store) as session:
            session.driver.get(...)
"""
import io
import json
import os
import platform
import subprocess
import time
import uuid
from contextlib import contextmanager

import requests
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.common.exceptions import NoSuchElementException

import store_config

ROOT = os.path.dirname(os.path.abspath(__file__))

IS_WIN = platform.system() == "Windows"
IS_MAC = platform.system() == "Darwin"
IS_LINUX = platform.system() == "Linux"


class ZiniaoError(RuntimeError):
    pass


# re-exported so callers can catch scope errors without importing store_config
StoreNotAllowed = store_config.StoreNotAllowed


def load_config(path=None):
    path = path or os.path.join(ROOT, "config.json")
    if not os.path.exists(path):
        raise ZiniaoError(
            "config.json not found. Copy config.example.json and fill it in.")
    with io.open(path, encoding="utf-8") as f:
        return json.load(f)


def _find_client():
    """Best-effort auto-detect, mirroring bootstrap.py."""
    if IS_WIN:
        cands = [r"D:\ziniao\ziniao.exe", r"C:\ziniao\ziniao.exe",
                 os.path.expandvars(r"%LOCALAPPDATA%\Programs\ziniao\ziniao.exe"),
                 os.path.expandvars(r"%PROGRAMFILES%\ziniao\ziniao.exe")]
    elif IS_MAC:
        cands = ["/Applications/ziniao.app"]
    else:
        cands = ["/opt/ziniao/ziniaobrowser"]
    for c in cands:
        if c and os.path.exists(c):
            return c
    return None


class StoreSession(object):
    """One opened store: the IPC response plus an attached selenium driver."""

    def __init__(self, client, info, driver):
        self.client = client
        self.info = info
        self.driver = driver

    @property
    def oauth(self):
        return self.info.get("browserOauth") or self.info.get("browserId")

    @property
    def download_path(self):
        return self.info.get("downloadPath")

    @property
    def launcher_page(self):
        return self.info.get("launcherPage")

    def check_ip(self):
        """Open the proxy check page and report whether it passed.

        Always returns to launcher_page afterwards - leaving the browser on the
        chrome-extension:// page makes future reattaches fail (GUIDE.md 6.3).
        """
        url = self.info.get("ipDetectionPage")
        if not url:
            print("  [警告] 没有 ipDetectionPage，跳过 IP 检测")
            return True
        ok = False
        try:
            self.driver.get(url)
            time.sleep(3)
            self.driver.find_element(
                By.XPATH, '//button[contains(@class, "styles_btn--success")]')
            ok = True
        except NoSuchElementException:
            print("  [警告] 未找到 IP 检测的成功标识")
        except Exception as e:
            print("  [警告] IP 检测出错：%s" % str(e)[:100])
        finally:
            if self.launcher_page:
                self.driver.get(self.launcher_page)
                time.sleep(3)
        return ok


class ZiniaoClient(object):
    def __init__(self, config=None, config_path=None):
        cfg = config or load_config(config_path)
        self.cfg = cfg
        self.user_info = cfg["user_info"]
        self.port = cfg.get("socket_port", 16851)
        self.version = cfg.get("client_version", "v6")
        self.client_path = cfg.get("client_path") or _find_client()
        if not self.client_path:
            raise ZiniaoError("ziniao client not found; set client_path in config.json")
        d = cfg.get("driver_folder_path") or "webdriver"
        self.driver_folder = d if os.path.isabs(d) else os.path.join(ROOT, d)
        self._opened = []

    # ---------------- IPC ----------------

    def send(self, data, timeout=120):
        payload = dict(data)
        payload.update(self.user_info)
        payload.setdefault("requestId", str(uuid.uuid4()))
        try:
            r = requests.post("http://127.0.0.1:%d" % self.port,
                              json.dumps(payload).encode("utf-8"), timeout=timeout)
            return json.loads(r.text)
        except Exception:
            return None

    def _require_ok(self, resp, what):
        if resp is None:
            raise ZiniaoError("%s: no response from client" % what)
        code = resp.get("statusCode")
        if code == 0:
            return resp
        if code == -10003:
            raise ZiniaoError("%s: login/permission error: %s"
                              % (what, json.dumps(resp, ensure_ascii=False)[:300]))
        raise ZiniaoError("%s failed: %s"
                          % (what, json.dumps(resp, ensure_ascii=False)[:300]))

    # ---------------- lifecycle ----------------

    def is_up(self):
        return self.send({"action": "getBrowserList"}, timeout=15) is not None

    def kill_existing(self):
        """Terminate a running client. This closes any stores a human has open."""
        if IS_WIN:
            name = "SuperBrowser.exe" if self.version == "v5" else "ziniao.exe"
            os.system("taskkill /f /t /im " + name + " >nul 2>&1")
        elif IS_MAC:
            os.system("killall ziniao 2>/dev/null")
        else:
            os.system("killall ziniaobrowser 2>/dev/null")
        time.sleep(4)

    def launch(self):
        args = ["--run_type=web_driver", "--ipc_type=http", "--port=%d" % self.port]
        if IS_WIN:
            cmd = [self.client_path] + args
        elif IS_MAC:
            cmd = ["open", "-a", self.client_path, "--args"] + args
        else:
            cmd = [self.client_path, "--no-sandbox"] + args

        # GUIDE.md 6.1 - inherited ELECTRON_RUN_AS_NODE makes the client exit
        # immediately with "bad option: --run_type=web_driver".
        env = os.environ.copy()
        env.pop("ELECTRON_RUN_AS_NODE", None)
        subprocess.Popen(cmd, env=env)

    def wait_ready(self, timeout=120):
        deadline = time.time() + timeout
        while time.time() < deadline:
            if self.is_up():
                return True
            time.sleep(2)
        return False

    def start(self, restart=True):
        """Ensure the client is running in webdriver/http mode."""
        if self.is_up():
            print("客户端已在端口 %d 上监听" % self.port)
            return
        if restart:
            print("正在结束已运行的紫鸟客户端…")
            self.kill_existing()
        print("以 webdriver 模式启动客户端（端口 %d）…" % self.port)
        self.launch()
        if not self.wait_ready():
            raise ZiniaoError("client did not open port %d in time" % self.port)
        print("客户端就绪")

    def update_core(self, max_attempts=15):
        """Pre-download cores. Non-fatal: on some installs this never returns 0
        (GUIDE.md 6.2), and startBrowser works regardless."""
        for _ in range(max_attempts):
            r = self.send({"action": "updateCore"}, timeout=120)
            if r is None:
                time.sleep(2)
                continue
            if r.get("statusCode") == 0:
                print("内核已是最新")
                return True
            time.sleep(2)
        print("[警告] updateCore 未正常返回，继续执行")
        return False

    # ---------------- stores ----------------

    def list_stores(self):
        r = self._require_ok(self.send({"action": "getBrowserList"}), "getBrowserList")
        return r.get("browserList") or []

    def find_store(self, name=None):
        """Resolve a store name to its record.

        Every path to `startBrowser` goes through here, so this is where the
        allowlist is enforced - and it is enforced BEFORE the lookup, so an
        unintended name never reaches Ziniao at all.

        Note that `getBrowserList` returns every store the enterprise account
        can see (93 on this account) and carries no permission field, so the
        client cannot tell which are authorised for WebDriver. The allowlist in
        config.json is the only local guard against opening the wrong one.
        """
        name = name or store_config.default_store()
        store_config.require_allowed(name)
        for b in self.list_stores():
            if (b.get("browserName") or "") == name:
                return b
        raise ZiniaoError("store %r not found on this account" % name)

    def _resolve_driver(self, info):
        bpath = info.get("browserPath") or ""
        if bpath.lower().endswith(("superbrowser.exe", "superbrowser")):
            bpath = os.path.dirname(bpath)
        if bpath:
            name = "webdriver.exe" if IS_WIN else "webdriver"
            p = os.path.join(bpath, name)
            if os.path.exists(p):
                return p
        major = str(info.get("core_version", "")).split(".")[0]
        name = "chromedriver%s%s" % (major, ".exe" if IS_WIN else "")
        p = os.path.join(self.driver_folder, name)
        if os.path.exists(p):
            return p
        raise ZiniaoError(
            "no chromedriver for core %s; run: python bootstrap.py --drivers"
            % info.get("core_version"))

    @contextmanager
    def open_store(self, store, headless=False, close_when_done=True):
        oauth = store["browserOauth"] if isinstance(store, dict) else store
        name = store.get("browserName", oauth) if isinstance(store, dict) else oauth
        print("打开店铺：%s" % name)
        info = self._require_ok(self.send({
            "action": "startBrowser",
            "browserOauth": oauth,
            "isWaitPluginUpdate": 0,
            "isHeadless": 1 if headless else 0,
            "isWebDriverReadOnlyMode": 0,
            "cookieTypeLoad": 0,
            "cookieTypeSave": 0,
            "runMode": "1",
            "isLoadUserPlugin": False,
            "pluginIdType": 1,
            "privacyMode": 0,
        }, timeout=300), "startBrowser")

        drv_path = self._resolve_driver(info)
        opts = webdriver.ChromeOptions()
        opts.add_argument("--log-level=3")
        opts.add_experimental_option(
            "debuggerAddress", "127.0.0.1:%s" % info.get("debuggingPort"))
        driver = webdriver.Chrome(service=Service(drv_path), options=opts)
        driver.implicitly_wait(20)

        session = StoreSession(self, info, driver)
        self._opened.append(session)
        try:
            yield session
        finally:
            # never quit() - that would kill the browser we attached to.
            if close_when_done:
                self.close_store(session)

    def close_store(self, session):
        oauth = session.oauth if isinstance(session, StoreSession) else session
        print("关闭店铺 %s" % oauth)
        self.send({"action": "stopBrowser", "browserOauth": oauth, "duplicate": 0})
        if isinstance(session, StoreSession) and session in self._opened:
            self._opened.remove(session)

    def exit_client(self):
        self.send({"action": "exit"}, timeout=30)

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False
