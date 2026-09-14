# -*- coding: utf-8 -*-
"""
Open a store browser for manual poking - no Ziniao login screen.

    python open_store.py                          # BOCINA_SM, leave it open
    python open_store.py --store BOCINA_TA02
    python open_store.py --url https://www.mercadopago.com.mx/balance/reports

Credentials come from config.json over the client's HTTP IPC, so the Ziniao
login UI never appears. The browser is deliberately LEFT OPEN when this exits -
that is the whole point - and the debugging port is printed so a probe script
can attach to the same session:

    opts.add_experimental_option("debuggerAddress", "127.0.0.1:<port>")

If the client is already running in webdriver mode it is reused; otherwise it is
restarted into that mode, which closes any store a human has open.
"""
import argparse
import sys
import time

from ziniao_client import ZiniaoClient, ZiniaoError, StoreNotAllowed
from store_config import StoreConfigError
import console  # noqa: F401  中文输出编码保护
import store_config




def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--store", default=None,
                    help="store name; defaults to config.json stores.default")
    ap.add_argument("--url", default=None,
                    help="navigate here after opening (default: the seller panel)")
    ap.add_argument("--restart", action="store_true",
                    help="force a client restart even if one is already listening")
    ap.add_argument("--ip-check", action="store_true",
                    help="run the proxy check before handing over")
    args = ap.parse_args()

    client = ZiniaoClient()
    if client.is_up() and not args.restart:
        print("客户端已处于 webdriver 模式，端口 %d" % client.port)
    else:
        client.start(restart=True)

    store = client.find_store(args.store)   # None -> config default
    print("店铺：%s | %s | IP %s"
          % (store.get("browserName"), store.get("platform_name"),
             store.get("browserIp")))

    # close_when_done=False is what leaves the browser up after we exit
    with client.open_store(store, close_when_done=False) as s:
        if args.ip_check:
            print("IP 检测：%s" % ("通过" if s.check_ip() else "失败"))
        target = args.url or s.launcher_page
        s.driver.get(target)
        time.sleep(3)
        print("\n" + "=" * 58)
        print("  就绪 —— 浏览器保持打开")
        print("  调试端口   ： %s" % s.info.get("debuggingPort"))
        print("  下载目录   ： %s" % s.download_path)
        print("  当前地址   ： %s" % s.driver.current_url)
        print("=" * 58)
        print("\n之后关闭它：")
        print("  python -c \"from ziniao_client import ZiniaoClient;"
              "c=ZiniaoClient();c.close_store('%s')\"" % s.oauth)
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except StoreConfigError as e:
        print("\n配置错误：%s" % e)
        sys.exit(4)
    except StoreNotAllowed as e:
        # Scope error, not a failure: the store is not on the allowlist.
        print("\n店铺不在白名单：%s" % e)
        sys.exit(3)
    except ZiniaoError as e:
        print("\n紫鸟客户端错误：%s" % e)
        sys.exit(2)
