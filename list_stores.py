# -*- coding: utf-8 -*-
"""
List the stores on the Ziniao account.

    python list_stores.py                       # every store, grouped by platform
    python list_stores.py --platform MercadoLibre-墨西哥-本土
    python list_stores.py --mx                  # shorthand for the MX MeLi platform
    python list_stores.py --mx --names          # bare names, one per line
    python list_stores.py --tag "MX MeLi/AMZ"

`--names` prints only store names, so it can feed a loop:

    for /f %s in ('python list_stores.py --mx --names') do python run_downloads.py --store %s
"""
import argparse
import collections
import sys

from ziniao_client import ZiniaoClient, ZiniaoError, StoreNotAllowed
from store_config import StoreConfigError
import console  # noqa: F401  中文输出编码保护
import store_config

MX_MELI = "MercadoLibre-墨西哥-本土"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--platform", default=None, help="exact platform_name filter")
    ap.add_argument("--mx", action="store_true",
                    help="shorthand for --platform %s" % MX_MELI)
    ap.add_argument("--tag", default=None, help="only stores carrying this tag")
    ap.add_argument("--names", action="store_true",
                    help="print bare store names only, for scripting")
    ap.add_argument("--include-expired", action="store_true")
    args = ap.parse_args()

    client = ZiniaoClient()
    if not client.is_up():
        client.start()
    stores = client.list_stores()

    if not args.include_expired:
        stores = [s for s in stores if not s.get("isExpired")]
    platform = MX_MELI if args.mx else args.platform
    if platform:
        stores = [s for s in stores if (s.get("platform_name") or "") == platform]
    if args.tag:
        stores = [s for s in stores if args.tag in (s.get("tags") or [])]
    stores.sort(key=lambda s: (s.get("platform_name") or "", s.get("browserName") or ""))

    if args.names:
        for s in stores:
            print(s.get("browserName"))
        return 0

    if not platform and not args.tag:
        by_plat = collections.Counter(s.get("platform_name") or "?" for s in stores)
        print("共 %d 家店铺，分布在 %d 个平台\n" % (len(stores), len(by_plat)))
        for name, n in by_plat.most_common():
            print("  %-34s %d" % (name, n))
        print("\n用 --platform / --mx / --tag 可列出具体店铺")
        return 0

    allowed = store_config.allowed_stores()
    print("%d 家店铺   配置：%s\n" % (len(stores), store_config.describe()))
    print("  %-3s %-22s %-16s %-22s %s"
          % ("ok", "name", "ip", "tags", "store_username"))
    print("  " + "-" * 92)
    for s in stores:
        # Several stores share one proxy IP - that matters when deciding how
        # aggressively to run them back to back.
        # "*" marks a store this tooling is actually permitted to open.
        mark = "*" if (not allowed or s.get("browserName") in allowed) else " "
        print("  %-3s %-22s %-16s %-22s %s" % (
            mark,
            (s.get("browserName") or "")[:22],
            s.get("browserIp") or "-",
            ",".join(s.get("tags") or [])[:22],
            (s.get("store_username") or "")[:30]))

    ips = collections.Counter(s.get("browserIp") for s in stores)
    shared = {ip: n for ip, n in ips.items() if n > 1}
    if shared:
        print("\n  共用代理出口 IP 的店铺：")
        for ip, n in sorted(shared.items(), key=lambda kv: -kv[1]):
            print("    %-16s %d 家店铺" % (ip, n))
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
