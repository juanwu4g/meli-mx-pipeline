# -*- coding: utf-8 -*-
"""
Single source of truth for which store the tooling touches.

The store name used to be hardcoded in six places (`run_downloads.py`,
`open_store.py`, both analysis scripts and twice in `meli_data.py`), which meant
there was no way to see - or limit - what the tooling could open. It now lives
in `config.json`:

    "stores": {
      "default": "BOCINA_SM",
      "allowed": ["BOCINA_SM", "EWTTO_SM"],
      "batch": []
    }

`default`  the store used when no --store is given.
`allowed`  the ONLY stores that may be opened. An empty list means no
           restriction, which is the previous behaviour; fill it in to make the
           scope explicit and auditable.
`batch`    which stores `run_batch.py` walks, in order. Empty (or absent)
           means "every store in `allowed`", which is what you usually want -
           it exists for the case where the allowlist is deliberately wider
           than the nightly run.

Deliberately dependency-free: `meli_data.py` imports this, and it is in turn
imported by the analysis scripts, which have no business loading selenium or
requests just to learn a store name.
"""
import io
import console  # noqa: F401  中文店名输出保护
import json
import os

ROOT = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(ROOT, "config.json")



class StoreNotAllowed(Exception):
    """Raised before any request reaches Ziniao."""


class StoreConfigError(Exception):
    """config.json is unreadable, so the allowlist cannot be trusted."""


def _load():
    """Read config.json, or fail loudly.

    This MUST NOT swallow a parse error. An earlier version returned {} on
    malformed JSON, which made `allowed` fall back to [] - meaning
    "unrestricted" - so a single missing comma silently unlocked all 93 stores.
    A guard that fails open is not a guard. Fail closed instead.
    """
    try:
        with io.open(CONFIG_PATH, encoding="utf-8") as f:
            return json.load(f)
    except ValueError as e:
        raise StoreConfigError(
            "config.json is not valid JSON (%s). Refusing to run: a malformed "
            "config cannot be trusted to limit which stores may be opened." % e)
    except (IOError, OSError) as e:
        raise StoreConfigError("cannot read %s (%s)" % (CONFIG_PATH, e))


def _stores():
    cfg = _load().get("stores")
    return cfg if isinstance(cfg, dict) else {}


def default_store():
    """The store used when nothing is specified."""
    name = _stores().get("default")
    if not name:
        # 绝不内置一个店名兜底。早先这里写死了某家店，config.json 缺少
        # stores.default 时会静默打开它 —— 白名单为空（不限制）的情况下
        # 就是在没人要求的店上跑了一整轮。缺配置要报错，不要猜。
        raise StoreConfigError(
            "config.json 里没有 stores.default。请在其中指定默认店铺，"
            "或在命令行用 --store / --stores 明确指定。")
    return name


def allowed_stores():
    """Names that may be opened, or [] meaning unrestricted."""
    allowed = _stores().get("allowed")
    return list(allowed) if isinstance(allowed, list) else []


def batch_stores():
    """Stores run_batch.py should walk, in order.

    Defaults to the whole allowlist. Returns [] when the allowlist is empty -
    i.e. unrestricted - because "walk every store on the account" is never an
    intended batch; callers must treat [] as "nothing configured" and refuse
    rather than iterating 93 stores.
    """
    batch = _stores().get("batch")
    if isinstance(batch, list) and batch:
        return [require_allowed(n) for n in batch]
    return allowed_stores()


def is_allowed(name):
    allowed = allowed_stores()
    return True if not allowed else name in allowed


def require_allowed(name):
    """Raise unless `name` may be opened. Call before touching the client.

    Failing here rather than at startBrowser means an unintended store is
    rejected locally, without a request ever reaching Ziniao.
    """
    if is_allowed(name):
        return name
    raise StoreNotAllowed(
        "store %r is not in config.json -> stores.allowed (%s). "
        "Add it there if you meant to open it."
        % (name, ", ".join(allowed_stores()) or "empty"))


def describe():
    allowed = allowed_stores()
    return "default=%s allowed=%s" % (
        _stores().get("default") or "(未配置)",
        ", ".join(allowed) if allowed else "(不限制)")
