"""Small helpers for talking to late-bound COM proxies."""

from __future__ import annotations

from typing import Any


def com_get(obj: Any, member: str, default: Any = None) -> Any:
    """Read a COM property, returning ``default`` if it is missing or raises.

    Preferred over ``hasattr(obj, member)`` followed by a read. On a late-bound
    proxy that probe is a separate ``GetIDsOfNames`` round trip whose failure
    mode is version dependent, and it reports False for a member that exists
    but whose getter raises, which turns a real error into a misleading one.

    Reading once and catching is both cheaper and honest. Pair it with
    ``tests/unit/test_com_members.py``, which checks that the member name is
    real, so a typo fails the suite instead of silently returning ``default``.
    """
    try:
        return getattr(obj, member)
    except Exception:
        return default
