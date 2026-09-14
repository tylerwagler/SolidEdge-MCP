"""Small helpers for talking to late-bound COM proxies."""

from __future__ import annotations

from typing import Any

from .constants import DOCUMENT_TYPE_NAMES, FEATURE_TYPE_NAMES


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


def describe_feature_type(raw: Any) -> dict[str, Any]:
    """A feature's type as both the raw value and a readable name.

    Feature.Type is a FeatureTypeConstants value. 462094706 on its own tells
    a caller nothing; "extruded protrusion" does.
    """
    return _describe(raw, FEATURE_TYPE_NAMES)


def describe_document_type(raw: Any) -> dict[str, Any]:
    """A document's type as both the raw value and a readable name."""
    return _describe(raw, DOCUMENT_TYPE_NAMES)


def _describe(raw: Any, names: dict[int, str]) -> dict[str, Any]:
    """Name a COM enum value, keeping the number when there is no name."""
    if not isinstance(raw, int) or isinstance(raw, bool):
        return {"type": "Unknown"}
    name = names.get(raw)
    if name is None:
        return {"type": raw}
    return {"type": name, "type_code": raw}
