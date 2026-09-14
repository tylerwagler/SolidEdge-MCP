"""Small helpers for talking to late-bound COM proxies."""

from __future__ import annotations

import contextlib
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


#: Styles this server creates and is therefore free to modify. A style Solid
#: Edge ships may be shared by every body and face using it, so writing to one
#: would change the whole document.
OWNED_STYLE_PREFIX = "MCP "


def face_style_named(doc: Any, name: str) -> tuple[Any, dict[str, Any] | None]:
    """Fetch or create a FaceStyle by name on this document.

    ``FaceStyles.Item`` raises rather than returning None for a name that is
    not there, so the lookup has to be guarded.

    Returns ``(style, error_dict)``; the error is set only when the document
    has no FaceStyles collection at all.
    """
    styles = com_get(doc, "FaceStyles")
    if styles is None:
        return None, {
            "error": "This document has no FaceStyles collection, so appearance cannot be set."
        }
    style = None
    with contextlib.suppress(Exception):
        style = styles.Item(name)
    if style is None:
        style = styles.Add(name, "")
    return style, None


def owned_style_for(doc: Any, holder: Any, label: str) -> tuple[Any, dict[str, Any] | None]:
    """A FaceStyle this server owns, assigned to ``holder``.

    Colour, opacity, reflectivity and texture are all properties of one
    ``FaceStyle``. Writing each to a style of its own means the last call wins
    and the others vanish, so they share one; and because the style a body or
    face already carries may be a stock Solid Edge style shared across the
    document, that one is never written to. The holder gets its own, seeded
    from whatever it had.

    ``holder`` is anything with a ``Style`` property: a ``Body`` or a ``Face``.
    ``label`` distinguishes one holder's style from another's.
    """
    current = com_get(holder, "Style")
    current_name = com_get(current, "StyleName", "") or ""
    if current_name.startswith(OWNED_STYLE_PREFIX):
        return current, None

    style, err = face_style_named(doc, f"{OWNED_STYLE_PREFIX}{label}")
    if err:
        return None, err

    # Carry over what the holder already looked like, so this reads as a change
    # to one property rather than a reset of all of them.
    if current is not None:
        with contextlib.suppress(Exception):
            style.SetDiffuse(*current.GetDiffuse()[:3])
        for prop in ("Opacity", "Reflectivity"):
            with contextlib.suppress(Exception):
                setattr(style, prop, getattr(current, prop))

    holder.Style = style
    return style, None
