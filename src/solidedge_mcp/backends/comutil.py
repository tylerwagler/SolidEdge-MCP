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


def profile_origin_element(profile: Any) -> tuple[Any, int] | None:
    """The 2D element and keypoint that place ``profile`` for SweptSurfaces.Add.

    Unlike Models.AddSweptProtrusion, which pairs each cross-section with a
    coordinate array, SweptSurfaces.Add declares Origins as
    SAFEARRAY(VT_DISPATCH) and OriginRefs as a KeyPointType: a sketch element
    and which of its keypoints is the origin. Verified on Solid Edge 2026: a
    circle with igKeyPointCenter builds the surface, while None, the profile
    object, and the circle with igKeyPointStart all fail with E_FAIL. Lines
    and arcs are placed by their start point, as profile_origin does.

    Returns None when the profile holds none of the four element kinds.
    """
    from .constants import KeyPointTypeConstants

    for collection, keypoint in (
        ("Lines2d", KeyPointTypeConstants.igKeyPointStart),
        ("Arcs2d", KeyPointTypeConstants.igKeyPointStart),
        ("Circles2d", KeyPointTypeConstants.igKeyPointCenter),
        ("Ellipses2d", KeyPointTypeConstants.igKeyPointCenter),
    ):
        try:
            items = getattr(profile, collection)
            if items.Count:
                return items.Item(1), keypoint
        except Exception:  # noqa: BLE001 - try the next kind of geometry
            continue
    return None


def profile_origin(profile: Any) -> tuple[float, float]:
    """A point on ``profile`` that Solid Edge accepts as its loft/sweep origin.

    The Origins array is how a loft or sweep pairs its cross-sections up, and
    each entry has to be a point that actually lies on its own section.
    Verified on Solid Edge 2026: two rectangles lofted with their real corner
    points build a 6-faced solid, and the same call with (0, 0) builds nothing
    -- no error, no geometry, just a silent no-op.

    A start point serves for lines and arcs, a centre for circles. Circles
    centred on the sketch origin work either way, which is how a hardcoded
    (0, 0) survived every earlier sweep: their profiles were centred circles.
    """
    for collection, getter in (
        ("Lines2d", "GetStartPoint"),
        ("Arcs2d", "GetStartPoint"),
        ("Circles2d", "GetCenterPoint"),
        ("Ellipses2d", "GetCenterPoint"),
    ):
        try:
            items = getattr(profile, collection)
            if items.Count:
                point = getattr(items.Item(1), getter)()
                return float(point[0]), float(point[1])
        except Exception:  # noqa: BLE001 - try the next kind of geometry
            continue
    return 0.0, 0.0
