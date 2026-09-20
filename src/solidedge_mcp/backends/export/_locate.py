"""Find the draft element a sheet coordinate points at.

Every dimension API in ``fwksupp.tlb`` takes the *object* being dimensioned:
``Dimensions.AddRadius(Object)``, ``AddDistanceBetweenObjects(Object1, x1, y1,
z1, keyPoint1, Object2, ...)``, ``AddAngleBetweenObjects(...)``. None of them
accepts bare coordinates. Several tools here were written as though a
coordinate-only overload existed -- ``AddRadial``, ``AddDiameter``,
``AddOrdinate``, ``AddDistanceBetweenPoints`` -- and none of those names is in
any Solid Edge type library, so those calls always raised.

Rather than declare a coordinate-driven dimension impossible, this module does
what an operator does with a mouse: it looks at the point and picks the
nearest element. That makes the object-based APIs reachable from the
coordinates an MCP caller actually has.

Geometry lives in two places. Sheet-level draft geometry (``sheet.Lines2d``
and friends) is already in sheet coordinates. Geometry derived from a model
lives on each drawing view (``view.DVLines2d`` and friends) in *view*
coordinates, which ``DrawingView.ViewToSheet`` converts. Verified against
Solid Edge 2026: a view placed at (0.2, 0.15) reports its first line running
from (-0.04, 0.01) to (0.04, 0.01), and ``ViewToSheet(0, 0)`` returns
(0.2, 0.15).
"""

from __future__ import annotations

import contextlib
import math
from collections.abc import Callable, Iterator
from dataclasses import dataclass
from typing import Any

from ..comutil import com_get

#: Sheet-level collections, mapped to the kind of curve they hold.
_SHEET_COLLECTIONS: dict[str, str] = {
    "Lines2d": "line",
    "Circles2d": "circle",
    "Arcs2d": "arc",
    "Ellipses2d": "curve",
    "EllipticalArcs2d": "curve",
    "LineStrings2d": "curve",
    "BsplineCurves2d": "curve",
    "Points2d": "point",
}

#: The same geometry as projected into a drawing view, in view coordinates.
_VIEW_COLLECTIONS: dict[str, str] = {
    "DVLines2d": "line",
    "DVCircles2d": "circle",
    "DVArcs2d": "arc",
    "DVEllipses2d": "curve",
    "DVEllipticalArcs2d": "curve",
    "DVLineStrings2d": "curve",
    "DVBSplineCurves2d": "curve",
    "DVPoints2d": "point",
}

#: Stop walking after this many elements. A busy sheet would otherwise cost
#: thousands of cross-process COM calls for a single dimension.
_SCAN_LIMIT = 4000

Mapper = Callable[[float, float], "tuple[float, float]"]


@dataclass(frozen=True)
class Hit:
    """An element found near a query point, with everything a caller needs."""

    obj: Any
    kind: str
    distance: float
    source: str
    #: Which collection it came from and where in it. pywin32 hands back a
    #: fresh wrapper for every Item() call, so ``hit_a.obj is hit_b.obj`` is
    #: False even for the same element. This triple is the real identity.
    collection: str = ""
    index: int = 0
    center: tuple[float, float] | None = None
    radius: float | None = None

    @property
    def identity(self) -> tuple[str, str, int]:
        """What makes two hits the same drawing element."""
        return (self.source, self.collection, self.index)

    def describe(self) -> dict[str, Any]:
        """A JSON-safe summary, for reporting what a dimension attached to."""
        info: dict[str, Any] = {
            "kind": self.kind,
            "source": self.source,
            "index": self.index - 1,  # COM is 1-indexed; tools are 0-indexed
            "distance": round(self.distance, 9),
        }
        if self.center is not None:
            info["center"] = [self.center[0], self.center[1]]
        if self.radius is not None:
            info["radius"] = self.radius
        return info


def _point(obj: Any, member: str) -> tuple[float, float] | None:
    """Call a pure-[out] accessor such as GetStartPoint and return (x, y)."""
    try:
        value = getattr(obj, member)()
    except Exception:
        return None
    try:
        return (float(value[0]), float(value[1]))
    except Exception:
        return None


def _distance_to_segment(
    px: float, py: float, a: tuple[float, float], b: tuple[float, float]
) -> float:
    ax, ay = a
    bx, by = b
    dx, dy = bx - ax, by - ay
    span = dx * dx + dy * dy
    if span == 0.0:
        return math.hypot(px - ax, py - ay)
    t = max(0.0, min(1.0, ((px - ax) * dx + (py - ay) * dy) / span))
    return math.hypot(px - (ax + t * dx), py - (ay + t * dy))


def _keypoint_distance(obj: Any, px: float, py: float, to_sheet: Mapper) -> float | None:
    """Fall back to the element's own keypoints for shapes we cannot solve.

    Every 2D element exposes ``KeyPointCount`` and ``GetKeyPoint(Index)``, so
    a spline or an ellipse still gets a usable, if coarser, distance.
    """
    count = com_get(obj, "KeyPointCount", 0) or 0
    best: float | None = None
    for index in range(int(count)):
        try:
            keypoint = obj.GetKeyPoint(index)
            point = to_sheet(float(keypoint[0]), float(keypoint[1]))
        except Exception:
            continue
        gap = math.hypot(px - point[0], py - point[1])
        if best is None or gap < best:
            best = gap
    return best


def _measure(
    obj: Any,
    kind: str,
    px: float,
    py: float,
    to_sheet: Mapper,
    source: str,
    collection: str = "",
    index: int = 0,
) -> Hit | None:
    """Distance from (px, py) to one element, in sheet coordinates."""
    tag: dict[str, Any] = {"collection": collection, "index": index}
    if kind == "line":
        start = _point(obj, "GetStartPoint")
        end = _point(obj, "GetEndPoint")
        if start is None or end is None:
            return None
        gap = _distance_to_segment(px, py, to_sheet(*start), to_sheet(*end))
        return Hit(obj=obj, kind=kind, distance=gap, source=source, **tag)

    if kind in ("circle", "arc"):
        centre = _point(obj, "GetCenterPoint")
        radius = com_get(obj, "Radius")
        if centre is None or radius is None:
            return None
        sheet_centre = to_sheet(*centre)
        # Distance to the curve, not to the centre: a click lands on the ring.
        gap = abs(math.hypot(px - sheet_centre[0], py - sheet_centre[1]) - float(radius))
        return Hit(
            obj=obj,
            kind=kind,
            distance=gap,
            source=source,
            center=sheet_centre,
            radius=float(radius),
            **tag,
        )

    keypoint_gap = _keypoint_distance(obj, px, py, to_sheet)
    if keypoint_gap is None:
        return None
    return Hit(obj=obj, kind=kind, distance=keypoint_gap, source=source, **tag)


def _identity(x: float, y: float) -> tuple[float, float]:
    return (x, y)


def _view_mapper(view: Any) -> Mapper:
    """A view-to-sheet converter, falling back to the view origin."""

    def convert(x: float, y: float) -> tuple[float, float]:
        try:
            sheet_point = view.ViewToSheet(x, y)
            return (float(sheet_point[0]), float(sheet_point[1]))
        except Exception:
            origin = _point(view, "GetOrigin") or (0.0, 0.0)
            return (origin[0] + x, origin[1] + y)

    return convert


def _sources(sheet: Any) -> Iterator[tuple[Any, str, Mapper, str, str]]:
    """Yield (collection, kind, to_sheet, source, name) for the whole sheet."""
    for name, kind in _SHEET_COLLECTIONS.items():
        collection = com_get(sheet, name)
        if collection is not None:
            yield collection, kind, _identity, "sheet", name

    views = com_get(sheet, "DrawingViews")
    count = int(com_get(views, "Count", 0) or 0)
    for index in range(1, count + 1):
        try:
            view = views.Item(index)
        except Exception:
            continue
        convert = _view_mapper(view)
        label = f"view {index - 1}"
        for name, kind in _VIEW_COLLECTIONS.items():
            collection = com_get(view, name)
            if collection is not None:
                yield collection, kind, convert, label, name


def nearest_element(
    sheet: Any,
    x: float,
    y: float,
    kinds: tuple[str, ...] | None = None,
    tolerance: float | None = None,
) -> Hit | None:
    """The draft element closest to (x, y) in sheet coordinates.

    Args:
        sheet: The active draft Sheet.
        x: Query point X, in meters of sheet space.
        y: Query point Y, in meters of sheet space.
        kinds: Restrict to these kinds ('line', 'circle', 'arc', 'curve',
            'point'). None accepts any.
        tolerance: Reject anything further away than this, in meters. None
            accepts the nearest element however far it is, which is what a
            caller naming an approximate position usually wants.

    Returns:
        The closest Hit, or None when the sheet holds nothing eligible.
    """
    best: Hit | None = None
    scanned = 0
    for collection, kind, to_sheet, source, name in _sources(sheet):
        if kinds is not None and kind not in kinds:
            continue
        count = int(com_get(collection, "Count", 0) or 0)
        for index in range(1, count + 1):
            if scanned >= _SCAN_LIMIT:
                break
            scanned += 1
            try:
                obj = collection.Item(index)
            except Exception:
                continue
            hit = _measure(obj, kind, x, y, to_sheet, source, collection=name, index=index)
            if hit is None:
                continue
            if best is None or hit.distance < best.distance:
                best = hit
    if best is None:
        return None
    if tolerance is not None and best.distance > tolerance:
        return None
    return best


def no_element_error(x: float, y: float, what: str) -> dict[str, Any]:
    """The error to return when nothing on the sheet can carry a dimension."""
    return {
        "error": (
            f"No {what} found near ({x}, {y}) on the active sheet. Solid Edge "
            f"dimensions attach to a drawing element, so the sheet needs "
            f"geometry there: draft 2D geometry, or a drawing view of a model."
        ),
        "point": [x, y],
    }


def element_count(sheet: Any) -> int:
    """How many elements a locate would consider. Used in error messages."""
    total = 0
    with contextlib.suppress(Exception):
        for collection, _kind, _to_sheet, _source, _name in _sources(sheet):
            total += int(com_get(collection, "Count", 0) or 0)
    return total
