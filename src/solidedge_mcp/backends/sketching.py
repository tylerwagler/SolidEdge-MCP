"""
Solid Edge Sketching Operations

Handles creating and manipulating 2D sketches.
"""

import contextlib
import math
from typing import Any

from solidedge_mcp.backends.errors import describe_exception, error_result

from .comutil import com_get
from .constants import FaceQueryConstants, ProfileValidationConstants
from .features._base import verifies_collection_growth
from .logging import get_logger

_logger = get_logger(__name__)


def ref_planes_of(doc: Any) -> Any:
    """The document's base reference planes, whatever it calls them.

    A part, sheet metal or weldment document holds them in ``RefPlanes``. An
    assembly has no such member at all: its three are ``AsmRefPlanes``, in the
    same 1=Top/XY, 2=Right/YZ, 3=Front/XZ order. Reaching only for RefPlanes
    made every sketch in an assembly fail with a bare ``<unknown>.RefPlanes``,
    and since each assembly-level creator consumes an accumulated profile,
    all six of them could only ever answer "No profiles available" -- there
    was no way to give them one. Verified on Solid Edge 2026, where
    AsmRefPlanes with ProfileSets builds an assembly profile normally.

    Returns None when the document has neither, which is what a draft is.
    """
    planes = com_get(doc, "RefPlanes")
    if planes is not None:
        return planes
    return com_get(doc, "AsmRefPlanes")


#: A document that holds neither RefPlanes nor AsmRefPlanes cannot be
#: sketched on. A draft is the case: its geometry lives on Sheets.
_NO_REF_PLANES = (
    "This document has no reference planes to sketch on. Parts, sheet metal "
    "and assemblies do; a draft does not -- draw on its sheet instead."
)


def _corner_points(
    line1: Any, line2: Any
) -> tuple[tuple[float, float], tuple[float, float]] | None:
    """Where two lines meet, and a point just inside that corner.

    AddAsFillet and AddAsChamfer take a point saying which corner to work on.
    The fillet accepts the vertex itself; the chamfer rejects it with
    E_INVALIDARG and wants a point nudged into the corner, verified against
    Solid Edge 2026. The inward point steps from the vertex toward the middle
    of the two far ends, which needs no knowledge of the shape.

    Returns (corner, inward), or None when the lines do not meet.
    """
    try:
        ends1 = (line1.GetStartPoint(), line1.GetEndPoint())
        ends2 = (line2.GetStartPoint(), line2.GetEndPoint())
    except Exception:
        return None

    best = None
    for a in ends1:
        for b in ends2:
            gap = abs(a[0] - b[0]) + abs(a[1] - b[1])
            if best is None or gap < best[0]:
                best = (gap, a, b)
    if best is None or best[0] > 1e-6:
        return None

    _, near1, _ = best
    corner = (float(near1[0]), float(near1[1]))
    far1 = max(ends1, key=lambda e: abs(e[0] - corner[0]) + abs(e[1] - corner[1]))
    far2 = max(ends2, key=lambda e: abs(e[0] - corner[0]) + abs(e[1] - corner[1]))
    toward = ((float(far1[0]) + float(far2[0])) / 2, (float(far1[1]) + float(far2[1])) / 2)

    dx, dy = toward[0] - corner[0], toward[1] - corner[1]
    length = math.hypot(dx, dy)
    if length < 1e-12:
        return corner, corner
    step = min(1e-3, length / 4)
    inward = (corner[0] + dx / length * step, corner[1] + dy / length * step)
    return corner, inward


#: Orientation argument of Ellipses2d.AddByCenter: 1 sweeps counterclockwise.
CURVE_COUNTERCLOCKWISE = 1

#: CleanProfileOptions default for CleanGeometry2d: no special handling.
_CLEAN_PROFILE_DEFAULT = 0

#: Profile collections that together make up the sketch's 2D geometry.
_GEOMETRY_2D_COLLECTIONS = (
    "Lines2d",
    "Arcs2d",
    "Circles2d",
    "Ellipses2d",
    "EllipticalArcs2d",
    "BSplineCurves2d",
    "Conics2d",
)


def _r8_array(size: int) -> Any:
    """Buffer for a COM ``SAFEARRAY(VT_R8)*`` ``[in, out]`` parameter.

    A plain list, not a VARIANT: see the note on
    ``solidedge_mcp.backends.query._base.r8_array``.
    """
    return [0.0] * size


def _dispatch_array(items: Any) -> Any:
    """Wrap a sequence of COM objects as a ``SAFEARRAY(VT_DISPATCH)``."""
    return list(items)


#: Every 2D collection a profile can hold, for counting and selecting.
_PROFILE_COLLECTIONS = (
    "Lines2d",
    "Circles2d",
    "Arcs2d",
    "Ellipses2d",
    "EllipticalArcs2d",
    "BSplineCurves2d",
    "Conics2d",
)


def _xy(obj: Any, member: str) -> tuple[float, float] | None:
    """Read a pure-[out] 2D accessor such as GetStartPoint as (x, y)."""
    try:
        point = getattr(obj, member)()
        return (float(point[0]), float(point[1]))
    except Exception:
        return None


def _mirror_across_x(x: float, y: float) -> tuple[float, float]:
    return (x, -y)


def _mirror_across_y(x: float, y: float) -> tuple[float, float]:
    return (-x, y)


def _element_count(profile: Any) -> int:
    """How much geometry a profile holds, across every 2D collection."""
    total = 0
    for name in _PROFILE_COLLECTIONS:
        collection = com_get(profile, name)
        total += int(com_get(collection, "Count", 0) or 0)
    return total


def _select_all(profile: Any, doc: Any = None) -> int:
    """Select every element in a profile. Offset2d acts on the selection.

    The select set has to be emptied first. It is document-wide and survives
    between calls, so a stale selection from an earlier sketch silently makes
    Offset2d do nothing at all -- verified on Solid Edge 2026, where offsetting
    a rectangle works on a clean select set and produces nothing on a dirty
    one.
    """
    select_set = com_get(doc, "SelectSet") if doc is not None else None
    if select_set is not None:
        with contextlib.suppress(Exception):
            select_set.RemoveAll()

    selected = 0
    for name in _PROFILE_COLLECTIONS:
        collection = com_get(profile, name)
        count = int(com_get(collection, "Count", 0) or 0)
        for i in range(1, count + 1):
            try:
                collection.Item(i).Select()
                selected += 1
            except Exception:
                continue
    return selected


def _offset_side_point(profile: Any, distance: float) -> tuple[float, float] | None:
    """A point saying which side to offset towards.

    Offset2d wants a coordinate, not a sign. Taking the centre of the
    geometry and stepping out past its extent puts the point outside for a
    positive distance and at the centre for a negative one.
    """
    xs: list[float] = []
    ys: list[float] = []
    for name in ("Lines2d", "Arcs2d", "Circles2d"):
        collection = com_get(profile, name)
        count = int(com_get(collection, "Count", 0) or 0)
        for i in range(1, count + 1):
            element = collection.Item(i)
            for member in ("GetStartPoint", "GetEndPoint", "GetCenterPoint"):
                point = _xy(element, member)
                if point is not None:
                    xs.append(point[0])
                    ys.append(point[1])
    if not xs:
        return None

    centre_x = (min(xs) + max(xs)) / 2
    centre_y = (min(ys) + max(ys)) / 2
    if distance < 0:
        return (centre_x, centre_y)
    reach = max(max(xs) - min(xs), max(ys) - min(ys)) or abs(distance)
    return (centre_x, centre_y + reach)


#: How to name the elements a constraint applies to. Both spellings work; the
#: element1_*/element2_* one used to be read only by the keypoint branch, so
#: naming elements that way sent an empty list and the call asked for elements
#: it had just been given.
_HOW_TO_NAME = (
    ' Name them with elements=[["line", 1], ["line", 2]] or with '
    "element1_type/element1_index and element2_type/element2_index; indices "
    "are 1-based."
)


class SketchManager:
    """Manages sketch creation and 2D geometry"""

    def __init__(self, document_manager: Any) -> None:
        self.doc_manager = document_manager
        self.active_sketch: Any | None = None
        self.active_profile: Any | None = None
        self.active_plane_index: int | None = None  # 1-based RefPlanes index
        self.active_refaxis: Any | None = None  # Reference axis for revolve operations
        self.accumulated_profiles: list[Any] = []  # For loft/sweep multi-profile operations
        self._last_document_handle: Any | None = None  # Track which document we're working with
        #: The active 3D sketch and every Line3D drawn into it, in order. A
        #: structural frame takes them by index.
        self.active_sketch3d: Any = None
        self.lines_3d: list[Any] = []

    def clear_state(self) -> None:
        """Clear all sketch state. Call this when switching documents."""
        _logger.debug("Clearing sketch state")
        self.active_sketch = None
        self.active_profile = None
        self.active_plane_index = None
        self.active_refaxis = None
        self.accumulated_profiles.clear()
        self.active_sketch3d = None
        self.lines_3d.clear()
        self._last_document_handle = None

    @verifies_collection_growth("ProfileSets")
    def create_sketch(self, plane: str = "Top") -> dict[str, Any]:
        """
        Create a new sketch on a reference plane.

        Args:
            plane: Plane name - 'Top', 'Front', 'Right', 'XY', 'XZ', 'YZ'

        Returns:
            Dict with status and sketch info
        """
        try:
            doc = self.doc_manager.get_active_document()

            # Get reference planes
            ref_planes = ref_planes_of(doc)
            if ref_planes is None:
                return {"error": _NO_REF_PLANES}

            # Map plane names to indices
            plane_map = {
                "Top": 1,  # XY plane (top view)
                "Right": 2,  # YZ plane (right view)
                "Front": 3,  # XZ plane (front view)
                "XY": 1,
                "YZ": 2,
                "XZ": 3,
            }

            plane_index = plane_map.get(plane)
            if plane_index is None:
                return {
                    "error": f"Invalid plane: {plane}. "
                    "Use 'Top', 'Front', 'Right', "
                    "'XY', 'XZ', or 'YZ'"
                }

            ref_plane = ref_planes.Item(plane_index)

            # Get ProfileSets collection
            profile_sets = doc.ProfileSets

            # Add a new profile set
            profile_set = profile_sets.Add()

            # Create a profile on the reference plane
            profiles = profile_set.Profiles
            profile = profiles.Add(ref_plane)

            self.active_sketch = profile_set
            self.active_profile = profile
            self.active_plane_index = plane_index
            self.active_refaxis = None  # Clear any previous axis

            _logger.info(f"Sketch created on plane: {plane}")
            return {
                "status": "created",
                "plane": plane,
                "sketch_id": com_get(profile_set, "Name", "sketch"),
            }
        except Exception as e:
            _logger.error(f"Failed to create sketch on plane {plane}: {e}")
            return error_result(e)

    @verifies_collection_growth("ProfileSets")
    def create_sketch_on_plane_index(self, plane_index: int) -> dict[str, Any]:
        """
        Create a new sketch on a reference plane by its 1-based index.

        Useful for sketching on user-created offset planes (index > 3).

        Args:
            plane_index: 1-based index of the reference plane

        Returns:
            Dict with status and sketch info
        """
        try:
            doc = self.doc_manager.get_active_document()
            ref_planes = ref_planes_of(doc)
            if ref_planes is None:
                return {"error": _NO_REF_PLANES}

            if plane_index < 1 or plane_index > ref_planes.Count:
                return {"error": f"Invalid plane index: {plane_index}. Count: {ref_planes.Count}"}

            ref_plane = ref_planes.Item(plane_index)

            profile_sets = doc.ProfileSets
            profile_set = profile_sets.Add()
            profiles = profile_set.Profiles
            profile = profiles.Add(ref_plane)

            self.active_sketch = profile_set
            self.active_profile = profile
            self.active_plane_index = plane_index
            self.active_refaxis = None

            return {
                "status": "created",
                "plane_index": plane_index,
                "sketch_id": com_get(profile_set, "Name", "sketch"),
            }
        except Exception as e:
            return error_result(e)

    def draw_line(self, x1: float, y1: float, x2: float, y2: float) -> dict[str, Any]:
        """Draw a line in the active sketch"""
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            # Get Lines2d collection
            lines = self.active_profile.Lines2d

            # Add line
            lines.AddBy2Points(x1, y1, x2, y2)

            return {"status": "created", "type": "line", "start": [x1, y1], "end": [x2, y2]}
        except Exception as e:
            return error_result(e)

    def draw_circle(self, center_x: float, center_y: float, radius: float) -> dict[str, Any]:
        """Draw a circle in the active sketch"""
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            # Get Circles2d collection
            circles = self.active_profile.Circles2d

            # Add circle by center and radius
            circles.AddByCenterRadius(center_x, center_y, radius)

            return {
                "status": "created",
                "type": "circle",
                "center": [center_x, center_y],
                "radius": radius,
            }
        except Exception as e:
            return error_result(e)

    def draw_rectangle(self, x1: float, y1: float, x2: float, y2: float) -> dict[str, Any]:
        """Draw a rectangle in the active sketch"""
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            # Rectangles are typically drawn as 4 lines
            lines = self.active_profile.Lines2d

            # Draw 4 sides of rectangle
            lines.AddBy2Points(x1, y1, x2, y1)  # Bottom
            lines.AddBy2Points(x2, y1, x2, y2)  # Right
            lines.AddBy2Points(x2, y2, x1, y2)  # Top
            lines.AddBy2Points(x1, y2, x1, y1)  # Left

            return {
                "status": "created",
                "type": "rectangle",
                "corner1": [x1, y1],
                "corner2": [x2, y2],
                "lines": 4,
            }
        except Exception as e:
            return error_result(e)

    def draw_arc(
        self, center_x: float, center_y: float, radius: float, start_angle: float, end_angle: float
    ) -> dict[str, Any]:
        """
        Draw an arc in the active sketch.

        Args:
            center_x, center_y: Arc center coordinates
            radius: Arc radius
            start_angle, end_angle: Angles in degrees (0 = right, 90 = up)
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            # Convert angles to radians
            start_rad = math.radians(start_angle)
            end_rad = math.radians(end_angle)

            # Calculate start and end points
            start_x = center_x + radius * math.cos(start_rad)
            start_y = center_y + radius * math.sin(start_rad)
            end_x = center_x + radius * math.cos(end_rad)
            end_y = center_y + radius * math.sin(end_rad)

            # Get Arcs2d collection
            arcs = self.active_profile.Arcs2d

            # Add arc by center and endpoints
            arcs.AddByCenterStartEnd(center_x, center_y, start_x, start_y, end_x, end_y)

            return {
                "status": "created",
                "type": "arc",
                "center": [center_x, center_y],
                "radius": radius,
                "start_angle": start_angle,
                "end_angle": end_angle,
            }
        except Exception as e:
            return error_result(e)

    def draw_polygon(
        self, center_x: float, center_y: float, radius: float, sides: int
    ) -> dict[str, Any]:
        """Draw a regular polygon"""
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            if sides < 3:
                return {"error": "Polygon must have at least 3 sides"}

            lines = self.active_profile.Lines2d

            # Calculate vertices
            angle_step = 2 * math.pi / sides
            points = []

            for i in range(sides):
                angle = i * angle_step
                x = center_x + radius * math.cos(angle)
                y = center_y + radius * math.sin(angle)
                points.append((x, y))

            # Draw lines connecting vertices
            for i in range(sides):
                x1, y1 = points[i]
                x2, y2 = points[(i + 1) % sides]
                lines.AddBy2Points(x1, y1, x2, y2)

            return {
                "status": "created",
                "type": "polygon",
                "center": [center_x, center_y],
                "radius": radius,
                "sides": sides,
            }
        except Exception as e:
            return error_result(e)

    def draw_ellipse(
        self,
        center_x: float,
        center_y: float,
        major_radius: float,
        minor_radius: float,
        angle: float = 0.0,
    ) -> dict[str, Any]:
        """
        Draw an ellipse in the active sketch.

        Args:
            center_x, center_y: Ellipse center coordinates
            major_radius: Major axis radius
            minor_radius: Minor axis radius
            angle: Rotation angle in degrees (default 0)
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            # Get Ellipses2d collection
            ellipses = self.active_profile.Ellipses2d

            if major_radius <= 0 or minor_radius <= 0:
                return {"error": "major_radius and minor_radius must be positive"}

            # fwksupp.tlb: Ellipses2d.AddByCenter(xCenter, yCenter, xMajor,
            # yMajor, Ratio, Orientation). xMajor/yMajor is a POINT at the end
            # of the major axis, Ratio is minor/major, and Orientation is the
            # sweep direction. Passing two radii and an axis vector instead
            # raised E_FAIL, so no ellipse could ever be drawn.
            angle_rad = math.radians(angle)
            major_x = center_x + major_radius * math.cos(angle_rad)
            major_y = center_y + major_radius * math.sin(angle_rad)
            ratio = minor_radius / major_radius

            ellipses.AddByCenter(
                center_x,
                center_y,
                major_x,
                major_y,
                ratio,
                CURVE_COUNTERCLOCKWISE,
            )

            return {
                "status": "created",
                "type": "ellipse",
                "center": [center_x, center_y],
                "major_radius": major_radius,
                "minor_radius": minor_radius,
                "angle": angle,
            }
        except Exception as e:
            return error_result(e)

    def draw_spline(self, points: list[list[float]]) -> dict[str, Any]:
        """
        Draw a B-spline curve through a list of points.

        Args:
            points: List of [x, y] coordinate pairs

        Returns:
            Dict with status and spline info
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            if len(points) < 2:
                return {"error": "Spline requires at least 2 points"}

            # Get BSplineCurves2d collection
            splines = self.active_profile.BSplineCurves2d

            # Convert points list to flat array format expected by COM
            point_array = []
            for point in points:
                if len(point) != 2:
                    return {"error": f"Invalid point format: {point}. Expected [x, y]"}
                point_array.extend(point)

            # Add spline by points
            # AddByPoints takes positional args: Order, NumPoints, PointArray
            splines.AddByPoints(
                3,  # Order (cubic spline)
                len(points),  # NumPoints
                tuple(point_array),  # PointArray (flattened x,y,x,y,...)
            )

            return {
                "status": "created",
                "type": "spline",
                "points": points,
                "num_points": len(points),
            }
        except Exception as e:
            return error_result(e)

    def draw_arc_by_3_points(
        self,
        start_x: float,
        start_y: float,
        along_x: float,
        along_y: float,
        end_x: float,
        end_y: float,
    ) -> dict[str, Any]:
        """
        Draw an arc through three points: start, a point along it, and end.

        fwksupp.tlb Arcs2d offers AddByStartAlongEnd and AddByCenterStartEnd.
        There is no AddByStartCenterEnd, which is what this used to call, so
        the arc always raised. The middle point is a point the arc passes
        through, not the centre: feeding a mid point to a centre-based API
        fails whenever start and end are not equidistant from it.

        Args:
            start_x, start_y: Arc start point (meters)
            along_x, along_y: A point the arc passes through (meters)
            end_x, end_y: Arc end point (meters)

        Returns:
            Dict with status and arc info
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            arcs = self.active_profile.Arcs2d
            arcs.AddByStartAlongEnd(start_x, start_y, along_x, along_y, end_x, end_y)

            return {
                "status": "created",
                "type": "arc",
                "start": [start_x, start_y],
                "along": [along_x, along_y],
                "end": [end_x, end_y],
                "method": "start_along_end",
            }
        except Exception as e:
            return error_result(e)

    def draw_circle_by_2_points(self, x1: float, y1: float, x2: float, y2: float) -> dict[str, Any]:
        """
        Draw a circle defined by two diametrically opposite points.

        The two points define the diameter of the circle.

        Args:
            x1, y1: First point on circle (meters)
            x2, y2: Second point on circle, diametrically opposite (meters)

        Returns:
            Dict with status and circle info
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            center_x = (x1 + x2) / 2
            center_y = (y1 + y2) / 2
            radius = math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2) / 2

            if radius <= 0:
                return {"error": "The two points must differ; they define a diameter."}

            # fwksupp.tlb Circles2d exposes only AddByCenterRadius and
            # AddBy3Points. AddBy2Points does not exist, so derive the centre
            # and radius from the diameter endpoints ourselves.
            circles = self.active_profile.Circles2d
            circles.AddByCenterRadius(center_x, center_y, radius)

            return {
                "status": "created",
                "type": "circle",
                "center": [center_x, center_y],
                "radius": radius,
                "method": "2_points",
            }
        except Exception as e:
            return error_result(e)

    def draw_circle_by_3_points(
        self, x1: float, y1: float, x2: float, y2: float, x3: float, y3: float
    ) -> dict[str, Any]:
        """
        Draw a circle through three points.

        The circle passes through all three specified points.

        Args:
            x1, y1: First point on circle (meters)
            x2, y2: Second point on circle (meters)
            x3, y3: Third point on circle (meters)

        Returns:
            Dict with status and circle info
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            circles = self.active_profile.Circles2d
            circles.AddBy3Points(x1, y1, x2, y2, x3, y3)

            return {
                "status": "created",
                "type": "circle",
                "point1": [x1, y1],
                "point2": [x2, y2],
                "point3": [x3, y3],
                "method": "3_points",
            }
        except Exception as e:
            return error_result(e)

    def mirror_spline(
        self, axis_x1: float, axis_y1: float, axis_x2: float, axis_y2: float, copy: bool = True
    ) -> dict[str, Any]:
        """
        Mirror B-spline curves across a line defined by two points.

        Mirrors all B-spline curves in the active sketch across the
        specified axis line.

        Args:
            axis_x1, axis_y1: Start point of mirror axis (meters)
            axis_x2, axis_y2: End point of mirror axis (meters)
            copy: If True, create a mirrored copy. If False, move the original.

        Returns:
            Dict with status and count of mirrored splines
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            profile = self.active_profile
            splines = profile.BSplineCurves2d

            if splines.Count == 0:
                return {"error": "No B-spline curves to mirror"}

            mirror_count = 0
            for i in range(1, splines.Count + 1):
                try:
                    spline = splines.Item(i)
                    spline.Mirror(axis_x1, axis_y1, axis_x2, axis_y2, copy)
                    mirror_count += 1
                except Exception:
                    pass

            return {
                "status": "created",
                "type": "mirror_spline",
                "mirror_axis": [[axis_x1, axis_y1], [axis_x2, axis_y2]],
                "copy": copy,
                "mirrored_count": mirror_count,
            }
        except Exception as e:
            return error_result(e)

    def hide_profile(self, visible: bool = False) -> dict[str, Any]:
        """
        Show or hide the active sketch profile.

        Hiding a profile makes it invisible in the 3D view but it
        remains functional for feature operations.

        Args:
            visible: True to show, False to hide

        Returns:
            Dict with status
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            self.active_profile.Visible = visible

            return {"status": "updated", "visible": visible}
        except Exception as e:
            return error_result(e)

    def draw_point(self, x: float, y: float) -> dict[str, Any]:
        """
        Draw a construction point in the active sketch.

        Creates a point at the specified coordinates. Points are useful
        as reference geometry for constraints and as hole center locations.

        Args:
            x: X coordinate (meters)
            y: Y coordinate (meters)

        Returns:
            Dict with status and point info
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            # Use Holes2d.Add to place a point (standard sketch point method)
            try:
                holes = self.active_profile.Holes2d
                point = holes.Add(x, y)
                return {
                    "status": "created",
                    "type": "point",
                    "position": [x, y],
                    "method": "Holes2d",
                }
            except Exception:
                pass

            # Fallback: use a zero-radius circle as a point marker
            circles = self.active_profile.Circles2d
            point = circles.AddByCenterRadius(x, y, 0.0001)  # Very small circle
            self.active_profile.ToggleConstruction(point)
            return {
                "status": "created",
                "type": "point",
                "position": [x, y],
                "method": "construction_circle",
            }
        except Exception as e:
            return error_result(e)

    def set_axis_of_revolution(self, x1: float, y1: float, x2: float, y2: float) -> dict[str, Any]:
        """
        Draw an axis of revolution line in the active sketch for revolve operations.

        The axis line is drawn as a construction line and set as the revolution axis.
        This must be called before close_sketch() when preparing a revolve feature.

        Args:
            x1, y1: Start point of axis line (meters)
            x2, y2: End point of axis line (meters)

        Returns:
            Dict with status and axis info
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            # Draw the axis line
            lines = self.active_profile.Lines2d
            axis_line = lines.AddBy2Points(x1, y1, x2, y2)

            # Mark as construction geometry
            self.active_profile.ToggleConstruction(axis_line)

            # Set as axis of revolution
            self.active_refaxis = self.active_profile.SetAxisOfRevolution(axis_line)

            return {
                "status": "axis_set",
                "start": [x1, y1],
                "end": [x2, y2],
                "note": "Axis of revolution set. Close sketch and use create_revolve().",
            }
        except Exception as e:
            return error_result(e)

    def _get_sketch_element(self, element_type: str, index: int) -> Any:
        """
        Resolve a sketch element by type name and 1-based index.

        Args:
            element_type: 'line', 'circle', 'arc', 'ellipse', 'spline'
            index: 1-based index within that collection

        Returns:
            COM object for the sketch element

        Raises:
            ValueError: If type or index is invalid
        """
        profile = self.active_profile
        type_map = {
            "line": "Lines2d",
            "circle": "Circles2d",
            "arc": "Arcs2d",
            "ellipse": "Ellipses2d",
            "spline": "BSplineCurves2d",
        }
        collection_name = type_map.get(element_type.lower())
        if not collection_name:
            valid_types = ", ".join(type_map.keys())
            raise ValueError(f"Unknown element type: '{element_type}'. Use: {valid_types}")

        collection = getattr(profile, collection_name)
        if index < 1 or index > collection.Count:
            raise ValueError(
                f"Index {index} out of range for {element_type} (count: {collection.Count})"
            )

        return collection.Item(index)

    def add_constraint(
        self, constraint_type: str, elements: list[list[str | int]]
    ) -> dict[str, Any]:
        """
        Add a geometric constraint to sketch elements.

        Elements are specified as [type, index] pairs where type is
        'line', 'circle', 'arc', 'ellipse', or 'spline' and index is
        1-based within that collection.

        Args:
            constraint_type: 'Horizontal', 'Vertical', 'Parallel', 'Perpendicular',
                           'Equal', 'Concentric', 'Tangent'
            elements: List of [type, index] pairs, e.g. [["line", 1], ["line", 2]]

        Returns:
            Dict with status
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch"}

            relations = self.active_profile.Relations2d

            # Resolve element references
            objs = []
            for elem in elements:
                if not isinstance(elem, (list, tuple)) or len(elem) != 2:
                    return {"error": f"Each element must be [type, index], got: {elem}"}
                elem_type, elem_index = elem[0], elem[1]
                if not isinstance(elem_type, str) or not isinstance(elem_index, int):
                    return {"error": f"Element must be [str, int], got: {elem}"}
                objs.append(self._get_sketch_element(elem_type, elem_index))

            ct = constraint_type.lower()

            # Single-element constraints
            if ct == "horizontal":
                if len(objs) < 1:
                    return {"error": "Horizontal constraint needs 1 element." + _HOW_TO_NAME}
                relations.AddHorizontal(objs[0])
            elif ct == "vertical":
                if len(objs) < 1:
                    return {"error": "Vertical constraint needs 1 element." + _HOW_TO_NAME}
                relations.AddVertical(objs[0])
            # Two-element constraints
            elif ct == "parallel":
                if len(objs) < 2:
                    return {"error": "Parallel constraint needs 2 elements." + _HOW_TO_NAME}
                relations.AddParallel(objs[0], objs[1])
            elif ct == "perpendicular":
                if len(objs) < 2:
                    return {"error": "Perpendicular constraint needs 2 elements." + _HOW_TO_NAME}
                relations.AddPerpendicular(objs[0], objs[1])
            elif ct == "equal":
                if len(objs) < 2:
                    return {"error": "Equal constraint needs 2 elements." + _HOW_TO_NAME}
                relations.AddEqual(objs[0], objs[1])
            elif ct == "concentric":
                if len(objs) < 2:
                    return {"error": "Concentric constraint needs 2 elements." + _HOW_TO_NAME}
                relations.AddConcentric(objs[0], objs[1])
            elif ct == "tangent":
                if len(objs) < 2:
                    return {"error": "Tangent constraint needs 2 elements." + _HOW_TO_NAME}
                relations.AddTangent(objs[0], objs[1])
            else:
                return {
                    "error": f"Unknown constraint type: "
                    f"'{constraint_type}'. Use: "
                    "Horizontal, Vertical, Parallel, "
                    "Perpendicular, Equal, "
                    "Concentric, Tangent"
                }

            return {"status": "constraint_added", "type": constraint_type, "elements": elements}
        except ValueError as e:
            return {"error": str(e)}
        except Exception as e:
            return error_result(e)

    def add_keypoint_constraint(
        self,
        element1_type: str,
        element1_index: int,
        keypoint1: int,
        element2_type: str,
        element2_index: int,
        keypoint2: int,
    ) -> dict[str, Any]:
        """
        Add a keypoint constraint connecting two sketch elements at specific points.

        Keypoint indices are GEOMETRIC, 0-based: 0=start, 1=end, 2=midpoint for
        lines/arcs; 0=center for circles. (Verified live: AddKeypoint(line, 1,
        other, 0) welds line's end to other's start.) These are NOT the
        KeyPointType enum values igKeyPointStart=1/igKeyPointEnd=2 -- that enum
        belongs to other (extent/revolve) APIs and must not be used here.

        Args:
            element1_type: Type of first element ('line', 'circle', 'arc', etc.)
            element1_index: 1-based index of first element
            keypoint1: Keypoint index on first element (0=start, 1=end)
            element2_type: Type of second element
            element2_index: 1-based index of second element
            keypoint2: Keypoint index on second element (0=start, 1=end)

        Returns:
            Dict with status
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch"}

            obj1 = self._get_sketch_element(element1_type, element1_index)
            obj2 = self._get_sketch_element(element2_type, element2_index)

            relations = self.active_profile.Relations2d
            relations.AddKeypoint(obj1, keypoint1, obj2, keypoint2)

            return {
                "status": "constraint_added",
                "type": "Keypoint",
                "element1": [element1_type, element1_index, keypoint1],
                "element2": [element2_type, element2_index, keypoint2],
            }
        except ValueError as e:
            return {"error": str(e)}
        except Exception as e:
            return error_result(e)

    def close_sketch(self, closed: bool = True) -> dict[str, Any]:
        """Close/finish the active sketch and report whether it validated.

        Profile.End(flags) returns 0 on success and a negative status code
        when the profile fails validation (e.g. an open loop). We surface that
        code instead of swallowing it, so callers can tell BEFORE building a
        feature whether the profile actually forms a region.

        Args:
            closed: If True (default) the profile is validated as a CLOSED
                region (igProfileClosed). This is required for solid features
                and, crucially, welds the coincident endpoints of a polyline
                (e.g. a rectangle drawn as 4 separate lines) into a single
                region -- with the old igProfileDefault flag such polylines
                silently produced no geometry. Set False for intentionally
                open profiles (sweep paths, open surfaces).
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch to close"}

            # Choose validation flags. A revolve profile already implies closed
            # (igProfileForRevolve = igProfileClosed | igProfileRefAxisRequired).
            if self.active_refaxis is not None:
                end_flags = ProfileValidationConstants.igProfileForRevolve  # 17
            elif closed:
                end_flags = ProfileValidationConstants.igProfileClosed  # 1
            else:
                end_flags = ProfileValidationConstants.igProfileDefault  # 0

            # Validate the profile. End() returns a status: 0 = OK, < 0 = invalid.
            # Do NOT suppress -- the validation result is the whole point.
            try:
                validation_code = self.active_profile.End(end_flags)
            except Exception as e:
                _logger.error(f"Profile.End({end_flags}) raised: {e}")
                return error_result(e, context="Profile validation failed", end_flags=end_flags)

            # Queue for loft/sweep, which consume several profiles. Guard against
            # a second close_sketch() on the same profile: that used to queue the
            # profile twice and silently corrupt the next multi-profile feature.
            # Identity, not ==, because COM proxies do not compare reliably.
            already_queued = any(p is self.active_profile for p in self.accumulated_profiles)
            if not already_queued:
                self.accumulated_profiles.append(self.active_profile)

            sketch_id = "sketch"
            if self.active_sketch is not None:
                sketch_id = com_get(self.active_sketch, "Name", "sketch")
            result: dict[str, Any] = {
                "status": "closed",
                "validation_code": validation_code,
                "sketch_id": sketch_id,
                "has_revolution_axis": self.active_refaxis is not None,
                "accumulated_profiles": len(self.accumulated_profiles),
            }
            # End() returns 0 for a cleanly closed profile. A non-zero code is
            # NOT a reliable pass/fail signal: igProfileClosed auto-connects a
            # polyline's coincident endpoints and reports a non-zero code (e.g.
            # -103) even though the resulting region extrudes fine -- the same
            # code also appears for a genuinely open profile that builds nothing.
            # So we surface the code as a hint and tell callers to verify that
            # the downstream feature actually produced geometry.
            if isinstance(validation_code, int) and validation_code != 0:
                result["note"] = (
                    f"Profile.End returned {validation_code} (0 = clean close). "
                    f"Often benign (e.g. auto-connected polyline endpoints), but "
                    f"verify the next feature actually created geometry."
                )

            # NOTE: We keep active_profile valid after closing so it can be used
            # by feature operations (extrude, revolve, etc.). The profile object
            # remains valid even after End() is called.
            # Only clear it when a new sketch is created.

            _logger.info(
                f"Sketch closed (end_flags={end_flags}, code={validation_code}, "
                f"accumulated={len(self.accumulated_profiles)})"
            )
            return result
        except Exception as e:
            _logger.error(f"Failed to close sketch: {e}")
            return error_result(e)

    def get_sketch_info(self) -> dict[str, Any]:
        """
        Get information about the active sketch.

        Returns element counts for each geometry type in the active sketch.

        Returns:
            Dict with sketch geometry counts
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            profile = self.active_profile
            info: dict[str, Any] = {"status": "active"}

            # Count elements in each collection
            collections = {
                "lines": "Lines2d",
                "circles": "Circles2d",
                "arcs": "Arcs2d",
                "ellipses": "Ellipses2d",
                "splines": "BSplineCurves2d",
                # Points2d holds sketch points. Holes2d, which this counted, is
                # the hole-position collection, so a sketch full of points
                # reported none and one with hole positions reported points.
                "points": "Points2d",
                "hole_positions": "Holes2d",
            }

            total = 0
            for key, collection_name in collections.items():
                try:
                    coll = getattr(profile, collection_name)
                    count = coll.Count
                    info[key] = count
                    total += count
                except Exception:
                    info[key] = 0

            info["total_elements"] = total

            return info
        except Exception as e:
            return error_result(e)

    def draw_line_3d(
        self,
        x1: float,
        y1: float,
        z1: float,
        x2: float,
        y2: float,
        z2: float,
        new_sketch: bool = False,
    ) -> dict[str, Any]:
        """Draw a 3D sketch line, in meters, in the active part or assembly.

        ``Sketches3D.Add()`` opens a 3D sketch and ``Lines3D.Add(x1, y1, z1,
        x2, y2, z2)`` draws into it -- verified on Solid Edge 2026 in a part
        and an assembly, where a structural frame then ran along the line.
        Lines accumulate in one sketch until ``new_sketch`` starts another.
        """
        try:
            doc = self.doc_manager.get_active_document()
            sketches = com_get(doc, "Sketches3D")
            if sketches is None:
                return {
                    "error": (
                        "This document has no Sketches3D collection; 3D sketch lines "
                        "live in parts and assemblies."
                    )
                }
            if new_sketch or self.active_sketch3d is None:
                self.active_sketch3d = sketches.Add()
            lines = self.active_sketch3d.Lines3D
            before = com_get(lines, "Count")
            line = lines.Add(x1, y1, z1, x2, y2, z2)
            after = com_get(lines, "Count")
            if type(before) is int and type(after) is int and after != before + 1:
                return {
                    "error": "Lines3D.Add returned without adding a line to the 3D sketch.",
                    "lines_before": before,
                    "lines_after": after,
                }
            self.lines_3d.append(line)
            return {
                "status": "created",
                "type": "line_3d",
                "index": len(self.lines_3d) - 1,
                "start": [x1, y1, z1],
                "end": [x2, y2, z2],
                "length": com_get(line, "Length"),
                "sketch_lines": after,
            }
        except Exception as e:
            return error_result(e)

    def get_active_sketch(self) -> Any | None:
        """Get the active sketch object"""
        return self.active_profile

    def get_active_refaxis(self) -> Any | None:
        """Get the active reference axis for revolve operations"""
        return self.active_refaxis

    def get_accumulated_profiles(self) -> list[Any]:
        """Get the list of accumulated closed profiles (for loft/sweep)."""
        return list(self.accumulated_profiles)

    def clear_accumulated_profiles(self) -> None:
        """Clear the accumulated profiles list."""
        self.accumulated_profiles.clear()

    def sketch_fillet(self, radius: float) -> dict[str, Any]:
        """
        Add fillet (round) to sketch corners.

        Rounds all sharp corners in the active sketch by the given radius.
        Works on Line2d intersections.

        Args:
            radius: Fillet radius in meters

        Returns:
            Dict with status and count of fillets added
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            profile = self.active_profile
            lines = profile.Lines2d

            if lines.Count < 2:
                return {"error": "Need at least 2 lines to create a fillet"}

            if radius <= 0:
                return {"error": f"radius must be positive (got {radius}); it is in meters"}

            fillet_count = 0
            failures: list[str] = []
            # Arcs2d.AddAsFillet(Obj1, Obj2, Radius, xDirection, yDirection).
            # There is no AddByFillet, which is what this used to call, so every
            # pair raised and the count silently stayed at zero.
            for i in range(1, lines.Count):
                line1 = lines.Item(i)
                line2 = lines.Item(i + 1)
                points = _corner_points(line1, line2)
                if points is None:
                    continue  # these two do not meet; nothing to round
                corner, _inward = points
                try:
                    profile.Arcs2d.AddAsFillet(line1, line2, radius, corner[0], corner[1])
                    fillet_count += 1
                except Exception as exc:
                    failures.append(f"lines {i - 1}/{i}: {describe_exception(exc)}")

            if not fillet_count:
                return {
                    "error": (
                        "No fillet could be created. Consecutive lines must meet at a "
                        "corner and the radius must fit between them."
                    ),
                    "radius": radius,
                    "failures": failures[:5],
                }

            result: dict[str, Any] = {
                "status": "created",
                "type": "sketch_fillet",
                "radius": radius,
                "fillet_count": fillet_count,
            }
            if failures:
                result["partial_failures"] = failures[:5]
            return result
        except Exception as e:
            return error_result(e)

    def sketch_chamfer(self, distance: float) -> dict[str, Any]:
        """
        Add chamfer to sketch corners.

        Chamfers sharp corners in the active sketch at the given distance.

        Args:
            distance: Chamfer setback distance in meters

        Returns:
            Dict with status
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            profile = self.active_profile
            lines = profile.Lines2d

            if lines.Count < 2:
                return {"error": "Need at least 2 lines to create a chamfer"}

            if distance <= 0:
                return {"error": f"distance must be positive (got {distance}); it is in meters"}

            chamfer_count = 0
            failures: list[str] = []
            # Lines2d.AddAsChamfer(Obj1, Obj2, xDirection, yDirection, SetBackA,
            # SetBackB). There is no AddByChamfer.
            for i in range(1, lines.Count):
                line1 = lines.Item(i)
                line2 = lines.Item(i + 1)
                points = _corner_points(line1, line2)
                if points is None:
                    continue
                _corner, inward = points
                try:
                    profile.Lines2d.AddAsChamfer(
                        line1, line2, inward[0], inward[1], distance, distance
                    )
                    chamfer_count += 1
                except Exception as exc:
                    failures.append(f"lines {i - 1}/{i}: {describe_exception(exc)}")

            if not chamfer_count:
                return {
                    "error": (
                        "No chamfer could be created. Consecutive lines must meet at a "
                        "corner and the setback must fit along both."
                    ),
                    "distance": distance,
                    "failures": failures[:5],
                }

            result: dict[str, Any] = {
                "status": "created",
                "type": "sketch_chamfer",
                "distance": distance,
                "chamfer_count": chamfer_count,
            }
            if failures:
                result["partial_failures"] = failures[:5]
            return result
        except Exception as e:
            return error_result(e)

    def sketch_offset(self, distance: float) -> dict[str, Any]:
        """Offset the active sketch profile by a distance.

        ``Profile.OffsetProfile`` is in no Solid Edge type library, so that
        call always raised, and the manual fallback read ``line.StartPoint.X``,
        which Line2d does not have either. Every element failed silently and
        the result said "created" with a count of zero.

        The real call is ``Profile.Offset2d(offsetSideX, offsetSideY,
        offsetDistance)``, and it only does anything once the geometry is
        selected: with nothing selected it returns cleanly and creates nothing.
        Verified on Solid Edge 2026.

        Solid Edge offsets one connected chain at a time, so a sketch holding
        two separate shapes offsets neither.

        Args:
            distance: Offset distance in meters. Positive offsets away from
                the sketch centre, negative towards it.

        Returns:
            Dict with status and how many elements the sketch gained.
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}
            if distance == 0:
                return {"error": "distance must not be zero; it is in meters"}

            profile = self.active_profile
            before = _element_count(profile)
            if not before:
                return {"error": "No sketch geometry to offset"}

            side = _offset_side_point(profile, distance)
            if side is None:
                return {
                    "error": (
                        "The sketch geometry could not be measured, so there is no "
                        "way to say which side to offset towards."
                    )
                }

            document = None
            with contextlib.suppress(Exception):
                document = self.doc_manager.get_active_document()
            selected = _select_all(profile, document)
            if not selected:
                return {
                    "error": (
                        "None of the sketch geometry could be selected, and "
                        "Profile.Offset2d only acts on a selection."
                    )
                }

            profile.Offset2d(side[0], side[1], abs(distance))
            after = _element_count(profile)

            if after <= before:
                return {
                    "error": (
                        f"Solid Edge created no offset geometry. It offsets one "
                        f"connected chain at a time, so a sketch holding separate "
                        f"shapes offsets none of them; this one has {before} "
                        f"elements. An offset of {distance} m may also be too "
                        f"large for the shape to survive."
                    ),
                    "distance": distance,
                    "elements": before,
                }

            return {
                "status": "created",
                "type": "sketch_offset",
                "distance": distance,
                "elements_created": after - before,
            }
        except Exception as e:
            return error_result(e)

    def sketch_mirror(self, axis: str = "X") -> dict[str, Any]:
        """Mirror the sketch geometry about the X or Y axis.

        This read ``line.StartPoint.X`` and ``circle.CenterPoint.X``. Line2d
        and Circle2d have neither; the accessors are ``GetStartPoint``,
        ``GetEndPoint`` and ``GetCenterPoint``, all pure [out]. Every element
        raised inside a bare except, so the result said "created" with a count
        of zero. Arcs were not handled at all.

        Args:
            axis: 'X' mirrors about the X axis, flipping Y. 'Y' mirrors about
                the Y axis, flipping X.

        Returns:
            Dict with status and how many elements were mirrored.
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            axis_upper = axis.upper()
            if axis_upper not in ("X", "Y"):
                return {"error": f"Invalid axis: {axis}. Use 'X' or 'Y'."}

            profile = self.active_profile
            flip = _mirror_across_x if axis_upper == "X" else _mirror_across_y
            mirror_count = 0
            failures: list[str] = []

            lines = profile.Lines2d
            for i in range(1, lines.Count + 1):
                line = lines.Item(i)
                start = _xy(line, "GetStartPoint")
                end = _xy(line, "GetEndPoint")
                if start is None or end is None:
                    failures.append(f"line {i - 1}: no endpoints")
                    continue
                sx, sy = flip(*start)
                ex, ey = flip(*end)
                try:
                    profile.Lines2d.AddBy2Points(sx, sy, ex, ey)
                    mirror_count += 1
                except Exception as exc:
                    failures.append(f"line {i - 1}: {describe_exception(exc)}")

            circles = profile.Circles2d
            for i in range(1, circles.Count + 1):
                circle = circles.Item(i)
                centre = _xy(circle, "GetCenterPoint")
                radius = com_get(circle, "Radius")
                if centre is None or radius is None:
                    failures.append(f"circle {i - 1}: no centre")
                    continue
                cx, cy = flip(*centre)
                try:
                    profile.Circles2d.AddByCenterRadius(cx, cy, float(radius))
                    mirror_count += 1
                except Exception as exc:
                    failures.append(f"circle {i - 1}: {describe_exception(exc)}")

            arcs = profile.Arcs2d
            for i in range(1, arcs.Count + 1):
                arc = arcs.Item(i)
                centre = _xy(arc, "GetCenterPoint")
                start = _xy(arc, "GetStartPoint")
                end = _xy(arc, "GetEndPoint")
                if centre is None or start is None or end is None:
                    failures.append(f"arc {i - 1}: incomplete")
                    continue
                cx, cy = flip(*centre)
                sx, sy = flip(*start)
                ex, ey = flip(*end)
                try:
                    # Mirroring reverses the sweep, so the endpoints swap.
                    profile.Arcs2d.AddByCenterStartEnd(cx, cy, ex, ey, sx, sy)
                    mirror_count += 1
                except Exception as exc:
                    failures.append(f"arc {i - 1}: {describe_exception(exc)}")

            if not mirror_count:
                return {
                    "error": (
                        "Nothing was mirrored. The sketch needs lines, circles or arcs to mirror."
                    ),
                    "axis": axis_upper,
                    "failures": failures[:5],
                }

            result: dict[str, Any] = {
                "status": "created",
                "type": "sketch_mirror",
                "axis": axis_upper,
                "mirrored_elements": mirror_count,
            }
            if failures:
                result["partial_failures"] = failures[:5]
            return result
        except Exception as e:
            return error_result(e)

    def draw_construction_line(self, x1: float, y1: float, x2: float, y2: float) -> dict[str, Any]:
        """
        Draw a construction line in the active sketch.

        Construction lines are reference geometry that doesn't form part
        of the profile. Useful for symmetry axes, alignment references, etc.

        Args:
            x1, y1: Start point coordinates (meters)
            x2, y2: End point coordinates (meters)

        Returns:
            Dict with creation status
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            profile = self.active_profile
            line = profile.Lines2d.AddBy2Points(x1, y1, x2, y2)

            # ToggleConstruction flips the element in place; it stays in
            # Lines2d and IsConstructionElement reports the new state.
            # Suppressing the toggle silently left an ordinary line that would
            # be treated as part of the profile by the next feature.
            try:
                profile.ToggleConstruction(line)
            except Exception as exc:
                with contextlib.suppress(Exception):
                    line.Delete()
                return {
                    "error": (
                        "The line was drawn but could not be turned into "
                        f"construction geometry, so it was removed rather than "
                        f"left to be extruded: {describe_exception(exc)}"
                    )
                }

            is_construction = None
            with contextlib.suppress(Exception):
                is_construction = bool(profile.IsConstructionElement(line))
            if is_construction is False:
                return {
                    "error": (
                        "Solid Edge accepted the toggle but the line is still "
                        "ordinary profile geometry, so the next feature would "
                        "try to use it."
                    )
                }

            return {
                "status": "created",
                "type": "construction_line",
                "start": [x1, y1],
                "end": [x2, y2],
                "is_construction": is_construction,
            }
        except Exception as e:
            return error_result(e)

    def project_edge(self, face_index: int, edge_index: int) -> dict[str, Any]:
        """
        Project a 3D body edge into the active sketch.

        Uses Profile.ProjectEdge(EdgeToProject) to project a body edge
        onto the sketch plane as a reference curve.

        Args:
            face_index: 0-based face index (from get_body_faces)
            edge_index: 0-based edge index on that face

        Returns:
            Dict with projection status
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            profile = self.active_profile
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            if models.Count == 0:
                return {"error": "No model exists"}

            model = models.Item(1)
            body = model.Body

            # Get the edge from face
            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Count: {faces.Count}"}

            face = faces.Item(face_index + 1)
            edges = face.Edges
            if edge_index < 0 or edge_index >= edges.Count:
                return {"error": f"Invalid edge index: {edge_index}. Count: {edges.Count}"}

            edge = edges.Item(edge_index + 1)

            projected = profile.ProjectEdge(edge)

            return {
                "status": "projected",
                "face_index": face_index,
                "edge_index": edge_index,
                "projected_geometry": str(type(projected).__name__) if projected else "unknown",
            }
        except Exception as e:
            return error_result(e)

    def include_edge(self, face_index: int, edge_index: int) -> dict[str, Any]:
        """
        Include a 3D body edge in the active sketch.

        Uses Profile.IncludeEdge(Edge, Geometry2d) to include a body edge
        as sketch geometry. Unlike ProjectEdge, IncludeEdge maintains an
        associative link to the original edge.

        Args:
            face_index: 0-based face index (from get_body_faces)
            edge_index: 0-based edge index on that face

        Returns:
            Dict with include status
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            profile = self.active_profile
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            if models.Count == 0:
                return {"error": "No model exists"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Count: {faces.Count}"}

            face = faces.Item(face_index + 1)
            edges = face.Edges
            if edge_index < 0 or edge_index >= edges.Count:
                return {"error": f"Invalid edge index: {edge_index}. Count: {edges.Count}"}

            edge = edges.Item(edge_index + 1)

            # IncludeEdge takes Edge and returns Geometry2d via out-param
            result = profile.IncludeEdge(edge)

            return {
                "status": "included",
                "face_index": face_index,
                "edge_index": edge_index,
                "geometry_2d": str(type(result).__name__) if result else "unknown",
            }
        except Exception as e:
            return error_result(e)

    def project_ref_plane(self, plane_index: int) -> dict[str, Any]:
        """
        Project a reference plane into the active sketch.

        Uses Profile.ProjectRefPlane(ReferencePlane) to project a reference plane
        as a construction line in the sketch.

        Args:
            plane_index: 1-based index of the reference plane to project

        Returns:
            Dict with projection status
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            profile = self.active_profile
            doc = self.doc_manager.get_active_document()
            ref_planes = ref_planes_of(doc)
            if ref_planes is None:
                return {"error": _NO_REF_PLANES}

            if plane_index < 1 or plane_index > ref_planes.Count:
                return {"error": f"Invalid plane_index: {plane_index}. Count: {ref_planes.Count}"}

            ref_plane = ref_planes.Item(plane_index)
            result = profile.ProjectRefPlane(ref_plane)

            return {
                "status": "projected",
                "plane_index": plane_index,
                "projected_geometry": str(type(result).__name__) if result else "unknown",
            }
        except Exception as e:
            return error_result(e)

    def offset_sketch_2d(
        self, offset_side_x: float, offset_side_y: float, offset_distance: float
    ) -> dict[str, Any]:
        """
        Offset the active sketch profile in 2D.

        Uses Profile.Offset2d(offsetSideX, offsetSideY, offsetDistance).
        The side parameters control the offset direction.

        Args:
            offset_side_x: X component of the offset direction
            offset_side_y: Y component of the offset direction
            offset_distance: Offset distance in meters

        Returns:
            Dict with offset status
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            profile = self.active_profile
            profile.Offset2d(offset_side_x, offset_side_y, offset_distance)

            return {
                "status": "offset",
                "offset_side_x": offset_side_x,
                "offset_side_y": offset_side_y,
                "offset_distance": offset_distance,
            }
        except Exception as e:
            return error_result(e)

    def sketch_rotate(
        self, center_x: float, center_y: float, angle_degrees: float
    ) -> dict[str, Any]:
        """Rotate every element in the active sketch about a point.

        Solid Edge has no profile-level rotate, so the geometry is read,
        transformed and rebuilt. This used to read ``line.StartPoint.X`` and
        ``circle.CenterPoint.X``, neither of which exists, inside a bare
        except -- and then delete the originals regardless. Rotating a sketch
        erased it and reported success with a count of zero.

        Args:
            center_x: Centre of rotation X, in meters.
            center_y: Centre of rotation Y, in meters.
            angle_degrees: Rotation in degrees, counterclockwise.

        Returns:
            Dict with status and how many elements moved.
        """
        angle_rad = math.radians(angle_degrees)
        cos_a = math.cos(angle_rad)
        sin_a = math.sin(angle_rad)

        def rotate_point(x: float, y: float) -> tuple[float, float]:
            dx, dy = x - center_x, y - center_y
            return (
                center_x + dx * cos_a - dy * sin_a,
                center_y + dx * sin_a + dy * cos_a,
            )

        return self._transform_sketch(
            rotate_point,
            radius_scale=1.0,
            result_type="sketch_rotate",
            extra={"center": [center_x, center_y], "angle_degrees": angle_degrees},
        )

    def sketch_scale(self, center_x: float, center_y: float, scale_factor: float) -> dict[str, Any]:
        """Scale every element in the active sketch about a point.

        Same rebuild as :meth:`sketch_rotate`, and it carried the same defect:
        the coordinate reads raised and the originals were deleted anyway, so
        scaling a sketch erased it.

        Args:
            center_x: Centre of scaling X, in meters.
            center_y: Centre of scaling Y, in meters.
            scale_factor: Multiplier. Must be positive.

        Returns:
            Dict with status and how many elements moved.
        """
        if scale_factor <= 0:
            return {"error": f"Scale factor must be positive (got {scale_factor})"}

        def scale_point(x: float, y: float) -> tuple[float, float]:
            return (
                center_x + (x - center_x) * scale_factor,
                center_y + (y - center_y) * scale_factor,
            )

        return self._transform_sketch(
            scale_point,
            radius_scale=scale_factor,
            result_type="sketch_scale",
            extra={"center": [center_x, center_y], "scale_factor": scale_factor},
        )

    def _transform_sketch(
        self,
        move: Any,
        radius_scale: float,
        result_type: str,
        extra: dict[str, Any],
    ) -> dict[str, Any]:
        """Rebuild every sketch element under a point transform.

        Read everything first, and refuse to delete anything unless all of it
        was read. The old order -- read, delete, rebuild -- destroyed the
        sketch whenever a read failed.
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            profile = self.active_profile
            lines, circles, arcs = profile.Lines2d, profile.Circles2d, profile.Arcs2d

            new_lines: list[tuple[float, float, float, float]] = []
            new_circles: list[tuple[float, float, float]] = []
            new_arcs: list[tuple[float, float, float, float, float, float]] = []
            unreadable: list[str] = []

            for i in range(1, lines.Count + 1):
                element = lines.Item(i)
                start = _xy(element, "GetStartPoint")
                end = _xy(element, "GetEndPoint")
                if start is None or end is None:
                    unreadable.append(f"line {i - 1}")
                    continue
                new_lines.append((*move(*start), *move(*end)))

            for i in range(1, circles.Count + 1):
                element = circles.Item(i)
                centre = _xy(element, "GetCenterPoint")
                radius = com_get(element, "Radius")
                if centre is None or radius is None:
                    unreadable.append(f"circle {i - 1}")
                    continue
                new_circles.append((*move(*centre), float(radius) * radius_scale))

            for i in range(1, arcs.Count + 1):
                element = arcs.Item(i)
                centre = _xy(element, "GetCenterPoint")
                start = _xy(element, "GetStartPoint")
                end = _xy(element, "GetEndPoint")
                if centre is None or start is None or end is None:
                    unreadable.append(f"arc {i - 1}")
                    continue
                new_arcs.append((*move(*centre), *move(*start), *move(*end)))

            total = len(new_lines) + len(new_circles) + len(new_arcs)
            if unreadable:
                # Deleting now would lose whatever could not be read.
                return {
                    "error": (
                        "Some sketch geometry could not be read, so nothing was "
                        "changed rather than risk deleting it: "
                        f"{', '.join(unreadable[:5])}."
                    ),
                    "readable": total,
                }
            if not total:
                return {"error": "No sketch geometry to transform"}

            for i in range(lines.Count, 0, -1):
                lines.Item(i).Delete()
            for i in range(circles.Count, 0, -1):
                circles.Item(i).Delete()
            for i in range(arcs.Count, 0, -1):
                arcs.Item(i).Delete()

            for coords in new_lines:
                lines.AddBy2Points(*coords)
            for circle in new_circles:
                circles.AddByCenterRadius(*circle)
            for arc in new_arcs:
                arcs.AddByCenterStartEnd(*arc)

            return {
                "status": "transformed",
                "type": result_type,
                "elements": total,
                **extra,
            }
        except Exception as e:
            return error_result(e)

    def get_sketch_matrix(self) -> dict[str, Any]:
        """
        Get the coordinate system matrix of the active sketch profile.

        Returns the transformation matrix from sketch 2D space to model 3D space.

        Part.tlb Profile.GetMatrix(Matrix SAFEARRAY(VT_R8)* [in,out]) - the
        caller supplies the 4x4 (16 element) buffer and reads the filled matrix
        back out of the return value.

        Returns:
            Dict with matrix elements
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            result = self.active_profile.GetMatrix(_r8_array(16))

            if isinstance(result, tuple):
                return {"status": "ok", "matrix": list(result)}
            else:
                return {"status": "ok", "matrix": result}
        except Exception as e:
            return error_result(e)

    def clean_sketch_geometry(
        self,
        clean_points: bool = True,
        clean_splines: bool = True,
        clean_identical: bool = True,
        clean_small: bool = True,
        small_tolerance: float = 0.0001,
    ) -> dict[str, Any]:
        """
        Clean up duplicate, small, or invalid geometry in the active sketch.

        Uses Profile.CleanGeometry2d to remove duplicate points, simplify
        splines, remove identical curves, and delete small elements.

        Args:
            clean_points: Remove duplicate/coincident points
            clean_splines: Simplify spline curves
            clean_identical: Remove identical overlapping curves
            clean_small: Remove very small elements
            small_tolerance: Size threshold for small element removal (meters)

        Returns:
            Dict with status
        """
        # No reachable COM call to pass the options to; see below.
        del clean_points, clean_splines, clean_identical, clean_small, small_tolerance
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            # Profile.CleanGeometry2d takes seven arguments (the Sheet
            # overload's nine include an element array, which Profile's does
            # not). Every shape was tried against Solid Edge 2026 -- five,
            # six and seven arguments, the layer as None, 0 and Missing -- and
            # each returns 0x80070057 E_INVALIDARG. The options argument is a
            # CleanProfileOptions value the type library does not enumerate,
            # so there is nothing left to guess at.
            return {
                "error": (
                    "Sketch cleanup is not reachable through COM automation. "
                    "Profile.CleanGeometry2d rejects every documented argument form "
                    "with E_INVALIDARG. Use Tools > Clean Geometry in Solid Edge, or "
                    "avoid creating duplicates: draw() reports the element count so a "
                    "repeated call is visible."
                ),
                "unsupported": True,
            }
        except Exception as e:
            return error_result(e)

    def get_sketch_constraints(self) -> dict[str, Any]:
        """
        Get information about constraints in the active sketch.

        Returns:
            Dict with constraint count and types
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            profile = self.active_profile
            constraints: list[dict[str, Any]] = []

            try:
                relations = profile.Relations2d
                for i in range(1, relations.Count + 1):
                    try:
                        rel = relations.Item(i)
                        constraint_info: dict[str, Any] = {"index": i - 1}
                        with contextlib.suppress(Exception):
                            constraint_info["type"] = rel.Type
                        with contextlib.suppress(Exception):
                            constraint_info["name"] = rel.Name
                        constraints.append(constraint_info)
                    except Exception:
                        constraints.append({"index": i - 1, "type": "unknown"})
            except Exception:
                pass

            return {"constraints": constraints, "count": len(constraints)}
        except Exception as e:
            return error_result(e)

    def project_silhouette_edges(self) -> dict[str, Any]:
        """
        Project silhouette edges of the body onto the active sketch.

        Part.tlb Profile.ProjectSilhouetteEdges(
            FaceToProject VT_DISPATCH [in],
            Geometry2dCount VT_I4* [out],
            Geometry2d SAFEARRAY(VT_DISPATCH)* [in,out])

        The first argument is the specific face whose silhouette is projected.
        This server has no way to pick that face, and there is no meaningful
        default, so the call is refused rather than guessed at.

        Returns:
            Dict with an unsupported error
        """
        profile = self.active_profile
        if not profile:
            return {"error": "No active sketch profile"}

        return {
            "error": (
                "Projecting silhouette edges needs the specific Face to project, "
                "which this server cannot select. Use include_region_faces() with "
                "a face index, or the Solid Edge UI."
            ),
            "unsupported": True,
        }

    def include_region_faces(self, face_indices: list[int]) -> dict[str, Any]:
        """
        Include faces as regions in the active sketch.

        Part.tlb Profile.IncludeRegionFaces(
            NumberOfRegionFaces VT_I4 [in],
            RegionFaces SAFEARRAY(VT_DISPATCH)* [in])

        Both arguments are required; the count must precede the array.

        Args:
            face_indices: List of 0-based face indices to include

        Returns:
            Dict with status and count of included faces
        """
        try:
            profile = self.active_profile
            if not profile:
                return {"error": "No active sketch profile"}

            if not face_indices:
                return {"error": "No face indices provided"}

            doc = self.doc_manager.get_active_document()
            models = doc.Models
            if models.Count == 0:
                return {"error": "No model exists"}

            model = models.Item(1)
            body = model.Body
            faces = body.Faces(FaceQueryConstants.igQueryAll)

            face_list = []
            for fi in face_indices:
                if fi < 0 or fi >= faces.Count:
                    return {"error": f"Invalid face index: {fi}. Count: {faces.Count}"}
                face_list.append(faces.Item(fi + 1))

            profile.IncludeRegionFaces(len(face_list), _dispatch_array(face_list))

            return {
                "status": "included",
                "type": "region_faces",
                "face_count": len(face_list),
                "face_indices": face_indices,
            }
        except Exception as e:
            return error_result(e)

    def sketch_paste(self) -> dict[str, Any]:
        """
        Paste clipboard content into the active sketch.

        Uses Profile.Paste() to paste geometry from the clipboard into the
        active sketch profile. The clipboard must contain valid sketch
        geometry (e.g., from a previous Copy operation).

        Returns:
            Dict with status
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            self.active_profile.Paste()

            return {"status": "pasted", "type": "sketch_paste"}
        except Exception as e:
            return error_result(e)

    def get_ordered_geometry(self) -> dict[str, Any]:
        """
        Get the ordered geometry elements from the active sketch.

        Part.tlb Profile.OrderedGeometry(
            NumElements VT_I4* [out],
            Elements SAFEARRAY(VT_DISPATCH)* [in,out])

        NumElements precedes the buffer, so Elements is passed by keyword; the
        buffer is pre-sized from the profile's own 2D geometry collections.

        Returns:
            Dict with status, element count, and element details
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            buffer = self._collect_geometry_2d(self.active_profile)
            result = self.active_profile.OrderedGeometry(Elements=_dispatch_array(buffer))

            # OrderedGeometry returns (NumElements, Elements) as out params
            if isinstance(result, tuple) and len(result) == 2:
                num_elements = result[0]
                elements = result[1]
            else:
                # Fallback: single return value might be the collection
                num_elements = 0
                elements = result if result else []

            element_list = []
            if elements:
                items = elements if isinstance(elements, (list, tuple)) else [elements]
                for i, elem in enumerate(items):
                    info: dict[str, Any] = {"index": i}

                    with contextlib.suppress(Exception):
                        info["type"] = str(type(elem).__name__)

                    # StartPoint, EndPoint and CenterPoint are not properties
                    # of a 2D element; the accessors are pure-[out] methods,
                    # so every entry came back holding nothing but an index.
                    start = _xy(elem, "GetStartPoint")
                    if start is not None:
                        info["start_x"], info["start_y"] = start

                    end = _xy(elem, "GetEndPoint")
                    if end is not None:
                        info["end_x"], info["end_y"] = end

                    centre = _xy(elem, "GetCenterPoint")
                    if centre is not None:
                        info["center_x"], info["center_y"] = centre

                    with contextlib.suppress(Exception):
                        info["radius"] = elem.Radius

                    with contextlib.suppress(Exception):
                        info["length"] = elem.Length

                    element_list.append(info)

            count = num_elements if isinstance(num_elements, int) else len(element_list)
            return {
                "status": "ok",
                "num_elements": count,
                "elements": element_list,
            }
        except Exception as e:
            return error_result(e)

    def chain_locate(self, x: float, y: float, tolerance: float = 0.001) -> dict[str, Any]:
        """
        Find a chain of connected sketch elements at a location.

        Part.tlb Profile.ChainLocate(x VT_R8 [in], y VT_R8 [in]) - the API takes
        the point only; ``tolerance`` is echoed back for the caller but is not
        passed to Solid Edge, which uses its own locate tolerance.

        Args:
            x: X coordinate to search near (meters)
            y: Y coordinate to search near (meters)
            tolerance: Reported back only; Solid Edge uses its own tolerance

        Returns:
            Dict with status and chain info
        """
        try:
            profile = self.active_profile
            if not profile:
                return {"error": "No active sketch profile"}

            result = profile.ChainLocate(x, y)

            return {
                "status": "located",
                "type": "chain",
                "x": x,
                "y": y,
                "tolerance": tolerance,
                "chain_result": str(type(result).__name__) if result else "none",
            }
        except Exception as e:
            return error_result(e)

    def convert_to_curve(self) -> dict[str, Any]:
        """
        Convert sketch geometry to a curve.

        Part.tlb Profile.ConvertToCurve(
            NumberOfElements VT_I4 [in],
            ElementArray SAFEARRAY(VT_DISPATCH)* [in],
            NumConvertedElements VT_VARIANT* [out,optional],
            ConvertedElements VT_VARIANT* [in,out,optional])

        The count and element array are required; they are gathered from the
        active profile's own 2D geometry collections.

        Returns:
            Dict with status
        """
        try:
            profile = self.active_profile
            if not profile:
                return {"error": "No active sketch profile"}

            elements = self._collect_geometry_2d(profile)
            if not elements:
                return {"error": "Active sketch has no 2D geometry to convert"}

            result = profile.ConvertToCurve(len(elements), _dispatch_array(elements))

            return {
                "status": "converted",
                "type": "curve",
                "element_count": len(elements),
                "curve_result": str(type(result).__name__) if result else "none",
            }
        except Exception as e:
            return error_result(e)

    @staticmethod
    def _collect_geometry_2d(profile: Any) -> list[Any]:
        """Gather every 2D element of a profile across its geometry collections.

        Profile has no single "all geometry" collection, so the individual
        Lines2d/Arcs2d/... collections are concatenated.
        """
        elements: list[Any] = []
        for coll_name in _GEOMETRY_2D_COLLECTIONS:
            try:
                collection = getattr(profile, coll_name)
                for i in range(1, collection.Count + 1):
                    with contextlib.suppress(Exception):
                        elements.append(collection.Item(i))
            except Exception:
                continue
        return elements

    def get_active_plane_index(self) -> int | None:
        """Return the 1-based reference plane index of the open sketch.

        1=Top/XY, 2=Right/YZ, 3=Front/XZ, 4+ user-created planes. Returns None
        when no sketch is open or the plane cannot be resolved.
        """
        if self.active_profile is None:
            return None
        if self.active_plane_index is not None:
            return self.active_plane_index

        # Fall back to matching the profile's plane against the document's
        # RefPlanes collection by name.
        try:
            plane_name = self.active_profile.Plane.Name
            ref_planes = ref_planes_of(self.doc_manager.get_active_document())
            for i in range(1, ref_planes.Count + 1):
                if ref_planes.Item(i).Name == plane_name:
                    return i
        except Exception:
            return None
        return None
