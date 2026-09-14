"""Annotation operations (dimensions, marks, symbols, text, leaders, balloons, 2D queries)."""

import contextlib
from typing import Any

from solidedge_mcp.backends.errors import error_result

from ..comutil import com_get
from ..logging import get_logger
from ._base import NOT_A_DRAFT
from ._locate import nearest_element, no_element_error

_logger = get_logger(__name__)

# WeldSymbol.TopType values, from Program/constant.tlb > DimWeldTypeConstants.
# The previous 0..4 map was invented: 0 is igDimWeldTypeNone, so every weld
# symbol came out with the wrong glyph.
_WELD_TYPE_MAP = {
    "fillet": 1,  # igDimWeldTopFillet
    "groove": 5,  # igDimWeldTopVGroove
    "plug": 6,  # igDimWeldTopSlot
    "spot": 2,  # igDimWeldTopSpot
    "seam": 3,  # igDimWeldTopSeam
}


def _xy(obj: Any, member: str) -> list[float] | None:
    """Read a pure-[out] 2D accessor such as GetCenterPoint as [x, y]."""
    try:
        point = getattr(obj, member)()
        return [float(point[0]), float(point[1])]
    except Exception:
        return None


def _dimension_value(dimension: Any) -> float | None:
    """The measured value of a Dimension, or None if it will not give one.

    Dimensions.Add* returns a Dimension whose default property is its value,
    so pywin32 hands back a float for most of them.
    """
    if isinstance(dimension, int | float):
        return float(dimension)
    for member in ("Value", "DimensionValue"):
        try:
            return float(getattr(dimension, member))
        except Exception:
            continue
    return None


class AnnotationsMixin:
    """Mixin providing annotation and dimension methods."""

    # =================================================================
    # BASIC ANNOTATIONS (text boxes, leaders, dimensions, balloons, notes)
    # =================================================================

    def add_text_box(self, x: float, y: float, text: str, height: float = 0.005) -> dict[str, Any]:
        """Add a text box annotation to the active draft sheet.

        ``height`` is the box height, which Solid Edge clamps upward to fit the
        text: asking for less than the text needs leaves it at the natural
        height, and the returned ``height`` is what Solid Edge actually holds
        rather than what was asked for. It is not the font size -- TextBox has
        no text-height property, only ``TextScale``, a multiplier whose
        relationship to a height in meters did not hold up on Solid Edge 2026,
        so it is left alone rather than guessed at.

        This used to write ``TextHeight``, which is real on DrawingView and
        CuttingPlane but not on TextBox, inside a suppress -- so the height was
        silently dropped and the caller was told it had been set.

        Args:
            x: X position on sheet (meters)
            y: Y position on sheet (meters)
            text: Text content
            height: Box height in meters (default 0.005 = 5mm)

        Returns:
            Dict with status and text box info
        """
        try:
            doc = self.doc_manager.get_active_document()

            err = self._require_draft(doc)
            if err:
                return err

            sheet = doc.ActiveSheet

            # TextBoxes.Add(x, y, 0) for 2D sheets
            text_boxes = sheet.TextBoxes
            text_box = text_boxes.Add(x, y, 0)

            # Set the text content
            text_box.Text = text

            # TextBox has Height, not TextHeight -- TextHeight is real but
            # belongs to DrawingView and CuttingPlane. The write went to a
            # member TextBox does not have, inside a suppress, so every text
            # box came out at the document default and nothing said so.
            text_box.Height = height

            return {
                "status": "added",
                "type": "text_box",
                "text": text,
                "position": [x, y],
                "height": com_get(text_box, "Height"),
            }
        except Exception as e:
            return error_result(e)

    def add_leader(
        self, x1: float, y1: float, x2: float, y2: float, text: str = ""
    ) -> dict[str, Any]:
        """Add a leader (a plain arrow) to the active draft sheet.

        A Solid Edge Leader carries no text. ``Leader.Text`` is on no
        interface -- Solid Edge 2026 answers "Property 'Add.Text' can not be
        set." -- and the write sat inside a suppress, so the text vanished and
        the result reported it anyway. Text with a leader is a Balloon or a
        Note, so a caller asking for one is told that instead of being handed
        a bare arrow and a false success.

        Args:
            x1: Arrow tip X (meters)
            y1: Arrow tip Y (meters)
            x2: Tail end X (meters)
            y2: Tail end Y (meters)
            text: Must be empty; see above.

        Returns:
            Dict with status
        """
        if text:
            return {
                "error": (
                    "A Solid Edge leader carries no text of its own. Use "
                    "add_annotation(type='balloon', x, y, text, leader_x, "
                    "leader_y) for text on a leader, or type='note' for text "
                    "alone, and call this without text for a plain arrow."
                ),
                "unsupported": True,
            }
        try:
            doc = self.doc_manager.get_active_document()

            err = self._require_draft(doc)
            if err:
                return err

            sheet = doc.ActiveSheet

            leaders = sheet.Leaders
            leaders.Add(x1, y1, 0, x2, y2, 0)

            return {
                "status": "added",
                "type": "leader",
                "start": [x1, y1],
                "end": [x2, y2],
            }
        except Exception as e:
            return error_result(e)

    # =================================================================
    # DIMENSION HELPERS
    #
    # Every Dimensions.Add* method takes the object being dimensioned. These
    # turn the coordinates an MCP caller has into that object.
    # =================================================================

    def _distance_between_points(
        self, x1: float, y1: float, x2: float, y2: float
    ) -> dict[str, Any]:
        """Dimension the distance between two points, via the elements at them."""
        try:
            doc = self.doc_manager.get_active_document()
            err = self._require_draft(doc)
            if err:
                return err
            sheet = doc.ActiveSheet

            first = nearest_element(sheet, x1, y1)
            if first is None:
                return no_element_error(x1, y1, "element to dimension from")
            second = nearest_element(sheet, x2, y2)
            if second is None:
                return no_element_error(x2, y2, "element to dimension to")

            # keyPoint=True snaps each end to the element keypoint nearest the
            # coordinate, which makes this a true point-to-point distance.
            value = sheet.Dimensions.AddDistanceBetweenObjects(
                first.obj, x1, y1, 0.0, True, second.obj, x2, y2, 0.0, True
            )

            return {
                "status": "created",
                "type": "distance_dimension",
                "point1": [x1, y1],
                "point2": [x2, y2],
                "value": _dimension_value(value),
                "attached_to": [first.describe(), second.describe()],
            }
        except Exception as e:
            return error_result(e)

    def _angle_between_points(
        self, x1: float, y1: float, x2: float, y2: float, x3: float, y3: float
    ) -> dict[str, Any]:
        """Dimension the angle between the lines at two of the three points."""
        try:
            doc = self.doc_manager.get_active_document()
            err = self._require_draft(doc)
            if err:
                return err
            sheet = doc.ActiveSheet

            first = nearest_element(sheet, x1, y1, kinds=("line",))
            if first is None:
                return no_element_error(x1, y1, "line to measure the angle from")
            second = nearest_element(sheet, x3, y3, kinds=("line",))
            if second is None:
                return no_element_error(x3, y3, "line to measure the angle to")
            if first.identity == second.identity:
                return {
                    "error": (
                        "Both points landed on the same line, so there is no angle "
                        "to dimension. Give a point on each of the two lines."
                    ),
                    "point1": [x1, y1],
                    "point2": [x3, y3],
                }

            value = sheet.Dimensions.AddAngleBetweenObjects(
                first.obj, x1, y1, 0.0, False, second.obj, x3, y3, 0.0, False
            )

            return {
                "status": "created",
                "type": "angular_dimension",
                "vertex": [x2, y2],
                "value_radians": _dimension_value(value),
                "attached_to": [first.describe(), second.describe()],
            }
        except Exception as e:
            return error_result(e)

    def _circular_dimension(
        self,
        center_x: float,
        center_y: float,
        point_x: float,
        point_y: float,
        want: str,
        result_type: str,
    ) -> dict[str, Any]:
        """Dimension the radius or diameter of the curve at the given point."""
        try:
            doc = self.doc_manager.get_active_document()
            err = self._require_draft(doc)
            if err:
                return err
            sheet = doc.ActiveSheet

            kinds = ("circle", "arc")
            hit = nearest_element(sheet, point_x, point_y, kinds=kinds)
            if hit is None:
                # A caller may only know the centre; that locates it too.
                hit = nearest_element(sheet, center_x, center_y, kinds=kinds)
            if hit is None:
                return no_element_error(point_x, point_y, "circle or arc")

            dims = sheet.Dimensions
            if want == "radius":
                value = dims.AddRadius(hit.obj)
            elif hit.kind == "circle":
                value = dims.AddCircularDiameter(hit.obj)
            else:
                # An arc has no full circle to place a diameter across.
                value = dims.AddRadialDiameter(hit.obj)

            return {
                "status": "created",
                "type": result_type,
                "center": [center_x, center_y],
                "point": [point_x, point_y],
                "value": _dimension_value(value),
                "attached_to": hit.describe(),
            }
        except Exception as e:
            return error_result(e)

    def add_dimension(
        self,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        dim_x: float | None = None,
        dim_y: float | None = None,
    ) -> dict[str, Any]:
        """Add a linear dimension between two points on the active draft sheet.

        The ``Dimensions`` collection is object-based:
        ``AddDistanceBetweenObjects(Object1, x1, y1, z1, keyPoint1, Object2,
        x2, y2, z2, keyPoint2)``. There is no coordinate-only overload, and
        the ``AddDistanceBetweenPoints`` this once called is in no Solid Edge
        type library. So each point is resolved to the nearest element on the
        sheet first, exactly as a mouse pick would, and ``keyPoint`` is set so
        the dimension snaps to that element's keypoint nearest the coordinate.
        Verified against Solid Edge 2026.

        Args:
            x1: First point X, in meters of sheet space.
            y1: First point Y, in meters of sheet space.
            x2: Second point X, in meters of sheet space.
            y2: Second point Y, in meters of sheet space.
            dim_x: Ignored. Solid Edge places the dimension text itself.
            dim_y: Ignored.

        Returns:
            Dict with status, the measured value and what each end attached to.
        """
        del dim_x, dim_y
        return self._distance_between_points(x1, y1, x2, y2)

    def add_balloon(
        self,
        x: float,
        y: float,
        text: str = "",
        leader_x: float | None = None,
        leader_y: float | None = None,
    ) -> dict[str, Any]:
        """
        Add a balloon annotation to the active draft sheet.

        Balloons are circular annotations typically used for BOM item numbers.

        Args:
            x: Balloon center X (meters)
            y: Balloon center Y (meters)
            text: Text inside the balloon
            leader_x: Leader arrow X (meters, optional)
            leader_y: Leader arrow Y (meters, optional)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            err = self._require_draft(doc)
            if err:
                return err

            sheet = doc.ActiveSheet

            balloons = sheet.Balloons

            # Balloons.Add(x1, y1, z1) places the balloon; it takes three
            # arguments, not the six of Leaders.Add. Extra leader vertices go on
            # afterwards via Balloon.AddVertex(x, y, z).
            balloon = balloons.Add(x, y, 0)

            if leader_x is not None and leader_y is not None:
                with contextlib.suppress(Exception):
                    balloon.AddVertex(leader_x, leader_y, 0)

            if text:
                with contextlib.suppress(Exception):
                    balloon.BalloonText = text

            return {"status": "added", "type": "balloon", "position": [x, y], "text": text}
        except Exception as e:
            return error_result(e)

    def add_note(self, x: float, y: float, text: str, height: float = 0.005) -> dict[str, Any]:
        """Add a note (free-standing text) to the active draft sheet.

        A note is a text box with plain content; ``height`` behaves exactly as
        it does in ``add_text_box``, clamped upward to fit the text and
        reported back as Solid Edge holds it.

        Args:
            x: Note X position (meters)
            y: Note Y position (meters)
            text: Note text content
            height: Box height in meters (default 5mm)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            err = self._require_draft(doc)
            if err:
                return err

            sheet = doc.ActiveSheet

            # Try TextBoxes first (most reliable)
            text_boxes = sheet.TextBoxes
            text_box = text_boxes.Add(x, y, 0)
            text_box.Text = text
            # See add_text_box: the height went to TextHeight, which TextBox
            # does not have, so this reported a height it had never set.
            text_box.Height = height

            return {
                "status": "added",
                "type": "note",
                "position": [x, y],
                "text": text,
                "height": com_get(text_box, "Height"),
            }
        except Exception as e:
            return error_result(e)

    # =================================================================
    # DIMENSION ANNOTATIONS
    # =================================================================

    def add_angular_dimension(
        self,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        x3: float,
        y3: float,
        dim_x: float | None = None,
        dim_y: float | None = None,
    ) -> dict[str, Any]:
        """Add an angular dimension between two lines on the active draft sheet.

        ``Dimensions`` has no ``AddAngular`` member. The real entry point is
        ``AddAngleBetweenObjects(ele1, x1, y1, z1, keyPoint1, ele2, ...)``,
        which takes the two elements forming the angle, so (x1, y1) and
        (x3, y3) are used to find them. The vertex is implied by where the two
        elements meet, so (x2, y2) only disambiguates which side is measured.

        Args:
            x1: A point on the first line, in meters of sheet space.
            y1: A point on the first line.
            x2: The vertex. Recorded in the result; Solid Edge derives it.
            y2: The vertex Y.
            x3: A point on the second line.
            y3: A point on the second line.
            dim_x: Ignored. Solid Edge places the dimension text itself.
            dim_y: Ignored.

        Returns:
            Dict with status, the measured angle in radians, and both picks.
        """
        del dim_x, dim_y
        return self._angle_between_points(x1, y1, x2, y2, x3, y3)

    def add_radial_dimension(
        self,
        center_x: float,
        center_y: float,
        point_x: float,
        point_y: float,
        dim_x: float | None = None,
        dim_y: float | None = None,
    ) -> dict[str, Any]:
        """Add a radius dimension to the circle or arc at the given point.

        ``Dimensions.AddRadius(Object)`` takes the curve being dimensioned.
        The ``AddRadial`` this once called is in no Solid Edge type library,
        so the point on the curve is used to find the circle or arc instead.

        Args:
            center_x: Curve center X, in meters of sheet space.
            center_y: Curve center Y.
            point_x: A point on the curve, in meters of sheet space.
            point_y: A point on the curve.
            dim_x: Ignored. Solid Edge places the dimension text itself.
            dim_y: Ignored.

        Returns:
            Dict with status, the radius, and the curve that was dimensioned.
        """
        del dim_x, dim_y
        return self._circular_dimension(
            center_x, center_y, point_x, point_y, "radius", "radial_dimension"
        )

    def add_diameter_dimension(
        self,
        center_x: float,
        center_y: float,
        point_x: float,
        point_y: float,
        dim_x: float | None = None,
        dim_y: float | None = None,
    ) -> dict[str, Any]:
        """Add a diameter dimension to the circle or arc at the given point.

        Circles get ``Dimensions.AddCircularDiameter(Object)`` and arcs get
        ``AddRadialDiameter(Object)``. The ``AddDiameter`` this once called is
        in no Solid Edge type library.

        Args:
            center_x: Curve center X, in meters of sheet space.
            center_y: Curve center Y.
            point_x: A point on the curve, in meters of sheet space.
            point_y: A point on the curve.
            dim_x: Ignored. Solid Edge places the dimension text itself.
            dim_y: Ignored.

        Returns:
            Dict with status, the diameter, and the curve that was dimensioned.
        """
        del dim_x, dim_y
        return self._circular_dimension(
            center_x, center_y, point_x, point_y, "diameter", "diameter_dimension"
        )

    def add_ordinate_dimension(
        self,
        origin_x: float,
        origin_y: float,
        point_x: float,
        point_y: float,
        dim_x: float | None = None,
        dim_y: float | None = None,
    ) -> dict[str, Any]:
        """Add an ordinate dimension from a datum origin to a measured point.

        Two calls, both object-based:
        ``AddCoordinateOrigin(Object, x, y, z, keyPoint)`` establishes the
        datum and ``AddCoordinate(Obj1, ..., Obj2, ...)`` measures to it. The
        ``AddOrdinate`` this once called is in no Solid Edge type library.
        Each coordinate is resolved to the nearest element, and ``keyPoint`` is
        set so both ends snap to real vertices.

        Args:
            origin_x: Datum origin X, in meters of sheet space.
            origin_y: Datum origin Y.
            point_x: Measured point X, in meters of sheet space.
            point_y: Measured point Y.
            dim_x: Ignored. Solid Edge places the dimension text itself.
            dim_y: Ignored.

        Returns:
            Dict with status, the measured value, and both picks.
        """
        del dim_x, dim_y
        try:
            doc = self.doc_manager.get_active_document()
            err = self._require_draft(doc)
            if err:
                return err
            sheet = doc.ActiveSheet

            origin_hit = nearest_element(sheet, origin_x, origin_y)
            if origin_hit is None:
                return no_element_error(origin_x, origin_y, "element to use as the datum")
            point_hit = nearest_element(sheet, point_x, point_y)
            if point_hit is None:
                return no_element_error(point_x, point_y, "element to measure to")

            dims = sheet.Dimensions
            dims.AddCoordinateOrigin(origin_hit.obj, origin_x, origin_y, 0.0, True)
            value = dims.AddCoordinate(
                origin_hit.obj,
                origin_x,
                origin_y,
                0.0,
                True,
                point_hit.obj,
                point_x,
                point_y,
                0.0,
                True,
            )

            return {
                "status": "created",
                "type": "ordinate_dimension",
                "origin": [origin_x, origin_y],
                "point": [point_x, point_y],
                "value": _dimension_value(value),
                "attached_to": [origin_hit.describe(), point_hit.describe()],
            }
        except Exception as e:
            return error_result(e)

    def add_distance_dimension(self, x1: float, y1: float, x2: float, y2: float) -> dict[str, Any]:
        """Add a distance dimension between two points on the active draft sheet.

        See :meth:`add_dimension`; this is the same operation reached through
        ``add_2d_dimension(type='distance')``.

        Args:
            x1: First point X, in meters of sheet space.
            y1: First point Y, in meters of sheet space.
            x2: Second point X, in meters of sheet space.
            y2: Second point Y, in meters of sheet space.

        Returns:
            Dict with status, the measured value and what each end attached to.
        """
        return self._distance_between_points(x1, y1, x2, y2)

    def add_length_dimension(self, object_index: int) -> dict[str, Any]:
        """
        Add a length dimension to a 2D line object on the active draft sheet.

        Gets the line from Lines2d collection by index and adds a length
        dimension via sheet.Dimensions.AddLength.

        Args:
            object_index: 0-based index into the Lines2d collection

        Returns:
            Dict with status and dimension type
        """
        try:
            doc = self.doc_manager.get_active_document()
            err = self._require_draft(doc, NOT_A_DRAFT)
            if err:
                return err
            sheet = doc.ActiveSheet
            lines2d = sheet.Lines2d

            if object_index < 0 or object_index >= lines2d.Count:
                return {"error": (f"Invalid line index: {object_index}. Count: {lines2d.Count}")}

            line = lines2d.Item(object_index + 1)  # COM is 1-indexed
            dims = sheet.Dimensions

            # AddLength(Object as VT_DISPATCH) dimensions the 2D element
            # itself. It does not take endpoint or text-placement coordinates.
            value = dims.AddLength(line)

            return {
                "status": "added",
                "dimension_type": "length",
                "object_index": object_index,
                "value": _dimension_value(value),
            }
        except Exception as e:
            return error_result(e)

    def add_radius_dimension_2d(
        self, object_index: int, object_type: str = "circle"
    ) -> dict[str, Any]:
        """Add a radius dimension to a circle or arc picked out of a collection.

        The curve comes from ``sheet.Circles2d`` or ``sheet.Arcs2d`` by index
        and is dimensioned with ``Dimensions.AddRadius(Object)``. The
        ``AddRadialDimension`` this once called is in no Solid Edge type
        library, and neither are the ``CenterX``/``CenterY`` accessors it read
        to place the text; Solid Edge places the text itself.

        Args:
            object_index: 0-based index into Circles2d or Arcs2d.
            object_type: 'circle' or 'arc', selecting the collection.

        Returns:
            Dict with status, the radius, and the object dimensioned.
        """
        try:
            doc = self.doc_manager.get_active_document()
            err = self._require_draft(doc, NOT_A_DRAFT)
            if err:
                return err
            sheet = doc.ActiveSheet

            if object_type == "circle":
                collection = sheet.Circles2d
            elif object_type == "arc":
                collection = sheet.Arcs2d
            else:
                return {"error": (f"Invalid object_type: {object_type}. Use 'circle' or 'arc'.")}

            if object_index < 0 or object_index >= collection.Count:
                return {"error": (f"Invalid index: {object_index}. Count: {collection.Count}")}

            obj = collection.Item(object_index + 1)  # COM 1-indexed
            value = sheet.Dimensions.AddRadius(obj)

            return {
                "status": "added",
                "dimension_type": "radius",
                "object_type": object_type,
                "object_index": object_index,
                "value": _dimension_value(value),
            }
        except Exception as e:
            return error_result(e)

    def add_angle_dimension_2d(
        self,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        x3: float,
        y3: float,
    ) -> dict[str, Any]:
        """Add an angle dimension between two lines on the active draft sheet.

        See :meth:`add_angular_dimension`; this is the same operation reached
        through ``add_2d_dimension(type='angle')``.

        Args:
            x1: A point on the first line, in meters of sheet space.
            y1: A point on the first line.
            x2: The vertex. Recorded in the result; Solid Edge derives it.
            y2: The vertex Y.
            x3: A point on the second line.
            y3: A point on the second line.

        Returns:
            Dict with status, the measured angle in radians, and both picks.
        """
        return self._angle_between_points(x1, y1, x2, y2, x3, y3)

    # =================================================================
    # SYMBOL ANNOTATIONS
    # =================================================================

    def add_center_mark(self, x: float, y: float) -> dict[str, Any]:
        """
        Add a center mark annotation at the specified coordinates.

        Args:
            x: Center mark X position (meters)
            y: Center mark Y position (meters)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            err = self._require_draft(doc)
            if err:
                return err

            sheet = doc.ActiveSheet
            center_marks = sheet.CenterMarks
            center_marks.Add(x, y, 0)

            return {
                "status": "added",
                "type": "center_mark",
                "position": [x, y],
            }
        except Exception as e:
            return error_result(e)

    def add_centerline(self, x1: float, y1: float, x2: float, y2: float) -> dict[str, Any]:
        """
        Add a centerline between two points on the active draft sheet.

        Args:
            x1: Start X (meters)
            y1: Start Y (meters)
            x2: End X (meters)
            y2: End Y (meters)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            err = self._require_draft(doc)
            if err:
                return err

            sheet = doc.ActiveSheet
            centerlines = sheet.CenterLines
            centerlines.Add(x1, y1, 0, x2, y2, 0)

            return {
                "status": "added",
                "type": "centerline",
                "start": [x1, y1],
                "end": [x2, y2],
            }
        except Exception as e:
            return error_result(e)

    def add_surface_finish_symbol(
        self, x: float, y: float, symbol_type: str = "machined"
    ) -> dict[str, Any]:
        """Surface finish symbols cannot be placed through COM automation.

        ``SurfaceFinishSymbols`` has no ``Add``; the real method is
        ``AddSurfaceFinishSymbol(AnnotInitData)``, and the AnnotInitData needs
        a terminator or connect element -- an edge or keypoint picked on the
        drawing -- which this server cannot select. Verified on Solid Edge
        2026: with only ``SetPlane(sheet)`` the call returns None and the
        collection stays empty.

        What this used to do was worse than failing. The ``Add`` call raised,
        and the fallback dropped a text box containing U+2327 on the sheet and
        reported ``{"status": "added", "type": "surface_finish_symbol"}`` --
        a stray glyph passed off as a symbol.

        Args:
            x: Symbol X position (meters), unused.
            y: Symbol Y position (meters), unused.
            symbol_type: 'machined', 'any' or 'prohibited', unused.

        Returns:
            Dict with an ``unsupported`` error.
        """
        del x, y, symbol_type
        return {
            "error": (
                "Surface finish symbols need an AnnotInitData carrying a terminator "
                "element picked on the drawing, which this server cannot select: "
                "SurfaceFinishSymbols has no Add, and AddSurfaceFinishSymbol with "
                "only a plane creates nothing. Place it in the Solid Edge UI."
            ),
            "unsupported": True,
        }

    def add_weld_symbol(self, x: float, y: float, weld_type: str = "fillet") -> dict[str, Any]:
        """
        Add a welding symbol to the active draft sheet.

        Args:
            x: Symbol X position (meters)
            y: Symbol Y position (meters)
            weld_type: Type of weld - 'fillet', 'groove', 'plug', 'spot', 'seam'

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            err = self._require_draft(doc)
            if err:
                return err

            type_value = _WELD_TYPE_MAP.get(weld_type.lower())
            if type_value is None:
                valid = ", ".join(_WELD_TYPE_MAP.keys())
                return {"error": f"Invalid weld_type: '{weld_type}'. Valid: {valid}"}

            sheet = doc.ActiveSheet

            # WeldSymbols.Add(x1, y1, z1) only places the symbol; the weld
            # type is the WeldSymbol.TopType property, not a fourth argument.
            # Both were verified against Solid Edge 2026. The fallback that
            # used to sit here drew a bare arrow -- it set Leader.Text, which
            # no interface has -- and still reported a weld symbol.
            symbol = sheet.WeldSymbols.Add(x, y, 0)
            symbol.TopType = type_value

            return {
                "status": "added",
                "type": "weld_symbol",
                "weld_type": weld_type,
                "position": [x, y],
                "top_type": com_get(symbol, "TopType"),
            }
        except Exception as e:
            return error_result(e)

    def add_geometric_tolerance(
        self, x: float, y: float, tolerance_text: str = ""
    ) -> dict[str, Any]:
        """
        Add a geometric tolerance (Feature Control Frame / GD&T) to the active draft.

        Args:
            x: FCF X position (meters)
            y: FCF Y position (meters)
            tolerance_text: Tolerance specification text (e.g., "0.05 A B")

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            err = self._require_draft(doc)
            if err:
                return err

            sheet = doc.ActiveSheet

            # draft.tlb names this collection FeatureControlFrames. sheet.FCFs
            # does not exist, so this always fell through to the text-box
            # fallback and no real feature control frame was ever created.
            frames = com_get(sheet, "FeatureControlFrames")
            if frames is None:
                return {
                    "error": (
                        "This sheet has no FeatureControlFrames collection, so a "
                        "geometric tolerance cannot be placed. Add it in the Solid "
                        "Edge UI."
                    )
                }
            frame = frames.Add(x, y, 0)
            if tolerance_text:
                with contextlib.suppress(Exception):
                    frame.Text = tolerance_text

            return {
                "status": "added",
                "type": "geometric_tolerance",
                "position": [x, y],
                "text": tolerance_text,
                "total_frames": com_get(frames, "Count"),
            }
        except Exception as e:
            return error_result(e)

    # =================================================================
    # 2D GEOMETRY COLLECTION ACCESS (Draft Sheets)
    # =================================================================

    def get_lines2d(self) -> dict[str, Any]:
        """
        List all 2D lines on the active draft sheet.

        Accesses the sheet.Lines2d collection and iterates to extract
        start/end vertex coordinates for each line.

        Returns:
            Dict with count and list of line info dicts
        """
        try:
            doc = self.doc_manager.get_active_document()
            err = self._require_draft(doc, NOT_A_DRAFT)
            if err:
                return err
            sheet = doc.ActiveSheet
            lines2d = sheet.Lines2d
            items = []
            for i in range(1, lines2d.Count + 1):
                line = lines2d.Item(i)
                info: dict[str, Any] = {"index": i - 1}
                # Line2d has no StartX/StartY/EndX/EndY. GetStartPoint and
                # GetEndPoint are pure [out], so pywin32 returns the pair.
                start = _xy(line, "GetStartPoint")
                if start is not None:
                    info["start"] = start
                end = _xy(line, "GetEndPoint")
                if end is not None:
                    info["end"] = end
                items.append(info)
            return {"count": len(items), "lines": items}
        except Exception as e:
            return error_result(e)

    def get_circles2d(self) -> dict[str, Any]:
        """
        List all 2D circles on the active draft sheet.

        Accesses the sheet.Circles2d collection and iterates to extract
        center coordinates and radius for each circle.

        Returns:
            Dict with count and list of circle info dicts
        """
        try:
            doc = self.doc_manager.get_active_document()
            err = self._require_draft(doc, NOT_A_DRAFT)
            if err:
                return err
            sheet = doc.ActiveSheet
            circles2d = sheet.Circles2d
            items = []
            for i in range(1, circles2d.Count + 1):
                circle = circles2d.Item(i)
                info: dict[str, Any] = {"index": i - 1}
                # Circle2d has no CenterX/CenterY; GetCenterPoint is the accessor.
                center = _xy(circle, "GetCenterPoint")
                if center is not None:
                    info["center"] = center
                with contextlib.suppress(Exception):
                    info["radius"] = circle.Radius
                    info["diameter"] = circle.Diameter
                items.append(info)
            return {"count": len(items), "circles": items}
        except Exception as e:
            return error_result(e)

    def get_arcs2d(self) -> dict[str, Any]:
        """
        List all 2D arcs on the active draft sheet.

        Accesses the sheet.Arcs2d collection and iterates to extract
        center coordinates, radius, and start/end angles for each arc.

        Returns:
            Dict with count and list of arc info dicts
        """
        try:
            doc = self.doc_manager.get_active_document()
            err = self._require_draft(doc, NOT_A_DRAFT)
            if err:
                return err
            sheet = doc.ActiveSheet
            arcs2d = sheet.Arcs2d
            items = []
            for i in range(1, arcs2d.Count + 1):
                arc = arcs2d.Item(i)
                info: dict[str, Any] = {"index": i - 1}
                # Arc2d has no CenterX/CenterY and no EndAngle. It reports
                # StartAngle and SweepAngle, both in radians, so the end angle
                # is their sum.
                center = _xy(arc, "GetCenterPoint")
                if center is not None:
                    info["center"] = center
                start = _xy(arc, "GetStartPoint")
                if start is not None:
                    info["start"] = start
                end = _xy(arc, "GetEndPoint")
                if end is not None:
                    info["end"] = end
                with contextlib.suppress(Exception):
                    info["radius"] = arc.Radius
                with contextlib.suppress(Exception):
                    start_angle = float(arc.StartAngle)
                    sweep = float(arc.SweepAngle)
                    info["start_angle"] = start_angle
                    info["sweep_angle"] = sweep
                    info["end_angle"] = start_angle + sweep
                items.append(info)
            return {"count": len(items), "arcs": items}
        except Exception as e:
            return error_result(e)

    # =================================================================
    # SHEET 2D GEOMETRY
    # =================================================================

    def draw_sheet_geometry(
        self,
        shape: str,
        x1: float = 0.0,
        y1: float = 0.0,
        x2: float = 0.0,
        y2: float = 0.0,
        x3: float = 0.0,
        y3: float = 0.0,
        center_x: float = 0.0,
        center_y: float = 0.0,
        radius: float = 0.0,
    ) -> dict[str, Any]:
        """Draw 2D geometry directly on the active draft sheet.

        The server could already read sheet geometry (query_sheet lines2d,
        circles2d, arcs2d) and dimension it, but had no way to create any, so
        those tools only worked on geometry a person had drawn by hand. These
        are the ``fwksupp.tlb`` collection methods: ``Lines2d.AddBy2Points``,
        ``Circles2d.AddByCenterRadius``, ``Circles2d.AddBy3Points`` and
        ``Arcs2d.AddByCenterStartEnd``.

        This is draft annotation geometry on the sheet itself. It is not a
        part sketch; use manage_sketch and draw for those.

        Args:
            shape: 'line' | 'rectangle' | 'circle' | 'circle_3point' | 'arc'.
            x1: First point X, in meters of sheet space.
            y1: First point Y.
            x2: Second point X. For a rectangle, the opposite corner.
            y2: Second point Y.
            x3: Third point X, for circle_3point.
            y3: Third point Y.
            center_x: Center X, for circle and arc.
            center_y: Center Y, for circle and arc.
            radius: Radius in meters, for circle.

        Returns:
            Dict with status and the new collection count.
        """
        try:
            doc = self.doc_manager.get_active_document()
            err = self._require_draft(doc, NOT_A_DRAFT)
            if err:
                return err
            sheet = doc.ActiveSheet

            match shape:
                case "line":
                    sheet.Lines2d.AddBy2Points(x1, y1, x2, y2)
                    created, total = 1, sheet.Lines2d.Count
                case "rectangle":
                    corners = [
                        (x1, y1, x2, y1),
                        (x2, y1, x2, y2),
                        (x2, y2, x1, y2),
                        (x1, y2, x1, y1),
                    ]
                    for ax, ay, bx, by in corners:
                        sheet.Lines2d.AddBy2Points(ax, ay, bx, by)
                    created, total = 4, sheet.Lines2d.Count
                case "circle":
                    if radius <= 0:
                        return {"error": f"radius must be positive (got {radius}); it is in meters"}
                    sheet.Circles2d.AddByCenterRadius(center_x, center_y, radius)
                    created, total = 1, sheet.Circles2d.Count
                case "circle_3point":
                    sheet.Circles2d.AddBy3Points(x1, y1, x2, y2, x3, y3)
                    created, total = 1, sheet.Circles2d.Count
                case "arc":
                    sheet.Arcs2d.AddByCenterStartEnd(center_x, center_y, x1, y1, x2, y2)
                    created, total = 1, sheet.Arcs2d.Count
                case _:
                    return {
                        "error": (
                            f"Unknown shape: {shape}. Use 'line', 'rectangle', "
                            f"'circle', 'circle_3point' or 'arc'."
                        )
                    }

            return {
                "status": "created",
                "shape": shape,
                "elements_created": created,
                "collection_count": total,
            }
        except Exception as e:
            return error_result(e)
