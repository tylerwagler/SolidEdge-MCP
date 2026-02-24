"""Annotation operations (dimensions, marks, symbols, text, leaders, balloons, 2D queries)."""

import contextlib
import traceback
from typing import Any

from ..logging import get_logger

_logger = get_logger(__name__)


class AnnotationsMixin:
    """Mixin providing annotation and dimension methods."""

    # =================================================================
    # BASIC ANNOTATIONS (text boxes, leaders, dimensions, balloons, notes)
    # =================================================================

    def add_text_box(self, x: float, y: float, text: str, height: float = 0.005) -> dict[str, Any]:
        """
        Add a text box annotation to the active draft sheet.

        Args:
            x: X position on sheet (meters)
            y: Y position on sheet (meters)
            text: Text content
            height: Text height in meters (default 0.005 = 5mm)

        Returns:
            Dict with status and text box info
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet

            # TextBoxes.Add(x, y, 0) for 2D sheets
            text_boxes = sheet.TextBoxes
            text_box = text_boxes.Add(x, y, 0)

            # Set the text content
            text_box.Text = text

            # Set text height if possible
            with contextlib.suppress(Exception):
                text_box.TextHeight = height

            return {"status": "added", "type": "text_box", "text": text, "position": [x, y]}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_leader(
        self, x1: float, y1: float, x2: float, y2: float, text: str = ""
    ) -> dict[str, Any]:
        """
        Add a leader annotation to the active draft sheet.

        A leader is an arrow pointing to geometry with optional text.

        Args:
            x1: Arrow start X (meters)
            y1: Arrow start Y (meters)
            x2: Text end X (meters)
            y2: Text end Y (meters)
            text: Optional text at the leader end

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet

            leaders = sheet.Leaders
            leader = leaders.Add(x1, y1, 0, x2, y2, 0)

            if text:
                with contextlib.suppress(Exception):
                    leader.Text = text

            return {
                "status": "added",
                "type": "leader",
                "start": [x1, y1],
                "end": [x2, y2],
                "text": text,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_dimension(
        self, x1: float, y1: float, x2: float, y2: float,
        dim_x: float | None = None, dim_y: float | None = None,
    ) -> dict[str, Any]:
        """
        Add a linear dimension between two points on the active draft sheet.

        Args:
            x1: First point X (meters)
            y1: First point Y (meters)
            x2: Second point X (meters)
            y2: Second point Y (meters)
            dim_x: Dimension text X position (meters, optional)
            dim_y: Dimension text Y position (meters, optional)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet

            # Default dimension text position to midpoint offset
            if dim_x is None:
                dim_x = (x1 + x2) / 2
            if dim_y is None:
                dim_y = max(y1, y2) + 0.02  # 20mm above

            dimensions = sheet.Dimensions
            dimensions.AddLength(x1, y1, 0, x2, y2, 0, dim_x, dim_y, 0)

            return {
                "status": "added",
                "type": "dimension",
                "point1": [x1, y1],
                "point2": [x2, y2],
                "text_position": [dim_x, dim_y],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_balloon(
        self, x: float, y: float, text: str = "",
        leader_x: float | None = None, leader_y: float | None = None,
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

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet

            balloons = sheet.Balloons

            if leader_x is not None and leader_y is not None:
                balloon = balloons.Add(leader_x, leader_y, 0, x, y, 0)
            else:
                balloon = balloons.Add(x, y, 0, x + 0.02, y + 0.02, 0)

            if text:
                try:
                    balloon.BalloonText = text
                except Exception:
                    with contextlib.suppress(Exception):
                        balloon.Text = text

            return {"status": "added", "type": "balloon", "position": [x, y], "text": text}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_note(self, x: float, y: float, text: str, height: float = 0.005) -> dict[str, Any]:
        """
        Add a note (free-standing text) to the active draft sheet.

        Similar to text box but simpler - just plain text annotation.

        Args:
            x: Note X position (meters)
            y: Note Y position (meters)
            text: Note text content
            height: Text height in meters (default 5mm)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet

            # Try TextBoxes first (most reliable)
            text_boxes = sheet.TextBoxes
            text_box = text_boxes.Add(x, y, 0)
            text_box.Text = text

            with contextlib.suppress(Exception):
                text_box.TextHeight = height

            return {
                "status": "added",
                "type": "note",
                "position": [x, y],
                "text": text,
                "height": height,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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
        """
        Add an angular dimension between three points on the active draft sheet.

        The angle is measured at the vertex (x2, y2) between the rays to
        (x1, y1) and (x3, y3).

        Args:
            x1: First ray endpoint X (meters)
            y1: First ray endpoint Y (meters)
            x2: Vertex X (meters)
            y2: Vertex Y (meters)
            x3: Second ray endpoint X (meters)
            y3: Second ray endpoint Y (meters)
            dim_x: Dimension text X position (meters, optional)
            dim_y: Dimension text Y position (meters, optional)

        Returns:
            Dict with status and dimension info
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet
            dims = sheet.Dimensions

            text_x = dim_x if dim_x is not None else (x1 + x3) / 2
            text_y = dim_y if dim_y is not None else (y1 + y3) / 2 + 0.02

            dims.AddAngular(x1, y1, x2, y2, x3, y3, text_x, text_y)

            return {
                "status": "created",
                "type": "angular_dimension",
                "vertex": [x2, y2],
                "text_position": [text_x, text_y],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_radial_dimension(
        self,
        center_x: float,
        center_y: float,
        point_x: float,
        point_y: float,
        dim_x: float | None = None,
        dim_y: float | None = None,
    ) -> dict[str, Any]:
        """
        Add a radial dimension on the active draft sheet.

        Args:
            center_x: Arc center X (meters)
            center_y: Arc center Y (meters)
            point_x: Point on arc X (meters)
            point_y: Point on arc Y (meters)
            dim_x: Dimension text X position (meters, optional)
            dim_y: Dimension text Y position (meters, optional)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet
            dims = sheet.Dimensions

            text_x = dim_x if dim_x is not None else (center_x + point_x) / 2
            text_y = dim_y if dim_y is not None else (center_y + point_y) / 2 + 0.02

            dims.AddRadial(center_x, center_y, point_x, point_y, text_x, text_y)

            return {
                "status": "created",
                "type": "radial_dimension",
                "center": [center_x, center_y],
                "point": [point_x, point_y],
                "text_position": [text_x, text_y],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_diameter_dimension(
        self,
        center_x: float,
        center_y: float,
        point_x: float,
        point_y: float,
        dim_x: float | None = None,
        dim_y: float | None = None,
    ) -> dict[str, Any]:
        """
        Add a diameter dimension on the active draft sheet.

        Args:
            center_x: Circle center X (meters)
            center_y: Circle center Y (meters)
            point_x: Point on circle X (meters)
            point_y: Point on circle Y (meters)
            dim_x: Dimension text X position (meters, optional)
            dim_y: Dimension text Y position (meters, optional)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet
            dims = sheet.Dimensions

            text_x = dim_x if dim_x is not None else center_x + (point_x - center_x) * 1.3
            text_y = dim_y if dim_y is not None else center_y + (point_y - center_y) * 1.3 + 0.02

            dims.AddDiameter(center_x, center_y, point_x, point_y, text_x, text_y)

            return {
                "status": "created",
                "type": "diameter_dimension",
                "center": [center_x, center_y],
                "point": [point_x, point_y],
                "text_position": [text_x, text_y],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_ordinate_dimension(
        self,
        origin_x: float,
        origin_y: float,
        point_x: float,
        point_y: float,
        dim_x: float | None = None,
        dim_y: float | None = None,
    ) -> dict[str, Any]:
        """
        Add an ordinate dimension on the active draft sheet.

        Ordinate dimensions show the distance from an origin to a point
        along a single axis.

        Args:
            origin_x: Datum origin X (meters)
            origin_y: Datum origin Y (meters)
            point_x: Measured point X (meters)
            point_y: Measured point Y (meters)
            dim_x: Dimension text X position (meters, optional)
            dim_y: Dimension text Y position (meters, optional)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet
            dims = sheet.Dimensions

            text_x = dim_x if dim_x is not None else point_x
            text_y = dim_y if dim_y is not None else point_y + 0.02

            dims.AddOrdinate(origin_x, origin_y, point_x, point_y, text_x, text_y)

            return {
                "status": "created",
                "type": "ordinate_dimension",
                "origin": [origin_x, origin_y],
                "point": [point_x, point_y],
                "text_position": [text_x, text_y],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_distance_dimension(self, x1: float, y1: float, x2: float, y2: float) -> dict[str, Any]:
        """
        Add a distance dimension between two points on the active draft sheet.

        Uses sheet.Dimensions.AddDistanceBetweenPoints to measure the distance
        between two coordinate pairs. The dimension text is placed at the
        midpoint offset above.

        Args:
            x1: First point X (meters)
            y1: First point Y (meters)
            x2: Second point X (meters)
            y2: Second point Y (meters)

        Returns:
            Dict with status and dimension type
        """
        try:
            doc = self.doc_manager.get_active_document()
            if not hasattr(doc, "ActiveSheet"):
                return {"error": "Active document is not a draft"}
            sheet = doc.ActiveSheet
            dims = sheet.Dimensions

            # Place dimension text at midpoint, offset above
            dim_x = (x1 + x2) / 2
            dim_y = max(y1, y2) + 0.02

            # AddDistanceBetweenPoints(x1, y1, z1, x2, y2, z2, dimX, dimY, dimZ)
            dims.AddDistanceBetweenPoints(x1, y1, 0.0, x2, y2, 0.0, dim_x, dim_y, 0.0)

            return {
                "status": "added",
                "dimension_type": "distance",
                "point1": [x1, y1],
                "point2": [x2, y2],
                "text_position": [dim_x, dim_y],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            if not hasattr(doc, "ActiveSheet"):
                return {"error": "Active document is not a draft"}
            sheet = doc.ActiveSheet
            lines2d = sheet.Lines2d

            if object_index < 0 or object_index >= lines2d.Count:
                return {"error": (f"Invalid line index: {object_index}. Count: {lines2d.Count}")}

            line = lines2d.Item(object_index + 1)  # COM is 1-indexed
            dims = sheet.Dimensions

            # Get line endpoints for dimension text placement
            x1, y1, x2, y2 = 0.0, 0.0, 0.0, 0.0
            with contextlib.suppress(Exception):
                x1, y1, x2, y2 = line.StartX, line.StartY, line.EndX, line.EndY
            dim_x = (x1 + x2) / 2
            dim_y = max(y1, y2) + 0.02

            dims.AddLength(x1, y1, 0.0, x2, y2, 0.0, dim_x, dim_y, 0.0)

            return {
                "status": "added",
                "dimension_type": "length",
                "object_index": object_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_radius_dimension_2d(
        self, object_index: int, object_type: str = "circle"
    ) -> dict[str, Any]:
        """
        Add a radius dimension to a circle or arc on the active draft sheet.

        Gets the object from Circles2d or Arcs2d by index and adds a radius
        dimension via sheet.Dimensions.AddRadialDimension.

        Args:
            object_index: 0-based index into Circles2d or Arcs2d collection
            object_type: 'circle' or 'arc'

        Returns:
            Dict with status and dimension type
        """
        try:
            doc = self.doc_manager.get_active_document()
            if not hasattr(doc, "ActiveSheet"):
                return {"error": "Active document is not a draft"}
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
            dims = sheet.Dimensions

            # Get center and radius for dimension placement
            cx, cy, radius = 0.0, 0.0, 0.0
            with contextlib.suppress(Exception):
                cx = obj.CenterX
                cy = obj.CenterY
                radius = obj.Radius

            # Place dimension text slightly outside the circle/arc
            dim_x = cx + radius + 0.01
            dim_y = cy + 0.01

            dims.AddRadialDimension(obj, dim_x, dim_y, 0.0)

            return {
                "status": "added",
                "dimension_type": "radius",
                "object_type": object_type,
                "object_index": object_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_angle_dimension_2d(
        self,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        x3: float,
        y3: float,
    ) -> dict[str, Any]:
        """
        Add an angle dimension between three points on the active draft sheet.

        Uses sheet.Dimensions.AddAngle to measure the angle at the vertex
        (x2, y2).

        Args:
            x1: First point X (meters) - start of first ray
            y1: First point Y (meters)
            x2: Vertex point X (meters) - the corner
            y2: Vertex point Y (meters)
            x3: Third point X (meters) - end of second ray
            y3: Third point Y (meters)

        Returns:
            Dict with status and dimension type
        """
        try:
            doc = self.doc_manager.get_active_document()
            if not hasattr(doc, "ActiveSheet"):
                return {"error": "Active document is not a draft"}
            sheet = doc.ActiveSheet
            dims = sheet.Dimensions

            # Place dimension text at the vertex with offset
            dim_x = x2 + 0.02
            dim_y = y2 + 0.02

            dims.AddAngle(x1, y1, 0.0, x2, y2, 0.0, x3, y3, 0.0, dim_x, dim_y, 0.0)

            return {
                "status": "added",
                "dimension_type": "angle",
                "vertex": [x2, y2],
                "text_position": [dim_x, dim_y],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet
            center_marks = sheet.CenterMarks
            center_marks.Add(x, y, 0)

            return {
                "status": "added",
                "type": "center_mark",
                "position": [x, y],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet
            centerlines = sheet.Centerlines
            centerlines.Add(x1, y1, 0, x2, y2, 0)

            return {
                "status": "added",
                "type": "centerline",
                "start": [x1, y1],
                "end": [x2, y2],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_surface_finish_symbol(
        self, x: float, y: float, symbol_type: str = "machined"
    ) -> dict[str, Any]:
        """
        Add a surface finish symbol to the active draft sheet.

        Args:
            x: Symbol X position (meters)
            y: Symbol Y position (meters)
            symbol_type: Type of surface finish - 'machined', 'any', 'prohibited'

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            type_map = {"machined": 1, "any": 0, "prohibited": 2}
            type_value = type_map.get(symbol_type.lower())
            if type_value is None:
                valid = ", ".join(type_map.keys())
                return {"error": f"Invalid symbol_type: '{symbol_type}'. Valid: {valid}"}

            sheet = doc.ActiveSheet

            try:
                sfs = sheet.SurfaceFinishSymbols
                sfs.Add(x, y, 0, type_value)
            except Exception:
                # Fallback: use TextBoxes with standard surface finish text
                text_boxes = sheet.TextBoxes
                symbols = {"machined": "\u2327", "any": "\u2328", "prohibited": "\u2329"}
                text_box = text_boxes.Add(x, y, 0)
                text_box.Text = symbols.get(symbol_type.lower(), "\u2327")

            return {
                "status": "added",
                "type": "surface_finish_symbol",
                "symbol_type": symbol_type,
                "position": [x, y],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            type_map = {"fillet": 0, "groove": 1, "plug": 2, "spot": 3, "seam": 4}
            type_value = type_map.get(weld_type.lower())
            if type_value is None:
                valid = ", ".join(type_map.keys())
                return {"error": f"Invalid weld_type: '{weld_type}'. Valid: {valid}"}

            sheet = doc.ActiveSheet

            try:
                ws = sheet.WeldSymbols
                ws.Add(x, y, 0, type_value)
            except Exception:
                # Fallback: use a leader with weld designation text
                leaders = sheet.Leaders
                leader = leaders.Add(x, y, 0, x + 0.02, y + 0.02, 0)
                with contextlib.suppress(Exception):
                    leader.Text = f"[{weld_type.upper()}]"

            return {
                "status": "added",
                "type": "weld_symbol",
                "weld_type": weld_type,
                "position": [x, y],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet

            try:
                fcfs = sheet.FCFs
                fcf = fcfs.Add(x, y, 0)
                if tolerance_text:
                    with contextlib.suppress(Exception):
                        fcf.Text = tolerance_text
            except Exception:
                # Fallback: use a text box with GD&T text
                text_boxes = sheet.TextBoxes
                text_box = text_boxes.Add(x, y, 0)
                text_box.Text = tolerance_text if tolerance_text else "[GD&T]"

            return {
                "status": "added",
                "type": "geometric_tolerance",
                "position": [x, y],
                "text": tolerance_text,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            if not hasattr(doc, "ActiveSheet"):
                return {"error": "Active document is not a draft"}
            sheet = doc.ActiveSheet
            lines2d = sheet.Lines2d
            items = []
            for i in range(1, lines2d.Count + 1):
                line = lines2d.Item(i)
                info: dict[str, Any] = {"index": i - 1}
                with contextlib.suppress(Exception):
                    info["start"] = [line.StartX, line.StartY]
                with contextlib.suppress(Exception):
                    info["end"] = [line.EndX, line.EndY]
                items.append(info)
            return {"count": len(items), "lines": items}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            if not hasattr(doc, "ActiveSheet"):
                return {"error": "Active document is not a draft"}
            sheet = doc.ActiveSheet
            circles2d = sheet.Circles2d
            items = []
            for i in range(1, circles2d.Count + 1):
                circle = circles2d.Item(i)
                info: dict[str, Any] = {"index": i - 1}
                with contextlib.suppress(Exception):
                    info["center"] = [circle.CenterX, circle.CenterY]
                with contextlib.suppress(Exception):
                    info["radius"] = circle.Radius
                items.append(info)
            return {"count": len(items), "circles": items}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            if not hasattr(doc, "ActiveSheet"):
                return {"error": "Active document is not a draft"}
            sheet = doc.ActiveSheet
            arcs2d = sheet.Arcs2d
            items = []
            for i in range(1, arcs2d.Count + 1):
                arc = arcs2d.Item(i)
                info: dict[str, Any] = {"index": i - 1}
                with contextlib.suppress(Exception):
                    info["center"] = [arc.CenterX, arc.CenterY]
                with contextlib.suppress(Exception):
                    info["radius"] = arc.Radius
                with contextlib.suppress(Exception):
                    info["start_angle"] = arc.StartAngle
                with contextlib.suppress(Exception):
                    info["end_angle"] = arc.EndAngle
                items.append(info)
            return {"count": len(items), "arcs": items}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
