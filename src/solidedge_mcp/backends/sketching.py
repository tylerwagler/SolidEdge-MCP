"""
Solid Edge Sketching Operations

Handles creating and manipulating 2D sketches.
"""

from typing import Dict, Any, Optional, Tuple
import traceback
import math
from .constants import RefPlaneConstants, ProfileValidationConstants, FaceQueryConstants


class SketchManager:
    """Manages sketch creation and 2D geometry"""

    def __init__(self, document_manager):
        self.doc_manager = document_manager
        self.active_sketch = None
        self.active_profile = None
        self.active_refaxis = None  # Reference axis for revolve operations
        self.accumulated_profiles = []  # For loft/sweep multi-profile operations

    def create_sketch(self, plane: str = "Top") -> Dict[str, Any]:
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
            ref_planes = doc.RefPlanes

            # Map plane names to indices
            plane_map = {
                "Top": 1,     # XZ plane (top view)
                "Front": 2,   # XY plane (front view)
                "Right": 3,   # YZ plane (right view)
                "XZ": 1,
                "XY": 2,
                "YZ": 3
            }

            plane_index = plane_map.get(plane)
            if plane_index is None:
                return {"error": f"Invalid plane: {plane}. Use 'Top', 'Front', 'Right', 'XY', 'XZ', or 'YZ'"}

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
            self.active_refaxis = None  # Clear any previous axis

            return {
                "status": "created",
                "plane": plane,
                "sketch_id": profile_set.Name if hasattr(profile_set, 'Name') else "sketch"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_sketch_on_plane_index(self, plane_index: int) -> Dict[str, Any]:
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
            ref_planes = doc.RefPlanes

            if plane_index < 1 or plane_index > ref_planes.Count:
                return {"error": f"Invalid plane index: {plane_index}. Count: {ref_planes.Count}"}

            ref_plane = ref_planes.Item(plane_index)

            profile_sets = doc.ProfileSets
            profile_set = profile_sets.Add()
            profiles = profile_set.Profiles
            profile = profiles.Add(ref_plane)

            self.active_sketch = profile_set
            self.active_profile = profile
            self.active_refaxis = None

            return {
                "status": "created",
                "plane_index": plane_index,
                "sketch_id": profile_set.Name if hasattr(profile_set, 'Name') else "sketch"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def draw_line(self, x1: float, y1: float, x2: float, y2: float) -> Dict[str, Any]:
        """Draw a line in the active sketch"""
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            # Get Lines2d collection
            lines = self.active_profile.Lines2d

            # Add line
            line = lines.AddBy2Points(x1, y1, x2, y2)

            return {
                "status": "created",
                "type": "line",
                "start": [x1, y1],
                "end": [x2, y2]
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def draw_circle(self, center_x: float, center_y: float, radius: float) -> Dict[str, Any]:
        """Draw a circle in the active sketch"""
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            # Get Circles2d collection
            circles = self.active_profile.Circles2d

            # Add circle by center and radius
            circle = circles.AddByCenterRadius(center_x, center_y, radius)

            return {
                "status": "created",
                "type": "circle",
                "center": [center_x, center_y],
                "radius": radius
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def draw_rectangle(self, x1: float, y1: float, x2: float, y2: float) -> Dict[str, Any]:
        """Draw a rectangle in the active sketch"""
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            # Rectangles are typically drawn as 4 lines
            lines = self.active_profile.Lines2d

            # Draw 4 sides of rectangle
            line1 = lines.AddBy2Points(x1, y1, x2, y1)  # Bottom
            line2 = lines.AddBy2Points(x2, y1, x2, y2)  # Right
            line3 = lines.AddBy2Points(x2, y2, x1, y2)  # Top
            line4 = lines.AddBy2Points(x1, y2, x1, y1)  # Left

            return {
                "status": "created",
                "type": "rectangle",
                "corner1": [x1, y1],
                "corner2": [x2, y2],
                "lines": 4
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def draw_arc(self, center_x: float, center_y: float, radius: float,
                 start_angle: float, end_angle: float) -> Dict[str, Any]:
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
            arc = arcs.AddByCenterStartEnd(
                center_x, center_y,
                start_x, start_y,
                end_x, end_y
            )

            return {
                "status": "created",
                "type": "arc",
                "center": [center_x, center_y],
                "radius": radius,
                "start_angle": start_angle,
                "end_angle": end_angle
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def draw_polygon(self, center_x: float, center_y: float, radius: float, sides: int) -> Dict[str, Any]:
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
                "sides": sides
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def draw_ellipse(self, center_x: float, center_y: float,
                     major_radius: float, minor_radius: float, angle: float = 0.0) -> Dict[str, Any]:
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

            # Convert angle to radians for axis calculation
            angle_rad = math.radians(angle)

            # AddByCenter takes 6 params: cx, cy, major_radius, minor_radius, axis_x, axis_y
            # The axis defines the direction of the major axis
            axis_x = math.cos(angle_rad)
            axis_y = math.sin(angle_rad)

            ellipse = ellipses.AddByCenter(
                center_x, center_y,
                major_radius, minor_radius,
                axis_x, axis_y
            )

            return {
                "status": "created",
                "type": "ellipse",
                "center": [center_x, center_y],
                "major_radius": major_radius,
                "minor_radius": minor_radius,
                "angle": angle
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def draw_spline(self, points: list) -> Dict[str, Any]:
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
            spline = splines.AddByPoints(
                3,  # Order (cubic spline)
                len(points),  # NumPoints
                tuple(point_array)  # PointArray (flattened x,y,x,y,...)
            )

            return {
                "status": "created",
                "type": "spline",
                "points": points,
                "num_points": len(points)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def draw_arc_by_3_points(self, start_x: float, start_y: float,
                             center_x: float, center_y: float,
                             end_x: float, end_y: float) -> Dict[str, Any]:
        """
        Draw an arc defined by start point, center point, and end point.

        Args:
            start_x, start_y: Arc start point (meters)
            center_x, center_y: Arc center point (meters)
            end_x, end_y: Arc end point (meters)

        Returns:
            Dict with status and arc info
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            arcs = self.active_profile.Arcs2d
            arc = arcs.AddByStartCenterEnd(
                start_x, start_y,
                center_x, center_y,
                end_x, end_y
            )

            return {
                "status": "created",
                "type": "arc",
                "start": [start_x, start_y],
                "center": [center_x, center_y],
                "end": [end_x, end_y],
                "method": "start_center_end"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def draw_circle_by_2_points(self, x1: float, y1: float,
                                 x2: float, y2: float) -> Dict[str, Any]:
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

            circles = self.active_profile.Circles2d
            circle = circles.AddBy2Points(x1, y1, x2, y2)

            center_x = (x1 + x2) / 2
            center_y = (y1 + y2) / 2
            radius = math.sqrt((x2 - x1) ** 2 + (y2 - y1) ** 2) / 2

            return {
                "status": "created",
                "type": "circle",
                "center": [center_x, center_y],
                "radius": radius,
                "method": "2_points"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def draw_circle_by_3_points(self, x1: float, y1: float,
                                 x2: float, y2: float,
                                 x3: float, y3: float) -> Dict[str, Any]:
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
            circle = circles.AddBy3Points(x1, y1, x2, y2, x3, y3)

            return {
                "status": "created",
                "type": "circle",
                "point1": [x1, y1],
                "point2": [x2, y2],
                "point3": [x3, y3],
                "method": "3_points"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def mirror_spline(self, axis_x1: float, axis_y1: float,
                       axis_x2: float, axis_y2: float,
                       copy: bool = True) -> Dict[str, Any]:
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
                "mirrored_count": mirror_count
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def hide_profile(self, visible: bool = False) -> Dict[str, Any]:
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

            return {
                "status": "updated",
                "visible": visible
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def draw_point(self, x: float, y: float) -> Dict[str, Any]:
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
                    "method": "Holes2d"
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
                "method": "construction_circle"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def set_axis_of_revolution(self, x1: float, y1: float, x2: float, y2: float) -> Dict[str, Any]:
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
                "note": "Axis of revolution set. Close sketch and use create_revolve()."
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def _get_sketch_element(self, element_type: str, index: int):
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
            raise ValueError(f"Unknown element type: '{element_type}'. Use: {', '.join(type_map.keys())}")

        collection = getattr(profile, collection_name)
        if index < 1 or index > collection.Count:
            raise ValueError(f"Index {index} out of range for {element_type} (count: {collection.Count})")

        return collection.Item(index)

    def add_constraint(self, constraint_type: str, elements: list) -> Dict[str, Any]:
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
                objs.append(self._get_sketch_element(elem[0], elem[1]))

            ct = constraint_type.lower()

            # Single-element constraints
            if ct == "horizontal":
                if len(objs) < 1:
                    return {"error": "Horizontal constraint requires 1 element"}
                relations.AddHorizontal(objs[0])
            elif ct == "vertical":
                if len(objs) < 1:
                    return {"error": "Vertical constraint requires 1 element"}
                relations.AddVertical(objs[0])
            # Two-element constraints
            elif ct == "parallel":
                if len(objs) < 2:
                    return {"error": "Parallel constraint requires 2 elements"}
                relations.AddParallel(objs[0], objs[1])
            elif ct == "perpendicular":
                if len(objs) < 2:
                    return {"error": "Perpendicular constraint requires 2 elements"}
                relations.AddPerpendicular(objs[0], objs[1])
            elif ct == "equal":
                if len(objs) < 2:
                    return {"error": "Equal constraint requires 2 elements"}
                relations.AddEqual(objs[0], objs[1])
            elif ct == "concentric":
                if len(objs) < 2:
                    return {"error": "Concentric constraint requires 2 elements"}
                relations.AddConcentric(objs[0], objs[1])
            elif ct == "tangent":
                if len(objs) < 2:
                    return {"error": "Tangent constraint requires 2 elements"}
                relations.AddTangent(objs[0], objs[1])
            else:
                return {"error": f"Unknown constraint type: '{constraint_type}'. "
                        "Use: Horizontal, Vertical, Parallel, Perpendicular, Equal, Concentric, Tangent"}

            return {
                "status": "constraint_added",
                "type": constraint_type,
                "elements": elements
            }
        except ValueError as e:
            return {"error": str(e)}
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def add_keypoint_constraint(self, element1_type: str, element1_index: int,
                                 keypoint1: int, element2_type: str,
                                 element2_index: int, keypoint2: int) -> Dict[str, Any]:
        """
        Add a keypoint constraint connecting two sketch elements at specific points.

        Keypoint indices: 0=start, 1=end for lines/arcs; 0=center for circles.

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
                "element2": [element2_type, element2_index, keypoint2]
            }
        except ValueError as e:
            return {"error": str(e)}
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def close_sketch(self) -> Dict[str, Any]:
        """Close/finish the active sketch"""
        try:
            if not self.active_profile:
                return {"error": "No active sketch to close"}

            # Use correct End() flags based on whether axis of revolution is set
            if self.active_refaxis is not None:
                # Revolve profile needs igProfileClosed | igProfileRefAxisRequired
                end_flags = ProfileValidationConstants.igProfileForRevolve  # 17
            else:
                # Standard profile (extrude, etc.)
                end_flags = ProfileValidationConstants.igProfileDefault  # 0

            # Validate the profile
            try:
                status = self.active_profile.End(end_flags)
            except:
                # Some versions use different methods
                pass

            # Add to accumulated profiles for loft/sweep operations
            self.accumulated_profiles.append(self.active_profile)

            result = {
                "status": "closed",
                "sketch_id": self.active_sketch.Name if hasattr(self.active_sketch, 'Name') else "sketch",
                "has_revolution_axis": self.active_refaxis is not None,
                "accumulated_profiles": len(self.accumulated_profiles)
            }

            # NOTE: We keep active_profile valid after closing so it can be used
            # by feature operations (extrude, revolve, etc.). The profile object
            # remains valid even after End() is called.
            # Only clear it when a new sketch is created.

            return result
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_sketch_info(self) -> Dict[str, Any]:
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
            info = {"status": "active"}

            # Count elements in each collection
            collections = {
                "lines": "Lines2d",
                "circles": "Circles2d",
                "arcs": "Arcs2d",
                "ellipses": "Ellipses2d",
                "splines": "BSplineCurves2d",
                "points": "Holes2d",
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
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_active_sketch(self):
        """Get the active sketch object"""
        return self.active_profile

    def get_active_refaxis(self):
        """Get the active reference axis for revolve operations"""
        return self.active_refaxis

    def get_accumulated_profiles(self):
        """Get the list of accumulated closed profiles (for loft/sweep)."""
        return list(self.accumulated_profiles)

    def clear_accumulated_profiles(self):
        """Clear the accumulated profiles list."""
        self.accumulated_profiles.clear()

    def sketch_fillet(self, radius: float) -> Dict[str, Any]:
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

            fillet_count = 0
            # Try to fillet between consecutive line pairs
            for i in range(1, lines.Count):
                try:
                    line1 = lines.Item(i)
                    line2 = lines.Item(i + 1)
                    profile.Arcs2d.AddByFillet(line1, line2, radius)
                    fillet_count += 1
                except Exception:
                    pass

            return {
                "status": "created",
                "type": "sketch_fillet",
                "radius": radius,
                "fillet_count": fillet_count
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def sketch_chamfer(self, distance: float) -> Dict[str, Any]:
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

            chamfer_count = 0
            for i in range(1, lines.Count):
                try:
                    line1 = lines.Item(i)
                    line2 = lines.Item(i + 1)
                    profile.Lines2d.AddByChamfer(line1, line2, distance, distance)
                    chamfer_count += 1
                except Exception:
                    pass

            return {
                "status": "created",
                "type": "sketch_chamfer",
                "distance": distance,
                "chamfer_count": chamfer_count
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def sketch_offset(self, distance: float) -> Dict[str, Any]:
        """
        Create an offset copy of the sketch profile.

        Offsets all geometry in the active sketch by the given distance.

        Args:
            distance: Offset distance in meters (positive = outward)

        Returns:
            Dict with status
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            profile = self.active_profile

            # Try using the profile offset method
            try:
                profile.OffsetProfile(distance)
                return {
                    "status": "created",
                    "type": "sketch_offset",
                    "distance": distance
                }
            except Exception:
                pass

            # Fallback: manual offset of lines
            lines = profile.Lines2d
            if lines.Count == 0:
                return {"error": "No sketch geometry to offset"}

            offset_count = 0
            for i in range(1, lines.Count + 1):
                try:
                    line = lines.Item(i)
                    x1 = line.StartPoint.X
                    y1 = line.StartPoint.Y
                    x2 = line.EndPoint.X
                    y2 = line.EndPoint.Y

                    # Calculate normal offset
                    dx = x2 - x1
                    dy = y2 - y1
                    length = math.sqrt(dx * dx + dy * dy)
                    if length > 0:
                        nx = -dy / length * distance
                        ny = dx / length * distance
                        profile.Lines2d.AddBy2Points(x1 + nx, y1 + ny, x2 + nx, y2 + ny)
                        offset_count += 1
                except Exception:
                    pass

            return {
                "status": "created",
                "type": "sketch_offset",
                "distance": distance,
                "offset_lines": offset_count
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def sketch_mirror(self, axis: str = "X") -> Dict[str, Any]:
        """
        Mirror sketch geometry about an axis.

        Creates mirrored copies of all sketch elements about the
        X or Y axis.

        Args:
            axis: 'X' (mirror about X-axis, flip Y) or 'Y' (mirror about Y-axis, flip X)

        Returns:
            Dict with status
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            profile = self.active_profile
            mirror_count = 0

            # Mirror lines
            lines = profile.Lines2d
            for i in range(1, lines.Count + 1):
                try:
                    line = lines.Item(i)
                    x1 = line.StartPoint.X
                    y1 = line.StartPoint.Y
                    x2 = line.EndPoint.X
                    y2 = line.EndPoint.Y

                    if axis.upper() == "X":
                        profile.Lines2d.AddBy2Points(x1, -y1, x2, -y2)
                    else:
                        profile.Lines2d.AddBy2Points(-x1, y1, -x2, y2)
                    mirror_count += 1
                except Exception:
                    pass

            # Mirror circles
            circles = profile.Circles2d
            for i in range(1, circles.Count + 1):
                try:
                    circle = circles.Item(i)
                    cx = circle.CenterPoint.X
                    cy = circle.CenterPoint.Y
                    r = circle.Radius

                    if axis.upper() == "X":
                        profile.Circles2d.AddByCenterRadius(cx, -cy, r)
                    else:
                        profile.Circles2d.AddByCenterRadius(-cx, cy, r)
                    mirror_count += 1
                except Exception:
                    pass

            return {
                "status": "created",
                "type": "sketch_mirror",
                "axis": axis,
                "mirrored_elements": mirror_count
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def draw_construction_line(self, x1: float, y1: float, x2: float, y2: float) -> Dict[str, Any]:
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

            try:
                profile.ToggleConstruction(line)
            except Exception:
                pass

            return {
                "status": "created",
                "type": "construction_line",
                "start": [x1, y1],
                "end": [x2, y2]
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def project_edge(self, face_index: int, edge_index: int) -> Dict[str, Any]:
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
                "projected_geometry": str(type(projected).__name__) if projected else "unknown"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def include_edge(self, face_index: int, edge_index: int) -> Dict[str, Any]:
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
                "geometry_2d": str(type(result).__name__) if result else "unknown"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_sketch_constraints(self) -> Dict[str, Any]:
        """
        Get information about constraints in the active sketch.

        Returns:
            Dict with constraint count and types
        """
        try:
            if not self.active_profile:
                return {"error": "No active sketch. Call create_sketch() first"}

            profile = self.active_profile
            constraints = []

            try:
                relations = profile.Relations2d
                for i in range(1, relations.Count + 1):
                    try:
                        rel = relations.Item(i)
                        constraint_info = {"index": i - 1}
                        try:
                            constraint_info["type"] = rel.Type
                        except Exception:
                            pass
                        try:
                            constraint_info["name"] = rel.Name
                        except Exception:
                            pass
                        constraints.append(constraint_info)
                    except Exception:
                        constraints.append({"index": i - 1, "type": "unknown"})
            except Exception:
                pass

            return {
                "constraints": constraints,
                "count": len(constraints)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }
