"""
Solid Edge Sketching Operations

Handles creating and manipulating 2D sketches.
"""

from typing import Dict, Any, Optional, Tuple
import traceback
import math
from .constants import RefPlaneConstants, ProfileValidationConstants


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

    def add_constraint(self, constraint_type: str, elements: list) -> Dict[str, Any]:
        """Add a geometric constraint to sketch elements"""
        try:
            if not self.active_profile:
                return {"error": "No active sketch"}

            relations = self.active_profile.Relations2d

            # This is a simplified version - actual implementation would vary
            # based on constraint type
            return {
                "status": "constraint_added",
                "type": constraint_type,
                "note": "Constraint functionality requires specific element references"
            }
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
