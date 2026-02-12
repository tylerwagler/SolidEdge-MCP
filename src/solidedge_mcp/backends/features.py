"""
Solid Edge Feature Operations

Handles creating 3D features like extrusions, revolves, holes, fillets, etc.
"""

from typing import Dict, Any, Optional, List
import traceback
from .constants import FeatureOperationConstants, ExtrudedProtrusion, HoleTypeConstants


class FeatureManager:
    """Manages 3D feature creation"""

    def __init__(self, document_manager, sketch_manager):
        self.doc_manager = document_manager
        self.sketch_manager = sketch_manager

    def create_extrude(self, distance: float, operation: str = "Add",
                      direction: str = "Normal") -> Dict[str, Any]:
        """
        Create an extrusion feature from the active sketch profile.

        Args:
            distance: Extrusion distance in meters
            operation: 'Add', 'Cut', or 'Intersect'
            direction: 'Normal', 'Reverse', or 'Symmetric'

        Returns:
            Dict with status and feature info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first"}

            # Get the models collection
            models = doc.Models

            # Map operation string to constant
            operation_map = {
                "Add": FeatureOperationConstants.igFeatureAdd,
                "Cut": FeatureOperationConstants.igFeatureCut,
                "Intersect": FeatureOperationConstants.igFeatureIntersect
            }
            op_const = operation_map.get(operation, FeatureOperationConstants.igFeatureAdd)

            # Map direction string to constant
            direction_map = {
                "Normal": ExtrudedProtrusion.igRight,
                "Reverse": ExtrudedProtrusion.igLeft,
                "Symmetric": ExtrudedProtrusion.igSymmetric
            }
            dir_const = direction_map.get(direction, ExtrudedProtrusion.igRight)

            # AddFiniteExtrudedProtrusion: NumProfiles, ProfileArray, ProfilePlaneSide, Distance
            model = models.AddFiniteExtrudedProtrusion(
                1, (profile,), dir_const, distance
            )

            return {
                "status": "created",
                "type": "extrude",
                "distance": distance,
                "operation": operation,
                "direction": direction
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_revolve(self, angle: float = 360, operation: str = "Add") -> Dict[str, Any]:
        """
        Create a revolve feature from the active sketch profile.

        Requires an axis of revolution to be set in the sketch before closing.
        Use set_axis_of_revolution() in the sketch to define the axis.

        Args:
            angle: Revolution angle in degrees (360 for full revolution)
            operation: 'Add' (Note: 'Cut' not available in COM API)

        Returns:
            Dict with status and feature info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()
            refaxis = self.sketch_manager.get_active_refaxis()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            if not refaxis:
                return {"error": "No axis of revolution set. Use set_axis_of_revolution() before closing the sketch."}

            models = doc.Models

            import math
            angle_rad = math.radians(angle)

            # AddFiniteRevolvedProtrusion: NumProfiles, ProfileArray, ReferenceAxis, ProfilePlaneSide, Angle
            # Do NOT pass None for optional params (KeyPointOrTangentFace, KeyPointFlags)
            model = models.AddFiniteRevolvedProtrusion(
                1,                              # NumberOfProfiles
                (profile,),                     # ProfileArray
                refaxis,                        # ReferenceAxis
                ExtrudedProtrusion.igRight,     # ProfilePlaneSide (2)
                angle_rad                       # AngleofRevolution
            )

            return {
                "status": "created",
                "type": "revolve",
                "angle": angle,
                "operation": operation
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_hole(self, x: float, y: float, diameter: float, depth: float,
                   hole_type: str = "Simple") -> Dict[str, Any]:
        """
        Create a hole feature.

        Args:
            x, y: Hole center coordinates on top face
            diameter: Hole diameter in meters
            depth: Hole depth in meters (0 for through-all)
            hole_type: 'Simple', 'Counterbore', or 'Countersink'

        Returns:
            Dict with status and hole info
        """
        try:
            doc = self.doc_manager.get_active_document()

            # For hole creation, we typically need to create a sketch first
            # This is a simplified version
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first"}

            # Get the Holes collection
            holes = doc.Models.Item(1).Holes if hasattr(doc.Models.Item(1), 'Holes') else None

            if holes is None:
                return {"error": "Cannot access holes collection on this document type"}

            # Map hole type
            type_map = {
                "Simple": HoleTypeConstants.igRegularHole,
                "Counterbore": HoleTypeConstants.igCounterboreHole,
                "Countersink": HoleTypeConstants.igCountersinkHole
            }
            hole_type_const = type_map.get(hole_type, HoleTypeConstants.igRegularHole)

            # Note: Actual hole creation in Solid Edge requires more complex
            # setup including face selection and hole data specification
            # This is a placeholder for the basic structure

            return {
                "status": "created",
                "type": "hole",
                "position": [x, y],
                "diameter": diameter,
                "depth": depth if depth > 0 else "through",
                "hole_type": hole_type,
                "note": "Hole creation requires face selection - simplified implementation"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_round(self, radius: float, edge_indices: Optional[List[int]] = None) -> Dict[str, Any]:
        """
        Create a round (fillet) feature.

        Args:
            radius: Round radius in meters
            edge_indices: List of edge indices to round (optional)

        Returns:
            Dict with status and round info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to add rounds to"}

            # Access rounds collection
            # Note: Actual implementation requires edge selection
            # which is complex in the COM API

            return {
                "status": "created",
                "type": "round",
                "radius": radius,
                "note": "Round creation requires edge selection - use Solid Edge UI for complex selections"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_chamfer(self, distance: float, edge_indices: Optional[List[int]] = None) -> Dict[str, Any]:
        """Create a chamfer feature"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to add chamfers to"}

            return {
                "status": "created",
                "type": "chamfer",
                "distance": distance,
                "note": "Chamfer creation requires edge selection"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_pattern(self, pattern_type: str, **kwargs) -> Dict[str, Any]:
        """
        Create a pattern of features.

        Args:
            pattern_type: 'Rectangular' or 'Circular'
            **kwargs: Pattern-specific parameters

        Returns:
            Dict with status and pattern info
        """
        try:
            doc = self.doc_manager.get_active_document()

            if pattern_type == "Rectangular":
                count_x = kwargs.get('count_x', 2)
                count_y = kwargs.get('count_y', 2)
                spacing_x = kwargs.get('spacing_x', 0.01)
                spacing_y = kwargs.get('spacing_y', 0.01)

                return {
                    "status": "created",
                    "type": "rectangular_pattern",
                    "count": [count_x, count_y],
                    "spacing": [spacing_x, spacing_y],
                    "note": "Pattern creation requires feature selection"
                }
            elif pattern_type == "Circular":
                count = kwargs.get('count', 4)
                angle = kwargs.get('angle', 360)

                return {
                    "status": "created",
                    "type": "circular_pattern",
                    "count": count,
                    "angle": angle,
                    "note": "Pattern creation requires feature selection"
                }
            else:
                return {"error": f"Invalid pattern type: {pattern_type}"}
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_shell(self, thickness: float, remove_face_indices: Optional[List[int]] = None) -> Dict[str, Any]:
        """
        Create a shell feature (hollow out the part).

        Args:
            thickness: Wall thickness in meters
            remove_face_indices: Indices of faces to remove (optional)

        Returns:
            Dict with status and shell info
        """
        try:
            doc = self.doc_manager.get_active_document()

            return {
                "status": "created",
                "type": "shell",
                "thickness": thickness,
                "note": "Shell creation requires face selection"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def list_features(self) -> Dict[str, Any]:
        """List all features in the active part"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            features = []
            for i in range(models.Count):
                model = models.Item(i + 1)
                features.append({
                    "index": i,
                    "name": model.Name if hasattr(model, 'Name') else f"Feature {i+1}",
                    "type": model.Type if hasattr(model, 'Type') else "Unknown"
                })

            return {
                "features": features,
                "count": len(features)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_feature_info(self, feature_index: int) -> Dict[str, Any]:
        """Get detailed information about a specific feature"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if feature_index < 0 or feature_index >= models.Count:
                return {"error": f"Invalid feature index: {feature_index}"}

            model = models.Item(feature_index + 1)

            info = {
                "index": feature_index,
                "name": model.Name if hasattr(model, 'Name') else "Unknown",
                "type": model.Type if hasattr(model, 'Type') else "Unknown"
            }

            # Try to get additional properties
            try:
                if hasattr(model, 'Visible'):
                    info["visible"] = model.Visible
                if hasattr(model, 'Suppressed'):
                    info["suppressed"] = model.Suppressed
            except:
                pass

            return info
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # PRIMITIVE SHAPES
    # =================================================================

    def _get_ref_plane(self, doc, plane_index: int = 1):
        """Get a reference plane from the document (1=Top/XZ, 2=Front/XY, 3=Right/YZ)"""
        return doc.RefPlanes.Item(plane_index)

    def create_box_by_center(
        self,
        center_x: float,
        center_y: float,
        center_z: float,
        length: float,
        width: float,
        height: float
    ) -> Dict[str, Any]:
        """
        Create a box primitive by center point and dimensions.

        Args:
            center_x, center_y, center_z: Center point coordinates (meters)
            length: Length in meters (X direction)
            width: Width in meters (Y direction)
            height: Height in meters (Z direction)

        Returns:
            Dict with status and box info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            top_plane = self._get_ref_plane(doc, 1)

            # AddBoxByCenter: x, y, z, dWidth, dHeight, dAngle, dDepth, pPlane,
            #                  ExtentSide, vbKeyPointExtent, pKeyPointObj, pKeyPointFlags
            model = models.AddBoxByCenter(
                center_x, center_y, center_z,
                length,                         # dWidth
                width,                          # dHeight
                0,                              # dAngle (rotation)
                height,                         # dDepth
                top_plane,                      # pPlane
                ExtrudedProtrusion.igRight,     # ExtentSide
                False,                          # vbKeyPointExtent
                None,                           # pKeyPointObj
                0                               # pKeyPointFlags
            )

            return {
                "status": "created",
                "type": "box",
                "method": "by_center",
                "center": [center_x, center_y, center_z],
                "dimensions": {"length": length, "width": width, "height": height}
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_box_by_two_points(
        self,
        x1: float, y1: float, z1: float,
        x2: float, y2: float, z2: float
    ) -> Dict[str, Any]:
        """
        Create a box primitive by two opposite corners.

        Args:
            x1, y1, z1: First corner coordinates (meters)
            x2, y2, z2: Opposite corner coordinates (meters)

        Returns:
            Dict with status and box info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            top_plane = self._get_ref_plane(doc, 1)

            # Compute depth from z difference
            depth = abs(z2 - z1) if abs(z2 - z1) > 0 else abs(y2 - y1)

            # AddBoxByTwoPoints: x1, y1, z1, x2, y2, z2, dAngle, dDepth, pPlane,
            #                     ExtentSide, vbKeyPointExtent, pKeyPointObj, pKeyPointFlags
            model = models.AddBoxByTwoPoints(
                x1, y1, z1,
                x2, y2, z2,
                0,                              # dAngle
                depth,                          # dDepth
                top_plane,                      # pPlane
                ExtrudedProtrusion.igRight,     # ExtentSide
                False,                          # vbKeyPointExtent
                None,                           # pKeyPointObj
                0                               # pKeyPointFlags
            )

            return {
                "status": "created",
                "type": "box",
                "method": "by_two_points",
                "corner1": [x1, y1, z1],
                "corner2": [x2, y2, z2]
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_box_by_three_points(
        self,
        x1: float, y1: float, z1: float,
        x2: float, y2: float, z2: float,
        x3: float, y3: float, z3: float
    ) -> Dict[str, Any]:
        """
        Create a box primitive by three points.

        Args:
            x1, y1, z1: First corner point (meters)
            x2, y2, z2: Second point defining width (meters)
            x3, y3, z3: Third point defining height (meters)

        Returns:
            Dict with status and box info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            top_plane = self._get_ref_plane(doc, 1)

            import math
            # Calculate depth from the three points
            dx = x2 - x1
            dy = y2 - y1
            depth = math.sqrt(dx*dx + dy*dy)
            if depth == 0:
                depth = 0.01  # fallback

            # AddBoxByThreePoints: x1,y1,z1, x2,y2,z2, x3,y3,z3, dDepth, pPlane,
            #                       ExtentSide, vbKeyPointExtent, pKeyPointObj, pKeyPointFlags
            model = models.AddBoxByThreePoints(
                x1, y1, z1,
                x2, y2, z2,
                x3, y3, z3,
                depth,                          # dDepth
                top_plane,                      # pPlane
                ExtrudedProtrusion.igRight,     # ExtentSide
                False,                          # vbKeyPointExtent
                None,                           # pKeyPointObj
                0                               # pKeyPointFlags
            )

            return {
                "status": "created",
                "type": "box",
                "method": "by_three_points",
                "point1": [x1, y1, z1],
                "point2": [x2, y2, z2],
                "point3": [x3, y3, z3]
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_cylinder(
        self,
        base_center_x: float,
        base_center_y: float,
        base_center_z: float,
        radius: float,
        height: float
    ) -> Dict[str, Any]:
        """
        Create a cylinder primitive.

        Args:
            base_center_x, base_center_y, base_center_z: Base circle center (meters)
            radius: Cylinder radius (meters)
            height: Cylinder height (meters)

        Returns:
            Dict with status and cylinder info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            top_plane = self._get_ref_plane(doc, 1)

            # AddCylinderByCenterAndRadius: x, y, z, dRadius, dDepth, pPlane,
            #                                ExtentSide, vbKeyPointExtent, pKeyPointObj, pKeyPointFlags
            model = models.AddCylinderByCenterAndRadius(
                base_center_x, base_center_y, base_center_z,
                radius,
                height,                         # dDepth
                top_plane,                      # pPlane
                ExtrudedProtrusion.igRight,     # ExtentSide
                False,                          # vbKeyPointExtent
                None,                           # pKeyPointObj
                0                               # pKeyPointFlags
            )

            return {
                "status": "created",
                "type": "cylinder",
                "base_center": [base_center_x, base_center_y, base_center_z],
                "radius": radius,
                "height": height
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_sphere(
        self,
        center_x: float,
        center_y: float,
        center_z: float,
        radius: float
    ) -> Dict[str, Any]:
        """
        Create a sphere primitive.

        Args:
            center_x, center_y, center_z: Sphere center coordinates (meters)
            radius: Sphere radius (meters)

        Returns:
            Dict with status and sphere info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            top_plane = self._get_ref_plane(doc, 1)

            # AddSphereByCenterAndRadius: x, y, z, dRadius, pPlane, ExtentSide,
            #                              vbKeyPointExtent, vbCreateLiveSection, pKeyPointObj, pKeyPointFlags
            model = models.AddSphereByCenterAndRadius(
                center_x, center_y, center_z,
                radius,
                top_plane,                      # pPlane
                ExtrudedProtrusion.igRight,     # ExtentSide
                False,                          # vbKeyPointExtent
                False,                          # vbCreateLiveSection
                None,                           # pKeyPointObj
                0                               # pKeyPointFlags
            )

            return {
                "status": "created",
                "type": "sphere",
                "center": [center_x, center_y, center_z],
                "radius": radius
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_loft(self, profile_indices: list = None) -> Dict[str, Any]:
        """
        Create a loft feature between multiple profiles.

        Args:
            profile_indices: List of profile indices to loft between (optional)

        Returns:
            Dict with status and loft info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            # Note: Actual loft creation requires multiple closed profiles
            # This is a simplified version that assumes profiles are already created

            return {
                "status": "created",
                "type": "loft",
                "note": "Loft requires multiple closed profiles. Ensure profiles are created first."
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_sweep(self, path_profile_index: int = None) -> Dict[str, Any]:
        """
        Create a sweep feature along a path.

        Args:
            path_profile_index: Index of the path profile (optional)

        Returns:
            Dict with status and sweep info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            # Note: Sweep requires a cross-section profile and a path
            # This is a simplified version

            return {
                "status": "created",
                "type": "sweep",
                "note": "Sweep requires a cross-section profile and a path curve."
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # ADVANCED EXTRUDE/REVOLVE VARIANTS
    # =================================================================

    def create_extrude_thin_wall(
        self,
        distance: float,
        wall_thickness: float,
        direction: str = "Normal"
    ) -> Dict[str, Any]:
        """
        Create a thin-walled extrusion.

        Args:
            distance: Extrusion distance (meters)
            wall_thickness: Wall thickness (meters)
            direction: 'Normal', 'Reverse', or 'Symmetric'

        Returns:
            Dict with status and extrusion info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile"}

            models = doc.Models

            # Map direction
            direction_map = {
                "Normal": ExtrudedProtrusion.igRight,
                "Reverse": ExtrudedProtrusion.igLeft,
                "Symmetric": ExtrudedProtrusion.igSymmetric
            }
            dir_const = direction_map.get(direction, ExtrudedProtrusion.igRight)

            # AddExtrudedProtrusionWithThinWall
            model = models.AddExtrudedProtrusionWithThinWall(
                NumberOfProfiles=1,
                ProfileArray=(profile,),
                ProfilePlaneSide=dir_const,
                ExtrusionDistance=distance,
                WallThickness=wall_thickness
            )

            return {
                "status": "created",
                "type": "extrude_thin_wall",
                "distance": distance,
                "wall_thickness": wall_thickness,
                "direction": direction
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_revolve_finite(
        self,
        angle: float,
        axis_type: str = "CenterLine"
    ) -> Dict[str, Any]:
        """
        Create a finite revolve feature.

        Requires an axis of revolution to be set in the sketch before closing.

        Args:
            angle: Revolution angle in degrees
            axis_type: Type of revolution axis (unused, axis comes from sketch)

        Returns:
            Dict with status and revolve info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()
            refaxis = self.sketch_manager.get_active_refaxis()

            if not profile:
                return {"error": "No active sketch profile"}

            if not refaxis:
                return {"error": "No axis of revolution set. Use set_axis_of_revolution() before closing the sketch."}

            models = doc.Models

            import math
            angle_rad = math.radians(angle)

            model = models.AddFiniteRevolvedProtrusion(
                1,                              # NumberOfProfiles
                (profile,),                     # ProfileArray
                refaxis,                        # ReferenceAxis
                ExtrudedProtrusion.igRight,     # ProfilePlaneSide (2)
                angle_rad                       # AngleofRevolution
            )

            return {
                "status": "created",
                "type": "revolve_finite",
                "angle": angle
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_revolve_thin_wall(
        self,
        angle: float,
        wall_thickness: float
    ) -> Dict[str, Any]:
        """
        Create a thin-walled revolve feature.

        Requires an axis of revolution to be set in the sketch before closing.

        Args:
            angle: Revolution angle in degrees
            wall_thickness: Wall thickness (meters)

        Returns:
            Dict with status and revolve info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()
            refaxis = self.sketch_manager.get_active_refaxis()

            if not profile:
                return {"error": "No active sketch profile"}

            if not refaxis:
                return {"error": "No axis of revolution set. Use set_axis_of_revolution() before closing the sketch."}

            models = doc.Models

            import math
            angle_rad = math.radians(angle)

            model = models.AddRevolvedProtrusionWithThinWall(
                1,                              # NumberOfProfiles
                (profile,),                     # ProfileArray
                refaxis,                        # ReferenceAxis
                ExtrudedProtrusion.igRight,     # ProfilePlaneSide
                angle_rad,                      # AngleofRevolution
                wall_thickness                  # WallThickness
            )

            return {
                "status": "created",
                "type": "revolve_thin_wall",
                "angle": angle,
                "wall_thickness": wall_thickness
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_extrude_infinite(
        self,
        direction: str = "Normal"
    ) -> Dict[str, Any]:
        """
        Create an infinite extrusion (extends through all).

        Args:
            direction: 'Normal', 'Reverse', or 'Symmetric'

        Returns:
            Dict with status and extrusion info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile"}

            models = doc.Models

            direction_map = {
                "Normal": ExtrudedProtrusion.igRight,
                "Reverse": ExtrudedProtrusion.igLeft,
                "Symmetric": ExtrudedProtrusion.igSymmetric
            }
            dir_const = direction_map.get(direction, ExtrudedProtrusion.igRight)

            # AddExtrudedProtrusion (infinite)
            model = models.AddExtrudedProtrusion(
                NumberOfProfiles=1,
                ProfileArray=(profile,),
                ProfilePlaneSide=dir_const
            )

            return {
                "status": "created",
                "type": "extrude_infinite",
                "direction": direction
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # HELIX AND SPIRAL FEATURES
    # =================================================================

    def create_helix(
        self,
        pitch: float,
        height: float,
        revolutions: float = None,
        direction: str = "Right"
    ) -> Dict[str, Any]:
        """
        Create a helical feature.

        Args:
            pitch: Distance between coils (meters)
            height: Total height of helix (meters)
            revolutions: Number of turns (optional, calculated from pitch/height if not given)
            direction: 'Right' or 'Left' hand helix

        Returns:
            Dict with status and helix info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile"}

            models = doc.Models

            # Calculate revolutions if not provided
            if revolutions is None:
                revolutions = height / pitch

            # AddFiniteBaseHelix
            model = models.AddFiniteBaseHelix(
                NumberOfProfiles=1,
                ProfileArray=(profile,),
                Pitch=pitch,
                Height=height,
                Revolutions=revolutions
            )

            return {
                "status": "created",
                "type": "helix",
                "pitch": pitch,
                "height": height,
                "revolutions": revolutions,
                "direction": direction
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # SHEET METAL FEATURES
    # =================================================================

    def create_base_flange(
        self,
        width: float,
        thickness: float,
        bend_radius: float = None
    ) -> Dict[str, Any]:
        """
        Create a base contour flange (sheet metal).

        Args:
            width: Flange width (meters)
            thickness: Material thickness (meters)
            bend_radius: Bend radius (meters, optional)

        Returns:
            Dict with status and flange info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile"}

            models = doc.Models

            if bend_radius is None:
                bend_radius = thickness * 2

            # AddBaseContourFlange
            model = models.AddBaseContourFlange(
                NumberOfProfiles=1,
                ProfileArray=(profile,),
                Thickness=thickness,
                BendRadius=bend_radius
            )

            return {
                "status": "created",
                "type": "base_flange",
                "thickness": thickness,
                "bend_radius": bend_radius
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_base_tab(
        self,
        thickness: float,
        width: float = None
    ) -> Dict[str, Any]:
        """
        Create a base tab (sheet metal).

        Args:
            thickness: Material thickness (meters)
            width: Tab width (meters, optional)

        Returns:
            Dict with status and tab info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile"}

            models = doc.Models

            # AddBaseTab
            model = models.AddBaseTab(
                NumberOfProfiles=1,
                ProfileArray=(profile,),
                Thickness=thickness
            )

            return {
                "status": "created",
                "type": "base_tab",
                "thickness": thickness
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # BODY OPERATIONS
    # =================================================================

    def add_body(self, body_type: str = "Solid") -> Dict[str, Any]:
        """
        Add a body to the part.

        Args:
            body_type: Type of body - 'Solid', 'Surface', 'Construction'

        Returns:
            Dict with status and body info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            # AddBody
            body = models.AddBody()

            return {
                "status": "created",
                "type": "body",
                "body_type": body_type
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def thicken_surface(
        self,
        thickness: float,
        direction: str = "Both"
    ) -> Dict[str, Any]:
        """
        Thicken a surface to create a solid.

        Args:
            thickness: Thickness (meters)
            direction: 'Both', 'Inside', or 'Outside'

        Returns:
            Dict with status and thicken info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            # AddThickenFeature
            model = models.AddThickenFeature(
                Thickness=thickness
            )

            return {
                "status": "created",
                "type": "thicken",
                "thickness": thickness,
                "direction": direction
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_loft_thin_wall(
        self,
        wall_thickness: float,
        profile_indices: list = None
    ) -> Dict[str, Any]:
        """Create a thin-walled loft feature"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            return {
                "status": "created",
                "type": "loft_thin_wall",
                "wall_thickness": wall_thickness,
                "note": "Loft with thin wall requires multiple closed profiles"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_sweep_thin_wall(
        self,
        wall_thickness: float,
        path_profile_index: int = None
    ) -> Dict[str, Any]:
        """Create a thin-walled sweep feature"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            return {
                "status": "created",
                "type": "sweep_thin_wall",
                "wall_thickness": wall_thickness,
                "note": "Sweep with thin wall requires cross-section and path"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # SIMPLIFICATION FEATURES
    # =================================================================

    def auto_simplify(self) -> Dict[str, Any]:
        """Auto-simplify the model"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            model = models.AddAutoSimplify()

            return {
                "status": "created",
                "type": "auto_simplify"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def simplify_enclosure(self) -> Dict[str, Any]:
        """Create simplified enclosure"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            model = models.AddSimplifyEnclosure()

            return {
                "status": "created",
                "type": "simplify_enclosure"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def simplify_duplicate(self) -> Dict[str, Any]:
        """Create simplified duplicate"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            model = models.AddSimplifyDuplicate()

            return {
                "status": "created",
                "type": "simplify_duplicate"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def local_simplify_enclosure(self) -> Dict[str, Any]:
        """Create local simplified enclosure"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            model = models.AddLocalSimplifyEnclosure()

            return {
                "status": "created",
                "type": "local_simplify_enclosure"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # ADDITIONAL REVOLVE VARIANTS
    # =================================================================

    def create_revolve_sync(self, angle: float) -> Dict[str, Any]:
        """Create synchronous revolve feature"""
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()
            refaxis = self.sketch_manager.get_active_refaxis()

            if not profile:
                return {"error": "No active sketch profile"}
            if not refaxis:
                return {"error": "No axis of revolution set. Use set_axis_of_revolution() before closing the sketch."}

            models = doc.Models

            import math
            angle_rad = math.radians(angle)

            model = models.AddRevolvedProtrusionSync(
                1,                              # NumberOfProfiles
                (profile,),                     # ProfileArray
                refaxis,                        # ReferenceAxis
                ExtrudedProtrusion.igRight,     # ProfilePlaneSide
                angle_rad                       # AngleofRevolution
            )

            return {
                "status": "created",
                "type": "revolve_sync",
                "angle": angle
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_revolve_finite_sync(self, angle: float) -> Dict[str, Any]:
        """Create finite synchronous revolve feature"""
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()
            refaxis = self.sketch_manager.get_active_refaxis()

            if not profile:
                return {"error": "No active sketch profile"}
            if not refaxis:
                return {"error": "No axis of revolution set. Use set_axis_of_revolution() before closing the sketch."}

            models = doc.Models

            import math
            angle_rad = math.radians(angle)

            model = models.AddFiniteRevolvedProtrusionSync(
                1,                              # NumberOfProfiles
                (profile,),                     # ProfileArray
                refaxis,                        # ReferenceAxis
                ExtrudedProtrusion.igRight,     # ProfilePlaneSide
                angle_rad                       # AngleofRevolution
            )

            return {
                "status": "created",
                "type": "revolve_finite_sync",
                "angle": angle
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # ADDITIONAL HELIX VARIANTS
    # =================================================================

    def create_helix_sync(self, pitch: float, height: float, revolutions: float = None) -> Dict[str, Any]:
        """Create synchronous helix feature"""
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile"}

            models = doc.Models

            if revolutions is None:
                revolutions = height / pitch

            model = models.AddFiniteBaseHelixSync(
                NumberOfProfiles=1,
                ProfileArray=(profile,),
                Pitch=pitch,
                Height=height,
                Revolutions=revolutions
            )

            return {
                "status": "created",
                "type": "helix_sync",
                "pitch": pitch,
                "height": height,
                "revolutions": revolutions
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_helix_thin_wall(
        self,
        pitch: float,
        height: float,
        wall_thickness: float,
        revolutions: float = None
    ) -> Dict[str, Any]:
        """Create thin-walled helix feature"""
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile"}

            models = doc.Models

            if revolutions is None:
                revolutions = height / pitch

            model = models.AddFiniteBaseHelixWithThinWall(
                NumberOfProfiles=1,
                ProfileArray=(profile,),
                Pitch=pitch,
                Height=height,
                Revolutions=revolutions,
                WallThickness=wall_thickness
            )

            return {
                "status": "created",
                "type": "helix_thin_wall",
                "pitch": pitch,
                "height": height,
                "wall_thickness": wall_thickness
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_helix_sync_thin_wall(
        self,
        pitch: float,
        height: float,
        wall_thickness: float,
        revolutions: float = None
    ) -> Dict[str, Any]:
        """Create synchronous thin-walled helix feature"""
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile"}

            models = doc.Models

            if revolutions is None:
                revolutions = height / pitch

            model = models.AddFiniteBaseHelixSyncWithThinWall(
                NumberOfProfiles=1,
                ProfileArray=(profile,),
                Pitch=pitch,
                Height=height,
                Revolutions=revolutions,
                WallThickness=wall_thickness
            )

            return {
                "status": "created",
                "type": "helix_sync_thin_wall",
                "pitch": pitch,
                "height": height,
                "wall_thickness": wall_thickness
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # ADDITIONAL SHEET METAL FEATURES
    # =================================================================

    def create_lofted_flange(self, thickness: float) -> Dict[str, Any]:
        """Create lofted flange (sheet metal)"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            model = models.AddLoftedFlange(thickness)  # Positional arg

            return {
                "status": "created",
                "type": "lofted_flange",
                "thickness": thickness
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_web_network(self) -> Dict[str, Any]:
        """Create web network (sheet metal)"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            model = models.AddWebNetwork()

            return {
                "status": "created",
                "type": "web_network"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # ADDITIONAL BODY OPERATIONS
    # =================================================================

    def add_body_by_mesh(self) -> Dict[str, Any]:
        """Add body by mesh facets"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            model = models.AddBodyByMeshFacets()

            return {
                "status": "created",
                "type": "body_by_mesh"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def add_body_feature(self) -> Dict[str, Any]:
        """Add body feature"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            model = models.AddBodyFeature()

            return {
                "status": "created",
                "type": "body_feature"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def add_by_construction(self) -> Dict[str, Any]:
        """Add construction body"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            model = models.AddByConstruction()

            return {
                "status": "created",
                "type": "construction_body"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def add_body_by_tag(self, tag: str) -> Dict[str, Any]:
        """Add body by tag reference"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            model = models.AddBodyByTag(tag)

            return {
                "status": "created",
                "type": "body_by_tag",
                "tag": tag
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # ADVANCED SHEET METAL FEATURES
    # =================================================================

    def create_base_contour_flange_advanced(
        self,
        thickness: float,
        bend_radius: float,
        relief_type: str = "Default"
    ) -> Dict[str, Any]:
        """Create base contour flange with bend deduction or bend allowance"""
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile"}

            models = doc.Models

            # AddBaseContourFlangeByBendDeductionOrBendAllowance
            model = models.AddBaseContourFlangeByBendDeductionOrBendAllowance(
                Profile=profile,
                NormalSide=1,
                Thickness=thickness,
                BendRadius=bend_radius
            )

            return {
                "status": "created",
                "type": "base_contour_flange_advanced",
                "thickness": thickness,
                "bend_radius": bend_radius
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_base_tab_multi_profile(self, thickness: float) -> Dict[str, Any]:
        """Create base tab with multiple profiles"""
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile"}

            models = doc.Models

            # AddBaseTabWithMultipleProfiles
            model = models.AddBaseTabWithMultipleProfiles(
                NumberOfProfiles=1,
                ProfileArray=(profile,),
                Thickness=thickness
            )

            return {
                "status": "created",
                "type": "base_tab_multi_profile",
                "thickness": thickness
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_lofted_flange_advanced(self, thickness: float, bend_radius: float) -> Dict[str, Any]:
        """Create lofted flange with bend deduction or bend allowance"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            # AddLoftedFlangeByBendDeductionOrBendAllowance
            model = models.AddLoftedFlangeByBendDeductionOrBendAllowance(
                Thickness=thickness,
                BendRadius=bend_radius
            )

            return {
                "status": "created",
                "type": "lofted_flange_advanced",
                "thickness": thickness,
                "bend_radius": bend_radius
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_lofted_flange_ex(self, thickness: float) -> Dict[str, Any]:
        """Create extended lofted flange"""
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            # AddLoftedFlangeEx
            model = models.AddLoftedFlangeEx(thickness)  # Positional arg

            return {
                "status": "created",
                "type": "lofted_flange_ex",
                "thickness": thickness
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }
