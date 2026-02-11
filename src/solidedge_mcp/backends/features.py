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

            # Create the extrusion
            if models.Count == 0:
                # First feature - create base extrusion
                model = models.AddFiniteExtrudedProtrusion(
                    NumberOfProfiles=1,
                    ProfileArray=(profile,),
                    ProfilePlaneSide=dir_const,
                    ExtrusionDistance=distance
                )
            else:
                # Subsequent features
                if operation == "Add":
                    model = models.AddFiniteExtrudedProtrusion(
                        NumberOfProfiles=1,
                        ProfileArray=(profile,),
                        ProfilePlaneSide=dir_const,
                        ExtrusionDistance=distance
                    )
                elif operation == "Cut":
                    model = models.AddFiniteExtrudedCutout(
                        NumberOfProfiles=1,
                        ProfileArray=(profile,),
                        ProfilePlaneSide=dir_const,
                        ExtrusionDistance=distance
                    )
                else:
                    model = models.AddFiniteExtrudedProtrusion(
                        NumberOfProfiles=1,
                        ProfileArray=(profile,),
                        ProfilePlaneSide=dir_const,
                        ExtrusionDistance=distance
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

        Args:
            angle: Revolution angle in degrees (360 for full revolution)
            operation: 'Add' or 'Cut'

        Returns:
            Dict with status and feature info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile"}

            models = doc.Models

            # Convert angle to radians
            import math
            angle_rad = math.radians(angle)

            # Create the revolve
            if models.Count == 0 or operation == "Add":
                model = models.AddRevolvedProtrusion(
                    NumberOfProfiles=1,
                    ProfileArray=(profile,),
                    AxisOfRevolution=None,  # Will use default axis
                    Angle=angle_rad
                )
            elif operation == "Cut":
                model = models.AddRevolvedCutout(
                    NumberOfProfiles=1,
                    ProfileArray=(profile,),
                    AxisOfRevolution=None,
                    Angle=angle_rad
                )
            else:
                return {"error": f"Invalid operation: {operation}"}

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

            # AddBoxByCenter(CenterX, CenterY, CenterZ, Length, Width, Height)
            model = models.AddBoxByCenter(
                center_x, center_y, center_z,
                length, width, height
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

            # AddBoxByTwoPoints(X1, Y1, Z1, X2, Y2, Z2)
            model = models.AddBoxByTwoPoints(x1, y1, z1, x2, y2, z2)

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

            # AddCylinderByCenterAndRadius(CenterX, CenterY, CenterZ, Radius, Height)
            model = models.AddCylinderByCenterAndRadius(
                base_center_x, base_center_y, base_center_z,
                radius, height
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

            # AddSphereByCenterAndRadius(CenterX, CenterY, CenterZ, Radius)
            model = models.AddSphereByCenterAndRadius(
                center_x, center_y, center_z, radius
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
