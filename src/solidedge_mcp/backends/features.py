"""
Solid Edge Feature Operations

Handles creating 3D features like extrusions, revolves, holes, fillets, etc.
"""

from typing import Dict, Any, Optional, List
import traceback
import pythoncom
from win32com.client import VARIANT
from .constants import (
    DirectionConstants,
    ExtentTypeConstants,
    FaceQueryConstants,
    FeatureOperationConstants,
    KeyPointExtentConstants,
    LoftSweepConstants,
    OffsetSideConstants,
    TreatmentTypeConstants,
    DraftSideConstants,
    TreatmentCrownTypeConstants,
    TreatmentCrownSideConstants,
    TreatmentCrownCurvatureSideConstants,
)


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
                "Normal": DirectionConstants.igRight,
                "Reverse": DirectionConstants.igLeft,
                "Symmetric": DirectionConstants.igSymmetric
            }
            dir_const = direction_map.get(direction, DirectionConstants.igRight)

            # AddFiniteExtrudedProtrusion: NumProfiles, ProfileArray, ProfilePlaneSide, Distance
            model = models.AddFiniteExtrudedProtrusion(
                1, (profile,), dir_const, distance
            )

            # Clear accumulated profiles (consumed by this feature)
            self.sketch_manager.clear_accumulated_profiles()

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
                DirectionConstants.igRight,     # ProfilePlaneSide (2)
                angle_rad                       # AngleofRevolution
            )

            # Clear accumulated profiles (consumed by this feature)
            self.sketch_manager.clear_accumulated_profiles()

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
                   hole_type: str = "Simple", plane_index: int = 1,
                   direction: str = "Normal") -> Dict[str, Any]:
        """
        Create a hole feature (circular cutout).

        Creates a circular cutout at (x, y) on a reference plane using
        ExtrudedCutouts.AddFiniteMulti for reliable geometry creation.

        Args:
            x, y: Hole center coordinates on the sketch plane (meters)
            diameter: Hole diameter in meters
            depth: Hole depth in meters
            hole_type: 'Simple' (only type currently supported)
            plane_index: Reference plane index (1=Top/XZ, 2=Front/XY, 3=Right/YZ)
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with status and hole info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            radius = diameter / 2.0

            # Map direction
            dir_const = DirectionConstants.igRight  # Normal
            if direction == "Reverse":
                dir_const = DirectionConstants.igLeft

            # Create a circular profile on the specified plane
            ps = doc.ProfileSets.Add()
            plane = doc.RefPlanes.Item(plane_index)
            profile = ps.Profiles.Add(plane)
            profile.Circles2d.AddByCenterRadius(x, y, radius)
            profile.End(0)

            # Use ExtrudedCutouts.AddFiniteMulti for reliable hole creation
            cutouts = model.ExtrudedCutouts
            cutout = cutouts.AddFiniteMulti(1, (profile,), dir_const, depth)

            return {
                "status": "created",
                "type": "hole",
                "position": [x, y],
                "diameter": diameter,
                "depth": depth,
                "hole_type": hole_type
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_round(self, radius: float) -> Dict[str, Any]:
        """
        Create a round (fillet) feature on all body edges.

        Rounds all edges of the body using model.Rounds.Add(). All edges
        are grouped as one edge set with a single radius value.

        Args:
            radius: Round radius in meters

        Returns:
            Dict with status and round info
        """
        try:
            from win32com.client import VARIANT
            import pythoncom

            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to add rounds to"}

            model = models.Item(1)
            body = model.Body

            # Collect all edges from all body faces
            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if faces.Count == 0:
                return {"error": "No faces found on body"}

            edge_list = []
            for fi in range(1, faces.Count + 1):
                face = faces.Item(fi)
                face_edges = face.Edges
                if not hasattr(face_edges, 'Count'):
                    continue
                for ei in range(1, face_edges.Count + 1):
                    edge_list.append(face_edges.Item(ei))

            if not edge_list:
                return {"error": "No edges found on body"}

            # Group all edges as one edge set with one radius
            edge_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, edge_list)
            radius_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [radius])

            rounds = model.Rounds
            rnd = rounds.Add(1, edge_arr, radius_arr)

            return {
                "status": "created",
                "type": "round",
                "radius": radius,
                "edge_count": len(edge_list)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_chamfer(self, distance: float) -> Dict[str, Any]:
        """
        Create an equal-setback chamfer on all body edges.

        Chamfers all edges of the body using model.Chamfers.AddEqualSetback().

        Args:
            distance: Chamfer setback distance in meters

        Returns:
            Dict with status and chamfer info
        """
        try:
            from win32com.client import VARIANT
            import pythoncom

            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to add chamfers to"}

            model = models.Item(1)
            body = model.Body

            # Collect all edges from all body faces
            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if faces.Count == 0:
                return {"error": "No faces found on body"}

            edge_list = []
            for fi in range(1, faces.Count + 1):
                face = faces.Item(fi)
                face_edges = face.Edges
                if not hasattr(face_edges, 'Count'):
                    continue
                for ei in range(1, face_edges.Count + 1):
                    edge_list.append(face_edges.Item(ei))

            if not edge_list:
                return {"error": "No edges found on body"}

            chamfers = model.Chamfers
            chamfer = chamfers.AddEqualSetback(len(edge_list), edge_list, distance)

            return {
                "status": "created",
                "type": "chamfer",
                "distance": distance,
                "edge_count": len(edge_list)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_pattern(self, pattern_type: str, **kwargs) -> Dict[str, Any]:
        """
        Create a pattern of features.

        Note: Feature patterns require SAFEARRAY(IDispatch) marshaling of feature
        objects which is not currently supported via COM late binding. Use assembly-level
        component patterns (pattern_component) instead.

        Args:
            pattern_type: 'Rectangular' or 'Circular'
            **kwargs: Pattern-specific parameters

        Returns:
            Dict with error explaining limitation
        """
        return {
            "error": "Feature patterns (model.Patterns) require SAFEARRAY marshaling of "
                     "feature objects which is not supported via COM late binding. "
                     "Use assembly-level pattern_component() for component patterns instead.",
            "pattern_type": pattern_type
        }

    def create_shell(self, thickness: float, remove_face_indices: Optional[List[int]] = None) -> Dict[str, Any]:
        """
        Create a shell feature (hollow out the part).

        Note: Shell (Thinwalls) requires face selection for open faces which cannot
        be reliably automated via COM late binding. The Thinwalls.Add method requires
        complex VARIANT parameters for face arrays.

        Args:
            thickness: Wall thickness in meters
            remove_face_indices: Indices of faces to remove (optional)

        Returns:
            Dict with error explaining limitation
        """
        return {
            "error": "Shell (Thinwalls) feature requires face selection for open faces "
                     "which cannot be reliably automated via COM. Use the Solid Edge UI "
                     "to create shell features.",
            "thickness": thickness
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
                DirectionConstants.igRight,     # ExtentSide
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
                DirectionConstants.igRight,     # ExtentSide
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
                DirectionConstants.igRight,     # ExtentSide
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
                DirectionConstants.igRight,     # ExtentSide
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
                DirectionConstants.igRight,     # ExtentSide
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

    def _make_loft_variant_arrays(self, profiles):
        """Create properly typed VARIANT arrays for loft/sweep COM calls.

        COM requires explicit VARIANT typing for SAFEARRAY parameters.
        Python's automatic marshaling does not produce correct types for nested arrays.

        Args:
            profiles: List of profile COM objects

        Returns:
            Tuple of (v_profiles, v_types, v_origins) VARIANT arrays
        """
        v_profiles = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, profiles)
        v_types = VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_I4,
            [LoftSweepConstants.igProfileBasedCrossSection] * len(profiles)
        )
        v_origins = VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
            [VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0])
             for _ in profiles]
        )
        return v_profiles, v_types, v_origins

    def create_loft(self, profile_indices: list = None) -> Dict[str, Any]:
        """
        Create a loft feature between multiple profiles.

        Uses accumulated profiles from close_sketch() calls. Create 2+ sketches
        on different parallel planes, close each one, then call create_loft().

        Args:
            profile_indices: Optional list of profile indices to select from
                accumulated profiles. If None, uses all accumulated profiles.

        Returns:
            Dict with status and loft info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            # Get accumulated profiles from sketch manager
            all_profiles = self.sketch_manager.get_accumulated_profiles()

            if profile_indices is not None:
                profiles = [all_profiles[i] for i in profile_indices]
            else:
                profiles = all_profiles

            if len(profiles) < 2:
                return {
                    "error": f"Loft requires at least 2 profiles, got {len(profiles)}. "
                    "Create sketches on different planes and close each one before calling create_loft()."
                }

            v_profiles, v_types, v_origins = self._make_loft_variant_arrays(profiles)

            # Try LoftedProtrusions.AddSimple first (works when a base feature exists)
            try:
                model = models.Item(1)
                lp = model.LoftedProtrusions
                loft = lp.AddSimple(
                    len(profiles), v_profiles, v_types, v_origins,
                    DirectionConstants.igRight, ExtentTypeConstants.igNone, ExtentTypeConstants.igNone
                )
                self.sketch_manager.clear_accumulated_profiles()
                return {
                    "status": "created",
                    "type": "loft",
                    "num_profiles": len(profiles),
                    "method": "LoftedProtrusions.AddSimple"
                }
            except Exception:
                pass

            # Fall back to models.AddLoftedProtrusion (works as initial feature)
            v_seg = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, [])
            model = models.AddLoftedProtrusion(
                len(profiles), v_profiles, v_types, v_origins,
                v_seg,                                  # SegmentMaps (empty)
                DirectionConstants.igRight,             # MaterialSide
                ExtentTypeConstants.igNone, 0.0, None,  # Start extent
                ExtentTypeConstants.igNone, 0.0, None,  # End extent
                ExtentTypeConstants.igNone, 0.0,         # Start tangent
                ExtentTypeConstants.igNone, 0.0,         # End tangent
            )

            self.sketch_manager.clear_accumulated_profiles()
            return {
                "status": "created",
                "type": "loft",
                "num_profiles": len(profiles),
                "method": "models.AddLoftedProtrusion"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_sweep(self, path_profile_index: int = None) -> Dict[str, Any]:
        """
        Create a sweep feature along a path.

        Requires at least 2 accumulated profiles: the first is the path (open profile),
        and the second is the cross-section (closed profile). Create the path sketch
        first (open, e.g. a line or arc), then create the cross-section sketch
        (closed, e.g. a circle) on a plane perpendicular to the path start.

        Args:
            path_profile_index: Index of the path profile in accumulated profiles
                (default: 0, the first accumulated profile)

        Returns:
            Dict with status and sweep info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            all_profiles = self.sketch_manager.get_accumulated_profiles()

            if len(all_profiles) < 2:
                return {
                    "error": f"Sweep requires at least 2 profiles (path + cross-section), "
                    f"got {len(all_profiles)}. Create a path sketch and a cross-section sketch first."
                }

            path_idx = path_profile_index if path_profile_index is not None else 0
            path_profile = all_profiles[path_idx]
            cross_sections = [p for i, p in enumerate(all_profiles) if i != path_idx]

            _CS = LoftSweepConstants.igProfileBasedCrossSection

            # Path arrays
            v_paths = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [path_profile])
            v_path_types = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I4, [_CS])

            # Cross-section arrays
            v_sections = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, cross_sections)
            v_section_types = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_I4,
                [_CS] * len(cross_sections)
            )
            v_origins = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
                [VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0])
                 for _ in cross_sections]
            )
            v_seg = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, [])

            # AddSweptProtrusion: 15 required params
            model = models.AddSweptProtrusion(
                1, v_paths, v_path_types,                   # Path (1 curve)
                len(cross_sections), v_sections, v_section_types, v_origins, v_seg,  # Sections
                DirectionConstants.igRight,                  # MaterialSide
                ExtentTypeConstants.igNone, 0.0, None,       # Start extent
                ExtentTypeConstants.igNone, 0.0, None,       # End extent
            )

            self.sketch_manager.clear_accumulated_profiles()
            return {
                "status": "created",
                "type": "sweep",
                "num_cross_sections": len(cross_sections),
                "method": "models.AddSweptProtrusion"
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
                "Normal": DirectionConstants.igRight,
                "Reverse": DirectionConstants.igLeft,
                "Symmetric": DirectionConstants.igSymmetric
            }
            dir_const = direction_map.get(direction, DirectionConstants.igRight)

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
                DirectionConstants.igRight,     # ProfilePlaneSide (2)
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
                DirectionConstants.igRight,     # ProfilePlaneSide
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
                "Normal": DirectionConstants.igRight,
                "Reverse": DirectionConstants.igLeft,
                "Symmetric": DirectionConstants.igSymmetric
            }
            dir_const = direction_map.get(direction, DirectionConstants.igRight)

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

    def create_extruded_surface(self, distance: float,
                                direction: str = "Normal",
                                end_caps: bool = True) -> Dict[str, Any]:
        """
        Create an extruded surface (construction geometry, not solid body).

        Extrudes the active sketch profile as a surface rather than a solid.
        Surfaces are useful as construction geometry for trimming, splitting,
        or as reference faces.

        Args:
            distance: Extrusion distance in meters
            direction: 'Normal' or 'Symmetric'
            end_caps: If True, close the surface ends

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile"}

            constructions = doc.Constructions
            extruded_surfaces = constructions.ExtrudedSurfaces

            # Build profile array
            profile_array = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

            depth1 = distance
            depth2 = distance if direction == "Symmetric" else 0.0
            side1 = DirectionConstants.igRight
            side2 = DirectionConstants.igLeft if direction == "Symmetric" else DirectionConstants.igRight

            extruded_surface = extruded_surfaces.Add(
                1,                      # NumberOfProfiles
                profile_array,          # ProfileArray
                ExtentTypeConstants.igFinite,  # ExtentType1
                side1,                  # ExtentSide1
                depth1,                                          # FiniteDepth1
                None,                                            # KeyPointOrTangentFace1
                KeyPointExtentConstants.igTangentNormal,         # KeyPointFlags1
                None,                                            # FromFaceOrRefPlane
                OffsetSideConstants.seOffsetNone,                # FromFaceOffsetSide
                0.0,                                             # FromFaceOffsetDistance
                TreatmentTypeConstants.seTreatmentNone,          # TreatmentType1
                DraftSideConstants.seDraftNone,                  # TreatmentDraftSide1
                0.0,                                             # TreatmentDraftAngle1
                TreatmentCrownTypeConstants.seTreatmentCrownByOffset,       # TreatmentCrownType1
                TreatmentCrownSideConstants.seTreatmentCrownSideInside,     # TreatmentCrownSide1
                TreatmentCrownCurvatureSideConstants.seTreatmentCrownCurvatureInside,  # TreatmentCrownCurvatureSide1
                0.0,                                             # TreatmentCrownRadiusOrOffset1
                0.0,                                             # TreatmentCrownTakeOffAngle1
                ExtentTypeConstants.igFinite,                    # ExtentType2
                side2,                                           # ExtentSide2
                depth2,                                          # FiniteDepth2
                None,                                            # KeyPointOrTangentFace2
                KeyPointExtentConstants.igTangentNormal,         # KeyPointFlags2
                None,                                            # ToFaceOrRefPlane
                OffsetSideConstants.seOffsetNone,                # ToFaceOffsetSide
                0.0,                                             # ToFaceOffsetDistance
                TreatmentTypeConstants.seTreatmentNone,          # TreatmentType2
                DraftSideConstants.seDraftNone,                  # TreatmentDraftSide2
                0.0,                                             # TreatmentDraftAngle2
                TreatmentCrownTypeConstants.seTreatmentCrownByOffset,       # TreatmentCrownType2
                TreatmentCrownSideConstants.seTreatmentCrownSideInside,     # TreatmentCrownSide2
                TreatmentCrownCurvatureSideConstants.seTreatmentCrownCurvatureInside,  # TreatmentCrownCurvatureSide2
                0.0,                                             # TreatmentCrownRadiusOrOffset2
                0.0,                                             # TreatmentCrownTakeOffAngle2
                end_caps                                         # WantEndCaps
            )

            return {
                "status": "created",
                "type": "extruded_surface",
                "distance": distance,
                "direction": direction,
                "end_caps": end_caps
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
        """
        Create a thin-walled loft feature between multiple profiles.

        Uses accumulated profiles from close_sketch() calls.

        Args:
            wall_thickness: Wall thickness in meters
            profile_indices: Optional list of profile indices to select

        Returns:
            Dict with status and loft info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            all_profiles = self.sketch_manager.get_accumulated_profiles()

            if profile_indices is not None:
                profiles = [all_profiles[i] for i in profile_indices]
            else:
                profiles = all_profiles

            if len(profiles) < 2:
                return {
                    "error": f"Loft requires at least 2 profiles, got {len(profiles)}."
                }

            v_profiles, v_types, v_origins = self._make_loft_variant_arrays(profiles)
            v_seg = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, [])

            model = models.AddLoftedProtrusionWithThinWall(
                len(profiles), v_profiles, v_types, v_origins,
                v_seg,                                          # SegmentMaps
                DirectionConstants.igRight,                     # MaterialSide
                ExtentTypeConstants.igNone, 0.0, None,          # Start extent
                ExtentTypeConstants.igNone, 0.0, None,          # End extent
                ExtentTypeConstants.igNone, 0.0,                # Start tangent
                ExtentTypeConstants.igNone, 0.0,                # End tangent
                wall_thickness,     # WallThickness
            )

            self.sketch_manager.clear_accumulated_profiles()
            return {
                "status": "created",
                "type": "loft_thin_wall",
                "wall_thickness": wall_thickness,
                "num_profiles": len(profiles)
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
        """
        Create a thin-walled sweep feature along a path.

        Uses accumulated profiles: first is path (open), rest are cross-sections (closed).

        Args:
            wall_thickness: Wall thickness in meters
            path_profile_index: Index of the path profile (default: 0)

        Returns:
            Dict with status and sweep info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            all_profiles = self.sketch_manager.get_accumulated_profiles()

            if len(all_profiles) < 2:
                return {
                    "error": f"Sweep requires at least 2 profiles (path + cross-section), "
                    f"got {len(all_profiles)}."
                }

            path_idx = path_profile_index if path_profile_index is not None else 0
            path_profile = all_profiles[path_idx]
            cross_sections = [p for i, p in enumerate(all_profiles) if i != path_idx]

            _CS = LoftSweepConstants.igProfileBasedCrossSection

            v_paths = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [path_profile])
            v_path_types = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I4, [_CS])
            v_sections = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, cross_sections)
            v_section_types = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_I4,
                [_CS] * len(cross_sections)
            )
            v_origins = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
                [VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0])
                 for _ in cross_sections]
            )
            v_seg = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, [])

            model = models.AddSweptProtrusionWithThinWall(
                1, v_paths, v_path_types,
                len(cross_sections), v_sections, v_section_types, v_origins, v_seg,
                DirectionConstants.igRight,
                ExtentTypeConstants.igNone, 0.0, None,
                ExtentTypeConstants.igNone, 0.0, None,
                wall_thickness,
            )

            self.sketch_manager.clear_accumulated_profiles()
            return {
                "status": "created",
                "type": "sweep_thin_wall",
                "wall_thickness": wall_thickness,
                "num_cross_sections": len(cross_sections)
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
                DirectionConstants.igRight,     # ProfilePlaneSide
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
                DirectionConstants.igRight,     # ProfilePlaneSide
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
    # CUTOUT OPERATIONS
    # =================================================================

    def create_extruded_cutout(self, distance: float, direction: str = "Normal") -> Dict[str, Any]:
        """
        Create an extruded cutout (cut) through the part using the active sketch profile.

        Uses model.ExtrudedCutouts.AddFiniteMulti(NumProfiles, ProfileArray, PlaneSide, Depth).
        Requires an existing base feature and a closed sketch profile.

        Args:
            distance: Cutout depth in meters
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with status and cutout info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            direction_map = {
                "Normal": DirectionConstants.igRight,
                "Reverse": DirectionConstants.igLeft,
            }
            dir_const = direction_map.get(direction, DirectionConstants.igRight)

            cutouts = model.ExtrudedCutouts
            cutout = cutouts.AddFiniteMulti(1, (profile,), dir_const, distance)

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "extruded_cutout",
                "distance": distance,
                "direction": direction
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_extruded_cutout_through_all(self, direction: str = "Normal") -> Dict[str, Any]:
        """
        Create an extruded cutout that goes through the entire part.

        Uses model.ExtrudedCutouts.AddThroughAllMulti(NumProfiles, ProfileArray, PlaneSide).
        Requires an existing base feature and a closed sketch profile.

        Args:
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with status and cutout info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            direction_map = {
                "Normal": DirectionConstants.igRight,
                "Reverse": DirectionConstants.igLeft,
            }
            dir_const = direction_map.get(direction, DirectionConstants.igRight)

            cutouts = model.ExtrudedCutouts
            cutout = cutouts.AddThroughAllMulti(1, (profile,), dir_const)

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "extruded_cutout_through_all",
                "direction": direction
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_revolved_cutout(self, angle: float = 360) -> Dict[str, Any]:
        """
        Create a revolved cutout (cut) in the part using the active sketch profile.

        Uses model.RevolvedCutouts.AddFiniteMulti(NumProfiles, ProfileArray, RefAxis, PlaneSide, Angle).
        Requires an existing base feature, a closed sketch profile, and an axis of revolution.

        Args:
            angle: Revolution angle in degrees (360 for full revolution)

        Returns:
            Dict with status and cutout info
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
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            import math
            angle_rad = math.radians(angle)

            cutouts = model.RevolvedCutouts
            cutout = cutouts.AddFiniteMulti(
                1,                              # NumberOfProfiles
                (profile,),                     # ProfileArray
                refaxis,                        # ReferenceAxis
                DirectionConstants.igRight,     # ProfilePlaneSide
                angle_rad                       # AngleOfRevolution
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "revolved_cutout",
                "angle": angle
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_normal_cutout(self, distance: float, direction: str = "Normal") -> Dict[str, Any]:
        """
        Create a normal cutout (cut) through the part using the active sketch profile.

        Uses model.NormalCutouts.AddFiniteMulti(NumProfiles, ProfileArray, PlaneSide, Depth).
        A normal cutout extrudes the profile perpendicular to the sketch plane face,
        following the surface normal rather than a fixed direction.
        Requires an existing base feature and a closed sketch profile.

        Args:
            distance: Cutout depth in meters
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with status and cutout info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            direction_map = {
                "Normal": DirectionConstants.igRight,
                "Reverse": DirectionConstants.igLeft,
            }
            dir_const = direction_map.get(direction, DirectionConstants.igRight)

            cutouts = model.NormalCutouts
            cutout = cutouts.AddFiniteMulti(1, (profile,), dir_const, distance)

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "normal_cutout",
                "distance": distance,
                "direction": direction
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_lofted_cutout(self, profile_indices: list = None) -> Dict[str, Any]:
        """
        Create a lofted cutout between multiple profiles.

        Uses accumulated profiles from close_sketch() calls. Create 2+ sketches
        on different parallel planes, close each one, then call create_lofted_cutout().
        Requires an existing base feature (cutout removes material).

        Uses model.LoftedCutouts.AddSimple(count, profiles, types, origins, side, startTan, endTan).

        Args:
            profile_indices: Optional list of profile indices to select from
                accumulated profiles. If None, uses all accumulated profiles.

        Returns:
            Dict with status and lofted cutout info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            # Get accumulated profiles from sketch manager
            all_profiles = self.sketch_manager.get_accumulated_profiles()

            if profile_indices is not None:
                profiles = [all_profiles[i] for i in profile_indices]
            else:
                profiles = all_profiles

            if len(profiles) < 2:
                return {
                    "error": f"Lofted cutout requires at least 2 profiles, got {len(profiles)}. "
                    "Create sketches on different planes and close each one before calling create_lofted_cutout()."
                }

            v_profiles, v_types, v_origins = self._make_loft_variant_arrays(profiles)

            lc = model.LoftedCutouts
            cutout = lc.AddSimple(
                len(profiles), v_profiles, v_types, v_origins,
                DirectionConstants.igRight, ExtentTypeConstants.igNone, ExtentTypeConstants.igNone
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "lofted_cutout",
                "num_profiles": len(profiles),
                "method": "LoftedCutouts.AddSimple"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # MIRROR COPY
    # =================================================================

    def create_mirror(self, feature_name: str, mirror_plane_index: int) -> Dict[str, Any]:
        """
        Create a mirror copy of a feature across a reference plane.

        Note: MirrorCopies via COM has known limitations. The ordered-mode
        Add() method creates a feature object but doesn't persist geometry.
        AddSync() persists the feature tree entry but may not compute geometry.
        This is a known Solid Edge COM API limitation.

        Args:
            feature_name: Name of the feature to mirror (from list_features)
            mirror_plane_index: 1-based index of the mirror plane
                (1=Top/XZ, 2=Front/XY, 3=Right/YZ, or higher for user planes)

        Returns:
            Dict with status and mirror info
        """
        try:
            import win32com.client as win32

            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists."}

            model = models.Item(1)

            # Find the feature by name in DesignEdgebarFeatures
            features = doc.DesignEdgebarFeatures
            target_feature = None
            for i in range(1, features.Count + 1):
                f = features.Item(i)
                if f.Name == feature_name:
                    target_feature = f
                    break

            if target_feature is None:
                names = []
                for i in range(1, features.Count + 1):
                    names.append(features.Item(i).Name)
                return {
                    "error": f"Feature '{feature_name}' not found.",
                    "available_features": names
                }

            # Get the mirror plane
            ref_planes = doc.RefPlanes
            if mirror_plane_index < 1 or mirror_plane_index > ref_planes.Count:
                return {"error": f"Invalid plane index: {mirror_plane_index}. Count: {ref_planes.Count}"}

            mirror_plane = ref_planes.Item(mirror_plane_index)

            # Use AddSync which persists the feature tree entry
            mc = win32.gencache.EnsureDispatch(model.MirrorCopies)
            mirror = mc.AddSync(1, [target_feature], mirror_plane, False)

            return {
                "status": "created",
                "type": "mirror_copy",
                "feature": feature_name,
                "mirror_plane": mirror_plane_index,
                "name": mirror.Name if hasattr(mirror, 'Name') else None,
                "note": "Mirror feature created via AddSync. Geometry may require manual verification in Solid Edge UI."
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # REFERENCE PLANE CREATION
    # =================================================================

    def create_ref_plane_by_offset(self, parent_plane_index: int, distance: float,
                                    normal_side: str = "Normal") -> Dict[str, Any]:
        """
        Create a reference plane parallel to an existing plane at an offset distance.

        Uses RefPlanes.AddParallelByDistance(ParentPlane, Distance, NormalSide).
        Useful for creating sketches at different heights/positions.

        Args:
            parent_plane_index: Index of parent plane (1=Top/XZ, 2=Front/XY, 3=Right/YZ,
                                or higher for user-created planes)
            distance: Offset distance in meters
            normal_side: 'Normal' (igRight=2) or 'Reverse' (igLeft=1)

        Returns:
            Dict with status and new plane index
        """
        try:
            doc = self.doc_manager.get_active_document()
            ref_planes = doc.RefPlanes

            if parent_plane_index < 1 or parent_plane_index > ref_planes.Count:
                return {"error": f"Invalid plane index: {parent_plane_index}. Count: {ref_planes.Count}"}

            parent = ref_planes.Item(parent_plane_index)

            side_map = {
                "Normal": DirectionConstants.igRight,
                "Reverse": DirectionConstants.igLeft,
            }
            side_const = side_map.get(normal_side, DirectionConstants.igRight)

            new_plane = ref_planes.AddParallelByDistance(parent, distance, side_const)

            return {
                "status": "created",
                "type": "reference_plane",
                "method": "parallel_by_distance",
                "parent_plane": parent_plane_index,
                "distance": distance,
                "normal_side": normal_side,
                "new_plane_index": ref_planes.Count
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # SELECTIVE ROUNDS AND CHAMFERS (ON SPECIFIC FACE)
    # =================================================================

    def create_round_on_face(self, radius: float, face_index: int) -> Dict[str, Any]:
        """
        Create a round (fillet) on edges of a specific face.

        Unlike create_round() which rounds all edges, this targets only
        the edges of a single face. Use get_body_faces() to find face indices.

        Args:
            radius: Round radius in meters
            face_index: 0-based face index (from get_body_faces)

        Returns:
            Dict with status and round info
        """
        try:
            from win32com.client import VARIANT
            import pythoncom

            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to add rounds to"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)
            face_edges = face.Edges
            if not hasattr(face_edges, 'Count') or face_edges.Count == 0:
                return {"error": f"Face {face_index} has no edges"}

            edge_list = []
            for ei in range(1, face_edges.Count + 1):
                edge_list.append(face_edges.Item(ei))

            edge_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, edge_list)
            radius_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [radius])

            rounds = model.Rounds
            rnd = rounds.Add(1, edge_arr, radius_arr)

            return {
                "status": "created",
                "type": "round",
                "radius": radius,
                "face_index": face_index,
                "edge_count": len(edge_list)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_chamfer_on_face(self, distance: float, face_index: int) -> Dict[str, Any]:
        """
        Create a chamfer on edges of a specific face.

        Unlike create_chamfer() which chamfers all edges, this targets only
        the edges of a single face. Use get_body_faces() to find face indices.

        Args:
            distance: Chamfer setback distance in meters
            face_index: 0-based face index (from get_body_faces)

        Returns:
            Dict with status and chamfer info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to add chamfers to"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)
            face_edges = face.Edges
            if not hasattr(face_edges, 'Count') or face_edges.Count == 0:
                return {"error": f"Face {face_index} has no edges"}

            edge_list = []
            for ei in range(1, face_edges.Count + 1):
                edge_list.append(face_edges.Item(ei))

            chamfers = model.Chamfers
            chamfer = chamfers.AddEqualSetback(len(edge_list), edge_list, distance)

            return {
                "status": "created",
                "type": "chamfer",
                "distance": distance,
                "face_index": face_index,
                "edge_count": len(edge_list)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def delete_faces(self, face_indices: List[int]) -> Dict[str, Any]:
        """
        Delete faces from the model body.

        Uses model.DeleteFaces collection to remove specified faces.
        Useful for creating openings or removing geometry.

        Args:
            face_indices: List of 0-based face indices to delete

        Returns:
            Dict with status and deletion info
        """
        try:
            from win32com.client import VARIANT
            import pythoncom

            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to delete faces from"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if faces.Count == 0:
                return {"error": "No faces on body"}

            face_objs = []
            for idx in face_indices:
                if idx < 0 or idx >= faces.Count:
                    return {"error": f"Invalid face index: {idx}. Body has {faces.Count} faces."}
                face_objs.append(faces.Item(idx + 1))

            delete_faces = model.DeleteFaces
            result = delete_faces.Add(len(face_objs), face_objs)

            return {
                "status": "created",
                "type": "delete_faces",
                "face_count": len(face_indices),
                "face_indices": face_indices
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

    # =================================================================
    # EMBOSS AND FLANGE
    # =================================================================

    def create_emboss(
        self,
        face_indices: list,
        clearance: float = 0.001,
        thickness: float = 0.0,
        thicken: bool = False,
        default_side: bool = True,
    ) -> Dict[str, Any]:
        """
        Create an emboss feature using face geometry as tools.

        Uses selected faces as embossing tool geometry on the target body.
        Requires an existing base feature.

        Args:
            face_indices: List of 0-based face indices to use as emboss tools
            clearance: Clearance in meters (default 0.001)
            thickness: Thickness in meters (default 0.0)
            thicken: Enable thickening (default False)
            default_side: Default side (default True)

        Returns:
            Dict with status and emboss info
        """
        try:
            from win32com.client import VARIANT
            import pythoncom

            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            body = model.Body

            if not face_indices:
                return {"error": "face_indices must contain at least one face index."}

            faces = body.Faces(FaceQueryConstants.igQueryAll)

            face_list = []
            for fi in face_indices:
                if fi < 0 or fi >= faces.Count:
                    return {"error": f"Invalid face index: {fi}. Body has {faces.Count} faces."}
                face_list.append(faces.Item(fi + 1))

            tools_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, face_list)

            emboss_features = model.EmbossFeatures
            emboss = emboss_features.Add(
                body, len(face_list), tools_arr,
                thicken, default_side, clearance, thickness
            )

            return {
                "status": "created",
                "type": "emboss",
                "face_count": len(face_list),
                "clearance": clearance,
                "thickness": thickness,
                "thicken": thicken,
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_flange(
        self,
        face_index: int,
        edge_index: int,
        flange_length: float,
        side: str = "Right",
        inside_radius: float = None,
        bend_angle: float = None,
    ) -> Dict[str, Any]:
        """
        Create a flange feature on a sheet metal edge.

        Adds a flange to the specified edge of a sheet metal body.
        Requires an existing sheet metal base feature.

        Args:
            face_index: 0-based face index containing the target edge
            edge_index: 0-based edge index within that face
            flange_length: Flange length in meters
            side: 'Left' (1), 'Right' (2), or 'Both' (6)
            inside_radius: Bend inside radius in meters (optional)
            bend_angle: Bend angle in degrees (optional)

        Returns:
            Dict with status and flange info
        """
        try:
            import math

            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists. Create a sheet metal base feature first."}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)
            face_edges = face.Edges
            if not hasattr(face_edges, 'Count') or face_edges.Count == 0:
                return {"error": f"Face {face_index} has no edges."}

            if edge_index < 0 or edge_index >= face_edges.Count:
                return {"error": f"Invalid edge index: {edge_index}. Face has {face_edges.Count} edges."}

            edge = face_edges.Item(edge_index + 1)

            side_map = {
                "Left": DirectionConstants.igLeft,
                "Right": DirectionConstants.igRight,
                "Both": DirectionConstants.igBoth,
            }
            side_const = side_map.get(side, DirectionConstants.igRight)

            flanges = model.Flanges

            # Build optional args
            args = [edge, side_const, flange_length]
            kwargs_to_add = []

            # Optional params: ThicknessSide, InsideRadius, DimSide, BRType, BRWidth,
            # BRLength, CRType, NeutralFactor, BnParamType, BendAngle
            if inside_radius is not None or bend_angle is not None:
                # ThicknessSide (skip) -> InsideRadius
                # We need to pass positional VT_VARIANT optional params
                # In late binding, pass them positionally
                if inside_radius is not None and bend_angle is not None:
                    bend_angle_rad = math.radians(bend_angle)
                    flange = flanges.Add(
                        edge, side_const, flange_length,
                        None, inside_radius, None, None, None, None, None, None, None,
                        bend_angle_rad
                    )
                elif inside_radius is not None:
                    flange = flanges.Add(
                        edge, side_const, flange_length,
                        None, inside_radius
                    )
                else:
                    bend_angle_rad = math.radians(bend_angle)
                    flange = flanges.Add(
                        edge, side_const, flange_length,
                        None, None, None, None, None, None, None, None, None,
                        bend_angle_rad
                    )
            else:
                flange = flanges.Add(edge, side_const, flange_length)

            result = {
                "status": "created",
                "type": "flange",
                "face_index": face_index,
                "edge_index": edge_index,
                "flange_length": flange_length,
                "side": side,
            }
            if inside_radius is not None:
                result["inside_radius"] = inside_radius
            if bend_angle is not None:
                result["bend_angle"] = bend_angle

            return result
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # ADDITIONAL FEATURE TYPES (Dimple, Etch, Rib, Lip, Slot, etc.)
    # =================================================================

    def create_dimple(self, depth: float, direction: str = "Normal") -> Dict[str, Any]:
        """
        Create a dimple feature (sheet metal).

        Creates a dimple from the active sketch profile on the sheet metal body.
        Requires an active sketch profile and an existing sheet metal base feature.

        Args:
            depth: Dimple depth in meters
            direction: 'Normal' or 'Reverse' for dimple direction

        Returns:
            Dict with status and dimple info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a sheet metal base feature first."}

            model = models.Item(1)

            # seDimpleDepthLeft=1, seDimpleDepthRight=2
            profile_side = 1 if direction == "Normal" else 2
            depth_side = 2 if direction == "Normal" else 1

            dimples = model.Dimples
            dimple = dimples.Add(profile, depth, profile_side, depth_side)

            return {
                "status": "created",
                "type": "dimple",
                "depth": depth,
                "direction": direction
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_etch(self) -> Dict[str, Any]:
        """
        Create an etch feature (sheet metal).

        Etches the active sketch profile into the sheet metal body.
        Requires an active sketch profile and an existing sheet metal base feature.

        Returns:
            Dict with status and etch info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a sheet metal base feature first."}

            model = models.Item(1)

            etches = model.Etches
            etch = etches.Add(profile)

            return {
                "status": "created",
                "type": "etch"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_rib(self, thickness: float, direction: str = "Normal") -> Dict[str, Any]:
        """
        Create a rib feature from the active sketch profile.

        Ribs are structural reinforcements that extend from a profile to
        existing geometry. Requires an active sketch profile.

        Args:
            thickness: Rib thickness in meters
            direction: 'Normal', 'Reverse', or 'Symmetric'

        Returns:
            Dict with status and rib info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            dir_map = {
                "Normal": DirectionConstants.igRight,
                "Reverse": DirectionConstants.igLeft,
                "Symmetric": DirectionConstants.igSymmetric,
            }
            side = dir_map.get(direction, DirectionConstants.igRight)

            ribs = model.Ribs
            rib = ribs.Add(profile, 1, 0, side, thickness)

            return {
                "status": "created",
                "type": "rib",
                "thickness": thickness,
                "direction": direction
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_lip(self, depth: float, direction: str = "Normal") -> Dict[str, Any]:
        """
        Create a lip feature from the active sketch profile.

        Lips are raised edges or ridges on plastic or sheet metal parts.
        Requires an active sketch profile and an existing base feature.

        Args:
            depth: Lip depth/height in meters
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with status and lip info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            side = DirectionConstants.igRight if direction == "Normal" else DirectionConstants.igLeft

            lips = model.Lips
            lip = lips.Add(profile, side, depth)

            return {
                "status": "created",
                "type": "lip",
                "depth": depth,
                "direction": direction
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_drawn_cutout(self, depth: float, direction: str = "Normal") -> Dict[str, Any]:
        """
        Create a drawn cutout feature (sheet metal).

        Creates a formed cutout from the active sketch profile. Unlike extruded
        cutouts, drawn cutouts follow the material's bend characteristics.
        Requires an active sketch profile and existing base feature.

        Args:
            depth: Cutout depth in meters
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with status and cutout info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            # igLeft=1, igRight=2
            side = 2 if direction == "Normal" else 1

            drawn_cutouts = model.DrawnCutouts
            cutout = drawn_cutouts.Add(profile, side, depth)

            return {
                "status": "created",
                "type": "drawn_cutout",
                "depth": depth,
                "direction": direction
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_bead(self, depth: float, direction: str = "Normal") -> Dict[str, Any]:
        """
        Create a bead feature (sheet metal stiffener).

        Beads are raised ridges used to stiffen sheet metal parts.
        Requires an active sketch profile and an existing sheet metal base feature.

        Args:
            depth: Bead depth in meters
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with status and bead info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            # igLeft=1, igRight=2
            side = 2 if direction == "Normal" else 1

            beads = model.Beads
            bead = beads.Add(profile, side, depth)

            return {
                "status": "created",
                "type": "bead",
                "depth": depth,
                "direction": direction
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_louver(self, depth: float, direction: str = "Normal") -> Dict[str, Any]:
        """
        Create a louver feature (sheet metal vent).

        Louvers are formed openings used for ventilation in sheet metal parts.
        Requires an active sketch profile and an existing sheet metal base feature.

        Args:
            depth: Louver depth in meters
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with status and louver info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            # igLeft=1, igRight=2
            side = 2 if direction == "Normal" else 1

            louvers = model.Louvers
            louver = louvers.Add(profile, side, depth)

            return {
                "status": "created",
                "type": "louver",
                "depth": depth,
                "direction": direction
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_gusset(self, thickness: float, direction: str = "Normal") -> Dict[str, Any]:
        """
        Create a gusset feature (sheet metal reinforcement).

        Gussets are triangular reinforcement plates used in sheet metal.
        Requires an active sketch profile and an existing sheet metal base feature.

        Args:
            thickness: Gusset thickness in meters
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with status and gusset info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            # igLeft=1, igRight=2
            side = 2 if direction == "Normal" else 1

            gussets = model.Gussets
            gusset = gussets.Add(profile, side, thickness)

            return {
                "status": "created",
                "type": "gusset",
                "thickness": thickness,
                "direction": direction
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_thread(self, face_index: int, pitch: float = 0.001,
                      thread_type: str = "External") -> Dict[str, Any]:
        """
        Create a thread feature on a cylindrical face.

        Adds cosmetic or modeled threads to a cylindrical face (hole or shaft).

        Args:
            face_index: 0-based index of the cylindrical face
            pitch: Thread pitch in meters (default 1mm)
            thread_type: 'External' (on shaft) or 'Internal' (in hole)

        Returns:
            Dict with status and thread info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            body = model.Body
            faces = body.Faces(FaceQueryConstants.igQueryAll)

            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Count: {faces.Count}"}

            face = faces.Item(face_index + 1)

            threads = model.Threads
            thread = threads.Add(face, pitch)

            return {
                "status": "created",
                "type": "thread",
                "face_index": face_index,
                "pitch": pitch,
                "pitch_mm": pitch * 1000,
                "thread_type": thread_type
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_slot(self, depth: float, direction: str = "Normal") -> Dict[str, Any]:
        """
        Create a slot feature from the active sketch profile.

        Slots are elongated cutouts typically used for fastener clearance.
        Requires an active sketch profile and an existing base feature.

        Args:
            depth: Slot depth in meters
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with status and slot info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            # igLeft=1, igRight=2
            side = 2 if direction == "Normal" else 1

            slots = model.Slots
            slot = slots.Add(profile, side, depth)

            return {
                "status": "created",
                "type": "slot",
                "depth": depth,
                "direction": direction
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_split(self, direction: str = "Normal") -> Dict[str, Any]:
        """
        Create a split feature to divide a body along the active sketch profile.

        Requires an active sketch profile and an existing base feature.

        Args:
            direction: 'Normal' or 'Reverse' - which side to keep

        Returns:
            Dict with status and split info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            # igLeft=1, igRight=2
            side = 2 if direction == "Normal" else 1

            splits = model.Splits
            split = splits.Add(profile, side)

            return {
                "status": "created",
                "type": "split",
                "direction": direction
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # UNEQUAL CHAMFER
    # =================================================================

    def create_chamfer_unequal(self, distance1: float, distance2: float,
                               face_index: int = 0) -> Dict[str, Any]:
        """
        Create a chamfer with two different setback distances.

        Unlike equal chamfer, this creates an asymmetric chamfer where each side
        has a different setback. Requires a reference face to determine direction.

        Args:
            distance1: First setback distance in meters
            distance2: Second setback distance in meters
            face_index: 0-based face index for the reference face

        Returns:
            Dict with status and chamfer info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to add chamfers to"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)
            face_edges = face.Edges
            if not hasattr(face_edges, 'Count') or face_edges.Count == 0:
                return {"error": f"Face {face_index} has no edges"}

            edge_list = []
            for ei in range(1, face_edges.Count + 1):
                edge_list.append(face_edges.Item(ei))

            chamfers = model.Chamfers
            chamfer = chamfers.AddUnequalSetback(
                face, len(edge_list), edge_list, distance1, distance2)

            return {
                "status": "created",
                "type": "chamfer_unequal",
                "distance1": distance1,
                "distance2": distance2,
                "face_index": face_index,
                "edge_count": len(edge_list)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # ANGLE CHAMFER
    # =================================================================

    def create_chamfer_angle(self, distance: float, angle: float,
                             face_index: int = 0) -> Dict[str, Any]:
        """
        Create a chamfer defined by a setback distance and an angle.

        Args:
            distance: Setback distance in meters
            angle: Chamfer angle in degrees
            face_index: 0-based face index for the reference face

        Returns:
            Dict with status and chamfer info
        """
        try:
            import math

            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to add chamfers to"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)
            face_edges = face.Edges
            if not hasattr(face_edges, 'Count') or face_edges.Count == 0:
                return {"error": f"Face {face_index} has no edges"}

            edge_list = []
            for ei in range(1, face_edges.Count + 1):
                edge_list.append(face_edges.Item(ei))

            chamfers = model.Chamfers
            angle_rad = math.radians(angle)
            chamfer = chamfers.AddSetbackAngle(
                face, len(edge_list), edge_list, distance, angle_rad)

            return {
                "status": "created",
                "type": "chamfer_angle",
                "distance": distance,
                "angle_degrees": angle,
                "face_index": face_index,
                "edge_count": len(edge_list)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # FACE ROTATE
    # =================================================================

    def create_face_rotate_by_edge(self, face_index: int, edge_index: int,
                                    angle: float) -> Dict[str, Any]:
        """
        Rotate a face around an edge axis.

        Tilts a face by rotating it around a specified edge. Useful for
        creating draft angles or adjusting face orientations.

        Args:
            face_index: 0-based face index to rotate
            edge_index: 0-based edge index to use as rotation axis
            angle: Rotation angle in degrees

        Returns:
            Dict with status and face rotate info
        """
        try:
            import math

            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to rotate faces on"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)

            # Get edge from the face
            face_edges = face.Edges
            if not hasattr(face_edges, 'Count') or face_edges.Count == 0:
                return {"error": f"Face {face_index} has no edges"}
            if edge_index < 0 or edge_index >= face_edges.Count:
                return {"error": f"Invalid edge index: {edge_index}. Face has {face_edges.Count} edges."}

            edge = face_edges.Item(edge_index + 1)

            angle_rad = math.radians(angle)

            face_rotates = model.FaceRotates
            # igFaceRotateByGeometry = 1, igFaceRotateRecreateBlends = 1, igFaceRotateAxisEnd = 2
            face_rotate = face_rotates.Add(
                face, 1, 1, None, None, edge, 2, angle_rad)

            return {
                "status": "created",
                "type": "face_rotate",
                "method": "by_edge",
                "face_index": face_index,
                "edge_index": edge_index,
                "angle_degrees": angle
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_face_rotate_by_points(self, face_index: int,
                                      vertex1_index: int, vertex2_index: int,
                                      angle: float) -> Dict[str, Any]:
        """
        Rotate a face around an axis defined by two vertex points.

        Args:
            face_index: 0-based face index to rotate
            vertex1_index: 0-based index of first vertex defining rotation axis
            vertex2_index: 0-based index of second vertex defining rotation axis
            angle: Rotation angle in degrees

        Returns:
            Dict with status and face rotate info
        """
        try:
            import math

            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to rotate faces on"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)

            # Get vertices from the face
            vertices = face.Vertices
            if vertex1_index < 0 or vertex1_index >= vertices.Count:
                return {"error": f"Invalid vertex1 index: {vertex1_index}. Face has {vertices.Count} vertices."}
            if vertex2_index < 0 or vertex2_index >= vertices.Count:
                return {"error": f"Invalid vertex2 index: {vertex2_index}. Face has {vertices.Count} vertices."}

            point1 = vertices.Item(vertex1_index + 1)
            point2 = vertices.Item(vertex2_index + 1)

            angle_rad = math.radians(angle)

            face_rotates = model.FaceRotates
            # igFaceRotateByPoints = 2, igFaceRotateRecreateBlends = 1, igFaceRotateNone = 0
            face_rotate = face_rotates.Add(
                face, 2, 1, point1, point2, None, 0, angle_rad)

            return {
                "status": "created",
                "type": "face_rotate",
                "method": "by_points",
                "face_index": face_index,
                "vertex1_index": vertex1_index,
                "vertex2_index": vertex2_index,
                "angle_degrees": angle
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # DRAFT ANGLE
    # =================================================================

    def create_draft_angle(self, face_index: int, angle: float,
                           plane_index: int = 1) -> Dict[str, Any]:
        """
        Add a draft angle to a face.

        Draft angles are used in injection molding to facilitate part removal
        from the mold. Uses the model.Drafts collection.

        Args:
            face_index: 0-based face index to apply draft to
            angle: Draft angle in degrees
            plane_index: 1-based reference plane index for draft direction (default: 1 = Top)

        Returns:
            Dict with status and draft info
        """
        try:
            import math

            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to add draft to"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)
            ref_plane = doc.RefPlanes.Item(plane_index)

            angle_rad = math.radians(angle)

            # igRight = 2 (draft direction side)
            drafts = model.Drafts
            draft = drafts.Add(ref_plane, 1, [face], [angle_rad], 2)

            return {
                "status": "created",
                "type": "draft_angle",
                "face_index": face_index,
                "angle_degrees": angle,
                "plane_index": plane_index
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # REFERENCE PLANE NORMAL TO CURVE
    # =================================================================

    def create_ref_plane_normal_to_curve(self, curve_end: str = "End",
                                         pivot_plane_index: int = 2) -> Dict[str, Any]:
        """
        Create a reference plane normal (perpendicular) to a curve at its endpoint.

        Useful for creating sweep cross-section sketches perpendicular to a path.
        Requires an active sketch profile that defines the curve.

        Args:
            curve_end: Which end of the curve to place the plane at - 'Start' or 'End'
            pivot_plane_index: 1-based index of the pivot reference plane (default: 2 = Front)

        Returns:
            Dict with status and new plane index
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            ref_planes = doc.RefPlanes
            pivot_plane = ref_planes.Item(pivot_plane_index)

            # igCurveEnd = 2, igCurveStart = 1
            curve_end_const = 2 if curve_end == "End" else 1
            # igPivotEnd = 2
            pivot_end_const = 2

            new_plane = ref_planes.AddNormalToCurve(
                profile, curve_end_const, pivot_plane, pivot_end_const, True)

            new_index = ref_planes.Count

            return {
                "status": "created",
                "type": "ref_plane_normal_to_curve",
                "curve_end": curve_end,
                "pivot_plane_index": pivot_plane_index,
                "new_plane_index": new_index
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # TIER 1: SWEPT CUTOUT
    # =================================================================

    def create_swept_cutout(self, path_profile_index: int = None) -> Dict[str, Any]:
        """
        Create a swept cutout (cut) along a path.

        Same workflow as create_sweep but removes material instead of adding it.
        Requires at least 2 accumulated profiles: path (open) + cross-section (closed).
        Uses model.SweptCutouts.Add() (type library: SweptCutouts collection).

        Args:
            path_profile_index: Index of the path profile in accumulated profiles
                (default: 0, the first accumulated profile)

        Returns:
            Dict with status and swept cutout info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            all_profiles = self.sketch_manager.get_accumulated_profiles()

            if len(all_profiles) < 2:
                return {
                    "error": f"Swept cutout requires at least 2 profiles (path + cross-section), "
                    f"got {len(all_profiles)}. Create a path sketch and a cross-section sketch first."
                }

            path_idx = path_profile_index if path_profile_index is not None else 0
            path_profile = all_profiles[path_idx]
            cross_sections = [p for i, p in enumerate(all_profiles) if i != path_idx]

            _CS = LoftSweepConstants.igProfileBasedCrossSection

            # Path arrays
            v_paths = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [path_profile])
            v_path_types = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I4, [_CS])

            # Cross-section arrays
            v_sections = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, cross_sections)
            v_section_types = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_I4,
                [_CS] * len(cross_sections)
            )
            v_origins = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
                [VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0])
                 for _ in cross_sections]
            )
            v_seg = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, [])

            # SweptCutouts.Add: same 15 params as SweptProtrusions
            swept_cutouts = model.SweptCutouts
            cutout = swept_cutouts.Add(
                1, v_paths, v_path_types,                        # Path (1 curve)
                len(cross_sections), v_sections, v_section_types, v_origins, v_seg,
                DirectionConstants.igRight,                      # MaterialSide
                ExtentTypeConstants.igNone, 0.0, None,           # Start extent
                ExtentTypeConstants.igNone, 0.0, None,           # End extent
            )

            self.sketch_manager.clear_accumulated_profiles()
            return {
                "status": "created",
                "type": "swept_cutout",
                "num_cross_sections": len(cross_sections),
                "method": "model.SweptCutouts.Add"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # TIER 1: HELIX CUTOUT
    # =================================================================

    def create_helix_cutout(
        self,
        pitch: float,
        height: float,
        revolutions: float = None,
        direction: str = "Right"
    ) -> Dict[str, Any]:
        """
        Create a helical cutout (cut) in the part.

        Same workflow as create_helix but removes material. Requires a closed sketch
        profile and an axis of revolution. Uses model.HelixCutouts.AddFinite().
        Type library: HelixCutouts.AddFinite(HelixAxis, AxisStart, NumCrossSections,
        CrossSectionArray, ProfileSide, Height, Pitch, NumberOfTurns, HelixDir, ...).

        Args:
            pitch: Distance between coils in meters
            height: Total height of helix in meters
            revolutions: Number of turns (optional, calculated from pitch/height)
            direction: 'Right' or 'Left' hand helix

        Returns:
            Dict with status and helix cutout info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            profile = self.sketch_manager.get_active_sketch()
            refaxis = self.sketch_manager.get_active_refaxis()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            if not refaxis:
                return {"error": "No axis of revolution set. Use set_axis_of_revolution() in the sketch."}

            if revolutions is None:
                revolutions = height / pitch

            axis_start = DirectionConstants.igRight
            dir_const = DirectionConstants.igRight if direction == "Right" else DirectionConstants.igLeft

            # Wrap cross-section profile in SAFEARRAY
            v_profiles = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

            helix_cutouts = model.HelixCutouts
            cutout = helix_cutouts.AddFinite(
                refaxis,        # HelixAxis
                axis_start,     # AxisStart
                1,              # NumCrossSections
                v_profiles,     # CrossSectionArray
                DirectionConstants.igRight,  # ProfileSide
                height,         # Height
                pitch,          # Pitch
                revolutions,    # NumberOfTurns
                dir_const,      # HelixDir
            )

            self.sketch_manager.clear_accumulated_profiles()
            return {
                "status": "created",
                "type": "helix_cutout",
                "pitch": pitch,
                "height": height,
                "revolutions": revolutions,
                "direction": direction,
                "method": "model.HelixCutouts.AddFinite"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # TIER 1: VARIABLE ROUND (FILLET)
    # =================================================================

    def create_variable_round(self, radii: list, face_index: int = None) -> Dict[str, Any]:
        """
        Create a variable-radius round (fillet) on body edges.

        Unlike create_round which applies a constant radius, this allows different
        radii at different points along the edge. Uses model.Rounds.AddVariable().
        Type library: Rounds.AddVariable(NumberOfEdgeSets, EdgeSetArray, RadiusArray).

        Args:
            radii: List of radius values in meters. Each edge gets a corresponding radius.
                   If fewer radii than edges, the last radius is repeated.
            face_index: 0-based face index to apply to (None = all edges)

        Returns:
            Dict with status and variable round info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to add variable rounds to"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if faces.Count == 0:
                return {"error": "No faces found on body"}

            # Collect edges from specified face or all faces
            edge_list = []
            if face_index is not None:
                if face_index < 0 or face_index >= faces.Count:
                    return {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."}
                face = faces.Item(face_index + 1)
                face_edges = face.Edges
                if hasattr(face_edges, 'Count'):
                    for ei in range(1, face_edges.Count + 1):
                        edge_list.append(face_edges.Item(ei))
            else:
                for fi in range(1, faces.Count + 1):
                    face = faces.Item(fi)
                    face_edges = face.Edges
                    if not hasattr(face_edges, 'Count'):
                        continue
                    for ei in range(1, face_edges.Count + 1):
                        edge_list.append(face_edges.Item(ei))

            if not edge_list:
                return {"error": "No edges found on body"}

            # Extend radii list to match edge count if needed
            radius_values = list(radii)
            while len(radius_values) < len(edge_list):
                radius_values.append(radius_values[-1])

            # VARIANT wrappers required for Rounds methods
            edge_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, edge_list)
            radius_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, radius_values[:len(edge_list)])

            rounds = model.Rounds
            rnd = rounds.AddVariable(1, edge_arr, radius_arr)

            return {
                "status": "created",
                "type": "variable_round",
                "edge_count": len(edge_list),
                "radii": radius_values[:len(edge_list)]
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # TIER 1: BLEND (FACE-TO-FACE FILLET)
    # =================================================================

    def create_blend(self, radius: float, face_index: int = None) -> Dict[str, Any]:
        """
        Create a blend (face-to-face fillet) feature.

        Uses model.Blends.Add(NumberOfEdgeSets, EdgeSetArray, RadiusArray).
        Same VARIANT wrapper pattern as Rounds. Unlike Rounds which fillets edges,
        Blends create smooth transitions between faces.

        Args:
            radius: Blend radius in meters
            face_index: 0-based face index to apply to (None = all edges)

        Returns:
            Dict with status and blend info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to add blends to"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if faces.Count == 0:
                return {"error": "No faces found on body"}

            # Collect edges
            edge_list = []
            if face_index is not None:
                if face_index < 0 or face_index >= faces.Count:
                    return {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."}
                face = faces.Item(face_index + 1)
                face_edges = face.Edges
                if hasattr(face_edges, 'Count'):
                    for ei in range(1, face_edges.Count + 1):
                        edge_list.append(face_edges.Item(ei))
            else:
                for fi in range(1, faces.Count + 1):
                    face = faces.Item(fi)
                    face_edges = face.Edges
                    if not hasattr(face_edges, 'Count'):
                        continue
                    for ei in range(1, face_edges.Count + 1):
                        edge_list.append(face_edges.Item(ei))

            if not edge_list:
                return {"error": "No edges found on body"}

            # VARIANT wrappers (same pattern as Rounds)
            edge_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, edge_list)
            radius_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [radius])

            blends = model.Blends
            blend = blends.Add(1, edge_arr, radius_arr)

            return {
                "status": "created",
                "type": "blend",
                "radius": radius,
                "edge_count": len(edge_list)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # TIER 1: REFERENCE PLANE BY ANGLE
    # =================================================================

    def create_ref_plane_by_angle(
        self,
        parent_plane_index: int,
        angle: float,
        normal_side: str = "Normal"
    ) -> Dict[str, Any]:
        """
        Create a reference plane at an angle to an existing plane.

        Uses RefPlanes.AddAngularByAngle(ParentPlane, Angle, NormalSide).
        Type library: AddAngularByAngle(ParentPlane: IDispatch, Angle: VT_R8,
        NormalSide: FeaturePropertyConstants, [Edge: VT_VARIANT]).

        Args:
            parent_plane_index: Index of parent plane (1=Top/XZ, 2=Front/XY, 3=Right/YZ)
            angle: Angle in degrees from the parent plane
            normal_side: 'Normal' (igRight=2) or 'Reverse' (igLeft=1)

        Returns:
            Dict with status and new plane index
        """
        try:
            import math

            doc = self.doc_manager.get_active_document()
            ref_planes = doc.RefPlanes

            if parent_plane_index < 1 or parent_plane_index > ref_planes.Count:
                return {"error": f"Invalid plane index: {parent_plane_index}. Count: {ref_planes.Count}"}

            parent = ref_planes.Item(parent_plane_index)

            side_map = {
                "Normal": DirectionConstants.igRight,
                "Reverse": DirectionConstants.igLeft,
            }
            side_const = side_map.get(normal_side, DirectionConstants.igRight)

            # Angle in radians for the COM API
            angle_rad = math.radians(angle)

            new_plane = ref_planes.AddAngularByAngle(parent, angle_rad, side_const)

            return {
                "status": "created",
                "type": "reference_plane",
                "method": "angular_by_angle",
                "parent_plane": parent_plane_index,
                "angle_degrees": angle,
                "normal_side": normal_side,
                "new_plane_index": ref_planes.Count
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # TIER 1: REFERENCE PLANE BY 3 POINTS
    # =================================================================

    def create_ref_plane_by_3_points(
        self,
        x1: float, y1: float, z1: float,
        x2: float, y2: float, z2: float,
        x3: float, y3: float, z3: float
    ) -> Dict[str, Any]:
        """
        Create a reference plane through 3 points in space.

        Uses RefPlanes.AddBy3Points(Point1X, Point1Y, Point1Z, ...).
        Type library: AddBy3Points(9x VT_R8 params) -> RefPlane*.

        Args:
            x1, y1, z1: First point coordinates (meters)
            x2, y2, z2: Second point coordinates (meters)
            x3, y3, z3: Third point coordinates (meters)

        Returns:
            Dict with status and new plane index
        """
        try:
            doc = self.doc_manager.get_active_document()
            ref_planes = doc.RefPlanes

            new_plane = ref_planes.AddBy3Points(
                x1, y1, z1,
                x2, y2, z2,
                x3, y3, z3
            )

            return {
                "status": "created",
                "type": "reference_plane",
                "method": "by_3_points",
                "point1": [x1, y1, z1],
                "point2": [x2, y2, z2],
                "point3": [x3, y3, z3],
                "new_plane_index": ref_planes.Count
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # TIER 1: REFERENCE PLANE MID-PLANE
    # =================================================================

    def create_ref_plane_midplane(
        self,
        plane1_index: int,
        plane2_index: int
    ) -> Dict[str, Any]:
        """
        Create a reference plane midway between two existing planes.

        Uses RefPlanes.AddMidPlane(Plane1, Plane2).
        Useful for symmetry operations.

        Args:
            plane1_index: Index of first plane (1=Top/XZ, 2=Front/XY, 3=Right/YZ)
            plane2_index: Index of second plane

        Returns:
            Dict with status and new plane index
        """
        try:
            doc = self.doc_manager.get_active_document()
            ref_planes = doc.RefPlanes

            if plane1_index < 1 or plane1_index > ref_planes.Count:
                return {"error": f"Invalid plane1 index: {plane1_index}. Count: {ref_planes.Count}"}
            if plane2_index < 1 or plane2_index > ref_planes.Count:
                return {"error": f"Invalid plane2 index: {plane2_index}. Count: {ref_planes.Count}"}

            plane1 = ref_planes.Item(plane1_index)
            plane2 = ref_planes.Item(plane2_index)

            new_plane = ref_planes.AddMidPlane(plane1, plane2)

            return {
                "status": "created",
                "type": "reference_plane",
                "method": "mid_plane",
                "plane1_index": plane1_index,
                "plane2_index": plane2_index,
                "new_plane_index": ref_planes.Count
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # TIER 1: HOLE THROUGH ALL
    # =================================================================

    def create_hole_through_all(
        self,
        x: float, y: float,
        diameter: float,
        plane_index: int = 1,
        direction: str = "Normal"
    ) -> Dict[str, Any]:
        """
        Create a hole that goes through the entire part.

        Creates a circular profile and uses ExtrudedCutouts.AddThroughAllMulti.
        Type library: ExtrudedCutouts.AddThroughAllMulti(NumProfiles, ProfileArray, PlaneSide).

        Args:
            x, y: Hole center coordinates on the sketch plane (meters)
            diameter: Hole diameter in meters
            plane_index: Reference plane index (1=Top/XZ, 2=Front/XY, 3=Right/YZ)
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with status and hole info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            radius = diameter / 2.0

            dir_const = DirectionConstants.igRight  # Normal
            if direction == "Reverse":
                dir_const = DirectionConstants.igLeft

            # Create a circular profile on the specified plane
            ps = doc.ProfileSets.Add()
            plane = doc.RefPlanes.Item(plane_index)
            profile = ps.Profiles.Add(plane)
            profile.Circles2d.AddByCenterRadius(x, y, radius)
            profile.End(0)

            # Use through-all cutout
            cutouts = model.ExtrudedCutouts
            cutout = cutouts.AddThroughAllMulti(1, (profile,), dir_const)

            return {
                "status": "created",
                "type": "hole_through_all",
                "position": [x, y],
                "diameter": diameter,
                "plane_index": plane_index,
                "direction": direction
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # TIER 1: BOX CUTOUT PRIMITIVE
    # =================================================================

    def create_box_cutout_by_two_points(
        self,
        x1: float, y1: float, z1: float,
        x2: float, y2: float, z2: float
    ) -> Dict[str, Any]:
        """
        Create a box-shaped cutout (boolean subtract) by two opposite corners.

        Uses BoxFeatures.AddCutoutByTwoPoints with same params as AddBoxByTwoPoints.
        Requires an existing base feature to cut from.
        Type library: AddCutoutByTwoPoints(6x VT_R8, dAngle, dDepth, pPlane,
        ExtentSide, vbKeyPointExtent, pKeyPointObj, pKeyPointFlags).

        Args:
            x1, y1, z1: First corner coordinates (meters)
            x2, y2, z2: Opposite corner coordinates (meters)

        Returns:
            Dict with status and cutout info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            top_plane = self._get_ref_plane(doc, 1)
            depth = abs(z2 - z1) if abs(z2 - z1) > 0 else abs(y2 - y1)

            # BoxFeatures is on the Models collection level
            box_features = models.BoxFeatures if hasattr(models, 'BoxFeatures') else None
            if box_features is None:
                # Try via the model object
                model = models.Item(1)
                box_features = model.BoxFeatures if hasattr(model, 'BoxFeatures') else None

            if box_features is None:
                return {"error": "BoxFeatures collection not accessible"}

            cutout = box_features.AddCutoutByTwoPoints(
                x1, y1, z1,
                x2, y2, z2,
                0,                              # dAngle
                depth,                          # dDepth
                top_plane,                      # pPlane
                DirectionConstants.igRight,     # ExtentSide
                False,                          # vbKeyPointExtent
                None,                           # pKeyPointObj
                0                               # pKeyPointFlags
            )

            return {
                "status": "created",
                "type": "box_cutout",
                "method": "by_two_points",
                "corner1": [x1, y1, z1],
                "corner2": [x2, y2, z2]
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # TIER 1: CYLINDER CUTOUT PRIMITIVE
    # =================================================================

    def create_cylinder_cutout(
        self,
        center_x: float,
        center_y: float,
        center_z: float,
        radius: float,
        height: float
    ) -> Dict[str, Any]:
        """
        Create a cylindrical cutout (boolean subtract) primitive.

        Uses CylinderFeatures.AddCutoutByCenterAndRadius.
        Requires an existing base feature to cut from.
        Type library: AddCutoutByCenterAndRadius(x, y, z, dRadius, dDepth, pPlane,
        ProfileSide, ExtentSide, vbKeyPointExtent, pKeyPointObj, pKeyPointFlags).

        Args:
            center_x, center_y, center_z: Center coordinates (meters)
            radius: Cylinder radius (meters)
            height: Cylinder/cut depth (meters)

        Returns:
            Dict with status and cutout info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            top_plane = self._get_ref_plane(doc, 1)

            # CylinderFeatures collection - try on Models first, then model
            cyl_features = models.CylinderFeatures if hasattr(models, 'CylinderFeatures') else None
            if cyl_features is None:
                model = models.Item(1)
                cyl_features = model.CylinderFeatures if hasattr(model, 'CylinderFeatures') else None

            if cyl_features is None:
                return {"error": "CylinderFeatures collection not accessible"}

            cutout = cyl_features.AddCutoutByCenterAndRadius(
                center_x, center_y, center_z,
                radius,
                height,                             # dDepth
                top_plane,                          # pPlane
                DirectionConstants.igRight,         # ProfileSide
                DirectionConstants.igRight,         # ExtentSide
                False,                              # vbKeyPointExtent
                None,                               # pKeyPointObj
                0                                   # pKeyPointFlags
            )

            return {
                "status": "created",
                "type": "cylinder_cutout",
                "center": [center_x, center_y, center_z],
                "radius": radius,
                "height": height
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # TIER 1: SPHERE CUTOUT PRIMITIVE
    # =================================================================

    def create_sphere_cutout(
        self,
        center_x: float,
        center_y: float,
        center_z: float,
        radius: float
    ) -> Dict[str, Any]:
        """
        Create a spherical cutout (boolean subtract) primitive.

        Uses SphereFeatures.AddCutoutByCenterAndRadius.
        Requires an existing base feature to cut from.
        Type library: AddCutoutByCenterAndRadius(x, y, z, dRadius, pPlane,
        ProfileSide, ExtentSide, vbKeyPointExtent, vbCreateLiveSection,
        pKeyPointObj, pKeyPointFlags).

        Args:
            center_x, center_y, center_z: Sphere center coordinates (meters)
            radius: Sphere radius (meters)

        Returns:
            Dict with status and cutout info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            top_plane = self._get_ref_plane(doc, 1)

            # SphereFeatures collection
            sph_features = models.SphereFeatures if hasattr(models, 'SphereFeatures') else None
            if sph_features is None:
                model = models.Item(1)
                sph_features = model.SphereFeatures if hasattr(model, 'SphereFeatures') else None

            if sph_features is None:
                return {"error": "SphereFeatures collection not accessible"}

            cutout = sph_features.AddCutoutByCenterAndRadius(
                center_x, center_y, center_z,
                radius,
                top_plane,                          # pPlane
                DirectionConstants.igRight,         # ProfileSide
                DirectionConstants.igRight,         # ExtentSide
                False,                              # vbKeyPointExtent
                False,                              # vbCreateLiveSection
                None,                               # pKeyPointObj
                0                                   # pKeyPointFlags
            )

            return {
                "status": "created",
                "type": "sphere_cutout",
                "center": [center_x, center_y, center_z],
                "radius": radius
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # TIER 2: EXTRUDED CUTOUT THROUGH NEXT
    # =================================================================

    def create_extruded_cutout_through_next(self, direction: str = "Normal") -> Dict[str, Any]:
        """
        Create an extruded cutout that cuts to the next face encountered.

        Uses model.ExtrudedCutouts.AddThroughNextMulti(NumProfiles, ProfileArray, PlaneSide).
        Cuts from the sketch plane to the first face it meets.

        Args:
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with status and cutout info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            direction_map = {
                "Normal": DirectionConstants.igRight,
                "Reverse": DirectionConstants.igLeft,
            }
            dir_const = direction_map.get(direction, DirectionConstants.igRight)

            cutouts = model.ExtrudedCutouts
            cutout = cutouts.AddThroughNextMulti(1, (profile,), dir_const)

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "extruded_cutout_through_next",
                "direction": direction
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # TIER 2: NORMAL CUTOUT THROUGH ALL
    # =================================================================

    def create_normal_cutout_through_all(self, direction: str = "Normal") -> Dict[str, Any]:
        """
        Create a normal cutout that goes through the entire part.

        Uses model.NormalCutouts.AddThroughAllMulti(NumProfiles, ProfileArray,
        PlaneSide, Method). Normal cutouts follow the surface normal.

        Args:
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with status and cutout info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            direction_map = {
                "Normal": DirectionConstants.igRight,
                "Reverse": DirectionConstants.igLeft,
            }
            dir_const = direction_map.get(direction, DirectionConstants.igRight)

            # igNormalCutoutMethod_Normal = 0 (default method)
            cutouts = model.NormalCutouts
            cutout = cutouts.AddThroughAllMulti(1, (profile,), dir_const, 0)

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "normal_cutout_through_all",
                "direction": direction
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # TIER 2: DELETE HOLES
    # =================================================================

    def create_delete_hole(self, max_diameter: float = 1.0,
                           hole_type: str = "All") -> Dict[str, Any]:
        """
        Delete/fill holes in the model body.

        Uses model.DeleteHoles.Add(HoleTypeToDelete, ThresholdHoleDiameter).
        Fills holes up to the specified diameter threshold.

        Args:
            max_diameter: Maximum hole diameter to delete (meters). Holes with
                diameter <= this value will be filled.
            hole_type: Type of holes to delete: 'All', 'Round', 'NonRound'

        Returns:
            Dict with status and deletion info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            # HoleTypeToDeleteConstants from type library
            type_map = {
                "All": 0,
                "Round": 1,
                "NonRound": 2,
            }
            hole_type_const = type_map.get(hole_type, 0)

            delete_holes = model.DeleteHoles
            result_feat = delete_holes.Add(hole_type_const, max_diameter)

            return {
                "status": "created",
                "type": "delete_hole",
                "max_diameter": max_diameter,
                "hole_type": hole_type
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # TIER 2: DELETE BLENDS
    # =================================================================

    def create_delete_blend(self, face_index: int) -> Dict[str, Any]:
        """
        Delete/remove a blend (fillet) from the model by specifying a face.

        Uses model.DeleteBlends.Add(BlendsToDelete). Removes the blend
        associated with the specified face.

        Args:
            face_index: 0-based face index of the blend face to remove

        Returns:
            Dict with status and deletion info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists."}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)

            delete_blends = model.DeleteBlends
            result_feat = delete_blends.Add(face)

            return {
                "status": "created",
                "type": "delete_blend",
                "face_index": face_index
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # TIER 2: REVOLVED SURFACE
    # =================================================================

    def create_revolved_surface(self, angle: float = 360,
                                 want_end_caps: bool = False) -> Dict[str, Any]:
        """
        Create a revolved construction surface from the active profile.

        Uses RevolvedSurfaces.AddFinite(NumProfiles, ProfileArray, RefAxis,
        ProfilePlaneSide, AngleOfRevolution, WantEndCaps).
        Requires a profile with an axis of revolution set.

        Args:
            angle: Revolution angle in degrees (360 for full revolution)
            want_end_caps: Whether to cap the ends of the surface

        Returns:
            Dict with status and surface info
        """
        try:
            import math

            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()
            refaxis = self.sketch_manager.get_active_refaxis()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            if not refaxis:
                return {"error": "No axis of revolution set. Use set_axis_of_revolution() first."}

            models = doc.Models
            angle_rad = math.radians(angle)

            v_profiles = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

            # Try collection-level API first (on model), then Models-level
            if models.Count > 0:
                model = models.Item(1)
                rev_surfaces = model.RevolvedSurfaces
                surface = rev_surfaces.AddFinite(
                    1, v_profiles, refaxis, DirectionConstants.igRight, angle_rad, want_end_caps
                )
            else:
                # First feature - use Models method if available
                surface = models.AddFiniteRevolvedSurface(
                    1, v_profiles, refaxis, DirectionConstants.igRight, angle_rad, want_end_caps
                )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "revolved_surface",
                "angle_degrees": angle,
                "want_end_caps": want_end_caps
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # TIER 2: LOFTED SURFACE
    # =================================================================

    def create_lofted_surface(self, want_end_caps: bool = False) -> Dict[str, Any]:
        """
        Create a lofted construction surface between multiple profiles.

        Uses LoftedSurfaces.Add with accumulated profiles. Same workflow as
        create_loft: create 2+ sketches on different planes, close each,
        then call this method.

        Args:
            want_end_caps: Whether to cap the ends of the surface

        Returns:
            Dict with status and surface info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            all_profiles = self.sketch_manager.get_accumulated_profiles()

            if len(all_profiles) < 2:
                return {
                    "error": f"Lofted surface requires at least 2 profiles, "
                    f"got {len(all_profiles)}."
                }

            _CS = LoftSweepConstants.igProfileBasedCrossSection

            v_sections = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, all_profiles)
            v_types = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_I4,
                [_CS] * len(all_profiles)
            )
            v_origins = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
                [VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0])
                 for _ in all_profiles]
            )

            if models.Count > 0:
                model = models.Item(1)
                loft_surfaces = model.LoftedSurfaces
                surface = loft_surfaces.Add(
                    len(all_profiles), v_sections, v_types, v_origins,
                    ExtentTypeConstants.igNone,     # StartExtentType
                    ExtentTypeConstants.igNone,     # EndExtentType
                    0, 0.0,     # StartTangentType, StartTangentMagnitude
                    0, 0.0,     # EndTangentType, EndTangentMagnitude
                    0, None,    # NumGuideCurves, GuideCurves
                    want_end_caps
                )
            else:
                return {"error": "Lofted surface requires an existing base feature."}

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "lofted_surface",
                "num_profiles": len(all_profiles),
                "want_end_caps": want_end_caps
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # =================================================================
    # TIER 2: SWEPT SURFACE
    # =================================================================

    def create_swept_surface(self, path_profile_index: int = None,
                              want_end_caps: bool = False) -> Dict[str, Any]:
        """
        Create a swept construction surface along a path.

        Same workflow as create_sweep: path profile (open) + cross-section (closed).
        Uses SweptSurfaces.Add.

        Args:
            path_profile_index: Index of the path profile (default: 0)
            want_end_caps: Whether to cap the ends of the surface

        Returns:
            Dict with status and surface info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            all_profiles = self.sketch_manager.get_accumulated_profiles()

            if len(all_profiles) < 2:
                return {
                    "error": f"Swept surface requires at least 2 profiles (path + cross-section), "
                    f"got {len(all_profiles)}."
                }

            path_idx = path_profile_index if path_profile_index is not None else 0
            path_profile = all_profiles[path_idx]
            cross_sections = [p for i, p in enumerate(all_profiles) if i != path_idx]

            _CS = LoftSweepConstants.igProfileBasedCrossSection

            v_paths = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [path_profile])
            v_sections = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, cross_sections)

            swept_surfaces = model.SweptSurfaces
            surface = swept_surfaces.Add(
                1, v_paths, _CS,                            # Path
                len(cross_sections), v_sections, _CS,       # Sections
                None, None,                                 # Origins, OriginRefs
                ExtentTypeConstants.igNone,                 # StartExtentType
                ExtentTypeConstants.igNone,                 # EndExtentType
                want_end_caps
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "swept_surface",
                "num_cross_sections": len(cross_sections),
                "want_end_caps": want_end_caps
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def convert_feature_type(self, feature_name: str, target_type: str) -> Dict[str, Any]:
        """
        Convert a feature between cutout and protrusion.

        Uses Feature.ConvertToCutout() or Feature.ConvertToProtrusion()
        to toggle a feature between adding and removing material.

        Args:
            feature_name: Name of the feature (from list_features)
            target_type: 'cutout' or 'protrusion'

        Returns:
            Dict with conversion status and new feature reference
        """
        try:
            doc = self.doc_manager.get_active_document()

            # Find the feature by name
            features = doc.DesignEdgebarFeatures
            target_feature = None
            for i in range(1, features.Count + 1):
                feat = features.Item(i)
                try:
                    if feat.Name == feature_name:
                        target_feature = feat
                        break
                except Exception:
                    continue

            if target_feature is None:
                return {"error": f"Feature '{feature_name}' not found"}

            target_type_lower = target_type.lower()
            if target_type_lower == "cutout":
                result = target_feature.ConvertToCutout()
                new_name = None
                try:
                    new_name = result.Name
                except Exception:
                    pass
                return {
                    "status": "converted",
                    "original_name": feature_name,
                    "target_type": "cutout",
                    "new_name": new_name
                }
            elif target_type_lower == "protrusion":
                result = target_feature.ConvertToProtrusion()
                new_name = None
                try:
                    new_name = result.Name
                except Exception:
                    pass
                return {
                    "status": "converted",
                    "original_name": feature_name,
                    "target_type": "protrusion",
                    "new_name": new_name
                }
            else:
                return {"error": f"Invalid target_type: {target_type}. Use 'cutout' or 'protrusion'"}
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }
