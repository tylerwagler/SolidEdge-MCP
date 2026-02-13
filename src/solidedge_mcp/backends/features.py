"""
Solid Edge Feature Operations

Handles creating 3D features like extrusions, revolves, holes, fillets, etc.
"""

from typing import Dict, Any, Optional, List
import traceback
import pythoncom
from win32com.client import VARIANT
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
                ExtrudedProtrusion.igRight,     # ProfilePlaneSide (2)
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
            dir_const = ExtrudedProtrusion.igRight  # Normal
            if direction == "Reverse":
                dir_const = ExtrudedProtrusion.igLeft

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
            faces = body.Faces(1)  # igQueryAll = 1
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
            faces = body.Faces(1)  # igQueryAll = 1
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

    def _make_loft_variant_arrays(self, profiles):
        """Create properly typed VARIANT arrays for loft/sweep COM calls.

        COM requires explicit VARIANT typing for SAFEARRAY parameters.
        Python's automatic marshaling does not produce correct types for nested arrays.

        Args:
            profiles: List of profile COM objects

        Returns:
            Tuple of (v_profiles, v_types, v_origins) VARIANT arrays
        """
        igProfileBasedCrossSection = 48
        v_profiles = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, profiles)
        v_types = VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_I4,
            [igProfileBasedCrossSection] * len(profiles)
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

            igRight = 2  # Material side
            igNone = 44  # No tangent control

            # Try LoftedProtrusions.AddSimple first (works when a base feature exists)
            try:
                model = models.Item(1)
                lp = model.LoftedProtrusions
                loft = lp.AddSimple(
                    len(profiles), v_profiles, v_types, v_origins,
                    igRight, igNone, igNone
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
                v_seg,          # SegmentMaps (empty)
                igRight,        # MaterialSide
                igNone, 0.0, None,  # Start extent
                igNone, 0.0, None,  # End extent
                igNone, 0.0,        # Start tangent
                igNone, 0.0,        # End tangent
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

            igProfileBasedCrossSection = 48
            igRight = 2
            igNone = 44

            # Path arrays
            v_paths = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [path_profile])
            v_path_types = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I4, [igProfileBasedCrossSection])

            # Cross-section arrays
            v_sections = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, cross_sections)
            v_section_types = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_I4,
                [igProfileBasedCrossSection] * len(cross_sections)
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
                igRight,                                     # MaterialSide
                igNone, 0.0, None,                          # Start extent
                igNone, 0.0, None,                          # End extent
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

            # Constants
            igFinite = 13
            igRight = 2
            igLeft = 1
            igTangentNormal = 0
            seOffsetNone = 0
            seTreatmentNone = 0
            seDraftNone = 0
            seTreatmentCrownByOffset = 1
            seTreatmentCrownSideInside = 0
            seTreatmentCrownCurvatureInside = 0

            depth1 = distance
            depth2 = distance if direction == "Symmetric" else 0.0
            side1 = igRight
            side2 = igLeft if direction == "Symmetric" else igRight

            extruded_surface = extruded_surfaces.Add(
                1,                      # NumberOfProfiles
                profile_array,          # ProfileArray
                igFinite,               # ExtentType1
                side1,                  # ExtentSide1
                depth1,                 # FiniteDepth1
                None,                   # KeyPointOrTangentFace1
                igTangentNormal,        # KeyPointFlags1
                None,                   # FromFaceOrRefPlane
                seOffsetNone,           # FromFaceOffsetSide
                0.0,                    # FromFaceOffsetDistance
                seTreatmentNone,        # TreatmentType1
                seDraftNone,            # TreatmentDraftSide1
                0.0,                    # TreatmentDraftAngle1
                seTreatmentCrownByOffset,       # TreatmentCrownType1
                seTreatmentCrownSideInside,     # TreatmentCrownSide1
                seTreatmentCrownCurvatureInside, # TreatmentCrownCurvatureSide1
                0.0,                    # TreatmentCrownRadiusOrOffset1
                0.0,                    # TreatmentCrownTakeOffAngle1
                igFinite,               # ExtentType2
                side2,                  # ExtentSide2
                depth2,                 # FiniteDepth2
                None,                   # KeyPointOrTangentFace2
                igTangentNormal,        # KeyPointFlags2
                None,                   # ToFaceOrRefPlane
                seOffsetNone,           # ToFaceOffsetSide
                0.0,                    # ToFaceOffsetDistance
                seTreatmentNone,        # TreatmentType2
                seDraftNone,            # TreatmentDraftSide2
                0.0,                    # TreatmentDraftAngle2
                seTreatmentCrownByOffset,       # TreatmentCrownType2
                seTreatmentCrownSideInside,     # TreatmentCrownSide2
                seTreatmentCrownCurvatureInside, # TreatmentCrownCurvatureSide2
                0.0,                    # TreatmentCrownRadiusOrOffset2
                0.0,                    # TreatmentCrownTakeOffAngle2
                end_caps                # WantEndCaps
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

            igRight = 2
            igNone = 44

            model = models.AddLoftedProtrusionWithThinWall(
                len(profiles), v_profiles, v_types, v_origins,
                v_seg,              # SegmentMaps
                igRight,            # MaterialSide
                igNone, 0.0, None,  # Start extent
                igNone, 0.0, None,  # End extent
                igNone, 0.0,        # Start tangent
                igNone, 0.0,        # End tangent
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

            igProfileBasedCrossSection = 48
            igRight = 2
            igNone = 44

            v_paths = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [path_profile])
            v_path_types = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I4, [igProfileBasedCrossSection])
            v_sections = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, cross_sections)
            v_section_types = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_I4,
                [igProfileBasedCrossSection] * len(cross_sections)
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
                igRight,
                igNone, 0.0, None,
                igNone, 0.0, None,
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
                "Normal": ExtrudedProtrusion.igRight,
                "Reverse": ExtrudedProtrusion.igLeft,
            }
            dir_const = direction_map.get(direction, ExtrudedProtrusion.igRight)

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
                "Normal": ExtrudedProtrusion.igRight,
                "Reverse": ExtrudedProtrusion.igLeft,
            }
            dir_const = direction_map.get(direction, ExtrudedProtrusion.igRight)

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
                ExtrudedProtrusion.igRight,     # ProfilePlaneSide
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
                "Normal": ExtrudedProtrusion.igRight,
                "Reverse": ExtrudedProtrusion.igLeft,
            }
            dir_const = direction_map.get(direction, ExtrudedProtrusion.igRight)

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

            igRight = 2   # Material side
            igNone = 44   # No tangent control

            lc = model.LoftedCutouts
            cutout = lc.AddSimple(
                len(profiles), v_profiles, v_types, v_origins,
                igRight, igNone, igNone
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
                "Normal": ExtrudedProtrusion.igRight,
                "Reverse": ExtrudedProtrusion.igLeft,
            }
            side_const = side_map.get(normal_side, ExtrudedProtrusion.igRight)

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

            faces = body.Faces(1)  # igQueryAll = 1
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

            faces = body.Faces(1)  # igQueryAll = 1
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

            faces = body.Faces(1)  # igQueryAll = 1
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

            # igRight=2, igLeft=1, igSymmetric=3
            dir_map = {"Normal": 2, "Reverse": 1, "Symmetric": 3}
            side = dir_map.get(direction, 2)

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

            # igRight=2, igLeft=1
            side = 2 if direction == "Normal" else 1

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
            faces = body.Faces(1)  # igQueryAll = 1

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

            faces = body.Faces(1)  # igQueryAll = 1
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

            faces = body.Faces(1)  # igQueryAll = 1
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

            faces = body.Faces(1)  # igQueryAll = 1
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

            faces = body.Faces(1)  # igQueryAll = 1
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

            faces = body.Faces(1)  # igQueryAll = 1
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
