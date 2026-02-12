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
                   hole_type: str = "Simple") -> Dict[str, Any]:
        """
        Create a hole feature.

        Creates a hole at (x, y) on the first reference plane. Requires an existing
        base feature. Uses HoleDataCollection + Holes.AddFinite from the COM API.

        Args:
            x, y: Hole center coordinates on the sketch plane (meters)
            diameter: Hole diameter in meters
            depth: Hole depth in meters
            hole_type: 'Simple', 'Counterbore', or 'Countersink'

        Returns:
            Dict with status and hole info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            # Map hole type
            type_map = {
                "Simple": HoleTypeConstants.igRegularHole,
                "Counterbore": HoleTypeConstants.igCounterboreHole,
                "Countersink": HoleTypeConstants.igCountersinkHole
            }
            hole_type_const = type_map.get(hole_type, HoleTypeConstants.igRegularHole)

            # Create HoleData via doc.HoleDataCollection
            hdc = doc.HoleDataCollection
            hole_data = hdc.Add(hole_type_const, diameter, 90)

            # Create a profile with a Holes2d point for placement
            ps = doc.ProfileSets.Add()
            plane = doc.RefPlanes.Item(1)  # Top plane
            profile = ps.Profiles.Add(plane)
            profile.Holes2d.Add(x, y)
            profile.End(0)

            # Holes.AddFinite(Profile, PlaneSide, FiniteDepth, HoleData)
            holes = model.Holes
            hole = holes.AddFinite(profile, ExtrudedProtrusion.igRight, depth, hole_data)

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

    def create_round(self, radius: float, edge_indices: Optional[List[int]] = None) -> Dict[str, Any]:
        """
        Create a round (fillet) feature on body edges.

        Rounds all edges of the first face, or specific edges if indices are given.
        Uses model.Rounds.Add(count, edgeArray, radiusArray).

        Args:
            radius: Round radius in meters
            edge_indices: Not used currently; rounds all edges of first face

        Returns:
            Dict with status and round info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to add rounds to"}

            model = models.Item(1)
            body = model.Body

            # Get edges from body faces (body.Edges() doesn't work in late binding)
            faces = body.Faces(6)  # igQueryAll = 6
            if faces.Count == 0:
                return {"error": "No faces found on body"}

            # Collect edges from all faces (deduplicated by COM identity)
            edge_list = []
            radius_list = []
            seen_edges = set()

            for fi in range(1, faces.Count + 1):
                face = faces.Item(fi)
                face_edges = face.Edges
                if not hasattr(face_edges, 'Count'):
                    continue
                for ei in range(1, face_edges.Count + 1):
                    edge = face_edges.Item(ei)
                    # Use id as dedup key (COM objects)
                    eid = id(edge)
                    if eid not in seen_edges:
                        seen_edges.add(eid)
                        edge_list.append(edge)
                        radius_list.append(radius)

            if not edge_list:
                return {"error": "No edges found on body"}

            rounds = model.Rounds
            rnd = rounds.Add(len(edge_list), edge_list, radius_list)

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

    def create_chamfer(self, distance: float, edge_indices: Optional[List[int]] = None) -> Dict[str, Any]:
        """
        Create a chamfer feature with equal setback on body edges.

        Chamfers all edges of the first face, or specific edges if indices are given.
        Uses model.Chamfers.AddEqualSetback(count, edgeArray, setbackDistance).

        Args:
            distance: Chamfer setback distance in meters
            edge_indices: Not used currently; chamfers all edges of first face

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

            # Get edges from body faces
            faces = body.Faces(6)  # igQueryAll = 6
            if faces.Count == 0:
                return {"error": "No faces found on body"}

            # Collect edges from all faces
            edge_list = []
            seen_edges = set()

            for fi in range(1, faces.Count + 1):
                face = faces.Item(fi)
                face_edges = face.Edges
                if not hasattr(face_edges, 'Count'):
                    continue
                for ei in range(1, face_edges.Count + 1):
                    edge = face_edges.Item(ei)
                    eid = id(edge)
                    if eid not in seen_edges:
                        seen_edges.add(eid)
                        edge_list.append(edge)

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
