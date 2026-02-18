"""Cutout feature operations (extruded, revolved, normal, lofted, swept, helix cutouts)."""

import traceback
from typing import Any

import pythoncom
from win32com.client import VARIANT

from ..constants import (
    DirectionConstants,
    ExtentTypeConstants,
    KeyPointExtentConstants,
    LoftSweepConstants,
)


class CutoutMixin:
    """Mixin providing cutout/removal feature methods."""

    def create_extruded_cutout_from_to(
        self, from_plane_index: int, to_plane_index: int
    ) -> dict[str, Any]:
        """
        Create an extruded cutout between two reference planes.

        Uses ExtrudedCutouts.AddFromToMulti(NumProfiles, ProfileArray,
        FromFaceOrRefPlane, ToFaceOrRefPlane) on the collection.

        Args:
            from_plane_index: 1-based index of the starting reference plane
            to_plane_index: 1-based index of the ending reference plane

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
            ref_planes = doc.RefPlanes

            if from_plane_index < 1 or from_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid from_plane_index: "
                    f"{from_plane_index}. "
                    f"Count: {ref_planes.Count}"
                }
            if to_plane_index < 1 or to_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid to_plane_index: {to_plane_index}. Count: {ref_planes.Count}"
                }

            from_plane = ref_planes.Item(from_plane_index)
            to_plane = ref_planes.Item(to_plane_index)

            cutouts = model.ExtrudedCutouts
            cutouts.AddFromToMulti(1, (profile,), from_plane, to_plane)

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "extruded_cutout_from_to",
                "from_plane_index": from_plane_index,
                "to_plane_index": to_plane_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_extruded_cutout(self, distance: float, direction: str = "Normal") -> dict[str, Any]:
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
            cutouts.AddFiniteMulti(1, (profile,), dir_const, distance)

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "extruded_cutout",
                "distance": distance,
                "direction": direction,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_extruded_cutout_through_all(self, direction: str = "Normal") -> dict[str, Any]:
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
            cutouts.AddThroughAllMulti(1, (profile,), dir_const)

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "extruded_cutout_through_all",
                "direction": direction,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_revolved_cutout(self, angle: float = 360) -> dict[str, Any]:
        """
        Create a revolved cutout (cut) in the part using the active sketch profile.

        Uses model.RevolvedCutouts.AddFiniteMulti(
        NumProfiles, ProfileArray, RefAxis,
        PlaneSide, Angle).
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
                return {
                    "error": "No axis of revolution set. "
                    "Use set_axis_of_revolution() before "
                    "closing the sketch."
                }

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            import math

            angle_rad = math.radians(angle)

            cutouts = model.RevolvedCutouts
            cutouts.AddFiniteMulti(
                1,  # NumberOfProfiles
                (profile,),  # ProfileArray
                refaxis,  # ReferenceAxis
                DirectionConstants.igRight,  # ProfilePlaneSide
                angle_rad,  # AngleOfRevolution
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {"status": "created", "type": "revolved_cutout", "angle": angle}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_normal_cutout(self, distance: float, direction: str = "Normal") -> dict[str, Any]:
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
            cutouts.AddFiniteMulti(1, (profile,), dir_const, distance)

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "normal_cutout",
                "distance": distance,
                "direction": direction,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_lofted_cutout(self, profile_indices: list = None) -> dict[str, Any]:
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
                    "error": "Lofted cutout requires at "
                    "least 2 profiles, got "
                    f"{len(profiles)}. Create sketches on "
                    "different planes and close each one "
                    "before calling create_lofted_cutout()."
                }

            v_profiles, v_types, v_origins = self._make_loft_variant_arrays(profiles)

            lc = model.LoftedCutouts
            lc.AddSimple(
                len(profiles),
                v_profiles,
                v_types,
                v_origins,
                DirectionConstants.igRight,
                ExtentTypeConstants.igNone,
                ExtentTypeConstants.igNone,
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "lofted_cutout",
                "num_profiles": len(profiles),
                "method": "LoftedCutouts.AddSimple",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_swept_cutout(self, path_profile_index: int = None) -> dict[str, Any]:
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
                    "error": "Swept cutout requires at "
                    "least 2 profiles (path + "
                    "cross-section), got "
                    f"{len(all_profiles)}. Create a path "
                    "sketch and a cross-section "
                    "sketch first."
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
                pythoncom.VT_ARRAY | pythoncom.VT_I4, [_CS] * len(cross_sections)
            )
            v_origins = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
                [VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0]) for _ in cross_sections],
            )
            v_seg = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, [])

            # SweptCutouts.Add: same 15 params as SweptProtrusions
            swept_cutouts = model.SweptCutouts
            swept_cutouts.Add(
                1,
                v_paths,
                v_path_types,  # Path (1 curve)
                len(cross_sections),
                v_sections,
                v_section_types,
                v_origins,
                v_seg,
                DirectionConstants.igRight,  # MaterialSide
                ExtentTypeConstants.igNone,
                0.0,
                None,  # Start extent
                ExtentTypeConstants.igNone,
                0.0,
                None,  # End extent
            )

            self.sketch_manager.clear_accumulated_profiles()
            return {
                "status": "created",
                "type": "swept_cutout",
                "num_cross_sections": len(cross_sections),
                "method": "model.SweptCutouts.Add",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_helix_cutout(
        self, pitch: float, height: float, revolutions: float = None, direction: str = "Right"
    ) -> dict[str, Any]:
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
                return {
                    "error": "No axis of revolution set. "
                    "Use set_axis_of_revolution() "
                    "in the sketch."
                }

            if revolutions is None:
                revolutions = height / pitch

            axis_start = DirectionConstants.igRight
            dir_const = (
                DirectionConstants.igRight if direction == "Right" else DirectionConstants.igLeft
            )

            # Wrap cross-section profile in SAFEARRAY
            v_profiles = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

            helix_cutouts = model.HelixCutouts
            helix_cutouts.AddFinite(
                refaxis,  # HelixAxis
                axis_start,  # AxisStart
                1,  # NumCrossSections
                v_profiles,  # CrossSectionArray
                DirectionConstants.igRight,  # ProfileSide
                height,  # Height
                pitch,  # Pitch
                revolutions,  # NumberOfTurns
                dir_const,  # HelixDir
            )

            self.sketch_manager.clear_accumulated_profiles()
            return {
                "status": "created",
                "type": "helix_cutout",
                "pitch": pitch,
                "height": height,
                "revolutions": revolutions,
                "direction": direction,
                "method": "model.HelixCutouts.AddFinite",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_extruded_cutout_through_next(self, direction: str = "Normal") -> dict[str, Any]:
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
            cutouts.AddThroughNextMulti(1, (profile,), dir_const)

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "extruded_cutout_through_next",
                "direction": direction,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_normal_cutout_through_all(self, direction: str = "Normal") -> dict[str, Any]:
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
            cutouts.AddThroughAllMulti(1, (profile,), dir_const, 0)

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "normal_cutout_through_all",
                "direction": direction,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_extruded_cutout_from_to_v2(
        self, from_plane_index: int, to_plane_index: int
    ) -> dict[str, Any]:
        """
        Create an extruded cutout between two reference planes (multi-profile API).

        Uses ExtrudedCutouts.AddFromToMulti(NumProfiles, ProfileArray,
        FromFaceOrRefPlane, ToFaceOrRefPlane).

        Args:
            from_plane_index: 1-based index of the starting reference plane
            to_plane_index: 1-based index of the ending reference plane

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
            ref_planes = doc.RefPlanes

            if from_plane_index < 1 or from_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid from_plane_index: "
                    f"{from_plane_index}. "
                    f"Count: {ref_planes.Count}"
                }
            if to_plane_index < 1 or to_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid to_plane_index: {to_plane_index}. Count: {ref_planes.Count}"
                }

            from_plane = ref_planes.Item(from_plane_index)
            to_plane = ref_planes.Item(to_plane_index)

            cutouts = model.ExtrudedCutouts
            cutouts.AddFromToMulti(1, (profile,), from_plane, to_plane)

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "extruded_cutout_from_to_v2",
                "from_plane_index": from_plane_index,
                "to_plane_index": to_plane_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_extruded_cutout_by_keypoint(self, direction: str = "Normal") -> dict[str, Any]:
        """
        Create an extruded cutout up to a keypoint extent.

        Uses ExtrudedCutouts.AddFiniteByKeyPointMulti(NumProfiles, ProfileArray, PlaneSide).

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
            side = direction_map.get(direction, DirectionConstants.igRight)

            cutouts = model.ExtrudedCutouts
            cutouts.AddFiniteByKeyPointMulti(1, (profile,), side)

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "extruded_cutout_by_keypoint",
                "direction": direction,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_revolved_cutout_sync(self, angle: float = 360.0) -> dict[str, Any]:
        """
        Create a synchronous revolved cutout.

        Uses RevolvedCutouts.AddFiniteMultiSync(NumProfiles, ProfileArray,
        RefAxis, PlaneSide, AngleOfRevolution).

        Args:
            angle: Revolution angle in degrees (360 for full revolution)

        Returns:
            Dict with status and cutout info
        """
        try:
            import math

            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()
            refaxis = self.sketch_manager.get_active_refaxis()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}
            if not refaxis:
                return {
                    "error": "No axis of revolution set. "
                    "Use set_axis_of_revolution() before "
                    "closing the sketch."
                }

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            angle_rad = math.radians(angle)

            cutouts = model.RevolvedCutouts
            cutouts.AddFiniteMultiSync(
                1,  # NumProfiles
                (profile,),  # ProfileArray
                refaxis,  # RefAxis
                DirectionConstants.igRight,  # PlaneSide
                angle_rad,  # AngleOfRevolution
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {"status": "created", "type": "revolved_cutout_sync", "angle": angle}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_revolved_cutout_by_keypoint(self) -> dict[str, Any]:
        """
        Create a revolved cutout up to a keypoint extent.

        Uses RevolvedCutouts.AddFiniteByKeyPointMulti(NumProfiles, ProfileArray,
        RefAxis, PlaneSide).

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
                return {
                    "error": "No axis of revolution set. "
                    "Use set_axis_of_revolution() before "
                    "closing the sketch."
                }

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            cutouts = model.RevolvedCutouts
            cutouts.AddFiniteByKeyPointMulti(
                1,  # NumProfiles
                (profile,),  # ProfileArray
                refaxis,  # RefAxis
                DirectionConstants.igRight,  # PlaneSide
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {"status": "created", "type": "revolved_cutout_by_keypoint"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_normal_cutout_from_to(
        self, from_plane_index: int, to_plane_index: int
    ) -> dict[str, Any]:
        """
        Create a normal cutout between two reference planes.

        Uses NormalCutouts.AddFromToMulti(NumProfiles, ProfileArray,
        FromFaceOrRefPlane, ToFaceOrRefPlane, Method).

        Args:
            from_plane_index: 1-based index of the starting reference plane
            to_plane_index: 1-based index of the ending reference plane

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
            ref_planes = doc.RefPlanes

            if from_plane_index < 1 or from_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid from_plane_index: "
                    f"{from_plane_index}. "
                    f"Count: {ref_planes.Count}"
                }
            if to_plane_index < 1 or to_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid to_plane_index: {to_plane_index}. Count: {ref_planes.Count}"
                }

            from_plane = ref_planes.Item(from_plane_index)
            to_plane = ref_planes.Item(to_plane_index)

            cutouts = model.NormalCutouts
            cutouts.AddFromToMulti(1, (profile,), from_plane, to_plane, 0)

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "normal_cutout_from_to",
                "from_plane_index": from_plane_index,
                "to_plane_index": to_plane_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_normal_cutout_through_next(self, direction: str = "Normal") -> dict[str, Any]:
        """
        Create a normal cutout through the next face.

        Uses NormalCutouts.AddThroughNextMulti(NumProfiles, ProfileArray, PlaneSide, Method).

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

            cutouts = model.NormalCutouts
            cutouts.AddThroughNextMulti(1, (profile,), dir_const, 0)

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "normal_cutout_through_next",
                "direction": direction,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_normal_cutout_by_keypoint(self, direction: str = "Normal") -> dict[str, Any]:
        """
        Create a normal cutout up to a keypoint extent.

        Uses NormalCutouts.AddFiniteByKeyPointMulti(NumProfiles, ProfileArray,
        PlaneSide, Method).

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
            side = direction_map.get(direction, DirectionConstants.igRight)

            cutouts = model.NormalCutouts
            cutouts.AddFiniteByKeyPointMulti(1, (profile,), side, 0)

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "normal_cutout_by_keypoint",
                "direction": direction,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_lofted_cutout_full(self, profile_indices: list = None) -> dict[str, Any]:
        """
        Create a lofted cutout with guide curves support.

        Uses LoftedCutouts.Add(NumCrossSections, CrossSectionArray, CrossSectionTypes,
        Origins, SegmentMaps, PlaneSide, StartExtent, ..., EndExtent, ...).
        Provides the full API with all extent and treatment parameters.

        Args:
            profile_indices: Optional list of profile indices to use from
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

            all_profiles = self.sketch_manager.get_accumulated_profiles()

            if profile_indices is not None:
                profiles = [all_profiles[i] for i in profile_indices]
            else:
                profiles = all_profiles

            if len(profiles) < 2:
                return {
                    "error": "Lofted cutout requires at "
                    "least 2 profiles, got "
                    f"{len(profiles)}. Create sketches on "
                    "different planes and close each one "
                    "before calling create_lofted_cutout_full()."
                }

            v_profiles, v_types, v_origins = self._make_loft_variant_arrays(profiles)
            v_seg = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, [])

            lc = model.LoftedCutouts
            lc.Add(
                len(profiles),
                v_profiles,
                v_types,
                v_origins,
                v_seg,
                DirectionConstants.igRight,
                ExtentTypeConstants.igNone,
                0.0,
                None,
                ExtentTypeConstants.igNone,
                0.0,
                None,
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "lofted_cutout_full",
                "num_profiles": len(profiles),
                "method": "LoftedCutouts.Add",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_swept_cutout_multi_body(self, path_profile_index: int = None) -> dict[str, Any]:
        """
        Create a swept cutout that supports multi-body operations.

        Uses SweptCutouts.AddMultiBody with the same parameters as SweptCutouts.Add.
        Allows the cutout to span across multiple bodies in the part.

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
                    "error": "Swept cutout requires at "
                    "least 2 profiles (path + "
                    "cross-section), got "
                    f"{len(all_profiles)}. Create a path "
                    "sketch and a cross-section "
                    "sketch first."
                }

            path_idx = path_profile_index if path_profile_index is not None else 0
            path_profile = all_profiles[path_idx]
            cross_sections = [p for i, p in enumerate(all_profiles) if i != path_idx]

            _CS = LoftSweepConstants.igProfileBasedCrossSection

            v_paths = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [path_profile])
            v_path_types = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I4, [_CS])

            v_sections = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, cross_sections)
            v_section_types = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_I4, [_CS] * len(cross_sections)
            )
            v_origins = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
                [VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0]) for _ in cross_sections],
            )
            v_seg = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_VARIANT, [])

            swept_cutouts = model.SweptCutouts
            swept_cutouts.AddMultiBody(
                1,
                v_paths,
                v_path_types,
                len(cross_sections),
                v_sections,
                v_section_types,
                v_origins,
                v_seg,
                DirectionConstants.igRight,
                ExtentTypeConstants.igNone,
                0.0,
                None,
                ExtentTypeConstants.igNone,
                0.0,
                None,
            )

            self.sketch_manager.clear_accumulated_profiles()
            return {
                "status": "created",
                "type": "swept_cutout_multi_body",
                "num_cross_sections": len(cross_sections),
                "method": "SweptCutouts.AddMultiBody",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_helix_cutout_sync(
        self, pitch: float, height: float, revolutions: float = None, direction: str = "Right"
    ) -> dict[str, Any]:
        """
        Create a synchronous helical cutout.

        Uses HelixCutouts.AddFiniteSync(HelixAxis, AxisStart, NumCrossSections,
        CrossSectionArray, ProfileSide, Height, Pitch, NumberOfTurns, HelixDir).

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
                return {
                    "error": "No axis of revolution set. "
                    "Use set_axis_of_revolution() before "
                    "closing the sketch."
                }

            if revolutions is None:
                revolutions = height / pitch

            dir_const = (
                DirectionConstants.igRight if direction == "Right" else DirectionConstants.igLeft
            )

            v_profiles = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

            helix_cutouts = model.HelixCutouts
            helix_cutouts.AddFiniteSync(
                refaxis,  # HelixAxis
                DirectionConstants.igRight,  # AxisStart
                1,  # NumCrossSections
                v_profiles,  # CrossSectionArray
                DirectionConstants.igRight,  # ProfileSide
                height,  # Height
                pitch,  # Pitch
                revolutions,  # NumberOfTurns
                dir_const,  # HelixDir
            )

            self.sketch_manager.clear_accumulated_profiles()
            return {
                "status": "created",
                "type": "helix_cutout_sync",
                "pitch": pitch,
                "height": height,
                "revolutions": revolutions,
                "direction": direction,
                "method": "HelixCutouts.AddFiniteSync",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_helix_cutout_from_to(
        self, from_plane_index: int, to_plane_index: int, pitch: float
    ) -> dict[str, Any]:
        """
        Create a helical cutout between two reference planes.

        Uses HelixCutouts.AddFromTo(HelixAxis, AxisStart, NumCrossSections,
        CrossSectionArray, ProfileSide, FromFace, ToFace, Pitch, HelixDir).

        Args:
            from_plane_index: 1-based index of the starting reference plane
            to_plane_index: 1-based index of the ending reference plane
            pitch: Distance between coils in meters

        Returns:
            Dict with status and helix cutout info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()
            refaxis = self.sketch_manager.get_active_refaxis()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}
            if not refaxis:
                return {
                    "error": "No axis of revolution set. "
                    "Use set_axis_of_revolution() before "
                    "closing the sketch."
                }

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            ref_planes = doc.RefPlanes

            if from_plane_index < 1 or from_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid from_plane_index: "
                    f"{from_plane_index}. "
                    f"Count: {ref_planes.Count}"
                }
            if to_plane_index < 1 or to_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid to_plane_index: {to_plane_index}. Count: {ref_planes.Count}"
                }

            from_plane = ref_planes.Item(from_plane_index)
            to_plane = ref_planes.Item(to_plane_index)

            v_profiles = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

            helix_cutouts = model.HelixCutouts
            helix_cutouts.AddFromTo(
                refaxis,  # HelixAxis
                DirectionConstants.igRight,  # AxisStart
                1,  # NumCrossSections
                v_profiles,  # CrossSectionArray
                DirectionConstants.igRight,  # ProfileSide
                from_plane,  # FromFace
                to_plane,  # ToFace
                pitch,  # Pitch
                DirectionConstants.igRight,  # HelixDir
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "helix_cutout_from_to",
                "from_plane_index": from_plane_index,
                "to_plane_index": to_plane_index,
                "pitch": pitch,
                "method": "HelixCutouts.AddFromTo",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_helix_cutout_from_to_sync(
        self, from_plane_index: int, to_plane_index: int, pitch: float
    ) -> dict[str, Any]:
        """
        Create a synchronous helical cutout between two reference planes.

        Uses HelixCutouts.AddFromToSync.

        Args:
            from_plane_index: 1-based index of the starting reference plane
            to_plane_index: 1-based index of the ending reference plane
            pitch: Distance between coils in meters

        Returns:
            Dict with status and helix cutout info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()
            refaxis = self.sketch_manager.get_active_refaxis()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}
            if not refaxis:
                return {
                    "error": "No axis of revolution set. "
                    "Use set_axis_of_revolution() before "
                    "closing the sketch."
                }

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            ref_planes = doc.RefPlanes

            if from_plane_index < 1 or from_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid from_plane_index: {from_plane_index}. "
                    f"Count: {ref_planes.Count}"
                }
            if to_plane_index < 1 or to_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid to_plane_index: {to_plane_index}. Count: {ref_planes.Count}"
                }

            from_plane = ref_planes.Item(from_plane_index)
            to_plane = ref_planes.Item(to_plane_index)

            v_profiles = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

            helix_cutouts = model.HelixCutouts
            helix_cutouts.AddFromToSync(
                refaxis,  # HelixAxis
                DirectionConstants.igRight,  # AxisStart
                1,  # NumCrossSections
                v_profiles,  # CrossSectionArray
                DirectionConstants.igRight,  # ProfileSide
                0.0,  # Height
                pitch,  # Pitch
                0.0,  # NumberOfTurns
                DirectionConstants.igRight,  # HelixDir
                from_plane,  # FromPlane
                to_plane,  # ToPlane
                0.0,  # TaperAngle
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "helix_cutout_from_to_sync",
                "from_plane_index": from_plane_index,
                "to_plane_index": to_plane_index,
                "pitch": pitch,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_extruded_cutout_through_next_single(
        self, direction: str = "Normal"
    ) -> dict[str, Any]:
        """
        Create a single-profile extruded cutout through the next face.

        Uses ExtrudedCutouts.AddThroughNext(Profile, ProfileSide,
        ProfilePlaneSide).

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
            dir_const = (
                DirectionConstants.igRight if direction == "Normal" else DirectionConstants.igLeft
            )

            cutouts = model.ExtrudedCutouts
            cutouts.AddThroughNext(
                profile,
                dir_const,  # ProfileSide
                dir_const,  # ProfilePlaneSide
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "extruded_cutout_through_next_single",
                "direction": direction,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_extruded_cutout_multi_body(
        self, distance: float, direction: str = "Normal"
    ) -> dict[str, Any]:
        """
        Create a multi-body extruded cutout.

        Uses ExtrudedCutouts.AddFiniteMultiBody(NumProfiles, ProfileArray,
        ProfileSide, ProfilePlaneSide, Depth, NumBodies, BodyArray).

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
            dir_const = (
                DirectionConstants.igRight if direction == "Normal" else DirectionConstants.igLeft
            )

            body = model.Body
            body_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [body])

            cutouts = model.ExtrudedCutouts
            cutouts.AddFiniteMultiBody(
                1,  # NumberOfProfiles
                (profile,),  # ProfileArray
                dir_const,  # ProfileSide
                dir_const,  # ProfilePlaneSide
                distance,  # Depth
                1,  # NumberOfBodies
                body_arr,  # BodyArray
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "extruded_cutout_multi_body",
                "distance": distance,
                "direction": direction,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_extruded_cutout_from_to_multi_body(
        self, from_plane_index: int, to_plane_index: int
    ) -> dict[str, Any]:
        """
        Create a multi-body extruded cutout between two reference planes.

        Uses ExtrudedCutouts.AddFromToMultiBody(NumProfiles, ProfileArray,
        ProfileSide, FromFace, ToFace, NumBodies, BodyArray).

        Args:
            from_plane_index: 1-based index of the starting reference plane
            to_plane_index: 1-based index of the ending reference plane

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
            ref_planes = doc.RefPlanes

            if from_plane_index < 1 or from_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid from_plane_index: {from_plane_index}. "
                    f"Count: {ref_planes.Count}"
                }
            if to_plane_index < 1 or to_plane_index > ref_planes.Count:
                return {
                    "error": f"Invalid to_plane_index: {to_plane_index}. Count: {ref_planes.Count}"
                }

            from_plane = ref_planes.Item(from_plane_index)
            to_plane = ref_planes.Item(to_plane_index)

            body = model.Body
            body_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [body])

            cutouts = model.ExtrudedCutouts
            cutouts.AddFromToMultiBody(
                1,  # NumberOfProfiles
                (profile,),  # ProfileArray
                DirectionConstants.igRight,  # ProfileSide
                from_plane,  # FromFace
                to_plane,  # ToFace
                1,  # NumberOfBodies
                body_arr,  # BodyArray
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "extruded_cutout_from_to_multi_body",
                "from_plane_index": from_plane_index,
                "to_plane_index": to_plane_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_extruded_cutout_through_all_multi_body(
        self, direction: str = "Normal"
    ) -> dict[str, Any]:
        """
        Create a multi-body extruded cutout through all material.

        Uses ExtrudedCutouts.AddThroughAllMultiBody(NumProfiles, ProfileArray,
        ProfileSide, ProfilePlaneSide, NumBodies, BodyArray).

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
            dir_const = (
                DirectionConstants.igRight if direction == "Normal" else DirectionConstants.igLeft
            )

            body = model.Body
            body_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [body])

            cutouts = model.ExtrudedCutouts
            cutouts.AddThroughAllMultiBody(
                1,  # NumberOfProfiles
                (profile,),  # ProfileArray
                dir_const,  # ProfileSide
                dir_const,  # ProfilePlaneSide
                1,  # NumberOfBodies
                body_arr,  # BodyArray
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "extruded_cutout_through_all_multi_body",
                "direction": direction,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_revolved_cutout_multi_body(self, angle: float = 360.0) -> dict[str, Any]:
        """
        Create a multi-body revolved cutout.

        Uses RevolvedCutouts.AddFiniteMultiBody.

        Args:
            angle: Revolution angle in degrees

        Returns:
            Dict with status and cutout info
        """
        try:
            import math

            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()
            refaxis = self.sketch_manager.get_active_refaxis()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}
            if not refaxis:
                return {
                    "error": "No axis of revolution set. "
                    "Use set_axis_of_revolution() before "
                    "closing the sketch."
                }

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            angle_rad = math.radians(angle)

            body = model.Body
            body_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [body])

            cutouts = model.RevolvedCutouts
            cutouts.AddFiniteMultiBody(
                1,  # NumberOfProfiles
                (profile,),  # ProfileArray
                refaxis,  # RefAxis
                DirectionConstants.igRight,  # ProfileSide
                DirectionConstants.igRight,  # ProfilePlaneSide
                angle_rad,  # AngleOfRevolution
                1,  # NumberOfBodies
                body_arr,  # BodyArray
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "revolved_cutout_multi_body",
                "angle": angle,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_revolved_cutout_full(self, angle: float = 360.0) -> dict[str, Any]:
        """
        Create a revolved cutout with full extent parameters.

        Uses RevolvedCutouts.Add with dual-extent params.

        Args:
            angle: Revolution angle in degrees

        Returns:
            Dict with status and cutout info
        """
        try:
            import math

            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()
            refaxis = self.sketch_manager.get_active_refaxis()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}
            if not refaxis:
                return {
                    "error": "No axis of revolution set. "
                    "Use set_axis_of_revolution() before "
                    "closing the sketch."
                }

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            angle_rad = math.radians(angle)

            profile_array = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

            cutouts = model.RevolvedCutouts
            cutouts.Add(
                1,
                profile_array,
                refaxis,
                DirectionConstants.igRight,
                ExtentTypeConstants.igFinite,
                DirectionConstants.igRight,
                angle_rad,
                None,
                KeyPointExtentConstants.igTangentNormal,
                ExtentTypeConstants.igNone,
                DirectionConstants.igRight,
                0.0,
                None,
                KeyPointExtentConstants.igTangentNormal,
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {"status": "created", "type": "revolved_cutout_full", "angle": angle}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_revolved_cutout_full_sync(self, angle: float = 360.0) -> dict[str, Any]:
        """
        Create a synchronous revolved cutout with full extent parameters.

        Uses RevolvedCutouts.AddSync with dual-extent params.

        Args:
            angle: Revolution angle in degrees

        Returns:
            Dict with status and cutout info
        """
        try:
            import math

            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()
            refaxis = self.sketch_manager.get_active_refaxis()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}
            if not refaxis:
                return {
                    "error": "No axis of revolution set. "
                    "Use set_axis_of_revolution() before "
                    "closing the sketch."
                }

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            angle_rad = math.radians(angle)

            profile_array = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

            cutouts = model.RevolvedCutouts
            cutouts.AddSync(
                1,
                profile_array,
                refaxis,
                DirectionConstants.igRight,
                ExtentTypeConstants.igFinite,
                DirectionConstants.igRight,
                angle_rad,
                None,
                KeyPointExtentConstants.igTangentNormal,
                ExtentTypeConstants.igNone,
                DirectionConstants.igRight,
                0.0,
                None,
                KeyPointExtentConstants.igTangentNormal,
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {"status": "created", "type": "revolved_cutout_full_sync", "angle": angle}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
