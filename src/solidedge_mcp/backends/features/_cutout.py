"""Cutout feature operations (extruded, revolved, normal, lofted, swept, helix cutouts)."""

from typing import Any

import pythoncom
from win32com.client import VARIANT

from solidedge_mcp.backends.errors import error_result

from ..comutil import com_get, profile_origin
from ..constants import (
    AxisEndConstants,
    DirectionConstants,
    DocumentTypeConstants,
    ExtentTypeConstants,
    KeyPointExtentConstants,
    LoftSweepConstants,
    NormalCutoutMethodConstants,
)
from ..logging import get_logger
from ._base import verify_geometry_on_creators

_logger = get_logger(__name__)


@verify_geometry_on_creators
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
            return error_result(e)

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
            return error_result(e)

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
            return error_result(e)

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
            return error_result(e)

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

            # The Method argument is igSMFaceCutout, a sheet metal face cutout.
            # On an ordinary part the call succeeds and removes nothing, which
            # @verifies_geometry then reports as an empty feature. Say why
            # first. Verified on Solid Edge 2026: finite works on a sheet metal
            # document and cuts nothing on a part, while the through_all and
            # through_next variants work on both.
            doc_type = com_get(doc, "Type")
            if doc_type is not None and doc_type != DocumentTypeConstants.igSheetMetalDocument:
                return {
                    "error": (
                        "A normal cutout to a finite depth is a sheet metal feature "
                        "and removes no material from an ordinary part. Use "
                        "create_extrude(operation='Cut'), or "
                        "create_normal_cutout(method='through_all'), which does work "
                        "on a part."
                    ),
                    "document_type": doc_type,
                }

            cutouts = model.NormalCutouts
            # AddFiniteMulti(NumProfiles, ProfileArray, ProfilePlaneSide, Depth, Method)
            # Method is required; omitting it raises 0x8002000F Parameter not optional.
            cutouts.AddFiniteMulti(
                1,
                (profile,),
                dir_const,
                distance,
                NormalCutoutMethodConstants.igSMFaceCutout,
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "normal_cutout",
                "distance": distance,
                "direction": direction,
            }
        except Exception as e:
            return error_result(e)

    def create_lofted_cutout(self, profile_indices: list[int] | None = None) -> dict[str, Any]:
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
            return error_result(e)

    def create_swept_cutout(self, path_profile_index: int | None = None) -> dict[str, Any]:
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
            v_paths = [path_profile]
            v_path_types = [_CS]

            # Cross-section arrays
            v_sections = cross_sections
            v_section_types = [_CS] * len(cross_sections)
            # A SAFEARRAY of SAFEARRAY(VT_R8): the inner VARIANTs are required, only
            # the outer wrapper is not. Dropping them broke the lofted cutout.
            v_origins = [
                VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, list(profile_origin(p)))
                for p in cross_sections
            ]
            v_seg: list[Any] = []

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
            return error_result(e)

    def create_helix_cutout(
        self,
        pitch: float,
        height: float,
        revolutions: float | None = None,
        direction: str = "Right",
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
            v_profiles = [profile]

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
            return error_result(e)

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
            return error_result(e)

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
            return error_result(e)

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
            return error_result(e)

    def create_extruded_cutout_by_keypoint(self, direction: str = "Normal") -> dict[str, Any]:
        """
        Create an extruded cutout up to a keypoint extent.

        Type library: ExtrudedCutouts.AddFiniteByKeyPointMulti(NumberOfProfiles,
        ProfileArray, ProfilePlaneSide, KeyPointOrTangentFace, KeyPointFlags) --
        five required arguments. The KeyPoint (or tangent face) is a selected
        model object that this server has no way to pick, so the call can never
        be formed; report that instead of failing inside COM.

        Args:
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with an explanatory error
        """
        return {
            "error": (
                "An extruded cutout to a keypoint needs a KeyPoint or tangent face "
                "object, which this server cannot select. Use "
                "create_extruded_cutout(distance), create_extruded_cutout_from_to(...) "
                "or the Solid Edge UI."
            ),
            "unsupported": True,
            "direction": direction,
        }

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
            return error_result(e)

    def create_revolved_cutout_by_keypoint(self) -> dict[str, Any]:
        """
        Create a revolved cutout up to a keypoint extent.

        Type library: RevolvedCutouts.AddFiniteByKeyPointMulti(NumberOfProfiles,
        ProfileArray, RefAxis, KeyPointOrTangentFace, KeyPointFlags,
        [ProfilePlaneSide]) -- five required arguments. The KeyPoint (or tangent
        face) is a selected model object that this server has no way to pick, so
        the call can never be formed; report that instead of failing inside COM.

        Returns:
            Dict with an explanatory error
        """
        return {
            "error": (
                "A revolved cutout to a keypoint needs a KeyPoint or tangent face "
                "object, which this server cannot select. Use "
                "create_revolved_cutout(angle) or the Solid Edge UI."
            ),
            "unsupported": True,
        }

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
            return error_result(e)

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
            return error_result(e)

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
            # AddFiniteByKeyPointMulti takes six arguments: NumProfiles,
            # ProfileArray, ProfilePlaneSide, KeyPointOrTangentFace, KeyPointFlags,
            # Method. The keypoint object cannot be selected through this API, so
            # the call can never be formed; say so instead of failing inside COM.
            del cutouts, side
            return {
                "error": (
                    "Normal cutout to a keypoint needs a KeyPoint or tangent face "
                    "object, which this server cannot select. Use "
                    "create_normal_cutout(distance) or the Solid Edge UI."
                ),
                "unsupported": True,
            }

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "normal_cutout_by_keypoint",
                "direction": direction,
            }
        except Exception as e:
            return error_result(e)

    def create_lofted_cutout_full(self, profile_indices: list[int] | None = None) -> dict[str, Any]:
        """
        Create a lofted cutout with guide curves support.

        Type library: LoftedCutouts.Add(NumSections, CrossSections,
        CrossSectionTypes, Origins, SegmentMaps, MaterialSide, StartExtentType,
        StartExtentDistance, StartSurfaceOrRefPlane, EndExtentType,
        EndExtentDistance, EndSurfaceOrRefPlane, StartTangentType,
        StartTangentMagnitude, EndTangentType, EndTangentMagnitude,
        [NumGuideCurves, GuideCurves]) -- sixteen required arguments; the four
        tangency ones were previously missing.

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
            v_seg: list[Any] = []

            lc = model.LoftedCutouts
            lc.Add(
                len(profiles),  # NumSections
                v_profiles,  # CrossSections
                v_types,  # CrossSectionTypes
                v_origins,  # Origins
                v_seg,  # SegmentMaps
                DirectionConstants.igRight,  # MaterialSide
                ExtentTypeConstants.igNone,  # StartExtentType
                0.0,  # StartExtentDistance
                None,  # StartSurfaceOrRefPlane
                ExtentTypeConstants.igNone,  # EndExtentType
                0.0,  # EndExtentDistance
                None,  # EndSurfaceOrRefPlane
                ExtentTypeConstants.igNone,  # StartTangentType
                0.0,  # StartTangentMagnitude
                ExtentTypeConstants.igNone,  # EndTangentType
                0.0,  # EndTangentMagnitude
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "lofted_cutout_full",
                "num_profiles": len(profiles),
                "method": "LoftedCutouts.Add",
            }
        except Exception as e:
            return error_result(e)

    def create_swept_cutout_multi_body(
        self,
        path_profile_index: int | None = None,
    ) -> dict[str, Any]:
        """
        Create a swept cutout that supports multi-body operations.

        Type library: SweptCutouts.AddMultiBody(NumCurves, TraceCurves,
        TraceCurveTypes, NumSections, CrossSections, CrossSectionTypes, Origins,
        SegmentMaps, MaterialSide, StartExtentType, StartExtentDistance,
        StartSurfaceOrRefPlane, EndExtentType, EndExtentDistance,
        EndSurfaceOrRefPlane, NumberOfBodies, BodyArray) -- seventeen required
        arguments; it is SweptCutouts.Add plus the trailing body list, which was
        previously missing.

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

            v_paths = [path_profile]
            v_path_types = [_CS]

            v_sections = cross_sections
            v_section_types = [_CS] * len(cross_sections)
            # A SAFEARRAY of SAFEARRAY(VT_R8): the inner VARIANTs are required, only
            # the outer wrapper is not. Dropping them broke the lofted cutout.
            v_origins = [
                VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, list(profile_origin(p)))
                for p in cross_sections
            ]
            v_seg: list[Any] = []

            body_arr = [model.Body]

            swept_cutouts = model.SweptCutouts
            swept_cutouts.AddMultiBody(
                1,  # NumCurves
                v_paths,  # TraceCurves
                v_path_types,  # TraceCurveTypes
                len(cross_sections),  # NumSections
                v_sections,  # CrossSections
                v_section_types,  # CrossSectionTypes
                v_origins,  # Origins
                v_seg,  # SegmentMaps
                DirectionConstants.igRight,  # MaterialSide
                ExtentTypeConstants.igNone,  # StartExtentType
                0.0,  # StartExtentDistance
                None,  # StartSurfaceOrRefPlane
                ExtentTypeConstants.igNone,  # EndExtentType
                0.0,  # EndExtentDistance
                None,  # EndSurfaceOrRefPlane
                1,  # NumberOfBodies
                body_arr,  # BodyArray
            )

            self.sketch_manager.clear_accumulated_profiles()
            return {
                "status": "created",
                "type": "swept_cutout_multi_body",
                "num_cross_sections": len(cross_sections),
                "method": "SweptCutouts.AddMultiBody",
            }
        except Exception as e:
            return error_result(e)

    def create_helix_cutout_sync(
        self,
        pitch: float,
        height: float,
        revolutions: float | None = None,
        direction: str = "Right",
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

            v_profiles = [profile]

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
            return error_result(e)

    def create_helix_cutout_from_to(
        self, from_plane_index: int, to_plane_index: int, pitch: float
    ) -> dict[str, Any]:
        """
        Create a helical cutout between two reference planes.

        Type library: HelixCutouts.AddFromTo(HelixAxis, AxisStart,
        NumCrossSections, CrossSectionArray, ProfileSide, Height, Pitch,
        NumberOfTurns, HelixDir, FromPlane, ToPlane, [6 optional taper/pitch
        parameters]) -- eleven required arguments. Height and NumberOfTurns are
        driven by the from/to planes, so they are passed as 0.0, matching the
        AddFromToSync call below.

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

            v_profiles = [profile]

            helix_cutouts = model.HelixCutouts
            helix_cutouts.AddFromTo(
                refaxis,  # HelixAxis
                AxisEndConstants.igStart,  # AxisStart
                1,  # NumCrossSections
                v_profiles,  # CrossSectionArray
                DirectionConstants.igRight,  # ProfileSide
                0.0,  # Height (driven by the from/to planes)
                pitch,  # Pitch
                0.0,  # NumberOfTurns (driven by the from/to planes)
                DirectionConstants.igRight,  # HelixDir
                from_plane,  # FromPlane
                to_plane,  # ToPlane
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
            return error_result(e)

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

            v_profiles = [profile]

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
            return error_result(e)

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
            return error_result(e)

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
            body_arr = [body]

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
            return error_result(e)

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
            body_arr = [body]

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
            return error_result(e)

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
            body_arr = [body]

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
            return error_result(e)

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
            body_arr = [body]

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
            return error_result(e)

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

            profile_array = [profile]

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
            return error_result(e)

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

            profile_array = [profile]

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
            return error_result(e)
