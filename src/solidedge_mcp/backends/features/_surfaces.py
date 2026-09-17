"""Surface feature operations."""

from typing import Any

import pythoncom
from win32com.client import VARIANT

from solidedge_mcp.backends.errors import error_result

from ..comutil import profile_origin
from ..constants import (
    DirectionConstants,
    DraftSideConstants,
    ExtentTypeConstants,
    KeyPointExtentConstants,
    LoftSweepConstants,
    OffsetSideConstants,
    TreatmentCrownCurvatureSideConstants,
    TreatmentCrownSideConstants,
    TreatmentCrownTypeConstants,
    TreatmentTypeConstants,
)
from ..logging import get_logger
from ._base import verify_collection_growth_on_creators

_logger = get_logger(__name__)


# A surface is construction geometry, not body material, so the face count
# that verifies_geometry watches never moves for one. Every creator here
# lands in one of five Constructions sub-collections, and which one depends
# on the call, so all five are summed: a surface that lands in a sibling
# collection is still growth, not a no-op.
@verify_collection_growth_on_creators(
    "Constructions.ExtrudedSurfaces",
    "Constructions.RevolvedSurfaces",
    "Constructions.LoftedSurfaces",
    "Constructions.SweptSurfaces",
    "Constructions.BlueSurfs",
)
class SurfacesMixin:
    """Mixin providing surface creation methods."""

    def create_extruded_surface(
        self, distance: float, direction: str = "Normal", end_caps: bool = True
    ) -> dict[str, Any]:
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
            profile_array = [profile]

            depth1 = distance
            symmetric = direction == "Symmetric"
            depth2 = distance if symmetric else 0.0
            # A second extent of igFinite with a depth of zero is rejected
            # with E_INVALIDARG, so a one-sided surface must say igNone.
            # Verified on Solid Edge 2026; the extrude and revolve calls
            # already do this. Every non-symmetric extruded surface failed.
            extent2 = ExtentTypeConstants.igFinite if symmetric else ExtentTypeConstants.igNone
            side1 = DirectionConstants.igRight
            side2 = (
                DirectionConstants.igLeft
                if direction == "Symmetric"
                else DirectionConstants.igRight
            )

            extruded_surfaces.Add(
                1,  # NumberOfProfiles
                profile_array,  # ProfileArray
                ExtentTypeConstants.igFinite,  # ExtentType1
                side1,  # ExtentSide1
                depth1,  # FiniteDepth1
                None,  # KeyPointOrTangentFace1
                KeyPointExtentConstants.igTangentNormal,  # KeyPointFlags1
                None,  # FromFaceOrRefPlane
                OffsetSideConstants.seOffsetNone,  # FromFaceOffsetSide
                0.0,  # FromFaceOffsetDistance
                TreatmentTypeConstants.seTreatmentNone,  # TreatmentType1
                DraftSideConstants.seDraftNone,  # TreatmentDraftSide1
                0.0,  # TreatmentDraftAngle1
                TreatmentCrownTypeConstants.seTreatmentCrownByOffset,  # TreatmentCrownType1
                TreatmentCrownSideConstants.seTreatmentCrownSideInside,  # TreatmentCrownSide1
                # TreatmentCrownCurvatureSide1
                TreatmentCrownCurvatureSideConstants.seTreatmentCrownCurvatureInside,
                0.0,  # TreatmentCrownRadiusOrOffset1
                0.0,  # TreatmentCrownTakeOffAngle1
                extent2,  # ExtentType2
                side2,  # ExtentSide2
                depth2,  # FiniteDepth2
                None,  # KeyPointOrTangentFace2
                KeyPointExtentConstants.igTangentNormal,  # KeyPointFlags2
                None,  # ToFaceOrRefPlane
                OffsetSideConstants.seOffsetNone,  # ToFaceOffsetSide
                0.0,  # ToFaceOffsetDistance
                TreatmentTypeConstants.seTreatmentNone,  # TreatmentType2
                DraftSideConstants.seDraftNone,  # TreatmentDraftSide2
                0.0,  # TreatmentDraftAngle2
                TreatmentCrownTypeConstants.seTreatmentCrownByOffset,  # TreatmentCrownType2
                TreatmentCrownSideConstants.seTreatmentCrownSideInside,  # TreatmentCrownSide2
                # TreatmentCrownCurvatureSide2
                TreatmentCrownCurvatureSideConstants.seTreatmentCrownCurvatureInside,
                0.0,  # TreatmentCrownRadiusOrOffset2
                0.0,  # TreatmentCrownTakeOffAngle2
                end_caps,  # WantEndCaps
            )

            return {
                "status": "created",
                "type": "extruded_surface",
                "distance": distance,
                "direction": direction,
                "end_caps": end_caps,
            }
        except Exception as e:
            return error_result(e)

    def create_revolved_surface(
        self, angle: float = 360, want_end_caps: bool = False
    ) -> dict[str, Any]:
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

            v_profiles = [profile]

            # Surfaces live on doc.Constructions. Model has no
            # RevolvedSurfaces property and Models has no
            # AddFiniteRevolvedSurface, so both branches raised; the extruded
            # surface methods in this file already take the right route.
            del models
            rev_surfaces = doc.Constructions.RevolvedSurfaces
            rev_surfaces.AddFinite(
                1, v_profiles, refaxis, DirectionConstants.igRight, angle_rad, want_end_caps
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "revolved_surface",
                "angle_degrees": angle,
                "want_end_caps": want_end_caps,
            }
        except Exception as e:
            return error_result(e)

    def create_lofted_surface(self, want_end_caps: bool = False) -> dict[str, Any]:
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

            all_profiles = self.sketch_manager.get_accumulated_profiles()

            if len(all_profiles) < 2:
                return {
                    "error": f"Lofted surface requires at least 2 profiles, "
                    f"got {len(all_profiles)}."
                }

            # Constructions.LoftedSurfaces.Add answers E_INVALIDARG to every
            # argument shape that builds Models.AddLoftedProtrusion from the
            # same two profiles (Solid Edge 2026: coordinate-array origins,
            # element origins, VARIANT-wrapped arrays, every extent and tangent
            # constant, with and without guide curves -- 13 shapes). Say so
            # rather than raise it.
            del doc
            return {
                "error": (
                    "LoftedSurfaces.Add rejects every argument form Solid Edge 2026 "
                    "accepts for a solid loft (E_INVALIDARG). Use create_loft for a "
                    "solid, or the Solid Edge UI for a lofted surface."
                ),
                "unsupported": True,
                "num_profiles": len(all_profiles),
                "want_end_caps": want_end_caps,
            }
        except Exception as e:
            return error_result(e)

    def create_swept_surface(
        self, path_profile_index: int | None = None, want_end_caps: bool = False
    ) -> dict[str, Any]:
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
        # Constructions.SweptSurfaces.Add takes Solid Edge 2026 down -- the
        # process, not the call -- once a session has been through a few
        # documents: three crashes in five calls, each after other work in the
        # same instance, with a circle section as well as a rectangle. On a
        # fresh instance the call builds, and the argument shape that does is
        # worth keeping on record: TraceCurveTypes and CrossSectionTypes are
        # single igProfileBasedCrossSection values, Origins is [element] where
        # element comes from profile_origin_element(section), and OriginRefs
        # is that element's keypoint (igKeyPointCenter for a circle). None for
        # the origins is E_FAIL. A crash is worse than a refusal, so no call.
        all_profiles = self.sketch_manager.get_accumulated_profiles()
        return {
            "error": (
                "SweptSurfaces.Add crashes Solid Edge 2026 after other work in the "
                "same session (3 of 5 calls). Use create_sweep for a solid, or the "
                "Solid Edge UI for a swept surface."
            ),
            "unsupported": True,
            "num_profiles": len(all_profiles),
            "path_profile_index": path_profile_index,
            "want_end_caps": want_end_caps,
        }

    def create_extruded_surface_from_to(
        self, from_plane_index: int, to_plane_index: int
    ) -> dict[str, Any]:
        """
        Create an extruded surface between two reference planes.

        Uses ExtrudedSurfaces.AddFromTo(NumberOfProfiles, ProfileArray,
        FromFaceOrRefPlane, ToFaceOrRefPlane, WantEndCaps).

        Args:
            from_plane_index: 1-based index of the start reference plane
            to_plane_index: 1-based index of the end reference plane

        Returns:
            Dict with status and surface info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

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

            profile_array = [profile]

            constructions = doc.Constructions
            extruded_surfaces = constructions.ExtrudedSurfaces

            extruded_surfaces.AddFromTo(
                1,  # NumberOfProfiles
                profile_array,  # ProfileArray
                from_plane,  # FromFaceOrRefPlane
                to_plane,  # ToFaceOrRefPlane
                True,  # WantEndCaps
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "extruded_surface_from_to",
                "from_plane_index": from_plane_index,
                "to_plane_index": to_plane_index,
            }
        except Exception as e:
            return error_result(e)

    def create_extruded_surface_by_keypoint(self, keypoint_type: str = "End") -> dict[str, Any]:
        """
        Create an extruded surface up to a keypoint extent.

        Uses ExtrudedSurfaces.AddFiniteByKeyPoint(NumberOfProfiles, ProfileArray,
        ProfilePlaneSide, KeyPointOrTangentFace, KeyPointFlags, WantEndCaps).

        Args:
            keypoint_type: 'Start' or 'End' keypoint for the extent

        Returns:
            Dict with status and surface info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            profile_array = [profile]

            constructions = doc.Constructions
            extruded_surfaces = constructions.ExtrudedSurfaces

            extruded_surfaces.AddFiniteByKeyPoint(
                1,  # NumberOfProfiles
                profile_array,  # ProfileArray
                DirectionConstants.igRight,  # ProfilePlaneSide
                None,  # KeyPointOrTangentFace
                KeyPointExtentConstants.igTangentNormal,  # KeyPointFlags
                True,  # WantEndCaps
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "extruded_surface_by_keypoint",
                "keypoint_type": keypoint_type,
            }
        except Exception as e:
            return error_result(e)

    def create_extruded_surface_by_curves(
        self, distance: float, direction: str = "Normal"
    ) -> dict[str, Any]:
        """
        Create an extruded surface by curves (extrude along curve path).

        Uses ExtrudedSurfaces.AddByCurves with full treatment params.
        This uses curves (profiles) rather than standard profile extrusion.

        Args:
            distance: Extrusion distance in meters
            direction: 'Normal' or 'Symmetric'

        Returns:
            Dict with status and surface info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            curve_array = [profile]

            depth1 = distance
            symmetric = direction == "Symmetric"
            depth2 = distance if symmetric else 0.0
            # A second extent of igFinite with a depth of zero is rejected
            # with E_INVALIDARG, so a one-sided surface must say igNone.
            # Verified on Solid Edge 2026; the extrude and revolve calls
            # already do this. Every non-symmetric extruded surface failed.
            extent2 = ExtentTypeConstants.igFinite if symmetric else ExtentTypeConstants.igNone
            side1 = DirectionConstants.igRight
            side2 = (
                DirectionConstants.igLeft
                if direction == "Symmetric"
                else DirectionConstants.igRight
            )

            constructions = doc.Constructions
            extruded_surfaces = constructions.ExtrudedSurfaces

            extruded_surfaces.AddByCurves(
                1,  # NumberOfCurves
                curve_array,  # CurveArray
                ExtentTypeConstants.igFinite,  # ExtentType1
                side1,  # ExtentSide1
                depth1,  # FiniteDepth1
                None,  # KeyPointOrTangentFace1
                KeyPointExtentConstants.igTangentNormal,  # KeyPointFlags1
                None,  # FromFaceOrRefPlane
                OffsetSideConstants.seOffsetNone,  # FromFaceOffsetSide
                0.0,  # FromFaceOffsetDistance
                TreatmentTypeConstants.seTreatmentNone,  # TreatmentType1
                DraftSideConstants.seDraftNone,  # TreatmentDraftSide1
                0.0,  # TreatmentDraftAngle1
                TreatmentCrownTypeConstants.seTreatmentCrownByOffset,  # TreatmentCrownType1
                TreatmentCrownSideConstants.seTreatmentCrownSideInside,  # TreatmentCrownSide1
                TreatmentCrownCurvatureSideConstants.seTreatmentCrownCurvatureInside,
                0.0,  # TreatmentCrownRadiusOrOffset1
                0.0,  # TreatmentCrownTakeOffAngle1
                extent2,  # ExtentType2
                side2,  # ExtentSide2
                depth2,  # FiniteDepth2
                None,  # KeyPointOrTangentFace2
                KeyPointExtentConstants.igTangentNormal,  # KeyPointFlags2
                None,  # ToFaceOrRefPlane
                OffsetSideConstants.seOffsetNone,  # ToFaceOffsetSide
                0.0,  # ToFaceOffsetDistance
                TreatmentTypeConstants.seTreatmentNone,  # TreatmentType2
                DraftSideConstants.seDraftNone,  # TreatmentDraftSide2
                0.0,  # TreatmentDraftAngle2
                TreatmentCrownTypeConstants.seTreatmentCrownByOffset,  # TreatmentCrownType2
                TreatmentCrownSideConstants.seTreatmentCrownSideInside,  # TreatmentCrownSide2
                TreatmentCrownCurvatureSideConstants.seTreatmentCrownCurvatureInside,
                0.0,  # TreatmentCrownRadiusOrOffset2
                0.0,  # TreatmentCrownTakeOffAngle2
                True,  # WantEndCaps
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "extruded_surface_by_curves",
                "distance": distance,
                "direction": direction,
            }
        except Exception as e:
            return error_result(e)

    def create_revolved_surface_sync(
        self, angle: float = 360.0, want_end_caps: bool = False
    ) -> dict[str, Any]:
        """
        Create a synchronous revolved construction surface.

        Uses RevolvedSurfaces.AddFiniteSync(NumberOfProfiles, ProfileArray,
        RefAxis, ProfilePlaneSide, AngleOfRevolution, WantEndCaps).

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

            angle_rad = math.radians(angle)

            v_profiles = [profile]

            # Surfaces live on doc.Constructions; Model has no
            # RevolvedSurfaces property, so this always raised.
            rev_surfaces = doc.Constructions.RevolvedSurfaces
            rev_surfaces.AddFiniteSync(
                1,  # NumberOfProfiles
                v_profiles,  # ProfileArray
                refaxis,  # RefAxis
                DirectionConstants.igRight,  # ProfilePlaneSide
                angle_rad,  # AngleOfRevolution
                want_end_caps,  # WantEndCaps
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "revolved_surface_sync",
                "angle_degrees": angle,
                "want_end_caps": want_end_caps,
            }
        except Exception as e:
            return error_result(e)

    def create_revolved_surface_by_keypoint(
        self, keypoint_type: str = "End", want_end_caps: bool = False
    ) -> dict[str, Any]:
        """
        Create a revolved construction surface up to a keypoint extent.

        Uses RevolvedSurfaces.AddFiniteByKeyPoint(NumberOfProfiles, ProfileArray,
        RefAxis, KeyPointOrTangentFace, KeyPointFlags, ProfilePlaneSide, WantEndCaps).

        Args:
            keypoint_type: 'Start' or 'End' keypoint for the extent
            want_end_caps: Whether to cap the ends of the surface

        Returns:
            Dict with status and surface info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()
            refaxis = self.sketch_manager.get_active_refaxis()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            if not refaxis:
                return {"error": "No axis of revolution set. Use set_axis_of_revolution() first."}

            v_profiles = [profile]

            # Surfaces live on doc.Constructions; Model has no
            # RevolvedSurfaces property, so this always raised.
            rev_surfaces = doc.Constructions.RevolvedSurfaces
            rev_surfaces.AddFiniteByKeyPoint(
                1,  # NumberOfProfiles
                v_profiles,  # ProfileArray
                refaxis,  # RefAxis
                None,  # KeyPointOrTangentFace
                KeyPointExtentConstants.igTangentNormal,  # KeyPointFlags
                DirectionConstants.igRight,  # ProfilePlaneSide
                want_end_caps,  # WantEndCaps
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "revolved_surface_by_keypoint",
                "keypoint_type": keypoint_type,
                "want_end_caps": want_end_caps,
            }
        except Exception as e:
            return error_result(e)

    def create_lofted_surface_v2(self, want_end_caps: bool = False) -> dict[str, Any]:
        """
        Create a lofted surface using the extended Add2 method.

        Uses LoftedSurfaces.Add2 which supports an additional OutputSurfaceType
        parameter compared to the basic Add method. Requires 2+ accumulated profiles.

        Args:
            want_end_caps: Whether to cap the ends of the surface

        Returns:
            Dict with status and surface info
        """
        try:
            doc = self.doc_manager.get_active_document()

            all_profiles = self.sketch_manager.get_accumulated_profiles()

            if len(all_profiles) < 2:
                return {
                    "error": f"Lofted surface requires at least 2 profiles, "
                    f"got {len(all_profiles)}."
                }

            _CS = LoftSweepConstants.igProfileBasedCrossSection

            v_sections = all_profiles
            v_types = [_CS] * len(all_profiles)
            # A SAFEARRAY of SAFEARRAY(VT_R8): the inner VARIANTs are required, only
            # the outer wrapper is not. Dropping them broke the lofted cutout.
            v_origins = [
                VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, list(profile_origin(p)))
                for p in all_profiles
            ]

            # Surfaces live on doc.Constructions; Model has no
            # LoftedSurfaces property, so this always raised.
            loft_surfaces = doc.Constructions.LoftedSurfaces
            loft_surfaces.Add2(
                len(all_profiles),  # NumSections
                v_sections,  # CrossSections
                v_types,  # CrossSectionTypes
                v_origins,  # Origins
                ExtentTypeConstants.igNone,  # StartExtentType
                ExtentTypeConstants.igNone,  # EndExtentType
                0,  # StartTangentType
                0.0,  # StartTangentMagnitude
                0,  # EndTangentType
                0.0,  # EndTangentMagnitude
                0,  # NumGuideCurves
                None,  # GuideCurves
                want_end_caps,  # WantEndCaps
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "lofted_surface_v2",
                "num_profiles": len(all_profiles),
                "want_end_caps": want_end_caps,
            }
        except Exception as e:
            return error_result(e)

    def create_swept_surface_ex(
        self, path_profile_index: int | None = None, want_end_caps: bool = False
    ) -> dict[str, Any]:
        """
        Create a swept surface using the extended AddEx method.

        Uses SweptSurfaces.AddEx which provides additional control via Origins
        and OriginRefs parameters. Requires 2+ accumulated profiles (path + sections).

        Args:
            path_profile_index: Index of the path profile in accumulated profiles
                (default: 0, the first accumulated profile)
            want_end_caps: Whether to cap the ends of the surface

        Returns:
            Dict with status and surface info
        """
        try:
            doc = self.doc_manager.get_active_document()
            all_profiles = self.sketch_manager.get_accumulated_profiles()

            if len(all_profiles) < 2:
                return {
                    "error": "Swept surface requires at least 2 profiles "
                    "(path + cross-section), got "
                    f"{len(all_profiles)}."
                }

            path_idx = path_profile_index if path_profile_index is not None else 0
            path_profile = all_profiles[path_idx]
            cross_sections = [p for i, p in enumerate(all_profiles) if i != path_idx]

            _CS = LoftSweepConstants.igProfileBasedCrossSection

            v_paths = [path_profile]
            v_sections = cross_sections
            # A SAFEARRAY of SAFEARRAY(VT_R8): the inner VARIANTs are required, only
            # the outer wrapper is not. Dropping them broke the lofted cutout.
            v_origins = [
                VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, list(profile_origin(p)))
                for p in cross_sections
            ]

            # Surfaces live on doc.Constructions; Model has no
            # SweptSurfaces property, so this always raised.
            swept_surfaces = doc.Constructions.SweptSurfaces
            swept_surfaces.AddEx(
                1,  # NumCurves
                v_paths,  # TraceCurves
                _CS,  # TraceCurveTypes
                len(cross_sections),  # NumSections
                v_sections,  # CrossSections
                _CS,  # CrossSectionTypes
                v_origins,  # Origins
                None,  # OriginRefs
                ExtentTypeConstants.igNone,  # StartExtentType
                ExtentTypeConstants.igNone,  # EndExtentType
                want_end_caps,  # WantEndCaps
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "swept_surface_ex",
                "num_cross_sections": len(cross_sections),
                "want_end_caps": want_end_caps,
            }
        except Exception as e:
            return error_result(e)

    def create_extruded_surface_full(
        self,
        distance: float,
        direction: str = "Normal",
        treatment_type: str = "None",
        draft_angle: float = 0.0,
    ) -> dict[str, Any]:
        """
        Create an extruded surface with full treatment parameters (crown, draft).

        Uses ExtrudedSurfaces.Add with all treatment params exposed for control
        over draft angles and crown shaping.

        Args:
            distance: Extrusion distance in meters
            direction: 'Normal' or 'Symmetric'
            treatment_type: 'None', 'Crown', 'Draft', or 'CrownAndDraft'
            draft_angle: Draft angle in degrees (used when treatment_type includes 'Draft')

        Returns:
            Dict with status and surface info
        """
        try:
            import math

            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first."}

            # Map treatment type
            treatment_map = {
                "None": TreatmentTypeConstants.seTreatmentNone,
                "Crown": TreatmentTypeConstants.seTreatmentCrown,
                "Draft": TreatmentTypeConstants.seTreatmentDraft,
                "CrownAndDraft": TreatmentTypeConstants.seTreatmentCrownAndDraft,
            }
            treat_const = treatment_map.get(treatment_type, TreatmentTypeConstants.seTreatmentNone)

            # Draft side defaults to outside when draft is active
            draft_side = (
                DraftSideConstants.seDraftOutside
                if treat_const
                in (
                    TreatmentTypeConstants.seTreatmentDraft,
                    TreatmentTypeConstants.seTreatmentCrownAndDraft,
                )
                else DraftSideConstants.seDraftNone
            )
            draft_angle_rad = math.radians(draft_angle)

            profile_array = [profile]

            depth1 = distance
            symmetric = direction == "Symmetric"
            depth2 = distance if symmetric else 0.0
            # A second extent of igFinite with a depth of zero is rejected
            # with E_INVALIDARG, so a one-sided surface must say igNone.
            # Verified on Solid Edge 2026; the extrude and revolve calls
            # already do this. Every non-symmetric extruded surface failed.
            extent2 = ExtentTypeConstants.igFinite if symmetric else ExtentTypeConstants.igNone
            side1 = DirectionConstants.igRight
            side2 = (
                DirectionConstants.igLeft
                if direction == "Symmetric"
                else DirectionConstants.igRight
            )

            constructions = doc.Constructions
            extruded_surfaces = constructions.ExtrudedSurfaces

            extruded_surfaces.Add(
                1,  # NumberOfProfiles
                profile_array,  # ProfileArray
                ExtentTypeConstants.igFinite,  # ExtentType1
                side1,  # ExtentSide1
                depth1,  # FiniteDepth1
                None,  # KeyPointOrTangentFace1
                KeyPointExtentConstants.igTangentNormal,  # KeyPointFlags1
                None,  # FromFaceOrRefPlane
                OffsetSideConstants.seOffsetNone,  # FromFaceOffsetSide
                0.0,  # FromFaceOffsetDistance
                treat_const,  # TreatmentType1
                draft_side,  # TreatmentDraftSide1
                draft_angle_rad,  # TreatmentDraftAngle1
                TreatmentCrownTypeConstants.seTreatmentCrownByOffset,  # TreatmentCrownType1
                TreatmentCrownSideConstants.seTreatmentCrownSideInside,  # TreatmentCrownSide1
                TreatmentCrownCurvatureSideConstants.seTreatmentCrownCurvatureInside,
                0.0,  # TreatmentCrownRadiusOrOffset1
                0.0,  # TreatmentCrownTakeOffAngle1
                extent2,  # ExtentType2
                side2,  # ExtentSide2
                depth2,  # FiniteDepth2
                None,  # KeyPointOrTangentFace2
                KeyPointExtentConstants.igTangentNormal,  # KeyPointFlags2
                None,  # ToFaceOrRefPlane
                OffsetSideConstants.seOffsetNone,  # ToFaceOffsetSide
                0.0,  # ToFaceOffsetDistance
                TreatmentTypeConstants.seTreatmentNone,  # TreatmentType2
                DraftSideConstants.seDraftNone,  # TreatmentDraftSide2
                0.0,  # TreatmentDraftAngle2
                TreatmentCrownTypeConstants.seTreatmentCrownByOffset,  # TreatmentCrownType2
                TreatmentCrownSideConstants.seTreatmentCrownSideInside,  # TreatmentCrownSide2
                TreatmentCrownCurvatureSideConstants.seTreatmentCrownCurvatureInside,
                0.0,  # TreatmentCrownRadiusOrOffset2
                0.0,  # TreatmentCrownTakeOffAngle2
                True,  # WantEndCaps
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "extruded_surface_full",
                "distance": distance,
                "direction": direction,
                "treatment_type": treatment_type,
                "draft_angle": draft_angle,
            }
        except Exception as e:
            return error_result(e)

    def create_revolved_surface_full(
        self, angle: float = 360.0, want_end_caps: bool = False
    ) -> dict[str, Any]:
        """
        Create a revolved surface with full extent parameters.

        Uses RevolvedSurfaces.Add with dual-extent params.

        Args:
            angle: Revolution angle in degrees
            want_end_caps: Whether to add end caps

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
                return {
                    "error": "No axis of revolution set. "
                    "Use set_axis_of_revolution() before "
                    "closing the sketch."
                }

            angle_rad = math.radians(angle)
            profile_array = [profile]

            # Surfaces live on doc.Constructions; Model has no
            # RevolvedSurfaces property, so this always raised.
            surfaces = doc.Constructions.RevolvedSurfaces
            surfaces.Add(
                1,
                profile_array,
                refaxis,
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
                want_end_caps,
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "revolved_surface_full",
                "angle": angle,
                "want_end_caps": want_end_caps,
            }
        except Exception as e:
            return error_result(e)

    def create_bounded_surface(
        self,
        want_end_caps: bool = True,
        periodic: bool = False,
    ) -> dict[str, Any]:
        """
        Create a bounded surface (BlueSurf) from accumulated profiles.

        Uses BlueSurfs.Add to create a surface through multiple cross-section
        profiles, similar to a lofted surface but using the BlueSurf interface
        which provides more control over guide curves and tangent continuity.

        Requires 2+ accumulated profiles on different planes.

        Args:
            want_end_caps: Whether to cap the ends of the surface
            periodic: Whether to create a periodic (closed-loop) surface

        Returns:
            Dict with status and surface info
        """
        # Constructions.BlueSurfs.Add answers E_INVALIDARG whatever it is
        # given for Origins -- the SAFEARRAY(VT_DISPATCH) it declares, filled
        # with the sections' circles, with the profiles themselves, or with
        # None -- from two profiles that Models.AddLoftedProtrusion lofts
        # (Solid Edge 2026). Say so rather than raise it.
        all_profiles = self.sketch_manager.get_accumulated_profiles()
        return {
            "error": (
                "BlueSurfs.Add rejects every argument form tried on Solid Edge 2026 "
                "(E_INVALIDARG). Use create_loft for a solid, or the Solid Edge UI "
                "for a bounded surface."
            ),
            "unsupported": True,
            "num_profiles": len(all_profiles),
            "want_end_caps": want_end_caps,
            "periodic": periodic,
        }

    def create_revolved_surface_full_sync(
        self, angle: float = 360.0, want_end_caps: bool = False
    ) -> dict[str, Any]:
        """
        Create a synchronous revolved surface with full extent parameters.

        Uses RevolvedSurfaces.AddSync with dual-extent params.

        Args:
            angle: Revolution angle in degrees
            want_end_caps: Whether to add end caps

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
                return {
                    "error": "No axis of revolution set. "
                    "Use set_axis_of_revolution() before "
                    "closing the sketch."
                }

            angle_rad = math.radians(angle)
            profile_array = [profile]

            # Surfaces live on doc.Constructions; Model has no
            # RevolvedSurfaces property, so this always raised.
            surfaces = doc.Constructions.RevolvedSurfaces
            surfaces.AddSync(
                1,
                profile_array,
                refaxis,
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
                want_end_caps,
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "revolved_surface_full_sync",
                "angle": angle,
                "want_end_caps": want_end_caps,
            }
        except Exception as e:
            return error_result(e)
