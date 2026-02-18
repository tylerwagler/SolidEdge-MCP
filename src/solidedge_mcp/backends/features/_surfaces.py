"""Surface feature operations."""

import traceback
from typing import Any

import pythoncom
from win32com.client import VARIANT

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
            profile_array = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

            depth1 = distance
            depth2 = distance if direction == "Symmetric" else 0.0
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
                ExtentTypeConstants.igFinite,  # ExtentType2
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
            return {"error": str(e), "traceback": traceback.format_exc()}

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

            v_profiles = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

            # Try collection-level API first (on model), then Models-level
            if models.Count > 0:
                model = models.Item(1)
                rev_surfaces = model.RevolvedSurfaces
                rev_surfaces.AddFinite(
                    1, v_profiles, refaxis, DirectionConstants.igRight, angle_rad, want_end_caps
                )
            else:
                # First feature - use Models method if available
                models.AddFiniteRevolvedSurface(
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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            models = doc.Models

            all_profiles = self.sketch_manager.get_accumulated_profiles()

            if len(all_profiles) < 2:
                return {
                    "error": f"Lofted surface requires at least 2 profiles, "
                    f"got {len(all_profiles)}."
                }

            _CS = LoftSweepConstants.igProfileBasedCrossSection

            v_sections = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, all_profiles)
            v_types = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I4, [_CS] * len(all_profiles))
            v_origins = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
                [VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0]) for _ in all_profiles],
            )

            if models.Count > 0:
                model = models.Item(1)
                loft_surfaces = model.LoftedSurfaces
                loft_surfaces.Add(
                    len(all_profiles),
                    v_sections,
                    v_types,
                    v_origins,
                    ExtentTypeConstants.igNone,  # StartExtentType
                    ExtentTypeConstants.igNone,  # EndExtentType
                    0,
                    0.0,  # StartTangentType, StartTangentMagnitude
                    0,
                    0.0,  # EndTangentType, EndTangentMagnitude
                    0,
                    None,  # NumGuideCurves, GuideCurves
                    want_end_caps,
                )
            else:
                return {"error": "Lofted surface requires an existing base feature."}

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "lofted_surface",
                "num_profiles": len(all_profiles),
                "want_end_caps": want_end_caps,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_swept_surface(
        self, path_profile_index: int = None, want_end_caps: bool = False
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
            swept_surfaces.Add(
                1,
                v_paths,
                _CS,  # Path
                len(cross_sections),
                v_sections,
                _CS,  # Sections
                None,
                None,  # Origins, OriginRefs
                ExtentTypeConstants.igNone,  # StartExtentType
                ExtentTypeConstants.igNone,  # EndExtentType
                want_end_caps,
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "swept_surface",
                "num_cross_sections": len(cross_sections),
                "want_end_caps": want_end_caps,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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

            profile_array = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

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
            return {"error": str(e), "traceback": traceback.format_exc()}

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

            profile_array = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

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
            return {"error": str(e), "traceback": traceback.format_exc()}

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

            curve_array = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

            depth1 = distance
            depth2 = distance if direction == "Symmetric" else 0.0
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
                ExtentTypeConstants.igFinite,  # ExtentType2
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
            return {"error": str(e), "traceback": traceback.format_exc()}

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

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            angle_rad = math.radians(angle)

            v_profiles = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

            rev_surfaces = model.RevolvedSurfaces
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
            return {"error": str(e), "traceback": traceback.format_exc()}

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

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            v_profiles = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

            rev_surfaces = model.RevolvedSurfaces
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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            models = doc.Models

            all_profiles = self.sketch_manager.get_accumulated_profiles()

            if len(all_profiles) < 2:
                return {
                    "error": f"Lofted surface requires at least 2 profiles, "
                    f"got {len(all_profiles)}."
                }

            if models.Count == 0:
                return {"error": "Lofted surface requires an existing base feature."}

            model = models.Item(1)

            _CS = LoftSweepConstants.igProfileBasedCrossSection

            v_sections = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, all_profiles)
            v_types = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_I4, [_CS] * len(all_profiles))
            v_origins = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
                [VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0]) for _ in all_profiles],
            )

            loft_surfaces = model.LoftedSurfaces
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
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_swept_surface_ex(
        self, path_profile_index: int = None, want_end_caps: bool = False
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
            models = doc.Models

            if models.Count == 0:
                return {"error": "Swept surface requires an existing base feature."}

            model = models.Item(1)

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

            v_paths = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [path_profile])
            v_sections = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, cross_sections)
            v_origins = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
                [VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0]) for _ in cross_sections],
            )

            swept_surfaces = model.SweptSurfaces
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
            return {"error": str(e), "traceback": traceback.format_exc()}

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

            profile_array = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

            depth1 = distance
            depth2 = distance if direction == "Symmetric" else 0.0
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
                ExtentTypeConstants.igFinite,  # ExtentType2
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
            return {"error": str(e), "traceback": traceback.format_exc()}

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

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}
            model = models.Item(1)

            angle_rad = math.radians(angle)
            profile_array = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

            surfaces = model.RevolvedSurfaces
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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models
            if models.Count == 0:
                return {"error": "Bounded surface requires an existing base feature."}
            model = models.Item(1)

            all_profiles = self.sketch_manager.get_accumulated_profiles()
            if len(all_profiles) < 2:
                return {
                    "error": f"Bounded surface requires at least 2 profiles, "
                    f"got {len(all_profiles)}."
                }

            v_sections = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, all_profiles)
            v_origins = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH,
                [None] * len(all_profiles),
            )

            blue_surfs = model.BlueSurfs
            blue_surfs.Add(
                len(all_profiles),  # NumSections
                v_sections,  # CrossSections
                v_origins,  # Origins
                0,  # SectionStartTangentType (igNone)
                0.0,  # SectionStartTangentMagnitude
                0,  # SectionEndTangentType (igNone)
                0.0,  # SectionEndTangentMagnitude
                0,  # NumGuideCurves
                None,  # GuideCurves
                0,  # GuideStartTangentType
                0.0,  # GuideStartTangentMagnitude
                0,  # GuideEndTangentType
                0.0,  # GuideEndTangentMagnitude
                want_end_caps,  # WantEndCaps
                periodic,  # Periodic
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "bounded_surface",
                "num_profiles": len(all_profiles),
                "want_end_caps": want_end_caps,
                "periodic": periodic,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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

            models = doc.Models
            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}
            model = models.Item(1)

            angle_rad = math.radians(angle)
            profile_array = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

            surfaces = model.RevolvedSurfaces
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
            return {"error": str(e), "traceback": traceback.format_exc()}
