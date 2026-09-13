"""Extrude feature operations."""

from typing import Any

from solidedge_mcp.backends.errors import error_result

from ..constants import (
    DirectionConstants,
    DraftSideConstants,
    ExtentTypeConstants,
    KeyPointExtentConstants,
    OffsetSideConstants,
    TreatmentCrownCurvatureSideConstants,
    TreatmentCrownSideConstants,
    TreatmentCrownTypeConstants,
    TreatmentTypeConstants,
)
from ..logging import get_logger
from ._base import verify_geometry_on_creators

_logger = get_logger(__name__)

_EXTRUDE_OPERATIONS = ("Add", "Cut", "Intersect")

# constant.tlb > FeaturePropertyConstants.igInside. Used as ThicknessSide of a
# thin-wall feature (the wall grows inside the profile). Not yet exposed by
# backends/constants.py.
_IG_INSIDE = 4


@verify_geometry_on_creators
class ExtrudeMixin:
    """Mixin providing extrude protrusion methods."""

    def create_extrude(
        self, distance: float, operation: str = "Add", direction: str = "Normal"
    ) -> dict[str, Any]:
        """
        Create a finite extrusion feature from the active sketch profile.

        'Add' calls Models.AddFiniteExtrudedProtrusion. 'Cut' removes material
        instead, via ExtrudedCutouts.AddFiniteMulti on the existing base body
        (delegates to create_extruded_cutout, so a base feature must already
        exist). 'Intersect' has no finite-extrude COM equivalent and is
        rejected.

        Args:
            distance: Extrusion distance in meters
            operation: 'Add' (default) or 'Cut'; 'Intersect' returns an error
            direction: 'Normal', 'Reverse', or 'Symmetric' ('Symmetric' is
                only valid for 'Add')

        Returns:
            Dict with status and feature info
        """
        try:
            if operation not in _EXTRUDE_OPERATIONS:
                return {
                    "error": f"Unknown operation: {operation!r}. "
                    f"Expected one of {', '.join(_EXTRUDE_OPERATIONS)}."
                }
            if operation == "Intersect":
                return {
                    "error": "Intersect is not supported by create_extrude: Solid Edge "
                    "COM automation exposes no finite extruded-intersect feature. "
                    "Use 'Add' or 'Cut'.",
                    "unsupported": True,
                }
            if operation == "Cut":
                if direction == "Symmetric":
                    return {
                        "error": "direction='Symmetric' is not supported for a Cut "
                        "extrusion. Use 'Normal' or 'Reverse'."
                    }
                # create_extruded_cutout already checks for a profile and an
                # existing base body, so a cut on an empty part returns a clear
                # error instead of a COM exception.
                result = self.create_extruded_cutout(distance, direction)
                if "error" in result:
                    return result
                _logger.info(f"Created extrusion cut: distance={distance}m, direction={direction}")
                return {
                    **result,
                    "type": "extrude",
                    "operation": "Cut",
                    "method": "ExtrudedCutouts.AddFiniteMulti",
                }

            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first"}

            # Get the models collection
            models = doc.Models

            # Map direction string to constant
            direction_map = {
                "Normal": DirectionConstants.igRight,
                "Reverse": DirectionConstants.igLeft,
                "Symmetric": DirectionConstants.igSymmetric,
            }
            dir_const = direction_map.get(direction, DirectionConstants.igRight)

            # AddFiniteExtrudedProtrusion: NumProfiles, ProfileArray, ProfilePlaneSide, Distance
            models.AddFiniteExtrudedProtrusion(1, (profile,), dir_const, distance)

            # Clear accumulated profiles (consumed by this feature)
            self.sketch_manager.clear_accumulated_profiles()

            _logger.info(f"Created extrusion: distance={distance}m, direction={direction}")
            return {
                "status": "created",
                "type": "extrude",
                "distance": distance,
                "operation": "Add",
                "direction": direction,
            }
        except Exception as e:
            _logger.error(f"Extrude failed: {e}")
            return error_result(e)

    def create_extrude_symmetric(self, distance: float) -> dict[str, Any]:
        """
        Create a symmetric extrusion (extends equally in both directions).

        Args:
            distance: Total extrusion distance in meters (half on each side)

        Returns:
            Dict with status and feature info
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()

            if not profile:
                return {"error": "No active sketch profile. Create and close a sketch first"}

            models = doc.Models

            models.AddFiniteExtrudedProtrusion(
                1, (profile,), DirectionConstants.igSymmetric, distance
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "extrude_symmetric",
                "distance": distance,
                "direction": "Symmetric",
            }
        except Exception as e:
            return error_result(e)

    def create_extrude_thin_wall(
        self, distance: float, wall_thickness: float, direction: str = "Normal"
    ) -> dict[str, Any]:
        """
        Create a thin-walled extrusion.

        Type library: Models.AddExtrudedProtrusionWithThinWall takes forty
        required arguments -- NumberOfProfiles, ProfileArray, ProfileSide, then
        the first extent (ExtentType1, ExtentSide1, FiniteDepth1,
        KeyPointOrTangentFace1, KeyPointFlags1, FromFaceOrRefPlane,
        FromFaceOffsetSide, FromFaceOffsetDistance) and its treatment block
        (TreatmentType1, TreatmentDraftSide1, TreatmentDraftAngle1,
        TreatmentCrownType1, TreatmentCrownSide1, TreatmentCrownCurvatureSide1,
        TreatmentCrownRadiusOrOffset1, TreatmentCrownTakeOffAngle1), the same
        two blocks again for the second extent (ending in ToFaceOrRefPlane,
        ToFaceOffsetSide, ToFaceOffsetDistance), and finally ThinWall,
        AddEndCaps, RemoveInsideMaterial, Thickness, ThicknessSide.

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
                "Symmetric": DirectionConstants.igSymmetric,
            }
            dir_const = direction_map.get(direction, DirectionConstants.igRight)

            models.AddExtrudedProtrusionWithThinWall(
                1,  # NumberOfProfiles
                (profile,),  # ProfileArray
                dir_const,  # ProfileSide
                ExtentTypeConstants.igFinite,  # ExtentType1
                dir_const,  # ExtentSide1
                distance,  # FiniteDepth1
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
                ExtentTypeConstants.igNone,  # ExtentType2 (single-sided)
                dir_const,  # ExtentSide2
                0.0,  # FiniteDepth2
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
                True,  # ThinWall
                False,  # AddEndCaps
                True,  # RemoveInsideMaterial
                wall_thickness,  # Thickness
                _IG_INSIDE,  # ThicknessSide
            )

            return {
                "status": "created",
                "type": "extrude_thin_wall",
                "distance": distance,
                "wall_thickness": wall_thickness,
                "direction": direction,
            }
        except Exception as e:
            return error_result(e)

    def create_extrude_infinite(self, direction: str = "Normal") -> dict[str, Any]:
        """
        Create an infinite extrusion (extends through all).

        Type library: Models.AddExtrudedProtrusion takes thirty-five required
        arguments -- NumberOfProfiles, ProfileArray, ProfileSide, then two
        extent blocks (ExtentType, ExtentSide, FiniteDepth,
        KeyPointOrTangentFace, KeyPointFlags, From/ToFaceOrRefPlane, offset
        side and distance) each followed by its eight treatment parameters.
        The first extent is igThroughAll, which is what makes this the
        "infinite" variant; the second extent is igNone.

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
                "Symmetric": DirectionConstants.igSymmetric,
            }
            dir_const = direction_map.get(direction, DirectionConstants.igRight)

            models.AddExtrudedProtrusion(
                1,  # NumberOfProfiles
                (profile,),  # ProfileArray
                dir_const,  # ProfileSide
                ExtentTypeConstants.igThroughAll,  # ExtentType1
                dir_const,  # ExtentSide1
                0.0,  # FiniteDepth1 (unused for through-all)
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
                ExtentTypeConstants.igNone,  # ExtentType2
                dir_const,  # ExtentSide2
                0.0,  # FiniteDepth2
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
            )

            return {"status": "created", "type": "extrude_infinite", "direction": direction}
        except Exception as e:
            return error_result(e)

    def create_extrude_through_next(self, direction: str = "Normal") -> dict[str, Any]:
        """
        Create an extrusion that extends to the next face encountered.

        Type library: ExtrudedProtrusions.AddThroughNext(Profile, ProfileSide,
        ProfilePlaneSide) -- three required arguments. Extrudes from the sketch
        plane until it meets the first face.

        Args:
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with status and extrusion info
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

            protrusions = model.ExtrudedProtrusions
            protrusions.AddThroughNext(
                profile,  # Profile
                dir_const,  # ProfileSide
                dir_const,  # ProfilePlaneSide
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {"status": "created", "type": "extrude_through_next", "direction": direction}
        except Exception as e:
            return error_result(e)

    def create_extrude_from_to(self, from_plane_index: int, to_plane_index: int) -> dict[str, Any]:
        """
        Create an extrusion between two reference planes.

        Type library: ExtrudedProtrusions.AddFromTo(Profile, ProfileSide,
        FromFaceOrRefPlane, ToFaceOrRefPlane) -- four required arguments.
        Extrudes from one plane to another.

        Args:
            from_plane_index: 1-based index of the starting reference plane
            to_plane_index: 1-based index of the ending reference plane

        Returns:
            Dict with status and extrusion info
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

            protrusions = model.ExtrudedProtrusions
            protrusions.AddFromTo(
                profile,  # Profile
                DirectionConstants.igRight,  # ProfileSide
                from_plane,  # FromFaceOrRefPlane
                to_plane,  # ToFaceOrRefPlane
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "extrude_from_to",
                "from_plane_index": from_plane_index,
                "to_plane_index": to_plane_index,
            }
        except Exception as e:
            return error_result(e)

    def create_extrude_through_next_v2(self, direction: str = "Normal") -> dict[str, Any]:
        """
        Create an extrusion through the next face (collection-level multi-profile API).

        Uses ExtrudedProtrusions.AddThroughNextMulti(NumProfiles, ProfileArray, PlaneSide)
        instead of AddThroughNext for multi-profile support.

        Args:
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with status and extrusion info
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

            protrusions = model.ExtrudedProtrusions
            protrusions.AddThroughNextMulti(1, (profile,), side)

            self.sketch_manager.clear_accumulated_profiles()

            return {"status": "created", "type": "extrude_through_next_v2", "direction": direction}
        except Exception as e:
            return error_result(e)

    def create_extrude_from_to_v2(
        self, from_plane_index: int, to_plane_index: int
    ) -> dict[str, Any]:
        """
        Create an extrusion between two reference planes (collection multi-profile API).

        Uses ExtrudedProtrusions.AddFromToMulti(NumProfiles, ProfileArray,
        FromFaceOrRefPlane, ToFaceOrRefPlane).

        Args:
            from_plane_index: 1-based index of the starting reference plane
            to_plane_index: 1-based index of the ending reference plane

        Returns:
            Dict with status and extrusion info
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

            protrusions = model.ExtrudedProtrusions
            protrusions.AddFromToMulti(1, (profile,), from_plane, to_plane)

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "extrude_from_to_v2",
                "from_plane_index": from_plane_index,
                "to_plane_index": to_plane_index,
            }
        except Exception as e:
            return error_result(e)

    def create_extrude_by_keypoint(self, direction: str = "Normal") -> dict[str, Any]:
        """
        Create an extrusion up to a keypoint extent.

        Type library: ExtrudedProtrusions.AddFiniteByKeyPoint(Profile,
        ProfileSide, ProfilePlaneSide, KeyPointOrTangentFace, KeyPointFlags) --
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
                "Extruding to a keypoint needs a KeyPoint or tangent face object, "
                "which this server cannot select. Use create_extrude(distance), "
                "create_extrude_from_to(...) or the Solid Edge UI."
            ),
            "unsupported": True,
            "direction": direction,
        }

    def create_extrude_from_to_single(
        self, from_plane_index: int, to_plane_index: int
    ) -> dict[str, Any]:
        """
        Create a single-profile extrusion between two reference planes.

        Uses ExtrudedProtrusions.AddFromTo(Profile, ProfileSide,
        FromFaceOrRefPlane, ToFaceOrRefPlane).

        Args:
            from_plane_index: 1-based index of the starting reference plane
            to_plane_index: 1-based index of the ending reference plane

        Returns:
            Dict with status and extrusion info
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

            protrusions = model.ExtrudedProtrusions
            protrusions.AddFromTo(
                profile,
                DirectionConstants.igRight,  # ProfileSide
                from_plane,
                to_plane,
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "extrude_from_to_single",
                "from_plane_index": from_plane_index,
                "to_plane_index": to_plane_index,
            }
        except Exception as e:
            return error_result(e)

    def create_extrude_through_next_single(self, direction: str = "Normal") -> dict[str, Any]:
        """
        Create a single-profile extrusion through the next face.

        Uses ExtrudedProtrusions.AddThroughNext(Profile, ProfileSide,
        ProfilePlaneSide).

        Args:
            direction: 'Normal' or 'Reverse'

        Returns:
            Dict with status and extrusion info
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

            protrusions = model.ExtrudedProtrusions
            protrusions.AddThroughNext(
                profile,
                dir_const,  # ProfileSide
                dir_const,  # ProfilePlaneSide
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "extrude_through_next_single",
                "direction": direction,
            }
        except Exception as e:
            return error_result(e)
