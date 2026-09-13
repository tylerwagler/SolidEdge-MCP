"""Revolve feature operations."""

from typing import Any

from solidedge_mcp.backends.errors import error_result

from ..constants import (
    DirectionConstants,
    ExtentTypeConstants,
    KeyPointExtentConstants,
)
from ..logging import get_logger
from ._base import verify_geometry_on_creators

_logger = get_logger(__name__)

_REVOLVE_OPERATIONS = ("Add", "Cut", "Intersect")

# constant.tlb > FeaturePropertyConstants.igInside. Used as ThicknessSide of a
# thin-wall feature (the wall grows inside the profile). Not yet exposed by
# backends/constants.py.
_IG_INSIDE = 4


@verify_geometry_on_creators
class RevolveMixin:
    """Mixin providing revolve protrusion methods."""

    def create_revolve(self, angle: float = 360, operation: str = "Add") -> dict[str, Any]:
        """
        Create a finite revolve feature from the active sketch profile.

        Requires an axis of revolution to be set in the sketch before closing.
        Use set_axis_of_revolution() in the sketch to define the axis.

        'Add' calls Models.AddFiniteRevolvedProtrusion. 'Cut' removes material
        instead, via RevolvedCutouts.AddFiniteMulti on the existing base body
        (delegates to create_revolved_cutout, so a base feature must already
        exist). 'Intersect' has no revolved COM equivalent and is rejected.

        Args:
            angle: Revolution angle in degrees (360 for full revolution)
            operation: 'Add' (default) or 'Cut'; 'Intersect' returns an error

        Returns:
            Dict with status and feature info
        """
        try:
            if operation not in _REVOLVE_OPERATIONS:
                return {
                    "error": f"Unknown operation: {operation!r}. "
                    f"Expected one of {', '.join(_REVOLVE_OPERATIONS)}."
                }
            if operation == "Intersect":
                return {
                    "error": "Intersect is not supported by create_revolve: Solid Edge "
                    "COM automation exposes no revolved-intersect feature. "
                    "Use 'Add' or 'Cut'.",
                    "unsupported": True,
                }
            if operation == "Cut":
                # create_revolved_cutout checks for a profile, an axis and an
                # existing base body, so a cut on an empty part returns a clear
                # error instead of a COM exception.
                result = self.create_revolved_cutout(angle)
                if "error" in result:
                    return result
                return {
                    **result,
                    "type": "revolve",
                    "operation": "Cut",
                    "method": "RevolvedCutouts.AddFiniteMulti",
                }

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

            import math

            angle_rad = math.radians(angle)

            # AddFiniteRevolvedProtrusion: NumProfiles,
            # ProfileArray, ReferenceAxis, ProfilePlaneSide, Angle
            # Do NOT pass None for optional params (KeyPointOrTangentFace, KeyPointFlags)
            models.AddFiniteRevolvedProtrusion(
                1,  # NumberOfProfiles
                (profile,),  # ProfileArray
                refaxis,  # ReferenceAxis
                DirectionConstants.igRight,  # ProfilePlaneSide (2)
                angle_rad,  # AngleofRevolution
            )

            # Clear accumulated profiles (consumed by this feature)
            self.sketch_manager.clear_accumulated_profiles()

            return {"status": "created", "type": "revolve", "angle": angle, "operation": "Add"}
        except Exception as e:
            return error_result(e)

    def create_revolve_finite(self, angle: float, axis_type: str = "CenterLine") -> dict[str, Any]:
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
                return {
                    "error": "No axis of revolution set. "
                    "Use set_axis_of_revolution() before "
                    "closing the sketch."
                }

            models = doc.Models

            import math

            angle_rad = math.radians(angle)

            models.AddFiniteRevolvedProtrusion(
                1,  # NumberOfProfiles
                (profile,),  # ProfileArray
                refaxis,  # ReferenceAxis
                DirectionConstants.igRight,  # ProfilePlaneSide (2)
                angle_rad,  # AngleofRevolution
            )

            return {"status": "created", "type": "revolve_finite", "angle": angle}
        except Exception as e:
            return error_result(e)

    def create_revolve_thin_wall(self, angle: float, wall_thickness: float) -> dict[str, Any]:
        """
        Create a thin-walled revolve feature.

        Type library: Models.AddRevolvedProtrusionWithThinWall(NumberOfProfiles,
        ProfileArray, RefAxis, ProfileSide, ExtentType1, ExtentSide1,
        FiniteAngle1, KeyPointOrTangentFace1, KeyPointFlags1, ExtentType2,
        ExtentSide2, FiniteAngle2, KeyPointOrTangentFace2, KeyPointFlags2,
        ThinWall, AddEndCaps, RemoveInsideMaterial, Thickness, ThicknessSide) --
        nineteen required arguments. Only the first extent is used.

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
                return {
                    "error": "No axis of revolution set. "
                    "Use set_axis_of_revolution() before "
                    "closing the sketch."
                }

            models = doc.Models

            import math

            angle_rad = math.radians(angle)

            models.AddRevolvedProtrusionWithThinWall(
                1,  # NumberOfProfiles
                (profile,),  # ProfileArray
                refaxis,  # RefAxis
                DirectionConstants.igRight,  # ProfileSide
                ExtentTypeConstants.igFinite,  # ExtentType1
                DirectionConstants.igRight,  # ExtentSide1
                angle_rad,  # FiniteAngle1
                None,  # KeyPointOrTangentFace1
                KeyPointExtentConstants.igTangentNormal,  # KeyPointFlags1
                ExtentTypeConstants.igNone,  # ExtentType2 (single-sided)
                DirectionConstants.igRight,  # ExtentSide2
                0.0,  # FiniteAngle2
                None,  # KeyPointOrTangentFace2
                KeyPointExtentConstants.igTangentNormal,  # KeyPointFlags2
                True,  # ThinWall
                False,  # AddEndCaps
                True,  # RemoveInsideMaterial
                wall_thickness,  # Thickness
                _IG_INSIDE,  # ThicknessSide
            )

            return {
                "status": "created",
                "type": "revolve_thin_wall",
                "angle": angle,
                "wall_thickness": wall_thickness,
            }
        except Exception as e:
            return error_result(e)

    def create_revolve_sync(self, angle: float) -> dict[str, Any]:
        """Create synchronous revolve feature.

        Type library: Models.AddRevolvedProtrusionSync(NumberOfProfiles,
        ProfileArray, RefAxis, ProfileSide, ExtentType1, ExtentSide1,
        FiniteAngle1, KeyPointOrTangentFace1, KeyPointFlags1, ExtentType2,
        ExtentSide2, FiniteAngle2, KeyPointOrTangentFace2, KeyPointFlags2) --
        fourteen required arguments. Only the first extent is used.
        """
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()
            refaxis = self.sketch_manager.get_active_refaxis()

            if not profile:
                return {"error": "No active sketch profile"}
            if not refaxis:
                return {
                    "error": "No axis of revolution set. "
                    "Use set_axis_of_revolution() before "
                    "closing the sketch."
                }

            models = doc.Models

            import math

            angle_rad = math.radians(angle)

            models.AddRevolvedProtrusionSync(
                1,  # NumberOfProfiles
                (profile,),  # ProfileArray
                refaxis,  # RefAxis
                DirectionConstants.igRight,  # ProfileSide
                ExtentTypeConstants.igFinite,  # ExtentType1
                DirectionConstants.igRight,  # ExtentSide1
                angle_rad,  # FiniteAngle1
                None,  # KeyPointOrTangentFace1
                KeyPointExtentConstants.igTangentNormal,  # KeyPointFlags1
                ExtentTypeConstants.igNone,  # ExtentType2 (single-sided)
                DirectionConstants.igRight,  # ExtentSide2
                0.0,  # FiniteAngle2
                None,  # KeyPointOrTangentFace2
                KeyPointExtentConstants.igTangentNormal,  # KeyPointFlags2
            )

            return {"status": "created", "type": "revolve_sync", "angle": angle}
        except Exception as e:
            return error_result(e)

    def create_revolve_finite_sync(self, angle: float) -> dict[str, Any]:
        """Create finite synchronous revolve feature"""
        try:
            doc = self.doc_manager.get_active_document()
            profile = self.sketch_manager.get_active_sketch()
            refaxis = self.sketch_manager.get_active_refaxis()

            if not profile:
                return {"error": "No active sketch profile"}
            if not refaxis:
                return {
                    "error": "No axis of revolution set. "
                    "Use set_axis_of_revolution() before "
                    "closing the sketch."
                }

            models = doc.Models

            import math

            angle_rad = math.radians(angle)

            models.AddFiniteRevolvedProtrusionSync(
                1,  # NumberOfProfiles
                (profile,),  # ProfileArray
                refaxis,  # ReferenceAxis
                DirectionConstants.igRight,  # ProfilePlaneSide
                angle_rad,  # AngleofRevolution
            )

            return {"status": "created", "type": "revolve_finite_sync", "angle": angle}
        except Exception as e:
            return error_result(e)

    def create_revolve_by_keypoint(self) -> dict[str, Any]:
        """
        Create a revolve up to a keypoint extent.

        Type library: RevolvedProtrusions.AddFiniteByKeyPoint(Profile, RefAxis,
        KeyPointOrTangentFace, KeyPointFlags, ProfileSide, [ProfilePlaneSide]).
        The KeyPoint (or tangent face) is a selected model object that this
        server has no way to pick, so the call can never be formed; report that
        instead of failing inside COM.

        Returns:
            Dict with an explanatory error
        """
        return {
            "error": (
                "Revolving to a keypoint needs a KeyPoint or tangent face object, "
                "which this server cannot select. Use create_revolve(angle) or the "
                "Solid Edge UI."
            ),
            "unsupported": True,
        }

    def create_revolve_full(
        self, angle: float = 360.0, treatment_type: str = "None"
    ) -> dict[str, Any]:
        """
        Create a revolve through the full collection-level API.

        Type library: RevolvedProtrusions.Add(NumberOfProfiles, ProfileArray,
        RefAxis, ProfileSide, ExtentType1, ExtentSide1, FiniteAngle1,
        KeyPointOrTangentFace1, KeyPointFlags1, ExtentType2, ExtentSide2,
        FiniteAngle2, KeyPointOrTangentFace2, KeyPointFlags2) -- fourteen
        required arguments, all of them extents. Unlike the extruded
        protrusion API this overload has no treatment (draft/crown)
        parameters, so treatment_type must be 'None'.

        Args:
            angle: Revolution angle in degrees (360 for full revolution)
            treatment_type: must be 'None' -- see above

        Returns:
            Dict with status and revolve info
        """
        try:
            import math

            if treatment_type != "None":
                return {
                    "error": (
                        f"treatment_type={treatment_type!r} is not supported: "
                        "RevolvedProtrusions.Add takes only extent parameters, with no "
                        "draft or crown treatment slots. Add the treatment as a "
                        "separate draft feature, or use the Solid Edge UI."
                    ),
                    "unsupported": True,
                }

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

            protrusions = model.RevolvedProtrusions
            protrusions.Add(
                1,  # NumberOfProfiles
                (profile,),  # ProfileArray
                refaxis,  # RefAxis
                DirectionConstants.igRight,  # ProfileSide
                ExtentTypeConstants.igFinite,  # ExtentType1
                DirectionConstants.igRight,  # ExtentSide1
                angle_rad,  # FiniteAngle1
                None,  # KeyPointOrTangentFace1
                KeyPointExtentConstants.igTangentNormal,  # KeyPointFlags1
                ExtentTypeConstants.igNone,  # ExtentType2 (single-sided)
                DirectionConstants.igRight,  # ExtentSide2
                0.0,  # FiniteAngle2
                None,  # KeyPointOrTangentFace2
                KeyPointExtentConstants.igTangentNormal,  # KeyPointFlags2
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "revolve_full",
                "angle": angle,
                "treatment_type": treatment_type,
            }
        except Exception as e:
            return error_result(e)

    def create_revolve_by_keypoint_sync(self) -> dict[str, Any]:
        """
        Create a synchronous revolve up to a keypoint extent.

        Type library: RevolvedProtrusions.AddFiniteByKeyPointSync(Profile,
        RefAxis, KeyPointOrTangentFace, KeyPointFlags, ProfileSide,
        [ProfilePlaneSide]). KeyPointOrTangentFace is a required object that
        this server cannot select -- passing None only moves the failure into
        COM -- so report it here, exactly as create_revolve_by_keypoint does.

        Returns:
            Dict with an explanatory error
        """
        return {
            "error": (
                "Revolving to a keypoint needs a KeyPoint or tangent face object, "
                "which this server cannot select. Use create_revolve_sync(angle) or "
                "the Solid Edge UI."
            ),
            "unsupported": True,
        }
