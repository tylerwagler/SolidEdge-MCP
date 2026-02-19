"""Revolve feature operations."""

import traceback
from typing import Any

import pythoncom
from win32com.client import VARIANT

from ..constants import (
    DirectionConstants,
    DraftSideConstants,
    KeyPointExtentConstants,
    TreatmentCrownCurvatureSideConstants,
    TreatmentCrownSideConstants,
    TreatmentCrownTypeConstants,
    TreatmentTypeConstants,
)
from ..logging import get_logger

_logger = get_logger(__name__)


class RevolveMixin:
    """Mixin providing revolve protrusion methods."""

    def create_revolve(self, angle: float = 360, operation: str = "Add") -> dict[str, Any]:
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

            return {"status": "created", "type": "revolve", "angle": angle, "operation": operation}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_revolve_thin_wall(self, angle: float, wall_thickness: float) -> dict[str, Any]:
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
                refaxis,  # ReferenceAxis
                DirectionConstants.igRight,  # ProfilePlaneSide
                angle_rad,  # AngleofRevolution
                wall_thickness,  # WallThickness
            )

            return {
                "status": "created",
                "type": "revolve_thin_wall",
                "angle": angle,
                "wall_thickness": wall_thickness,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_revolve_sync(self, angle: float) -> dict[str, Any]:
        """Create synchronous revolve feature"""
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
                refaxis,  # ReferenceAxis
                DirectionConstants.igRight,  # ProfilePlaneSide
                angle_rad,  # AngleofRevolution
            )

            return {"status": "created", "type": "revolve_sync", "angle": angle}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_revolve_by_keypoint(self) -> dict[str, Any]:
        """
        Create a revolve up to a keypoint extent.

        Uses RevolvedProtrusions.AddFiniteByKeyPoint(Profile, RefAxis, PlaneSide).
        Revolves the profile to the nearest keypoint.

        Returns:
            Dict with status and revolve info
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

            protrusions = model.RevolvedProtrusions
            protrusions.AddFiniteByKeyPoint(profile, refaxis, DirectionConstants.igRight)

            self.sketch_manager.clear_accumulated_profiles()

            return {"status": "created", "type": "revolve_by_keypoint"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_revolve_full(
        self, angle: float = 360.0, treatment_type: str = "None"
    ) -> dict[str, Any]:
        """
        Create a revolve with full treatment parameters.

        Uses RevolvedProtrusions.Add(NumProfiles, ProfileArray, RefAxis, PlaneSide,
        AngleOfRevolution, ...). Provides access to treatment options.

        Args:
            angle: Revolution angle in degrees (360 for full revolution)
            treatment_type: 'None', 'Draft', 'Crown', or 'CrownAndDraft'

        Returns:
            Dict with status and revolve info
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

            treatment_map = {
                "None": TreatmentTypeConstants.seTreatmentNone,
                "Draft": TreatmentTypeConstants.seTreatmentDraft,
                "Crown": TreatmentTypeConstants.seTreatmentCrown,
                "CrownAndDraft": TreatmentTypeConstants.seTreatmentCrownAndDraft,
            }
            treat_const = treatment_map.get(treatment_type, TreatmentTypeConstants.seTreatmentNone)

            profile_array = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, [profile])

            protrusions = model.RevolvedProtrusions
            protrusions.Add(
                1,  # NumProfiles
                profile_array,  # ProfileArray
                refaxis,  # RefAxis
                DirectionConstants.igRight,  # PlaneSide
                angle_rad,  # AngleOfRevolution
                treat_const,  # TreatmentType
                DraftSideConstants.seDraftNone,  # TreatmentDraftSide
                0.0,  # TreatmentDraftAngle
                TreatmentCrownTypeConstants.seTreatmentCrownByOffset,
                TreatmentCrownSideConstants.seTreatmentCrownSideInside,
                TreatmentCrownCurvatureSideConstants.seTreatmentCrownCurvatureInside,
                0.0,  # TreatmentCrownRadiusOrOffset
                0.0,  # TreatmentCrownTakeOffAngle
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "revolve_full",
                "angle": angle,
                "treatment_type": treatment_type,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_revolve_by_keypoint_sync(self) -> dict[str, Any]:
        """
        Create a synchronous revolve up to a keypoint extent.

        Uses RevolvedProtrusions.AddFiniteByKeyPointSync(Profile, RefAxis,
        KeyPointOrTangentFace, KeyPointFlags, ProfileSide).

        Returns:
            Dict with status and revolve info
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

            protrusions = model.RevolvedProtrusions
            protrusions.AddFiniteByKeyPointSync(
                profile,
                refaxis,
                None,  # KeyPointOrTangentFace
                KeyPointExtentConstants.igTangentNormal,  # KeyPointFlags
                DirectionConstants.igRight,  # ProfileSide
            )

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "revolve_by_keypoint_sync",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
