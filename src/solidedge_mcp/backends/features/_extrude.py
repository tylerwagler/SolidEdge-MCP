"""Extrude feature operations."""

import traceback
from typing import Any

from ..constants import (
    DirectionConstants,
    FeatureOperationConstants,
)
from ..logging import get_logger

_logger = get_logger(__name__)


class ExtrudeMixin:
    """Mixin providing extrude protrusion methods."""

    def create_extrude(
        self, distance: float, operation: str = "Add", direction: str = "Normal"
    ) -> dict[str, Any]:
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
                "Intersect": FeatureOperationConstants.igFeatureIntersect,
            }
            operation_map.get(operation, FeatureOperationConstants.igFeatureAdd)

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
                "operation": operation,
                "direction": direction,
            }
        except Exception as e:
            _logger.error(f"Extrude failed: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_extrude_thin_wall(
        self, distance: float, wall_thickness: float, direction: str = "Normal"
    ) -> dict[str, Any]:
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
                "Symmetric": DirectionConstants.igSymmetric,
            }
            dir_const = direction_map.get(direction, DirectionConstants.igRight)

            # AddExtrudedProtrusionWithThinWall
            models.AddExtrudedProtrusionWithThinWall(
                NumberOfProfiles=1,
                ProfileArray=(profile,),
                ProfilePlaneSide=dir_const,
                ExtrusionDistance=distance,
                WallThickness=wall_thickness,
            )

            return {
                "status": "created",
                "type": "extrude_thin_wall",
                "distance": distance,
                "wall_thickness": wall_thickness,
                "direction": direction,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_extrude_infinite(self, direction: str = "Normal") -> dict[str, Any]:
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
                "Symmetric": DirectionConstants.igSymmetric,
            }
            dir_const = direction_map.get(direction, DirectionConstants.igRight)

            # AddExtrudedProtrusion (infinite)
            models.AddExtrudedProtrusion(
                NumberOfProfiles=1, ProfileArray=(profile,), ProfilePlaneSide=dir_const
            )

            return {"status": "created", "type": "extrude_infinite", "direction": direction}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_extrude_through_next(self, direction: str = "Normal") -> dict[str, Any]:
        """
        Create an extrusion that extends to the next face encountered.

        Uses ExtrudedProtrusions.AddThroughNext(Profile, ProfilePlaneSide) on the
        collection. Extrudes from the sketch plane until it meets the first face.

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
            protrusions.AddThroughNext(profile, dir_const)

            self.sketch_manager.clear_accumulated_profiles()

            return {"status": "created", "type": "extrude_through_next", "direction": direction}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_extrude_from_to(self, from_plane_index: int, to_plane_index: int) -> dict[str, Any]:
        """
        Create an extrusion between two reference planes.

        Uses ExtrudedProtrusions.AddFromTo(Profile, FromFaceOrRefPlane, ToFaceOrRefPlane)
        on the collection. Extrudes from one plane to another.

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
            protrusions.AddFromTo(profile, from_plane, to_plane)

            self.sketch_manager.clear_accumulated_profiles()

            return {
                "status": "created",
                "type": "extrude_from_to",
                "from_plane_index": from_plane_index,
                "to_plane_index": to_plane_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_extrude_by_keypoint(self, direction: str = "Normal") -> dict[str, Any]:
        """
        Create an extrusion up to a keypoint extent.

        Uses ExtrudedProtrusions.AddFiniteByKeyPoint(Profile, PlaneSide).
        Extrudes to the nearest keypoint on adjacent geometry.

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
            protrusions.AddFiniteByKeyPoint(profile, side)

            self.sketch_manager.clear_accumulated_profiles()

            return {"status": "created", "type": "extrude_by_keypoint", "direction": direction}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}

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
            return {"error": str(e), "traceback": traceback.format_exc()}
