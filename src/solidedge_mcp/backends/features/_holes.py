"""Hole feature operations."""

import traceback
from typing import Any

from ..constants import (
    DirectionConstants,
    FaceQueryConstants,
)


class HolesMixin:
    """Mixin providing hole creation methods."""

    def create_hole(
        self,
        x: float,
        y: float,
        diameter: float,
        depth: float,
        hole_type: str = "Simple",
        plane_index: int = 1,
        direction: str = "Normal",
    ) -> dict[str, Any]:
        """
        Create a hole feature (circular cutout).

        Creates a circular cutout at (x, y) on a reference plane using
        ExtrudedCutouts.AddFiniteMulti for reliable geometry creation.

        Args:
            x, y: Hole center coordinates on the sketch plane (meters)
            diameter: Hole diameter in meters
            depth: Hole depth in meters
            hole_type: 'Simple' (only type currently supported)
            plane_index: Reference plane index (1=Top/XY, 2=Right/YZ, 3=Front/XZ)
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
            dir_const = DirectionConstants.igRight  # Normal
            if direction == "Reverse":
                dir_const = DirectionConstants.igLeft

            # Create a circular profile on the specified plane
            ps = doc.ProfileSets.Add()
            plane = doc.RefPlanes.Item(plane_index)
            profile = ps.Profiles.Add(plane)
            profile.Circles2d.AddByCenterRadius(x, y, radius)
            profile.End(0)

            # Use ExtrudedCutouts.AddFiniteMulti for reliable hole creation
            cutouts = model.ExtrudedCutouts
            cutouts.AddFiniteMulti(1, (profile,), dir_const, depth)

            return {
                "status": "created",
                "type": "hole",
                "position": [x, y],
                "diameter": diameter,
                "depth": depth,
                "hole_type": hole_type,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_hole_through_all(
        self, x: float, y: float, diameter: float, plane_index: int = 1, direction: str = "Normal"
    ) -> dict[str, Any]:
        """
        Create a hole that goes through the entire part.

        Creates a circular profile and uses ExtrudedCutouts.AddThroughAllMulti.
        Type library: ExtrudedCutouts.AddThroughAllMulti(NumProfiles, ProfileArray, PlaneSide).

        Args:
            x, y: Hole center coordinates on the sketch plane (meters)
            diameter: Hole diameter in meters
            plane_index: Reference plane index (1=Top/XY, 2=Right/YZ, 3=Front/XZ)
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

            dir_const = DirectionConstants.igRight  # Normal
            if direction == "Reverse":
                dir_const = DirectionConstants.igLeft

            # Create a circular profile on the specified plane
            ps = doc.ProfileSets.Add()
            plane = doc.RefPlanes.Item(plane_index)
            profile = ps.Profiles.Add(plane)
            profile.Circles2d.AddByCenterRadius(x, y, radius)
            profile.End(0)

            # Use through-all cutout
            cutouts = model.ExtrudedCutouts
            cutouts.AddThroughAllMulti(1, (profile,), dir_const)

            return {
                "status": "created",
                "type": "hole_through_all",
                "position": [x, y],
                "diameter": diameter,
                "plane_index": plane_index,
                "direction": direction,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def delete_hole_by_face(self, face_index: int) -> dict[str, Any]:
        """
        Delete a specific hole by selecting its face.

        Unlike create_delete_hole which deletes holes by type/size criteria,
        this targets a specific hole identified by its face index.

        Args:
            face_index: 0-based face index of the hole to delete

        Returns:
            Dict with status and deletion info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)

            delete_holes = model.DeleteHoles
            delete_holes.AddByFace(face)

            return {"status": "created", "type": "delete_hole_by_face", "face_index": face_index}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_hole_from_to(
        self,
        x: float,
        y: float,
        diameter: float,
        from_plane_index: int,
        to_plane_index: int,
    ) -> dict[str, Any]:
        """
        Create a hole between two reference planes.

        Uses ExtrudedCutouts.AddFromToMulti with a circular profile as a
        workaround for Holes.AddFinite not cutting geometry.

        Args:
            x, y: Hole center coordinates on the sketch plane (meters)
            diameter: Hole diameter in meters
            from_plane_index: 1-based index of the starting reference plane
            to_plane_index: 1-based index of the ending reference plane

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

            # Create a circular profile on the from_plane
            ps = doc.ProfileSets.Add()
            profile = ps.Profiles.Add(from_plane)
            profile.Circles2d.AddByCenterRadius(x, y, radius)
            profile.End(0)

            cutouts = model.ExtrudedCutouts
            cutouts.AddFromToMulti(1, (profile,), from_plane, to_plane)

            return {
                "status": "created",
                "type": "hole_from_to",
                "position": [x, y],
                "diameter": diameter,
                "from_plane_index": from_plane_index,
                "to_plane_index": to_plane_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_hole_through_next(
        self,
        x: float,
        y: float,
        diameter: float,
        direction: str = "Normal",
        plane_index: int = 1,
    ) -> dict[str, Any]:
        """
        Create a hole through the next face.

        Uses ExtrudedCutouts.AddThroughNextMulti with a circular profile.

        Args:
            x, y: Hole center coordinates on the sketch plane (meters)
            diameter: Hole diameter in meters
            direction: 'Normal' or 'Reverse'
            plane_index: Reference plane index (1=Top/XY, 2=Right/YZ, 3=Front/XZ)

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

            dir_const = DirectionConstants.igRight
            if direction == "Reverse":
                dir_const = DirectionConstants.igLeft

            ps = doc.ProfileSets.Add()
            plane = doc.RefPlanes.Item(plane_index)
            profile = ps.Profiles.Add(plane)
            profile.Circles2d.AddByCenterRadius(x, y, radius)
            profile.End(0)

            cutouts = model.ExtrudedCutouts
            cutouts.AddThroughNextMulti(1, (profile,), dir_const)

            return {
                "status": "created",
                "type": "hole_through_next",
                "position": [x, y],
                "diameter": diameter,
                "direction": direction,
                "plane_index": plane_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_hole_sync(
        self,
        x: float,
        y: float,
        diameter: float,
        depth: float,
        plane_index: int = 1,
    ) -> dict[str, Any]:
        """
        Create a synchronous hole feature.

        Uses Holes.AddSync(Profile, PlaneSide, Depth, HoleData).
        Note: Holes API may not cut geometry - if so, use circular cutout instead.

        Args:
            x, y: Hole center coordinates on the sketch plane (meters)
            diameter: Hole diameter in meters
            depth: Hole depth in meters
            plane_index: Reference plane index (1=Top/XY, 2=Right/YZ, 3=Front/XZ)

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

            ps = doc.ProfileSets.Add()
            plane = doc.RefPlanes.Item(plane_index)
            profile = ps.Profiles.Add(plane)
            profile.Circles2d.AddByCenterRadius(x, y, radius)
            profile.End(0)

            holes = model.Holes
            holes.AddSync(profile, DirectionConstants.igRight, depth, None)

            return {
                "status": "created",
                "type": "hole_sync",
                "position": [x, y],
                "diameter": diameter,
                "depth": depth,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_hole_finite_ex(
        self,
        x: float,
        y: float,
        diameter: float,
        depth: float,
        direction: str = "Normal",
        plane_index: int = 1,
    ) -> dict[str, Any]:
        """
        Create a finite hole using the extended API.

        Uses Holes.AddFiniteEx(Profile, PlaneSide, Depth, HoleData).

        Args:
            x, y: Hole center coordinates on the sketch plane (meters)
            diameter: Hole diameter in meters
            depth: Hole depth in meters
            direction: 'Normal' or 'Reverse'
            plane_index: Reference plane index (1=Top/XY, 2=Right/YZ, 3=Front/XZ)

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

            dir_const = DirectionConstants.igRight
            if direction == "Reverse":
                dir_const = DirectionConstants.igLeft

            ps = doc.ProfileSets.Add()
            plane = doc.RefPlanes.Item(plane_index)
            profile = ps.Profiles.Add(plane)
            profile.Circles2d.AddByCenterRadius(x, y, radius)
            profile.End(0)

            holes = model.Holes
            holes.AddFiniteEx(profile, dir_const, depth, None)

            return {
                "status": "created",
                "type": "hole_finite_ex",
                "position": [x, y],
                "diameter": diameter,
                "depth": depth,
                "direction": direction,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_hole_from_to_ex(
        self,
        x: float,
        y: float,
        diameter: float,
        from_plane_index: int,
        to_plane_index: int,
    ) -> dict[str, Any]:
        """
        Create a hole between two planes using the extended API.

        Uses Holes.AddFromToEx(Profile, FromFace, ToFace, HoleData).

        Args:
            x, y: Hole center coordinates on the sketch plane (meters)
            diameter: Hole diameter in meters
            from_plane_index: 1-based index of the starting reference plane
            to_plane_index: 1-based index of the ending reference plane

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

            ps = doc.ProfileSets.Add()
            profile = ps.Profiles.Add(from_plane)
            profile.Circles2d.AddByCenterRadius(x, y, radius)
            profile.End(0)

            holes = model.Holes
            holes.AddFromToEx(profile, from_plane, to_plane, None)

            return {
                "status": "created",
                "type": "hole_from_to_ex",
                "position": [x, y],
                "diameter": diameter,
                "from_plane_index": from_plane_index,
                "to_plane_index": to_plane_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_hole_through_next_ex(
        self,
        x: float,
        y: float,
        diameter: float,
        direction: str = "Normal",
        plane_index: int = 1,
    ) -> dict[str, Any]:
        """
        Create a hole through the next face using the extended API.

        Uses Holes.AddThroughNextEx(Profile, PlaneSide, HoleData).

        Args:
            x, y: Hole center coordinates on the sketch plane (meters)
            diameter: Hole diameter in meters
            direction: 'Normal' or 'Reverse'
            plane_index: Reference plane index (1=Top/XY, 2=Right/YZ, 3=Front/XZ)

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

            dir_const = DirectionConstants.igRight
            if direction == "Reverse":
                dir_const = DirectionConstants.igLeft

            ps = doc.ProfileSets.Add()
            plane = doc.RefPlanes.Item(plane_index)
            profile = ps.Profiles.Add(plane)
            profile.Circles2d.AddByCenterRadius(x, y, radius)
            profile.End(0)

            holes = model.Holes
            holes.AddThroughNextEx(profile, dir_const, None)

            return {
                "status": "created",
                "type": "hole_through_next_ex",
                "position": [x, y],
                "diameter": diameter,
                "direction": direction,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_hole_through_all_ex(
        self,
        x: float,
        y: float,
        diameter: float,
        plane_index: int = 1,
    ) -> dict[str, Any]:
        """
        Create a hole through all material using the extended API.

        Uses Holes.AddThroughAllEx(Profile, PlaneSide, HoleData).

        Args:
            x, y: Hole center coordinates on the sketch plane (meters)
            diameter: Hole diameter in meters
            plane_index: Reference plane index (1=Top/XY, 2=Right/YZ, 3=Front/XZ)

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

            ps = doc.ProfileSets.Add()
            plane = doc.RefPlanes.Item(plane_index)
            profile = ps.Profiles.Add(plane)
            profile.Circles2d.AddByCenterRadius(x, y, radius)
            profile.End(0)

            holes = model.Holes
            holes.AddThroughAllEx(profile, DirectionConstants.igRight, None)

            return {
                "status": "created",
                "type": "hole_through_all_ex",
                "position": [x, y],
                "diameter": diameter,
                "plane_index": plane_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_hole_sync_ex(
        self,
        x: float,
        y: float,
        diameter: float,
        depth: float,
        plane_index: int = 1,
    ) -> dict[str, Any]:
        """
        Create a synchronous hole using the extended API.

        Uses Holes.AddSyncEx(Profile, PlaneSide, Depth, HoleData).

        Args:
            x, y: Hole center coordinates on the sketch plane (meters)
            diameter: Hole diameter in meters
            depth: Hole depth in meters
            plane_index: Reference plane index (1=Top/XY, 2=Right/YZ, 3=Front/XZ)

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

            ps = doc.ProfileSets.Add()
            plane = doc.RefPlanes.Item(plane_index)
            profile = ps.Profiles.Add(plane)
            profile.Circles2d.AddByCenterRadius(x, y, radius)
            profile.End(0)

            holes = model.Holes
            holes.AddSyncEx(profile, DirectionConstants.igRight, depth, None)

            return {
                "status": "created",
                "type": "hole_sync_ex",
                "position": [x, y],
                "diameter": diameter,
                "depth": depth,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_hole_multi_body(
        self,
        x: float,
        y: float,
        diameter: float,
        depth: float,
        plane_index: int = 1,
        direction: str = "Normal",
    ) -> dict[str, Any]:
        """
        Create a hole that spans multiple bodies.

        Uses Holes.AddMultiBody(Profile, PlaneSide, Depth, HoleData).

        Args:
            x, y: Hole center coordinates on the sketch plane (meters)
            diameter: Hole diameter in meters
            depth: Hole depth in meters
            plane_index: Reference plane index (1=Top/XY, 2=Right/YZ, 3=Front/XZ)
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

            dir_const = DirectionConstants.igRight
            if direction == "Reverse":
                dir_const = DirectionConstants.igLeft

            ps = doc.ProfileSets.Add()
            plane = doc.RefPlanes.Item(plane_index)
            profile = ps.Profiles.Add(plane)
            profile.Circles2d.AddByCenterRadius(x, y, radius)
            profile.End(0)

            holes = model.Holes
            holes.AddMultiBody(profile, dir_const, depth, None)

            return {
                "status": "created",
                "type": "hole_multi_body",
                "position": [x, y],
                "diameter": diameter,
                "depth": depth,
                "direction": direction,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_hole_sync_multi_body(
        self,
        x: float,
        y: float,
        diameter: float,
        depth: float,
        plane_index: int = 1,
    ) -> dict[str, Any]:
        """
        Create a synchronous hole that spans multiple bodies.

        Uses Holes.AddSyncMultiBody(Profile, PlaneSide, Depth, HoleData).

        Args:
            x, y: Hole center coordinates on the sketch plane (meters)
            diameter: Hole diameter in meters
            depth: Hole depth in meters
            plane_index: Reference plane index (1=Top/XY, 2=Right/YZ, 3=Front/XZ)

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

            ps = doc.ProfileSets.Add()
            plane = doc.RefPlanes.Item(plane_index)
            profile = ps.Profiles.Add(plane)
            profile.Circles2d.AddByCenterRadius(x, y, radius)
            profile.End(0)

            holes = model.Holes
            holes.AddSyncMultiBody(profile, DirectionConstants.igRight, depth, None)

            return {
                "status": "created",
                "type": "hole_sync_multi_body",
                "position": [x, y],
                "diameter": diameter,
                "depth": depth,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
