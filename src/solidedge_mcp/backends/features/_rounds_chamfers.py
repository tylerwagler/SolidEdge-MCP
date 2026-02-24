"""Round, chamfer, and blend feature operations."""

import traceback
from typing import Any

import pythoncom
from win32com.client import VARIANT

from ..constants import (
    FaceQueryConstants,
)
from ..logging import get_logger

_logger = get_logger(__name__)


class RoundsChamfersMixin:
    """Mixin providing round, chamfer, and blend methods."""

    def create_round(self, radius: float) -> dict[str, Any]:
        """
        Create a round (fillet) feature on all body edges.

        Rounds all edges of the body using model.Rounds.Add(). All edges
        are grouped as one edge set with a single radius value.

        Args:
            radius: Round radius in meters

        Returns:
            Dict with status and round info
        """
        try:
            import pythoncom
            from win32com.client import VARIANT

            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to add rounds to"}

            model = models.Item(1)
            body = model.Body

            # Collect all edges from all body faces
            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if faces.Count == 0:
                return {"error": "No faces found on body"}

            edge_list = []
            for fi in range(1, faces.Count + 1):
                face = faces.Item(fi)
                face_edges = face.Edges
                if not hasattr(face_edges, "Count"):
                    continue
                for ei in range(1, face_edges.Count + 1):
                    edge_list.append(face_edges.Item(ei))

            if not edge_list:
                return {"error": "No edges found on body"}

            # Each edge as its own edge set, all with the same radius
            edge_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, edge_list)
            radius_arr = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_R8,
                [radius] * len(edge_list),
            )

            rounds = model.Rounds
            rounds.Add(len(edge_list), edge_arr, radius_arr)

            return {
                "status": "created",
                "type": "round",
                "radius": radius,
                "edge_count": len(edge_list),
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_round_on_face(self, radius: float, face_index: int) -> dict[str, Any]:
        """
        Create a round (fillet) on edges of a specific face.

        Unlike create_round() which rounds all edges, this targets only
        the edges of a single face. Use get_body_faces() to find face indices.

        Args:
            radius: Round radius in meters
            face_index: 0-based face index (from get_body_faces)

        Returns:
            Dict with status and round info
        """
        try:
            import pythoncom
            from win32com.client import VARIANT

            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to add rounds to"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)
            face_edges = face.Edges
            if not hasattr(face_edges, "Count") or face_edges.Count == 0:
                return {"error": f"Face {face_index} has no edges"}

            edge_list = []
            for ei in range(1, face_edges.Count + 1):
                edge_list.append(face_edges.Item(ei))

            edge_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, edge_list)
            radius_arr = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_R8,
                [radius] * len(edge_list),
            )

            rounds = model.Rounds
            rounds.Add(len(edge_list), edge_arr, radius_arr)

            return {
                "status": "created",
                "type": "round",
                "radius": radius,
                "face_index": face_index,
                "edge_count": len(edge_list),
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_variable_round(
        self, radii: list[float], face_index: int | None = None,
    ) -> dict[str, Any]:
        """
        Create a variable-radius round (fillet) on body edges.

        Unlike create_round which applies a constant radius, this allows different
        radii at different points along the edge. Uses model.Rounds.AddVariable().
        Type library: Rounds.AddVariable(NumberOfEdgeSets, EdgeSetArray, RadiusArray).

        Args:
            radii: List of radius values in meters. Each edge gets a corresponding radius.
                   If fewer radii than edges, the last radius is repeated.
            face_index: 0-based face index to apply to (None = all edges)

        Returns:
            Dict with status and variable round info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to add variable rounds to"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if faces.Count == 0:
                return {"error": "No faces found on body"}

            # Collect edges from specified face or all faces
            edge_list = []
            if face_index is not None:
                if face_index < 0 or face_index >= faces.Count:
                    return {
                        "error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."
                    }
                face = faces.Item(face_index + 1)
                face_edges = face.Edges
                if hasattr(face_edges, "Count"):
                    for ei in range(1, face_edges.Count + 1):
                        edge_list.append(face_edges.Item(ei))
            else:
                for fi in range(1, faces.Count + 1):
                    face = faces.Item(fi)
                    face_edges = face.Edges
                    if not hasattr(face_edges, "Count"):
                        continue
                    for ei in range(1, face_edges.Count + 1):
                        edge_list.append(face_edges.Item(ei))

            if not edge_list:
                return {"error": "No edges found on body"}

            # Extend radii list to match edge count if needed
            radius_values = list(radii)
            while len(radius_values) < len(edge_list):
                radius_values.append(radius_values[-1])

            # VARIANT wrappers required for Rounds methods
            edge_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, edge_list)
            radius_arr = VARIANT(
                pythoncom.VT_ARRAY | pythoncom.VT_R8, radius_values[: len(edge_list)]
            )

            rounds = model.Rounds
            rounds.AddVariable(1, edge_arr, radius_arr)

            return {
                "status": "created",
                "type": "variable_round",
                "edge_count": len(edge_list),
                "radii": radius_values[: len(edge_list)],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_chamfer(self, distance: float) -> dict[str, Any]:
        """
        Create an equal-setback chamfer on all body edges.

        Chamfers all edges of the body using model.Chamfers.AddEqualSetback().

        Args:
            distance: Chamfer setback distance in meters

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

            # Collect all edges from all body faces
            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if faces.Count == 0:
                return {"error": "No faces found on body"}

            edge_list = []
            for fi in range(1, faces.Count + 1):
                face = faces.Item(fi)
                face_edges = face.Edges
                if not hasattr(face_edges, "Count"):
                    continue
                for ei in range(1, face_edges.Count + 1):
                    edge_list.append(face_edges.Item(ei))

            if not edge_list:
                return {"error": "No edges found on body"}

            chamfers = model.Chamfers
            chamfers.AddEqualSetback(len(edge_list), edge_list, distance)

            return {
                "status": "created",
                "type": "chamfer",
                "distance": distance,
                "edge_count": len(edge_list),
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_chamfer_on_face(self, distance: float, face_index: int) -> dict[str, Any]:
        """
        Create a chamfer on edges of a specific face.

        Unlike create_chamfer() which chamfers all edges, this targets only
        the edges of a single face. Use get_body_faces() to find face indices.

        Args:
            distance: Chamfer setback distance in meters
            face_index: 0-based face index (from get_body_faces)

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

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)
            face_edges = face.Edges
            if not hasattr(face_edges, "Count") or face_edges.Count == 0:
                return {"error": f"Face {face_index} has no edges"}

            edge_list = []
            for ei in range(1, face_edges.Count + 1):
                edge_list.append(face_edges.Item(ei))

            chamfers = model.Chamfers
            chamfers.AddEqualSetback(len(edge_list), edge_list, distance)

            return {
                "status": "created",
                "type": "chamfer",
                "distance": distance,
                "face_index": face_index,
                "edge_count": len(edge_list),
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_chamfer_unequal(
        self, distance1: float, distance2: float, face_index: int = 0
    ) -> dict[str, Any]:
        """
        Create a chamfer with two different setback distances.

        Unlike equal chamfer, this creates an asymmetric chamfer where each side
        has a different setback. Requires a reference face to determine direction.

        Args:
            distance1: First setback distance in meters
            distance2: Second setback distance in meters
            face_index: 0-based face index for the reference face

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

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)
            face_edges = face.Edges
            if not hasattr(face_edges, "Count") or face_edges.Count == 0:
                return {"error": f"Face {face_index} has no edges"}

            edge_list = []
            for ei in range(1, face_edges.Count + 1):
                edge_list.append(face_edges.Item(ei))

            chamfers = model.Chamfers
            chamfers.AddUnequalSetback(face, len(edge_list), edge_list, distance1, distance2)

            return {
                "status": "created",
                "type": "chamfer_unequal",
                "distance1": distance1,
                "distance2": distance2,
                "face_index": face_index,
                "edge_count": len(edge_list),
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_chamfer_unequal_on_face(
        self, distance1: float, distance2: float, face_index: int
    ) -> dict[str, Any]:
        """
        Create an unequal-setback chamfer on all edges of a specific face.

        Args:
            distance1: First setback distance in meters
            distance2: Second setback distance in meters
            face_index: 0-based face index

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

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Count: {faces.Count}"}

            face = faces.Item(face_index + 1)
            edges = face.Edges
            if not hasattr(edges, "Count") or edges.Count == 0:
                return {"error": f"No edges found on face {face_index}"}

            edge_list = []
            for ei in range(1, edges.Count + 1):
                edge_list.append(edges.Item(ei))

            chamfers = model.Chamfers
            chamfers.AddUnequalSetback(len(edge_list), edge_list, distance1, distance2)

            return {
                "status": "created",
                "type": "chamfer_unequal_on_face",
                "distance1": distance1,
                "distance2": distance2,
                "face_index": face_index,
                "edge_count": len(edge_list),
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_chamfer_angle(
        self, distance: float, angle: float, face_index: int = 0
    ) -> dict[str, Any]:
        """
        Create a chamfer defined by a setback distance and an angle.

        Args:
            distance: Setback distance in meters
            angle: Chamfer angle in degrees
            face_index: 0-based face index for the reference face

        Returns:
            Dict with status and chamfer info
        """
        try:
            import math

            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to add chamfers to"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)
            face_edges = face.Edges
            if not hasattr(face_edges, "Count") or face_edges.Count == 0:
                return {"error": f"Face {face_index} has no edges"}

            edge_list = []
            for ei in range(1, face_edges.Count + 1):
                edge_list.append(face_edges.Item(ei))

            chamfers = model.Chamfers
            angle_rad = math.radians(angle)
            chamfers.AddSetbackAngle(face, len(edge_list), edge_list, distance, angle_rad)

            return {
                "status": "created",
                "type": "chamfer_angle",
                "distance": distance,
                "angle_degrees": angle,
                "face_index": face_index,
                "edge_count": len(edge_list),
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_blend(self, radius: float, face_index: int | None = None) -> dict[str, Any]:
        """
        Create a blend (face-to-face fillet) feature.

        Uses model.Blends.Add(NumberOfEdgeSets, EdgeSetArray, RadiusArray).
        Same VARIANT wrapper pattern as Rounds. Unlike Rounds which fillets edges,
        Blends create smooth transitions between faces.

        Args:
            radius: Blend radius in meters
            face_index: 0-based face index to apply to (None = all edges)

        Returns:
            Dict with status and blend info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to add blends to"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if faces.Count == 0:
                return {"error": "No faces found on body"}

            # Collect edges
            edge_list = []
            if face_index is not None:
                if face_index < 0 or face_index >= faces.Count:
                    return {
                        "error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."
                    }
                face = faces.Item(face_index + 1)
                face_edges = face.Edges
                if hasattr(face_edges, "Count"):
                    for ei in range(1, face_edges.Count + 1):
                        edge_list.append(face_edges.Item(ei))
            else:
                for fi in range(1, faces.Count + 1):
                    face = faces.Item(fi)
                    face_edges = face.Edges
                    if not hasattr(face_edges, "Count"):
                        continue
                    for ei in range(1, face_edges.Count + 1):
                        edge_list.append(face_edges.Item(ei))

            if not edge_list:
                return {"error": "No edges found on body"}

            # VARIANT wrappers (same pattern as Rounds)
            edge_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, edge_list)
            radius_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [radius])

            blends = model.Blends
            blends.Add(1, edge_arr, radius_arr)

            return {
                "status": "created",
                "type": "blend",
                "radius": radius,
                "edge_count": len(edge_list),
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_blend_variable(
        self, radius1: float, radius2: float, face_index: int | None = None
    ) -> dict[str, Any]:
        """
        Create a variable-radius blend feature.

        Uses Blends.AddVariable(NumberOfEdgeSets, EdgeSetArray, RadiusArray).
        Applies varying radius values from radius1 to radius2 along edges.

        Args:
            radius1: Starting blend radius in meters
            radius2: Ending blend radius in meters
            face_index: 0-based face index to apply to (None = all edges)

        Returns:
            Dict with status and blend info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to add variable blends to"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if faces.Count == 0:
                return {"error": "No faces found on body"}

            edge_list = []
            if face_index is not None:
                if face_index < 0 or face_index >= faces.Count:
                    return {
                        "error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."
                    }
                face = faces.Item(face_index + 1)
                face_edges = face.Edges
                if hasattr(face_edges, "Count"):
                    for ei in range(1, face_edges.Count + 1):
                        edge_list.append(face_edges.Item(ei))
            else:
                for fi in range(1, faces.Count + 1):
                    face = faces.Item(fi)
                    face_edges = face.Edges
                    if not hasattr(face_edges, "Count"):
                        continue
                    for ei in range(1, face_edges.Count + 1):
                        edge_list.append(face_edges.Item(ei))

            if not edge_list:
                return {"error": "No edges found on body"}

            edge_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, edge_list)
            radius_arr = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [radius1, radius2])

            blends = model.Blends
            blends.AddVariable(1, edge_arr, radius_arr)

            return {
                "status": "created",
                "type": "blend_variable",
                "radius1": radius1,
                "radius2": radius2,
                "edge_count": len(edge_list),
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_blend_surface(self, face_index1: int, face_index2: int) -> dict[str, Any]:
        """
        Create a surface blend between two faces.

        Uses Blends.AddSurfaceBlend(Face1, Face2). Creates a smooth
        surface blend transition between two specified faces.

        Args:
            face_index1: 0-based index of the first face
            face_index2: 0-based index of the second face

        Returns:
            Dict with status and blend info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to add surface blends to"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if faces.Count == 0:
                return {"error": "No faces found on body"}

            if face_index1 < 0 or face_index1 >= faces.Count:
                return {
                    "error": f"Invalid face_index1: {face_index1}. Body has {faces.Count} faces."
                }
            if face_index2 < 0 or face_index2 >= faces.Count:
                return {
                    "error": f"Invalid face_index2: {face_index2}. Body has {faces.Count} faces."
                }

            face1 = faces.Item(face_index1 + 1)
            face2 = faces.Item(face_index2 + 1)

            blends = model.Blends
            blends.AddSurfaceBlend(face1, face2)

            return {
                "status": "created",
                "type": "blend_surface",
                "face_index1": face_index1,
                "face_index2": face_index2,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_round_blend(
        self, face_index1: int, face_index2: int, radius: float
    ) -> dict[str, Any]:
        """
        Create a round blend between two faces.

        Uses Rounds.AddBlend(Face1, Face2, Radius). Creates a smooth
        fillet transition between two specified faces.

        Args:
            face_index1: 0-based index of the first face
            face_index2: 0-based index of the second face
            radius: Blend radius in meters

        Returns:
            Dict with status and blend info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to add round blends to"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if faces.Count == 0:
                return {"error": "No faces found on body"}

            if face_index1 < 0 or face_index1 >= faces.Count:
                return {
                    "error": f"Invalid face_index1: {face_index1}. Body has {faces.Count} faces."
                }
            if face_index2 < 0 or face_index2 >= faces.Count:
                return {
                    "error": f"Invalid face_index2: {face_index2}. Body has {faces.Count} faces."
                }

            face1 = faces.Item(face_index1 + 1)
            face2 = faces.Item(face_index2 + 1)

            rounds = model.Rounds
            rounds.AddBlend(face1, face2, radius)

            return {
                "status": "created",
                "type": "round_blend",
                "face_index1": face_index1,
                "face_index2": face_index2,
                "radius": radius,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_round_surface_blend(
        self, face_index1: int, face_index2: int, radius: float
    ) -> dict[str, Any]:
        """
        Create a round surface blend between two faces.

        Uses Rounds.AddSurfaceBlend(Face1, Face2, Radius). Creates a surface
        blend between two faces with finer control than AddBlend.

        Args:
            face_index1: 0-based index of the first face
            face_index2: 0-based index of the second face
            radius: Blend radius in meters

        Returns:
            Dict with status and blend info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No features exist to add surface blends to"}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if faces.Count == 0:
                return {"error": "No faces found on body"}

            if face_index1 < 0 or face_index1 >= faces.Count:
                return {
                    "error": f"Invalid face_index1: {face_index1}. Body has {faces.Count} faces."
                }
            if face_index2 < 0 or face_index2 >= faces.Count:
                return {
                    "error": f"Invalid face_index2: {face_index2}. Body has {faces.Count} faces."
                }

            face1 = faces.Item(face_index1 + 1)
            face2 = faces.Item(face_index2 + 1)

            rounds = model.Rounds
            rounds.AddSurfaceBlend(face1, face2, radius)

            return {
                "status": "created",
                "type": "round_surface_blend",
                "face_index1": face_index1,
                "face_index2": face_index2,
                "radius": radius,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_delete_hole(
        self, max_diameter: float = 1.0, hole_type: str = "All"
    ) -> dict[str, Any]:
        """
        Delete/fill holes in the model body.

        Uses model.DeleteHoles.Add(HoleTypeToDelete, ThresholdHoleDiameter).
        Fills holes up to the specified diameter threshold.

        Args:
            max_diameter: Maximum hole diameter to delete (meters). Holes with
                diameter <= this value will be filled.
            hole_type: Type of holes to delete: 'All', 'Round', 'NonRound'

        Returns:
            Dict with status and deletion info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists. Create a base feature first."}

            model = models.Item(1)

            # HoleTypeToDeleteConstants from type library
            type_map = {
                "All": 0,
                "Round": 1,
                "NonRound": 2,
            }
            hole_type_const = type_map.get(hole_type, 0)

            delete_holes = model.DeleteHoles
            delete_holes.Add(hole_type_const, max_diameter)

            return {
                "status": "created",
                "type": "delete_hole",
                "max_diameter": max_diameter,
                "hole_type": hole_type,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_delete_blend(self, face_index: int) -> dict[str, Any]:
        """
        Delete/remove a blend (fillet) from the model by specifying a face.

        Uses model.DeleteBlends.Add(BlendsToDelete). Removes the blend
        associated with the specified face.

        Args:
            face_index: 0-based face index of the blend face to remove

        Returns:
            Dict with status and deletion info
        """
        try:
            doc = self.doc_manager.get_active_document()
            models = doc.Models

            if models.Count == 0:
                return {"error": "No base feature exists."}

            model = models.Item(1)
            body = model.Body

            faces = body.Faces(FaceQueryConstants.igQueryAll)
            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."}

            face = faces.Item(face_index + 1)

            delete_blends = model.DeleteBlends
            delete_blends.Add(face)

            return {"status": "created", "type": "delete_blend", "face_index": face_index}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
