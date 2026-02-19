"""
Base class for QueryManager providing constructor and shared helpers.
"""

from ..constants import FaceQueryConstants
from ..logging import get_logger

_logger = get_logger(__name__)


class QueryManagerBase:
    """Base providing __init__ and helpers shared across query mixins."""

    def __init__(self, document_manager):
        self.doc_manager = document_manager

    def _get_first_model(self):
        """Get the first model from the active document."""
        doc = self.doc_manager.get_active_document()
        if not hasattr(doc, "Models"):
            raise Exception("Document does not have a Models collection")
        models = doc.Models
        if models.Count == 0:
            raise Exception("No features in document")
        return doc, models.Item(1)

    def _find_feature(self, feature_name: str):
        """Find a feature by name in DesignEdgebarFeatures. Returns (feature, doc) or raises."""
        doc = self.doc_manager.get_active_document()
        features = doc.DesignEdgebarFeatures
        for i in range(1, features.Count + 1):
            feat = features.Item(i)
            if hasattr(feat, "Name") and feat.Name == feature_name:
                return feat, doc
        return None, doc

    def _get_body(self):
        """Get the body from the first model of the active document."""
        doc, model = self._get_first_model()
        return doc, model, model.Body

    def _get_face(self, face_index: int):
        """Get a specific face by 0-based index. Returns (doc, model, body, face)."""
        doc, model, body = self._get_body()
        faces = body.Faces(FaceQueryConstants.igQueryAll)
        if face_index < 0 or face_index >= faces.Count:
            raise IndexError(f"Invalid face index: {face_index}. Body has {faces.Count} faces.")
        face = faces.Item(face_index + 1)
        return doc, model, body, face

    def _get_face_edge(self, face_index: int, edge_index: int):
        """Get a specific edge on a face. Returns (doc, model, body, face, edge)."""
        doc, model, body, face = self._get_face(face_index)
        edges = face.Edges
        if edge_index < 0 or edge_index >= edges.Count:
            raise IndexError(f"Invalid edge index: {edge_index}. Face has {edges.Count} edges.")
        edge = edges.Item(edge_index + 1)
        return doc, model, body, face, edge

    @staticmethod
    def _to_list(val):
        """Convert a COM value to a Python list (handles iterables and scalars)."""
        if hasattr(val, "__iter__") and not isinstance(val, str):
            return list(val)
        return [val]
