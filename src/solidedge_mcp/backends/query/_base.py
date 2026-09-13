"""
Base class for QueryManager providing constructor and shared helpers.
"""

from typing import Any

from ..constants import FaceQueryConstants
from ..logging import get_logger

_logger = get_logger(__name__)

#: Default number of entities returned by a paged collection query.
DEFAULT_PAGE_LIMIT = 200
#: Hard ceiling on ``limit`` so a single call can never walk an entire
#: imported model (one COM round trip per entity).
MAX_PAGE_LIMIT = 2000


def r8_array(size: int) -> Any:
    """Buffer for a COM ``SAFEARRAY(VT_R8)*`` ``[in, out]`` parameter.

    Pass a plain Python list. pywin32 marshals it into the SAFEARRAY and
    returns the filled values in the result tuple; the list itself is not
    updated in place. Verified against Solid Edge 2026: wrapping it in a
    ``VARIANT`` (with or without ``VT_BYREF``) raises "Objects for SAFEARRAYS
    must be sequences (of sequences), or a buffer object" on ``Body.GetRange``.
    """
    return [0.0] * size


def i4_array(size: int) -> Any:
    """Buffer for a COM ``SAFEARRAY(VT_I4)*`` ``[in, out]`` parameter."""
    return [0] * size


def bool_array(size: int) -> Any:
    """Buffer for a COM ``SAFEARRAY(VT_BOOL)*`` ``[in, out]`` parameter."""
    return [False] * size


def dispatch_array(items: Any) -> Any:
    """Wrap a sequence of COM objects as a ``SAFEARRAY(VT_DISPATCH)``."""
    return list(items)


def page_bounds(total: int, offset: int, limit: int) -> tuple[int, int, int]:
    """Clamp ``offset``/``limit`` against ``total``.

    Returns ``(start, stop, limit)`` with ``0 <= start <= stop <= total``.
    A negative or oversized ``limit`` is clamped into
    ``0..MAX_PAGE_LIMIT``; an ``offset`` past the end yields an empty page.
    """
    limit = min(max(int(limit), 0), MAX_PAGE_LIMIT)
    start = min(max(int(offset), 0), max(total, 0))
    stop = min(start + limit, max(total, 0))
    return start, stop, limit


def page_result(
    items: list[Any], total: int, start: int, limit: int, **extra: Any
) -> dict[str, Any]:
    """Build the standard paging envelope for a bounded collection query.

    ``truncated`` means "more entities follow this page", so a caller can keep
    requesting ``offset += limit`` until it is ``False``.
    """
    result: dict[str, Any] = {
        "total": total,
        "offset": start,
        "limit": limit,
        "items": items,
        "truncated": start + len(items) < total,
    }
    result.update(extra)
    return result


class QueryManagerBase:
    """Base providing __init__ and helpers shared across query mixins."""

    doc_manager: Any

    def __init__(self, document_manager: Any) -> None:
        self.doc_manager = document_manager

    def _get_first_model(self) -> tuple[Any, Any]:
        """Get the first model from the active document."""
        doc = self.doc_manager.get_active_document()
        if not hasattr(doc, "Models"):
            raise Exception("Document does not have a Models collection")
        models = doc.Models
        if models.Count == 0:
            raise Exception("No features in document")
        return doc, models.Item(1)

    def _find_feature(self, feature_name: str) -> tuple[Any, Any]:
        """Find a feature by name in DesignEdgebarFeatures. Returns (feature, doc) or raises."""
        doc = self.doc_manager.get_active_document()
        features = doc.DesignEdgebarFeatures
        for i in range(1, features.Count + 1):
            feat = features.Item(i)
            if hasattr(feat, "Name") and feat.Name == feature_name:
                return feat, doc
        return None, doc

    def _get_body(self) -> tuple[Any, Any, Any]:
        """Get the body from the first model of the active document."""
        doc, model = self._get_first_model()
        return doc, model, model.Body

    def _get_face(self, face_index: int) -> tuple[Any, Any, Any, Any]:
        """Get a specific face by 0-based index. Returns (doc, model, body, face)."""
        doc, model, body = self._get_body()
        faces = body.Faces(FaceQueryConstants.igQueryAll)
        if face_index < 0 or face_index >= faces.Count:
            raise IndexError(f"Invalid face index: {face_index}. Body has {faces.Count} faces.")
        face = faces.Item(face_index + 1)
        return doc, model, body, face

    def _get_face_edge(self, face_index: int, edge_index: int) -> tuple[Any, Any, Any, Any, Any]:
        """Get a specific edge on a face. Returns (doc, model, body, face, edge)."""
        doc, model, body, face = self._get_face(face_index)
        edges = face.Edges
        if edge_index < 0 or edge_index >= edges.Count:
            raise IndexError(f"Invalid edge index: {edge_index}. Face has {edges.Count} edges.")
        edge = edges.Item(edge_index + 1)
        return doc, model, body, face, edge

    @staticmethod
    def _to_list(val: Any) -> list[Any]:
        """Convert a COM value to a Python list (handles iterables and scalars)."""
        if hasattr(val, "__iter__") and not isinstance(val, str):
            return list(val)
        return [val]
