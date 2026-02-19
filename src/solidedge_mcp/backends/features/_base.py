"""
Base class for FeatureManager providing constructor and shared helpers.
"""

import contextlib
import traceback
from typing import Any

import pythoncom
from win32com.client import VARIANT

from ..constants import (
    FaceQueryConstants,
    LoftSweepConstants,
)
from ..logging import get_logger

_logger = get_logger(__name__)


class FeatureManagerBase:
    """Base providing __init__ and helpers shared across feature mixins."""

    def __init__(self, document_manager, sketch_manager):
        self.doc_manager = document_manager
        self.sketch_manager = sketch_manager

    def _get_ref_plane(self, doc, plane_index: int = 1):
        """Get a reference plane from the document (1=Top/XY, 2=Right/YZ, 3=Front/XZ)"""
        return doc.RefPlanes.Item(plane_index)

    def _make_loft_variant_arrays(self, profiles):
        """Create properly typed VARIANT arrays for loft/sweep COM calls.

        COM requires explicit VARIANT typing for SAFEARRAY parameters.
        Python's automatic marshaling does not produce correct types for nested arrays.

        Args:
            profiles: List of profile COM objects

        Returns:
            Tuple of (v_profiles, v_types, v_origins) VARIANT arrays
        """
        v_profiles = VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_DISPATCH, profiles)
        v_types = VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_I4,
            [LoftSweepConstants.igProfileBasedCrossSection] * len(profiles),
        )
        v_origins = VARIANT(
            pythoncom.VT_ARRAY | pythoncom.VT_VARIANT,
            [VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, [0.0, 0.0]) for _ in profiles],
        )
        return v_profiles, v_types, v_origins

    def _get_edge_from_face(
        self, face_index: int, edge_index: int = 0
    ) -> tuple[Any, Any, Any, dict[str, Any] | None]:
        """Helper to get an edge from a face on the first model body.

        Returns (model, face, edge, error_dict).
        If error_dict is not None, the caller should return it immediately.
        """
        doc = self.doc_manager.get_active_document()
        models = doc.Models

        if models.Count == 0:
            return (
                None,
                None,
                None,
                {"error": "No base feature exists. Create a sheet metal base feature first."},
            )

        model = models.Item(1)
        body = model.Body

        faces = body.Faces(FaceQueryConstants.igQueryAll)
        if face_index < 0 or face_index >= faces.Count:
            return (
                None,
                None,
                None,
                {"error": f"Invalid face index: {face_index}. Body has {faces.Count} faces."},
            )

        face = faces.Item(face_index + 1)
        face_edges = face.Edges
        if not hasattr(face_edges, "Count") or face_edges.Count == 0:
            return None, None, None, {"error": f"Face {face_index} has no edges."}

        if edge_index < 0 or edge_index >= face_edges.Count:
            return (
                None,
                None,
                None,
                {"error": f"Invalid edge index: {edge_index}. Face has {face_edges.Count} edges."},
            )

        edge = face_edges.Item(edge_index + 1)
        return model, face, edge, None

    def _find_feature_by_name(self, feature_name: str):
        """
        Find a feature by name in DesignEdgebarFeatures.

        Args:
            feature_name: Name of the feature to find

        Returns:
            Tuple of (feature_object, error_dict). If found, error_dict is None.
            If not found, feature_object is None and error_dict has error info.
        """
        doc = self.doc_manager.get_active_document()
        features = doc.DesignEdgebarFeatures
        target = None
        for i in range(1, features.Count + 1):
            f = features.Item(i)
            if hasattr(f, "Name") and f.Name == feature_name:
                target = f
                break

        if target is None:
            names = []
            for i in range(1, features.Count + 1):
                with contextlib.suppress(Exception):
                    names.append(features.Item(i).Name)
            return None, {
                "error": f"Feature '{feature_name}' not found.",
                "available_features": names,
            }

        return target, None

    def _get_feature_by_index(self, index: int):
        """Get a feature from DesignEdgebarFeatures by 0-based index.

        Returns:
            Tuple of (feature_object, error_dict).
        """
        doc = self.doc_manager.get_active_document()
        features = doc.DesignEdgebarFeatures
        com_index = index + 1  # Convert to 1-based

        if com_index < 1 or com_index > features.Count:
            return None, {
                "error": f"Invalid feature index: {index}. "
                f"Feature count: {features.Count}",
            }

        feat = features.Item(com_index)
        return feat, None

    def delete_feature(self, index: int) -> dict[str, Any]:
        """Delete a feature by 0-based index."""
        try:
            feat, err = self._get_feature_by_index(index)
            if err:
                return err
            name = getattr(feat, "Name", f"Feature_{index}")
            feat.Delete()
            _logger.info(f"Deleted feature: {name} (index={index})")
            return {"status": "deleted", "feature_name": name, "index": index}
        except Exception as e:
            _logger.error(f"Delete feature failed: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}

    def feature_suppress(self, index: int) -> dict[str, Any]:
        """Suppress a feature by 0-based index."""
        try:
            feat, err = self._get_feature_by_index(index)
            if err:
                return err
            name = getattr(feat, "Name", f"Feature_{index}")
            feat.Suppress()
            return {"status": "suppressed", "feature_name": name, "index": index}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def feature_unsuppress(self, index: int) -> dict[str, Any]:
        """Unsuppress a feature by 0-based index."""
        try:
            feat, err = self._get_feature_by_index(index)
            if err:
                return err
            name = getattr(feat, "Name", f"Feature_{index}")
            feat.Unsuppress()
            return {"status": "unsuppressed", "feature_name": name, "index": index}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def feature_reorder(
        self, index: int, target_index: int, after: bool = True
    ) -> dict[str, Any]:
        """Reorder a feature by moving it relative to another feature."""
        try:
            feat, err = self._get_feature_by_index(index)
            if err:
                return err
            target, err = self._get_feature_by_index(target_index)
            if err:
                return err
            name = getattr(feat, "Name", f"Feature_{index}")
            if after:
                feat.MoveAfter(target)
            else:
                feat.MoveBefore(target)
            return {
                "status": "reordered",
                "feature_name": name,
                "index": index,
                "target_index": target_index,
                "after": after,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def feature_rename(self, index: int, new_name: str) -> dict[str, Any]:
        """Rename a feature by 0-based index."""
        try:
            feat, err = self._get_feature_by_index(index)
            if err:
                return err
            old_name = getattr(feat, "Name", f"Feature_{index}")
            feat.Name = new_name
            return {
                "status": "renamed",
                "old_name": old_name,
                "new_name": new_name,
                "index": index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def convert_feature_type(
        self, feature_name: str, target_type: str
    ) -> dict[str, Any]:
        """Convert a feature to a different type."""
        try:
            feat, err = self._find_feature_by_name(feature_name)
            if err:
                return err
            feat.ConvertToType(target_type)
            return {
                "status": "converted",
                "feature_name": feature_name,
                "target_type": target_type,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
