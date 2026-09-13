"""
Base class for FeatureManager providing constructor and shared helpers.
"""

import contextlib
import functools
import traceback
from collections.abc import Callable
from typing import Any, Concatenate, ParamSpec

import pythoncom
from win32com.client import VARIANT

from ..constants import (
    FaceQueryConstants,
    LoftSweepConstants,
)
from ..logging import get_logger

_logger = get_logger(__name__)

_P = ParamSpec("_P")
# Decorated methods live on mixins that are not FeatureManagerBase subclasses
# statically, so ``self`` is typed as Any here.
_Creator = Callable[Concatenate[Any, _P], dict[str, Any]]


def verifies_geometry(fn: _Creator[_P]) -> _Creator[_P]:
    """Decorate a feature-creation method to confirm it actually built geometry.

    Several Solid Edge COM feature calls silently no-op -- e.g. when the active
    profile failed to close into a region -- yet they do not raise, so the
    wrapped method returns ``status='created'`` with nothing built. This
    decorator snapshots the body's geometry before/after and, on apparent
    success, downgrades the misleading result to an explicit error when the
    body did not change.

    The check keys on the body's FACE COUNT (and Models.Count for the first
    solid), NOT the feature-tree count -- a failed feature still adds a tree
    node, so DesignEdgebarFeatures.Count is not a reliable geometry signal,
    whereas a no-op leaves the body's face count unchanged. If those counts
    cannot be read as ints (e.g. mocked tests, or no body), the decorator
    passes the result through unchanged -- it never invents a failure it
    cannot prove.
    """

    @functools.wraps(fn)
    def wrapper(self: Any, /, *args: _P.args, **kwargs: _P.kwargs) -> dict[str, Any]:
        before = self._geometry_snapshot()
        result = fn(self, *args, **kwargs)
        if not (isinstance(result, dict) and "error" not in result):
            return result
        after = self._geometry_snapshot()
        if self._no_geometry_created(before, after):
            _logger.warning(
                f"{fn.__name__} reported success but the body did not change "
                f"(models {before[0]}->{after[0]}, faces {before[1]}->{after[1]}); "
                f"reporting as a no-geometry error."
            )
            return {
                "error": (
                    "Feature reported success but no geometry was created: the "
                    "body's face count did not change. The sketch profile is most "
                    "likely open/invalid -- close it as a closed region first -- "
                    "or the operation had no effect on the body."
                ),
                "attempted": result.get("type"),
                "models_before": before[0],
                "models_after": after[0],
                "faces_before": before[1],
                "faces_after": after[1],
            }
        return result

    return wrapper


def verify_geometry_on_creators(cls: type) -> type:
    """Class decorator: wrap every ``create_*`` method with @verifies_geometry.

    Apply only to mixins whose create_* methods ALL change the solid body
    (add or remove material) -- extrude, revolve, holes, cutout, primitives.
    Do NOT use on mixins with geometry-neutral creators (draft, cosmetic
    thread, surface/ref-plane builders): a successful such op leaves the face
    count unchanged and would be misreported as a no-op. Those need per-method
    decoration with a deliberate skip-list instead.
    """
    for name, attr in list(vars(cls).items()):
        if name.startswith("create_") and callable(attr):
            setattr(cls, name, verifies_geometry(attr))
    return cls


class FeatureManagerBase:
    """Base providing __init__ and helpers shared across feature mixins."""

    def __init__(self, document_manager: Any, sketch_manager: Any) -> None:
        self.doc_manager = document_manager
        self.sketch_manager = sketch_manager

    def _geometry_snapshot(self) -> tuple[int | None, int | None]:
        """Return (models_count, body_face_count); None for unreadable.

        Used by @verifies_geometry to detect feature calls that silently
        created nothing. face_count is 0 when no body exists yet, the body's
        face count when one does, and None when it cannot be read (e.g. mocked
        COM objects) so the check stays conservative. ``type(x) is int`` guards
        against MagicMock values in unit tests.
        """
        try:
            doc = self.doc_manager.get_active_document()
        except Exception:
            return (None, None)
        # Geometry verification needs a real Solid Edge document. Against a
        # unittest.mock double there is nothing real to measure (its counts are
        # arbitrary), so stay inert rather than invent failures in unit tests.
        if type(doc).__module__.startswith("unittest.mock"):
            return (None, None)
        models: int | None = None
        faces: int | None = None
        with contextlib.suppress(Exception):
            count = doc.Models.Count
            models = count if type(count) is int else None
        if models == 0:
            faces = 0  # no body yet -> definitively zero faces
        elif models is not None and models > 0:
            with contextlib.suppress(Exception):
                body = doc.Models.Item(1).Body
                fc = body.Faces(FaceQueryConstants.igQueryAll).Count
                faces = fc if type(fc) is int else None
        return (models, faces)

    @staticmethod
    def _no_geometry_created(
        before: tuple[int | None, int | None],
        after: tuple[int | None, int | None],
    ) -> bool:
        """True only when we can PROVE the operation changed no geometry.

        Success is a new solid (Models.Count grew, the base feature) or a change
        in the body's face count (any add/cut/round/chamfer/hole). When neither
        can be read we return False -- never claim a failure we cannot prove.
        """
        mb, ma = before[0], after[0]
        if mb is not None and ma is not None and ma > mb:
            return False  # base solid appeared
        fb, fa = before[1], after[1]
        if fb is not None and fa is not None:
            return fa == fb  # body unchanged -> nothing was built
        return False

    def _get_ref_plane(self, doc: Any, plane_index: int = 1) -> Any:
        """Get a reference plane from the document (1=Top/XY, 2=Right/YZ, 3=Front/XZ)"""
        return doc.RefPlanes.Item(plane_index)

    def _make_loft_variant_arrays(self, profiles: list[Any]) -> tuple[Any, Any, Any]:
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

    def _find_feature_by_name(self, feature_name: str) -> tuple[Any | None, dict[str, Any] | None]:
        """Find a feature by name in DesignEdgebarFeatures.

        Returns (feature, error_dict). If found, error_dict is None.
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

    def _get_feature_by_index(self, index: int) -> tuple[Any | None, dict[str, Any] | None]:
        """Get a feature from DesignEdgebarFeatures by 0-based index."""
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
            assert feat is not None
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
            assert feat is not None
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
            assert feat is not None
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
            assert feat is not None
            target, err = self._get_feature_by_index(target_index)
            if err:
                return err
            assert target is not None
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
            assert feat is not None
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
            assert feat is not None
            feat.ConvertToType(target_type)
            return {
                "status": "converted",
                "feature_name": feature_name,
                "target_type": target_type,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
