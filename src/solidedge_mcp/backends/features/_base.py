"""
Base class for FeatureManager providing constructor and shared helpers.
"""

import contextlib
import functools
from collections.abc import Callable
from typing import Any, Concatenate, ParamSpec, Protocol

import pythoncom
from win32com.client import VARIANT

from solidedge_mcp.backends.errors import error_result

from ..comutil import com_get, describe_feature_type, profile_origin
from ..constants import (
    FaceQueryConstants,
    LoftSweepConstants,
    ModelingModeConstants,
)
from ..logging import get_logger

_logger = get_logger(__name__)

_P = ParamSpec("_P")
# Decorated methods live on mixins that are not FeatureManagerBase subclasses
# statically, so ``self`` is typed as Any here.
_Creator = Callable[Concatenate[Any, _P], dict[str, Any]]


#: Words that mark a feature as consuming the active sketch. Anything else
#: works on existing edges or faces, where blaming the sketch profile for an
#: empty result is simply wrong.
_SKETCH_DRIVEN = (
    "extrude",
    "protrusion",
    "cutout",
    "revolve",
    "loft",
    "sweep",
    "helix",
    "rib",
    "web",
    "thicken",
    "surface",
    "slot",
    "emboss",
)


def _why_nothing(method_name: str) -> str:
    """The likely reason a feature built nothing, given what it works from."""
    if any(word in method_name for word in _SKETCH_DRIVEN):
        return (
            "The sketch profile is most likely open or invalid; close it as a "
            "closed region first. It can also mean the feature missed the body "
            "entirely, or removed nothing."
        )
    return (
        "This feature works on existing edges or faces rather than a sketch, so "
        "Solid Edge most likely found nothing it could apply to: the radius or "
        "distance may not fit, or the selected edges may already be consumed."
    )


def verifies_geometry(fn: _Creator[_P]) -> _Creator[_P]:
    """Decorate a feature-creation method to confirm it actually built geometry.

    Several Solid Edge COM feature calls silently no-op -- e.g. when the active
    profile failed to close into a region -- yet they do not raise, so the
    wrapped method returns ``status='created'`` with nothing built. This
    decorator snapshots the body's geometry before/after and, on apparent
    success, downgrades the misleading result to an explicit error when the
    body did not change.

    The check keys on the TOTAL FACE COUNT across every body in Models (and
    Models.Count for the first solid), NOT the feature-tree count -- a failed
    feature still adds a tree node, so DesignEdgebarFeatures.Count is not a
    reliable geometry signal, whereas a no-op leaves every body's face count
    unchanged. Summing over all bodies means a multi-body feature that only
    touches body 2+ is still recognised as real geometry. If those counts
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
                    f"{fn.__name__} reported success but created no geometry: the "
                    f"body's face count and volume did not change. {_why_nothing(fn.__name__)}"
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


class _Decorator(Protocol):
    """A method decorator that keeps the decorated signature."""

    def __call__(self, fn: _Creator[_P], /) -> _Creator[_P]: ...


def collection_count(*paths: str, root: str = "document") -> Callable[[Any], int | None]:
    """A snapshot: the summed ``Count`` of the named COM collections.

    This is the signal for creators that build no solid material -- a
    reference plane, a construction surface, a sketch, a document -- so a
    face count never moves for them and ``verifies_geometry`` has nothing to
    say. Each of them does land in a collection whose ``Count`` is readable.

    ``paths`` are dotted from the root: ``"RefPlanes"``,
    ``"Constructions.ExtrudedSurfaces"``, and ``"Models.*.Rounds"`` where
    ``*`` sums over every Item of a collection. ``root`` is ``"document"`` (the
    active document) or ``"application"``, for creators that run before there
    is a document at all, which is what the ones that make one do.

    Several paths are summed on purpose. A surface creator lands in whichever
    of five ``Constructions`` sub-collections suits it, and naming one per
    method would read a surface that landed in a sibling as a no-op. Summing
    the group is immune to that.

    It answers ``None`` -- "cannot read, do not judge" -- when the root is a
    mock, when a step of any path is missing, or when a leaf ``Count`` is not
    a plain ``int``. A MagicMock ``Count`` is truthy and not an int, and has
    already talked one check into claiming something Solid Edge never said.
    """

    def snapshot(self: Any) -> int | None:
        try:
            if root == "application":
                # DocumentManager holds the connection itself; everything
                # else reaches it through its doc_manager.
                holder = getattr(self, "doc_manager", self)
                base = holder.connection.get_application()
            else:
                base = self.doc_manager.get_active_document()
        except Exception:
            return None
        if base is None or type(base).__module__.startswith("unittest.mock"):
            return None
        total = 0
        for path in paths:
            summed = _count_along(base, path.split("."))
            if summed is None:
                return None
            total += summed
        return total

    return snapshot


def _count_along(obj: Any, steps: list[str]) -> int | None:
    """Sum ``Count`` at the end of ``steps``; a ``*`` step fans out over every Item.

    Most feature collections hang off a Model, not the document --
    ``Models.Item(n).Rounds`` -- and a part can hold more than one model, so
    ``"Models.*.Rounds"`` sums the rounds of every model, the way the face
    count check already sums the faces of every body. Any step that cannot be
    read, and any Count that is not a plain int, makes the whole path ``None``.
    """
    if not steps:
        count = com_get(obj, "Count")
        return count if type(count) is int else None
    step, rest = steps[0], steps[1:]
    if step == "*":
        n = com_get(obj, "Count")
        if type(n) is not int:
            return None
        total = 0
        for i in range(1, n + 1):
            try:
                item = obj.Item(i)
            except Exception:
                return None
            part = _count_along(item, rest)
            if part is None:
                return None
            total += part
        return total
    nxt = com_get(obj, step)
    if nxt is None:
        return None
    return _count_along(nxt, rest)


def verifies_change(snapshot: Callable[[Any], int | None], what: str) -> _Decorator:
    """Refuse to report success when ``snapshot`` reads the same before and after.

    The same shape as ``verifies_geometry`` with the measurement made
    pluggable: snapshot, call, and if the result claims success while the
    snapshot did not move, hand back an explicit error instead. Solid Edge
    accepts a great many calls, records the feature, and builds nothing; the
    result dict looks right and only the document disagrees.

    It never invents a failure it cannot prove. A ``None`` from the snapshot
    on either side passes the result through untouched, as does a result that
    already carries an error.
    """

    def decorate(fn: _Creator[_P]) -> _Creator[_P]:
        @functools.wraps(fn)
        def wrapper(self: Any, /, *args: _P.args, **kwargs: _P.kwargs) -> dict[str, Any]:
            before = snapshot(self)
            result = fn(self, *args, **kwargs)
            if not (isinstance(result, dict) and "error" not in result):
                return result
            if before is None:
                return result
            after = snapshot(self)
            if after is None or after != before:
                return result
            _logger.warning(
                "%s reported success but %s did not change (%d); reporting as a no-op error.",
                fn.__name__,
                what,
                before,
            )
            return {
                "error": (
                    f"{fn.__name__} reported success but added nothing: {what} "
                    f"still holds {before} item(s), exactly as before the call. "
                    f"Solid Edge accepted the call and built nothing."
                ),
                "attempted": result.get("type"),
                "count_before": before,
                "count_after": after,
            }

        return wrapper

    return decorate


def verifies_collection_growth(*paths: str, root: str = "document") -> _Decorator:
    """Method decorator: the named collection(s) must grow, or success is refused."""
    return verifies_change(collection_count(*paths, root=root), " + ".join(paths))


def verify_collection_growth_on_creators(
    *paths: str, root: str = "document"
) -> Callable[[type], type]:
    """Class decorator: every ``create_*`` on the mixin must grow the collection(s).

    The counterpart of ``verify_geometry_on_creators`` for a mixin whose
    creators all land in one collection -- every method on RefPlaneMixin ends
    in a ``RefPlanes.Add*``. Apply it only where that is true of every
    creator; one that legitimately replaces rather than adds would be
    misreported as a no-op.
    """
    decorator = verifies_collection_growth(*paths, root=root)

    def decorate(cls: type) -> type:
        for name, attr in list(vars(cls).items()):
            if name.startswith("create_") and callable(attr):
                setattr(cls, name, decorator(attr))
        return cls

    return decorate


def _reference_plane_names(doc: Any) -> set[str]:
    """Names of the document's reference planes, to keep them out of the list.

    They share the edgebar with real features, and an unnamed plane reports an
    empty string, so blanks are dropped too.
    """
    planes = com_get(doc, "RefPlanes")
    count = int(com_get(planes, "Count", 0) or 0)
    names = {""}
    for i in range(1, count + 1):
        name = com_get(planes.Item(i), "Name")
        if name:
            names.add(str(name))
    return names


def enumerate_features(doc: Any) -> list[dict[str, Any]]:
    """The part's features, in tree order, however the document holds them.

    This is the one enumeration feature indices refer to, and anything that
    hands a caller an index has to speak it. ``Models.Item(n).Features`` is the
    normal source and holds real features only, which is why indexing
    ``DesignEdgebarFeatures`` instead shifted every lookup past the reference
    planes.

    A part built in ordered mode and then switched to synchronous is the
    exception: its features stay in the Pathfinder but leave
    ``Models.Item(n).Features`` empty. Verified on Solid Edge 2026, where
    ``Body.Faces`` raises E_FAIL in the same state. Falling back to the
    edgebar keeps the feature list honest there, and each entry says which
    collection it came from so the index lookup follows.
    """
    features: list[dict[str, Any]] = []
    models = com_get(doc, "Models")
    count_models = int(com_get(models, "Count", 0) or 0)
    for m in range(1, count_models + 1):
        model = models.Item(m)
        collection = com_get(model, "Features")
        if collection is None:
            continue
        count = int(com_get(collection, "Count", 0) or 0)
        for i in range(1, count + 1):
            feature = collection.Item(i)
            features.append(
                {
                    "index": len(features),
                    "source": "model",
                    "body_index": m - 1,
                    "position_in_body": i,
                    "name": com_get(feature, "Name", f"Feature_{len(features) + 1}"),
                    **describe_feature_type(com_get(feature, "Type")),
                }
            )
    if features:
        return features
    return enumerate_edgebar_features(doc)


def enumerate_edgebar_features(doc: Any) -> list[dict[str, Any]]:
    """Real features from the Pathfinder tree, planes and datums excluded."""
    edgebar = com_get(doc, "DesignEdgebarFeatures")
    count = int(com_get(edgebar, "Count", 0) or 0)
    if not count:
        return []

    skip = _reference_plane_names(doc)
    features: list[dict[str, Any]] = []
    for i in range(1, count + 1):
        entry = edgebar.Item(i)
        name = com_get(entry, "Name")
        if name is None or name in skip:
            continue
        features.append(
            {
                "index": len(features),
                "source": "edgebar",
                "body_index": 0,
                "position_in_body": i,
                "name": name,
                **describe_feature_type(com_get(entry, "Type")),
            }
        )
    return features


def feature_index_of(doc: Any, feature_name: str) -> int | None:
    """Where ``feature_name`` sits in the index every index-taking tool speaks.

    A position in ``DesignEdgebarFeatures`` is not that index: the tree holds
    the reference planes too, so the same feature is three higher there, and
    handing a caller the tree position under the name "index" pointed them at
    a different feature entirely. Verified on Solid Edge 2026, where the status
    of ExtrudedProtrusion_1 reported index 3 and renaming index 3 hit
    ExtrudedCutout_3.
    """
    for entry in enumerate_features(doc):
        if entry.get("name") == feature_name:
            index = entry.get("index")
            return int(index) if isinstance(index, int) else None
    return None


class FeatureManagerBase:
    """Base providing __init__ and helpers shared across feature mixins."""

    def __init__(self, document_manager: Any, sketch_manager: Any) -> None:
        self.doc_manager = document_manager
        self.sketch_manager = sketch_manager

    def _geometry_snapshot(self) -> tuple[int | None, int | None, float | None]:
        """Return (models_count, total_face_count, total_volume); None for unreadable.

        Used by @verifies_geometry to detect feature calls that silently
        created nothing. total_face_count is 0 when no body exists yet, the
        SUM of the face counts of every body in Models (Item(1..Count)) when
        bodies do exist, and None when any of them cannot be read (e.g. mocked
        COM objects) so the check stays conservative. Summing across bodies is
        what lets multi-body creators that change only body 2+ register as
        real geometry. ``type(x) is int`` guards against MagicMock values in
        unit tests.

        total_volume is the sum of every body's Volume, for the changes a face
        count cannot see: a mirror of a box across its own face is a wider box
        with the same six faces (Solid Edge 2026, where that read as a no-op).
        """
        try:
            doc = self.doc_manager.get_active_document()
        except Exception:
            return (None, None, None)
        # Geometry verification needs a real Solid Edge document. Against a
        # unittest.mock double there is nothing real to measure (its counts are
        # arbitrary), so stay inert rather than invent failures in unit tests.
        if type(doc).__module__.startswith("unittest.mock"):
            return (None, None, None)
        models: int | None = None
        faces: int | None = None
        volume: float | None = None
        with contextlib.suppress(Exception):
            count = doc.Models.Count
            models = count if type(count) is int else None
        if models == 0:
            faces = 0  # no body yet -> definitively zero faces
        elif models is not None and models > 0:
            faces = self._total_face_count(doc, models)
            volume = self._total_volume(doc, models)
        return (models, faces, volume)

    @staticmethod
    def _total_face_count(doc: Any, models_count: int) -> int | None:
        """Sum the face counts of bodies Models.Item(1..models_count).

        Returns None as soon as any body's count is unreadable or not a plain
        int, so a partial sum can never be mistaken for a real measurement.
        """
        total = 0
        try:
            for i in range(1, models_count + 1):
                body = doc.Models.Item(i).Body
                fc = body.Faces(FaceQueryConstants.igQueryAll).Count
                if type(fc) is not int:
                    return None
                total += fc
        except Exception:
            return None
        return total

    @staticmethod
    def _total_volume(doc: Any, models_count: int) -> float | None:
        """Sum Body.Volume over Models.Item(1..models_count); None if any is unreadable."""
        total = 0.0
        try:
            for i in range(1, models_count + 1):
                volume = com_get(doc.Models.Item(i).Body, "Volume")
                if type(volume) is not float:
                    return None
                total += volume
        except Exception:
            return None
        return total

    @staticmethod
    def _no_geometry_created(
        before: tuple[int | None, int | None, float | None],
        after: tuple[int | None, int | None, float | None],
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
            if fa != fb:
                return False
            # Same faces: the volume decides when it can be read on both sides.
            vb, va = before[2], after[2]
            if vb is not None and va is not None:
                return abs(va - vb) <= 1e-12 * max(1.0, abs(vb))
            return True  # body unchanged -> nothing was built
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
        v_profiles = profiles
        v_types = [LoftSweepConstants.igProfileBasedCrossSection] * len(profiles)
        # A SAFEARRAY of SAFEARRAY(VT_R8): the inner VARIANTs are required, only
        # the outer wrapper is not. Dropping them broke the lofted cutout.
        #
        # Each origin must be a point on its own profile; see profile_origin.
        # These were hardcoded to (0, 0), which only lines up when every
        # profile happens to pass through the sketch origin -- otherwise the
        # loft silently built nothing.
        v_origins = [
            VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, list(profile_origin(profile)))
            for profile in profiles
        ]
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
        if com_get(face_edges, "Count", 0) == 0:
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
            if com_get(f, "Name") == feature_name:
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

    def _require_synchronous(self, doc: Any, allow_switch: bool = False) -> dict[str, Any] | None:
        """Ensure the document is in synchronous mode, or explain why not.

        Several feature APIs are synchronous-only: the Models.Add*ByCenter
        primitives and MirrorCopies.AddSync among them. In an ordered document
        they raise a bare 0x80070057 E_INVALIDARG or 0x80004005 E_FAIL with no
        clue why, which made all five primitives and mirror look broken.
        Verified against Solid Edge 2026: the identical calls succeed once
        ModelingMode is switched.

        An empty part is switched automatically, since nothing can be lost.
        A part that already has features is only switched when the caller asks
        for it, because ordered and synchronous rebuild differently and that
        choice belongs to the caller.
        """
        mode = com_get(doc, "ModelingMode")
        if mode is None or mode == ModelingModeConstants.seModelingModeSynchronous:
            return None

        has_features = bool(com_get(com_get(doc, "Models"), "Count", 0))
        if has_features and not allow_switch:
            return {
                "error": (
                    "This feature is synchronous-only and the part is in ordered mode "
                    "with existing features. Switching changes how they rebuild, so it "
                    "is not done for you: pass allow_mode_switch=true, set the document "
                    "to synchronous in Solid Edge, or build the shape from a sketch."
                ),
                "modeling_mode": "ordered",
                "unsupported": True,
            }

        try:
            doc.ModelingMode = ModelingModeConstants.seModelingModeSynchronous
        except Exception as exc:
            return error_result(exc, context="Could not switch to synchronous modeling")
        _logger.info("Switched the part to synchronous mode for a synchronous-only feature")
        return None

    def _enumerate_features(self, doc: Any) -> list[dict[str, Any]]:
        """The part's features, in tree order, however the document holds them."""
        return enumerate_features(doc)

    def _enumerate_edgebar_features(self, doc: Any) -> list[dict[str, Any]]:
        """Real features from the Pathfinder tree, planes and datums excluded."""
        return enumerate_edgebar_features(doc)

    def _get_feature_by_index(self, index: int) -> tuple[Any | None, dict[str, Any] | None]:
        """Resolve a 0-based index from list_features to the COM feature.

        It has to walk the same collection list_features reports.
        ``doc.DesignEdgebarFeatures`` is the whole Pathfinder tree, reference
        planes included, so indexing that directly shifted every lookup by
        however many planes came first and made delete_feature(0) remove a
        reference plane. Each entry records its source so a part whose ordered
        tree is out of reach still resolves.
        """
        doc = self.doc_manager.get_active_document()
        entries = self._enumerate_features(doc)

        if index < 0 or index >= len(entries):
            return None, {
                "error": (
                    f"Invalid feature index: {index}. The active part has "
                    f"{len(entries)} feature(s). Read them with list_features."
                ),
            }

        entry = entries[index]
        if entry.get("source") == "edgebar":
            return doc.DesignEdgebarFeatures.Item(entry["position_in_body"]), None
        model = doc.Models.Item(entry["body_index"] + 1)
        return model.Features.Item(entry["position_in_body"]), None

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
            return error_result(e)

    def feature_suppress(self, index: int) -> dict[str, Any]:
        """Suppress a feature by 0-based index."""
        try:
            feat, err = self._get_feature_by_index(index)
            if err:
                return err
            assert feat is not None
            name = getattr(feat, "Name", f"Feature_{index}")
            feat.Suppress = True
            return {"status": "suppressed", "feature_name": name, "index": index}
        except Exception as e:
            return error_result(e)

    def feature_unsuppress(self, index: int) -> dict[str, Any]:
        """Unsuppress a feature by 0-based index."""
        try:
            feat, err = self._get_feature_by_index(index)
            if err:
                return err
            assert feat is not None
            name = getattr(feat, "Name", f"Feature_{index}")
            feat.Suppress = False
            return {"status": "unsuppressed", "feature_name": name, "index": index}
        except Exception as e:
            return error_result(e)

    def feature_reorder(self, index: int, target_index: int, after: bool = True) -> dict[str, Any]:
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
            # Every feature type has Reorder(TargetFeature, InsertBefore).
            # MoveAfter and MoveBefore are in no Solid Edge type library, so
            # reordering a feature always raised.
            feat.Reorder(target, not after)
            return {
                "status": "reordered",
                "feature_name": name,
                "index": index,
                "target_index": target_index,
                "after": after,
            }
        except Exception as e:
            return error_result(e)

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
            return error_result(e)
