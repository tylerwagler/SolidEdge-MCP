"""
Base class for AssemblyManager providing constructor and shared helpers.
"""

import contextlib
import functools
from collections.abc import Callable
from typing import Any, TypeVar, cast

from ..comutil import com_get as com_get  # re-exported for this package
from ..constants import DocumentTypeConstants, FaceQueryConstants
from ..logging import get_logger

_logger = get_logger(__name__)

_F = TypeVar("_F", bound=Callable[..., dict[str, Any]])

#: Document types that expose ``Occurrences``/``Relations3d``.
#: From ``Program/constant.tlb > DocumentTypeConstants``.
ASSEMBLY_DOCUMENT_TYPES = frozenset(
    {
        DocumentTypeConstants.igAssemblyDocument,
        DocumentTypeConstants.igWeldmentAssemblyDocument,
    }
)

NOT_AN_ASSEMBLY = "Active document is not an assembly"


def occurrence_face_count(doc: Any) -> int | None:
    """Total faces across every occurrence's body, or None if unreadable.

    This is what an assembly-level feature changes. An AssemblyDocument has no
    ``Models`` collection of its own, so the part-level ``_geometry_snapshot``
    reads nothing there and its check is silently inert -- it can never prove a
    failure, which is exactly the state an assembly feature fails in.

    ``Occurrence.Body`` is the occurrence's own body and is where an assembly
    feature's effect would show; it raises until some assembly feature exists,
    so the occurrence's source document is the fallback. Returns None as soon
    as any occurrence cannot be read, so a partial sum is never mistaken for a
    measurement.
    """
    try:
        occurrences = doc.Occurrences
        count = occurrences.Count
        if type(count) is not int:
            return None
    except Exception:
        return None

    total = 0
    for i in range(1, count + 1):
        try:
            occurrence = occurrences.Item(i)
        except Exception:
            return None
        faces = _occurrence_faces(occurrence)
        if faces is None:
            return None
        total += faces
    return total


def _occurrence_faces(occurrence: Any) -> int | None:
    """One occurrence's face count, from its own body or its source document."""
    try:
        value = occurrence.Body.Faces(FaceQueryConstants.igQueryAll).Count
        if type(value) is int:
            return value
    except Exception:
        pass
    try:
        models = occurrence.OccurrenceDocument.Models
        total = 0
        for m in range(1, models.Count + 1):
            value = models.Item(m).Body.Faces(FaceQueryConstants.igQueryAll).Count
            if type(value) is not int:
                return None
            total += value
        return total
    except Exception:
        return None


def verifies_assembly_geometry(fn: _F) -> _F:
    """Refuse to report success when an assembly feature changed no geometry.

    ``AssemblyFeaturesExtrudedCutouts.Add`` and its siblings return cleanly on
    Solid Edge 2026 and add an entry to the feature collection, but remove or
    add no material: the occurrence's body keeps every face it had. Verified
    across all twelve side-constant combinations and both SAFEARRAY spellings
    the API accepts, with UpdateAll called afterwards. Without this the caller
    is told the cutout was created and the assembly is unchanged.

    Like the part-level decorator, it never invents a failure it cannot prove:
    an unreadable count on either side passes the result through untouched, so
    mocked tests and documents with no occurrences stay unaffected.
    """

    @functools.wraps(fn)
    def wrapper(self: Any, /, *args: Any, **kwargs: Any) -> dict[str, Any]:
        try:
            doc = self.doc_manager.get_active_document()
        except Exception:
            return fn(self, *args, **kwargs)
        if type(doc).__module__.startswith("unittest.mock"):
            return fn(self, *args, **kwargs)

        before = occurrence_face_count(doc)
        result = fn(self, *args, **kwargs)
        if not (isinstance(result, dict) and "error" not in result):
            return result
        if before is None:
            return result

        # An update we cannot run proves nothing either way, so its failure
        # is not evidence that the feature did or did not build.
        with_update = getattr(doc, "UpdateAll", None)
        if callable(with_update):
            with contextlib.suppress(Exception):
                with_update()

        after = occurrence_face_count(doc)
        if after is None or after != before:
            return result

        _logger.warning(
            "%s reported success but no occurrence body changed (faces %s); "
            "reporting as a no-geometry error.",
            fn.__name__,
            before,
        )
        return {
            "error": (
                f"{fn.__name__} reported success but changed no geometry: every "
                f"occurrence body still has the same {before} faces. Solid Edge "
                f"2026 accepts these assembly feature calls and records the "
                f"feature without cutting or adding material through COM. Build "
                f"the feature in the Solid Edge UI, or apply it to the part."
            ),
            "attempted": result.get("type"),
            "faces_before": before,
            "faces_after": after,
            "unsupported": True,
        }

    return cast("_F", wrapper)


class AssemblyManagerBase:
    """Base providing __init__ and helpers shared across assembly mixins."""

    def __init__(self, document_manager: Any, sketch_manager: Any | None = None) -> None:
        self.doc_manager = document_manager
        self.sketch_manager = sketch_manager

    def _require_assembly(self, doc: Any) -> dict[str, Any] | None:
        """Return an error dict when ``doc`` is not an assembly, else ``None``.

        Checks ``Document.Type`` against ``DocumentTypeConstants`` rather than
        probing for a member with ``hasattr``. On a late-bound COM proxy
        ``hasattr`` is a ``GetIDsOfNames`` round trip that also reports False
        when the member exists but its getter raises, which turned unrelated
        COM failures into "not an assembly".
        """
        try:
            doc_type = doc.Type
        except Exception:
            return {"error": NOT_AN_ASSEMBLY}
        if doc_type not in ASSEMBLY_DOCUMENT_TYPES:
            return {"error": NOT_AN_ASSEMBLY}
        return None

    def _get_occurrence_matrix(self, occurrence: Any) -> list[float]:
        """Read an occurrence's 4x4 transform as 16 floats.

        ``Occurrence.GetMatrix(Matrix as SAFEARRAY(VT_R8)*)`` declares its one
        parameter ``[in, out]``: pass a plain 16-element list and read the
        filled matrix out of the return value. A ``VARIANT`` wrapper is
        rejected with "Objects for SAFEARRAYS must be sequences".
        """
        values = occurrence.GetMatrix([0.0] * 16)
        return [float(v) for v in values]

    def _validate_occurrence_index(
        self, doc: Any, component_index: int
    ) -> tuple[Any, Any, dict[str, Any] | None]:
        """Validate a single occurrence index and return (occurrences, occurrence, error_dict).

        If error_dict is not None, caller should return it.
        """
        err = self._require_assembly(doc)
        if err:
            return None, None, err

        occurrences = doc.Occurrences

        if component_index < 0 or component_index >= occurrences.Count:
            return (
                None,
                None,
                {
                    "error": f"Invalid component index: "
                    f"{component_index}. Count: {occurrences.Count}"
                },
            )

        occurrence = occurrences.Item(component_index + 1)
        return occurrences, occurrence, None
