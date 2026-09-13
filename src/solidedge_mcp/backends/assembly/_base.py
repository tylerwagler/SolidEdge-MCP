"""
Base class for AssemblyManager providing constructor and shared helpers.
"""

from typing import Any

from ..comutil import com_get as com_get  # re-exported for this package
from ..constants import DocumentTypeConstants
from ..logging import get_logger

_logger = get_logger(__name__)

#: Document types that expose ``Occurrences``/``Relations3d``.
#: From ``Program/constant.tlb > DocumentTypeConstants``.
ASSEMBLY_DOCUMENT_TYPES = frozenset(
    {
        DocumentTypeConstants.igAssemblyDocument,
        DocumentTypeConstants.igWeldmentAssemblyDocument,
    }
)

NOT_AN_ASSEMBLY = "Active document is not an assembly"


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
