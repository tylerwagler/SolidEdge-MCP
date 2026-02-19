"""
Base class for AssemblyManager providing constructor and shared helpers.
"""

from typing import Any

from ..logging import get_logger

_logger = get_logger(__name__)


class AssemblyManagerBase:
    """Base providing __init__ and helpers shared across assembly mixins."""

    def __init__(self, document_manager, sketch_manager=None):
        self.doc_manager = document_manager
        self.sketch_manager = sketch_manager

    def _validate_occurrence_index(
        self, doc, component_index: int
    ) -> tuple[Any, Any, dict[str, Any] | None]:
        """Validate a single occurrence index and return (occurrences, occurrence, error_dict).

        If error_dict is not None, caller should return it.
        """
        if not hasattr(doc, "Occurrences"):
            return None, None, {"error": "Active document is not an assembly"}

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
