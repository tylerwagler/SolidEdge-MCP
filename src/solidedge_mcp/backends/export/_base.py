"""
Base class for ExportManager providing constructor and shared helpers.
"""

import contextlib
from typing import Any

from ..comutil import com_get as com_get  # re-exported for this package
from ..constants import DocumentTypeConstants
from ..logging import get_logger

_logger = get_logger(__name__)

NOT_A_DRAFT_DOCUMENT = "Active document is not a draft document"
NOT_A_DRAFT = "Active document is not a draft"


def resolve_view(
    doc: Any, no_window_message: str = "No window available"
) -> tuple[Any, dict[str, Any] | None]:
    """Return ``(view, error_dict)`` for the active window's View object.

    Reads ``Windows``/``View`` directly inside try/except instead of probing
    with ``hasattr`` first: on a late-bound proxy the probe is an extra
    ``GetIDsOfNames`` round trip and it reports False for a member whose
    getter raises.
    """
    try:
        windows = doc.Windows
        count = windows.Count
    except Exception:
        return None, {"error": no_window_message}
    if not count:
        return None, {"error": no_window_message}

    view_obj = com_get(windows.Item(1), "View")
    if not view_obj:
        return None, {"error": "Cannot access view object"}
    return view_obj, None


class ExportManagerBase:
    """Base providing __init__ and helpers shared across export mixins."""

    def __init__(self, document_manager: Any) -> None:
        self.doc_manager = document_manager

    def _require_draft(
        self, doc: Any, message: str = NOT_A_DRAFT_DOCUMENT
    ) -> dict[str, Any] | None:
        """Return an error dict when ``doc`` is not a draft, else ``None``.

        Checks ``Document.Type`` against ``DocumentTypeConstants`` rather than
        probing for ``Sheets``/``ActiveSheet`` with ``hasattr``. On a
        late-bound COM proxy that probe is a ``GetIDsOfNames`` round trip whose
        failure mode is version dependent, and it also reports False when the
        member exists but its getter raises — which reported an unrelated COM
        failure as "not a draft document".
        """
        try:
            doc_type = doc.Type
        except Exception:
            return {"error": message}
        if doc_type != DocumentTypeConstants.igDraftDocument:
            return {"error": message}
        return None

    def _resolve_view(
        self, doc: Any, no_window_message: str = "No window available"
    ) -> tuple[Any, dict[str, Any] | None]:
        """Return ``(view, error_dict)`` for the active window's View object."""
        return resolve_view(doc, no_window_message)

    def _get_drawing_views(self) -> Any:
        """Get the DrawingViews collection from the active sheet."""
        doc = self.doc_manager.get_active_document()
        if self._require_draft(doc) is not None:
            raise Exception(NOT_A_DRAFT_DOCUMENT)
        sheet = doc.ActiveSheet
        import win32com.client.dynamic

        dvs = sheet.DrawingViews
        # Force late binding to avoid Part type library mismatch
        with contextlib.suppress(Exception):
            dvs = win32com.client.dynamic.Dispatch(dvs._oleobj_)
        return dvs
