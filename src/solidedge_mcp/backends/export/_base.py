"""
Base class for ExportManager providing constructor and shared helpers.
"""

import contextlib


class ExportManagerBase:
    """Base providing __init__ and helpers shared across export mixins."""

    def __init__(self, document_manager):
        self.doc_manager = document_manager

    def _get_drawing_views(self):
        """Get the DrawingViews collection from the active sheet."""
        doc = self.doc_manager.get_active_document()
        if not hasattr(doc, "Sheets"):
            raise Exception("Active document is not a draft document")
        sheet = doc.ActiveSheet
        import win32com.client.dynamic

        dvs = sheet.DrawingViews
        # Force late binding to avoid Part type library mismatch
        with contextlib.suppress(Exception):
            dvs = win32com.client.dynamic.Dispatch(dvs._oleobj_)
        return dvs
