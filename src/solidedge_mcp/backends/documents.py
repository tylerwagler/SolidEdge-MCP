"""
Solid Edge Document Operations

Handles creating, opening, saving, and closing documents.
"""

import contextlib
import os
import traceback
from typing import Any

from .constants import DocumentTypeConstants
from .logging import get_logger

_logger = get_logger(__name__)


class DocumentManager:
    """Manages Solid Edge documents"""

    def __init__(
        self, connection: Any, sketch_manager: Any | None = None
    ) -> None:
        self.connection = connection
        self.active_document: Any | None = None
        # Optional reference to clear sketch state on doc switch
        self.sketch_manager = sketch_manager

    def _clear_sketch_state(self) -> None:
        """Clear sketch manager state to prevent stale profile references."""
        if self.sketch_manager:
            self.sketch_manager.clear_state()

    def create_part(self, template: str | None = None) -> dict[str, Any]:
        """Create a new part document"""
        try:
            app = self.connection.get_application()

            if template and os.path.exists(template):
                doc = app.Documents.Add(template)
            else:
                doc = app.Documents.Add("SolidEdge.PartDocument")

            self._clear_sketch_state()
            self.active_document = doc

            _logger.info(f"Created Part document: {doc.Name}")
            return {
                "status": "created",
                "type": "Part",
                "name": doc.Name,
                "path": doc.FullName if doc.FullName else "untitled",
            }
        except Exception as e:
            _logger.error(f"Failed to create Part document: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_assembly(self, template: str | None = None) -> dict[str, Any]:
        """Create a new assembly document"""
        try:
            app = self.connection.get_application()

            if template and os.path.exists(template):
                doc = app.Documents.Add(template)
            else:
                doc = app.Documents.Add("SolidEdge.AssemblyDocument")

            self._clear_sketch_state()
            self.active_document = doc

            _logger.info(f"Created Assembly document: {doc.Name}")
            return {
                "status": "created",
                "type": "Assembly",
                "name": doc.Name,
                "path": doc.FullName if doc.FullName else "untitled",
            }
        except Exception as e:
            _logger.error(f"Failed to create Assembly document: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_sheet_metal(self, template: str | None = None) -> dict[str, Any]:
        """Create a new sheet metal document"""
        try:
            app = self.connection.get_application()

            if template and os.path.exists(template):
                doc = app.Documents.Add(template)
            else:
                doc = app.Documents.Add("SolidEdge.SheetMetalDocument")

            self._clear_sketch_state()
            self.active_document = doc

            _logger.info(f"Created SheetMetal document: {doc.Name}")
            return {
                "status": "created",
                "type": "SheetMetal",
                "name": doc.Name,
                "path": doc.FullName if doc.FullName else "untitled",
            }
        except Exception as e:
            _logger.error(f"Failed to create SheetMetal document: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_draft(self, template: str | None = None) -> dict[str, Any]:
        """Create a new draft document"""
        try:
            app = self.connection.get_application()

            if template and os.path.exists(template):
                doc = app.Documents.Add(template)
            else:
                doc = app.Documents.Add("SolidEdge.DraftDocument")

            self._clear_sketch_state()
            self.active_document = doc

            _logger.info(f"Created Draft document: {doc.Name}")
            return {
                "status": "created",
                "type": "Draft",
                "name": doc.Name,
                "path": doc.FullName if doc.FullName else "untitled",
            }
        except Exception as e:
            _logger.error(f"Failed to create Draft document: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}

    def open_document(self, file_path: str) -> dict[str, Any]:
        """Open an existing document"""
        try:
            if not os.path.exists(file_path):
                return {"error": f"File not found: {file_path}"}

            app = self.connection.get_application()
            doc = app.Documents.Open(file_path)
            self._clear_sketch_state()
            self.active_document = doc

            _logger.info(f"Opened document: {file_path}")
            return {
                "status": "opened",
                "path": file_path,
                "name": doc.Name,
                "type": self._get_document_type(doc),
            }
        except Exception as e:
            _logger.error(f"Failed to open document {file_path}: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}

    def save_document(self, file_path: str | None = None) -> dict[str, Any]:
        """Save the active document"""
        try:
            if not self.active_document:
                return {"error": "No active document"}

            if file_path:
                self.active_document.SaveAs(file_path)
                _logger.info(f"Saved document to: {file_path}")
                return {"status": "saved", "path": file_path, "name": self.active_document.Name}
            else:
                self.active_document.Save()
                _logger.info(f"Saved document: {self.active_document.Name}")
                return {
                    "status": "saved",
                    "path": self.active_document.FullName,
                    "name": self.active_document.Name,
                }
        except Exception as e:
            _logger.error(f"Failed to save document: {e}")
            return {"error": str(e), "traceback": traceback.format_exc()}

    def close_document(self, save: bool = True) -> dict[str, Any]:
        """Close the active document"""
        try:
            if not self.active_document:
                return {"error": "No active document"}

            doc_name = self.active_document.Name
            app = self.connection.get_application()

            if save:
                try:
                    if not self.active_document.Saved:
                        self.active_document.Save()
                except Exception:
                    self.active_document.Save()
            else:
                # Suppress save dialog: disable alerts, mark saved, then close
                with contextlib.suppress(Exception):
                    app.DisplayAlerts = False
                with contextlib.suppress(Exception):
                    self.active_document.Saved = True

            self.active_document.Close()
            self.active_document = None

            # Clear sketch state since the document is gone
            if self.sketch_manager:
                self.sketch_manager.clear_state()

            # Re-enable alerts
            if not save:
                with contextlib.suppress(Exception):
                    app.DisplayAlerts = True

            _logger.info(f"Closed document: {doc_name} (saved={save})")
            return {"status": "closed", "document": doc_name, "saved": save}
        except Exception as e:
            # Make sure alerts are re-enabled
            try:
                app = self.connection.get_application()
                app.DisplayAlerts = True
            except Exception:
                pass
            return {"error": str(e), "traceback": traceback.format_exc()}

    def list_documents(self) -> dict[str, Any]:
        """List all open documents"""
        try:
            app = self.connection.get_application()
            documents = []

            for i in range(app.Documents.Count):
                doc = app.Documents.Item(i + 1)  # COM is 1-indexed
                documents.append(
                    {
                        "index": i,
                        "name": doc.Name,
                        "full_path": doc.FullName if doc.FullName else "untitled",
                        "type": self._get_document_type(doc),
                        "modified": not doc.Saved,
                        "read_only": doc.ReadOnly,
                    }
                )

            return {"documents": documents, "count": len(documents)}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def activate_document(self, name_or_index: str | int) -> dict[str, Any]:
        """
        Activate a specific open document by name or index.

        Note: This clears the sketch manager's active profile and accumulated profiles
        to prevent using stale sketch state from a previous document.

        Args:
            name_or_index: Document name (string) or 0-based index (int)

        Returns:
            Dict with activation status
        """
        try:
            app = self.connection.get_application()
            docs = app.Documents

            if docs.Count == 0:
                return {"error": "No documents are open"}

            doc = None

            if isinstance(name_or_index, int):
                idx = name_or_index
                if idx < 0 or idx >= docs.Count:
                    return {"error": f"Invalid index: {idx}. {docs.Count} documents open."}
                doc = docs.Item(idx + 1)  # COM is 1-indexed
            else:
                # Search by name
                for i in range(1, docs.Count + 1):
                    d = docs.Item(i)
                    if d.Name == name_or_index:
                        doc = d
                        break
                if doc is None:
                    return {"error": f"Document '{name_or_index}' not found"}

            doc.Activate()
            self._clear_sketch_state()
            self.active_document = doc

            _logger.info(f"Activated document: {doc.Name}")
            return {
                "status": "activated",
                "name": doc.Name,
                "path": doc.FullName if hasattr(doc, "FullName") else "untitled",
                "type": self._get_document_type(doc),
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def undo(self) -> dict[str, Any]:
        """
        Undo the last operation on the active document.

        Returns:
            Dict with undo status
        """
        try:
            doc = self.get_active_document()
            doc.Undo()
            return {"status": "undone"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def redo(self) -> dict[str, Any]:
        """
        Redo the last undone operation on the active document.

        Returns:
            Dict with redo status
        """
        try:
            doc = self.get_active_document()
            doc.Redo()
            return {"status": "redone"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_active_document(self) -> Any:
        """Get the active document object"""
        if not self.active_document:
            # Try to get active document from application
            try:
                app = self.connection.get_application()
                self.active_document = app.ActiveDocument
            except Exception as e:
                raise Exception("No active document") from e
        return self.active_document

    def get_active_document_type(self) -> dict[str, Any]:
        """
        Get the type of the currently active document.

        Returns:
            Dict with document type, name, and path
        """
        try:
            doc = self.get_active_document()
            doc_type = self._get_document_type(doc)

            return {
                "type": doc_type,
                "name": doc.Name if hasattr(doc, "Name") else "Unknown",
                "path": doc.FullName if hasattr(doc, "FullName") else "untitled",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_weldment(self, template: str | None = None) -> dict[str, Any]:
        """Create a new weldment document"""
        try:
            app = self.connection.get_application()

            if template and os.path.exists(template):
                doc = app.Documents.Add(template)
            else:
                doc = app.Documents.Add("SolidEdge.WeldmentDocument")

            self.active_document = doc

            return {
                "status": "created",
                "type": "Weldment",
                "name": doc.Name,
                "path": doc.FullName if doc.FullName else "untitled",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def import_file(self, file_path: str) -> dict[str, Any]:
        """
        Import an external CAD file (STEP, IGES, Parasolid, etc.).

        Opens the file using Solid Edge's import translators.

        Args:
            file_path: Path to the file to import

        Returns:
            Dict with import status
        """
        try:
            if not os.path.exists(file_path):
                return {"error": f"File not found: {file_path}"}

            app = self.connection.get_application()
            doc = app.Documents.Open(file_path)
            self.active_document = doc

            return {
                "status": "imported",
                "path": file_path,
                "name": doc.Name,
                "type": self._get_document_type(doc),
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_document_count(self) -> dict[str, Any]:
        """
        Get the count of open documents.

        Returns:
            Dict with document count
        """
        try:
            app = self.connection.get_application()
            return {"count": app.Documents.Count}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def open_in_background(self, file_path: str) -> dict[str, Any]:
        """
        Open a document in the background (no visible window).

        Uses the 0x8 flag to suppress the UI window. Useful for batch
        processing or reading data without user interaction.

        Args:
            file_path: Path to the document file

        Returns:
            Dict with open status
        """
        try:
            if not os.path.exists(file_path):
                return {"error": f"File not found: {file_path}"}

            app = self.connection.get_application()
            doc = app.Documents.Open(file_path, 0x8)
            self.active_document = doc

            return {
                "status": "opened_in_background",
                "path": file_path,
                "name": doc.Name,
                "type": self._get_document_type(doc),
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def close_all_documents(self, save: bool = False) -> dict[str, Any]:
        """
        Close all open documents.

        Args:
            save: If True, save each document before closing

        Returns:
            Dict with close status and count of closed documents
        """
        try:
            app = self.connection.get_application()
            docs = app.Documents
            count = docs.Count

            if count == 0:
                return {"status": "no_documents", "closed": 0}

            closed = 0
            errors = []

            if not save:
                with contextlib.suppress(Exception):
                    app.DisplayAlerts = False

            # Close in reverse order (COM collections shift on removal)
            for i in range(count, 0, -1):
                try:
                    doc = docs.Item(i)
                    if save:
                        with contextlib.suppress(Exception):
                            doc.Save()
                    else:
                        with contextlib.suppress(Exception):
                            doc.Saved = True
                    doc.Close()
                    closed += 1
                except Exception as e:
                    errors.append(str(e))

            if not save:
                with contextlib.suppress(Exception):
                    app.DisplayAlerts = True

            self.active_document = None

            result = {"status": "closed_all", "closed": closed, "total": count}
            if errors:
                result["errors"] = errors
            return result
        except Exception as e:
            try:
                app = self.connection.get_application()
                app.DisplayAlerts = True
            except Exception:
                pass
            return {"error": str(e), "traceback": traceback.format_exc()}

    def save_copy_as(self, file_path: str) -> dict[str, Any]:
        """
        Save a copy of the active document to a new file without changing the active file.

        Unlike SaveAs, this does not change the active document's filename.

        Args:
            file_path: Full path for the copy (must include extension, e.g. .par, .asm)

        Returns:
            Dict with status and file info
        """
        try:
            if not self.active_document:
                return {"error": "No active document"}

            self.active_document.SaveCopyAs(file_path)

            return {
                "status": "copy_saved",
                "path": file_path,
                "active_document": self.active_document.Name,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def open_with_template(self, file_path: str, template: str) -> dict[str, Any]:
        """
        Open a file and map it to a specific template.

        Uses Documents.OpenWithTemplate for more control over how a file
        is opened, particularly useful for imported files.

        Args:
            file_path: Path to the file to open
            template: Path to the template file

        Returns:
            Dict with open status
        """
        try:
            if not os.path.exists(file_path):
                return {"error": f"File not found: {file_path}"}

            app = self.connection.get_application()
            doc = app.Documents.OpenWithTemplate(file_path, template)
            self.active_document = doc

            return {
                "status": "opened_with_template",
                "path": file_path,
                "template": template,
                "name": doc.Name,
                "type": self._get_document_type(doc),
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def open_with_file_open_dialog(
        self, filename: str | None = None, dialog_title: str | None = None
    ) -> dict[str, Any]:
        """
        Open a file using Solid Edge's built-in file open dialog.

        Triggers the native file open dialog. Useful for interactive use cases
        where the user should select the file.

        Args:
            filename: Optional initial filename/filter for the dialog
            dialog_title: Optional custom title for the dialog

        Returns:
            Dict with status
        """
        try:
            app = self.connection.get_application()

            # Build optional variant args
            kwargs = {}
            if filename is not None:
                kwargs["Filename"] = filename
            if dialog_title is not None:
                kwargs["DialogTitle"] = dialog_title

            if kwargs:
                doc = app.Documents.OpenWithFileOpenDialog(**kwargs)
            else:
                doc = app.Documents.OpenWithFileOpenDialog()

            if doc is not None:
                self.active_document = doc
                return {
                    "status": "opened",
                    "name": doc.Name,
                    "type": self._get_document_type(doc),
                }
            else:
                return {"status": "cancelled", "message": "User cancelled the dialog"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def _get_document_type(self, doc: Any) -> str:
        """Determine document type"""
        try:
            doc_type = doc.Type
            type_map = {
                DocumentTypeConstants.igPartDocument: "Part",
                DocumentTypeConstants.igAssemblyDocument: "Assembly",
                DocumentTypeConstants.igDraftDocument: "Draft",
                DocumentTypeConstants.igSheetMetalDocument: "SheetMetal",
                DocumentTypeConstants.igWeldmentDocument: "Weldment",
                DocumentTypeConstants.igWeldmentAssemblyDocument: "WeldmentAssembly",
            }
            return type_map.get(doc_type, f"Unknown({doc_type})")
        except Exception:
            return "Unknown"
