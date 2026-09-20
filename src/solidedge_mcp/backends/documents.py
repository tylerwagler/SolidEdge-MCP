"""
Solid Edge Document Operations

Handles creating, opening, saving, and closing documents.
"""

import contextlib
import os
from typing import Any

from solidedge_mcp.backends.errors import error_result

from .comutil import com_get
from .constants import DocumentTypeConstants
from .features._base import verifies_collection_growth
from .logging import get_logger
from .validation import guard_overwrite

_logger = get_logger(__name__)


def _missing_template(template: str) -> dict[str, Any]:
    """A template the caller named but that is not on disk.

    This used to fall through silently to the default document, so a caller
    who asked for a template got a plain part and no word that their path was
    ignored. It is refused before any COM call is made.
    """
    return {"error": f"Template not found: {template}. Nothing was created."}


#: Why create_weldment needs an explicit template on this Solid Edge.
#: Documents.Add("SolidEdge.WeldmentDocument") is a registered ProgID, and
#: Solid Edge 2026 accepts it -- then goes looking for its default weldment
#: template, and on an install without the weldment environment that file
#: does not exist. It answers with a modal "Path not found" that
#: DisplayAlerts does not suppress, which blocks the one UI thread and hangs
#: this server until someone dismisses it by hand; only then does the call
#: return 0x80030003. Verified live. The server cannot know in advance whether
#: the template is present, so the ProgID route is never taken.
_NO_WELDMENT_TEMPLATE: dict[str, Any] = {
    "error": (
        "create_weldment needs a template path on this Solid Edge. Without one "
        'Documents.Add("SolidEdge.WeldmentDocument") looks for a default '
        "weldment template that is not installed, and raises a modal dialog "
        "that hangs the server. Pass template=<path to a .pwd> that exists."
    ),
    "unsupported": True,
}


class DocumentManager:
    """Manages Solid Edge documents"""

    def __init__(self, connection: Any, sketch_manager: Any | None = None) -> None:
        self.connection = connection
        self.active_document: Any | None = None
        # Optional reference to clear sketch state on doc switch
        self.sketch_manager = sketch_manager
        # Drop cached pointers when the connection layer detects SE went away
        if hasattr(connection, "on_disconnect"):
            connection.on_disconnect = self.on_connection_lost

    def _clear_sketch_state(self) -> None:
        """Clear sketch manager state to prevent stale profile references."""
        if self.sketch_manager:
            self.sketch_manager.clear_state()

    @verifies_collection_growth("Documents", root="application")
    def create_part(self, template: str | None = None) -> dict[str, Any]:
        """Create a new part document"""
        try:
            app = self.connection.get_application()

            if template:
                if not os.path.exists(template):
                    return _missing_template(template)
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
            return error_result(e)

    @verifies_collection_growth("Documents", root="application")
    def create_assembly(self, template: str | None = None) -> dict[str, Any]:
        """Create a new assembly document"""
        try:
            app = self.connection.get_application()

            if template:
                if not os.path.exists(template):
                    return _missing_template(template)
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
            return error_result(e)

    @verifies_collection_growth("Documents", root="application")
    def create_sheet_metal(self, template: str | None = None) -> dict[str, Any]:
        """Create a new sheet metal document"""
        try:
            app = self.connection.get_application()

            if template:
                if not os.path.exists(template):
                    return _missing_template(template)
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
            return error_result(e)

    @verifies_collection_growth("Documents", root="application")
    def create_draft(self, template: str | None = None) -> dict[str, Any]:
        """Create a new draft document"""
        try:
            app = self.connection.get_application()

            if template:
                if not os.path.exists(template):
                    return _missing_template(template)
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
            return error_result(e)

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
            return error_result(e)

    def save_document(
        self, file_path: str | None = None, overwrite: bool = False
    ) -> dict[str, Any]:
        """Save the active document, optionally to a new path.

        Saving over an existing file makes Solid Edge raise a modal "This file
        exists. Do you want to overwrite it?" prompt, which blocks the COM call
        until somebody clicks it. Application.DisplayAlerts does not suppress
        that one, so an automation client hangs with no error and no timeout.

        Rather than answer a prompt nobody can see, refuse up front unless the
        caller passes overwrite=True, in which case the existing file is
        removed before the save.

        Args:
            file_path: Target path. Omit to save in place.
            overwrite: Permit replacing an existing file at ``file_path``.
        """
        try:
            if not self.active_document:
                return {"error": "No active document"}

            if file_path:
                err = guard_overwrite(file_path, overwrite)
                if err:
                    return err
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
            return error_result(e)

    def close_document(self, save: bool = True) -> dict[str, Any]:
        """Close the active document"""
        try:
            if not self.active_document:
                return {"error": "No active document"}

            doc_name = self.active_document.Name
            app = self.connection.get_application()

            if save:
                # Document.Dirty is the real member; Document.Saved does not
                # exist on any Solid Edge document interface.
                try:
                    if self.active_document.Dirty:
                        self.active_document.Save()
                except Exception:
                    self.active_document.Save()
            else:
                # Clearing Dirty stops Solid Edge asking to save on close.
                # Alerts are already off from connect(), but a stale session
                # may predate that, so keep belt and braces.
                with contextlib.suppress(Exception):
                    app.DisplayAlerts = False
                with contextlib.suppress(Exception):
                    self.active_document.Dirty = False

            # Close(SaveChanges) is optional, and an omitted one does not mean
            # "discard": this used to close with it omitted and rely on the
            # Dirty=False above, a write that sits inside a suppressed except.
            # When that write failed, Solid Edge decided for itself and raised
            # the modal "This file exists. Do you want to overwrite it?", which
            # blocks every later COM call until somebody clicks it. Saying
            # which is meant costs nothing and cannot be missed.
            self.active_document.Close(bool(save))
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
            return error_result(e)

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
                        "modified": bool(doc.Dirty),
                        "read_only": doc.ReadOnly,
                    }
                )

            return {"documents": documents, "count": len(documents)}
        except Exception as e:
            return error_result(e)

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
                "path": com_get(doc, "FullName", "untitled"),
                "type": self._get_document_type(doc),
            }
        except Exception as e:
            return error_result(e)

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
            return error_result(e)

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
            return error_result(e)

    def get_active_document(self) -> Any:
        """Get the active document, tracking switches made in the Solid Edge UI.

        Re-reads ``Application.ActiveDocument`` on every call. If the user (or a
        previous call) activated a different document, cached sketch state is
        cleared so profiles from the old document are never reused.
        """
        current: Any = None
        try:
            app = self.connection.get_application()
            current = app.ActiveDocument
        except Exception as e:
            if self.active_document is not None:
                # Cannot reach the app right now; serve the cached document.
                return self.active_document
            raise Exception("No active document") from e

        if current is None:
            if self.active_document is not None:
                return self.active_document
            raise Exception("No active document")

        if self.active_document is not None and not self._same_document(
            self.active_document, current
        ):
            _logger.info("Active document changed outside the tool layer; clearing sketch state")
            if self.sketch_manager is not None:
                self.sketch_manager.clear_state()
        self.active_document = current
        return current

    @staticmethod
    def _document_key(doc: Any) -> Any:
        """Stable identity for a document proxy (COM proxies don't support ==)."""
        for attr in ("FullName", "Name"):
            try:
                value = getattr(doc, attr)
            except Exception:
                continue
            if value:
                return value
        return id(doc)

    def _same_document(self, a: Any, b: Any) -> bool:
        if a is b:
            return True
        return bool(self._document_key(a) == self._document_key(b))

    def on_connection_lost(self) -> None:
        """Drop every cached COM pointer after Solid Edge goes away."""
        self.active_document = None
        if self.sketch_manager is not None:
            self.sketch_manager.clear_state()

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
                "name": com_get(doc, "Name", "Unknown"),
                "path": com_get(doc, "FullName", "untitled"),
            }
        except Exception as e:
            return error_result(e)

    @verifies_collection_growth("Documents", root="application")
    def create_weldment(self, template: str | None = None) -> dict[str, Any]:
        """Create a new weldment document"""
        try:
            app = self.connection.get_application()

            if template:
                if not os.path.exists(template):
                    return _missing_template(template)
                doc = app.Documents.Add(template)
            else:
                return _NO_WELDMENT_TEMPLATE

            self.active_document = doc

            return {
                "status": "created",
                "type": "Weldment",
                "name": doc.Name,
                "path": doc.FullName if doc.FullName else "untitled",
            }
        except Exception as e:
            return error_result(e)

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
            return error_result(e)

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
            return error_result(e)

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
            return error_result(e)

    def close_all_documents(
        self, save: bool = False, discard_unsaved: bool = False
    ) -> dict[str, Any]:
        """
        Close every open document, including ones this server did not create.

        Closing everything without saving destroys unsaved work belonging to
        whoever is at the keyboard, and Solid Edge gives no warning once alerts
        are suppressed. So when any open document has unsaved changes and
        ``save`` is False, the call is refused and the documents are named;
        pass ``discard_unsaved`` to go ahead anyway.

        Args:
            save: Save each document before closing.
            discard_unsaved: Permit throwing away unsaved changes.

        Returns:
            Dict with close status and count of closed documents
        """
        try:
            app = self.connection.get_application()
            docs = app.Documents
            count = docs.Count

            if count == 0:
                return {"status": "no_documents", "closed": 0}

            if not save and not discard_unsaved:
                unsaved = []
                for i in range(1, count + 1):
                    doc = docs.Item(i)
                    if com_get(doc, "Dirty", False):
                        unsaved.append(str(com_get(doc, "Name", f"Document {i}")))
                if unsaved:
                    return {
                        "error": (
                            f"{len(unsaved)} open document(s) have unsaved changes: "
                            f"{', '.join(unsaved)}. Closing them all would discard that "
                            "work. Pass save=true to save first, discard_unsaved=true to "
                            "throw it away, or close_document(scope='active') for just "
                            "the current one."
                        ),
                        "unsaved_documents": unsaved,
                    }

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
                            doc.Dirty = False
                    # Always say whether to save; see close_document.
                    doc.Close(bool(save))
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
            return error_result(e)

    def save_copy_as(self, file_path: str, overwrite: bool = False) -> dict[str, Any]:
        """Write a copy of the active document, leaving the current file active.

        Unlike SaveAs this does not repoint the active document. Like every
        other write, it refuses an existing target rather than let Solid Edge
        raise the modal overwrite prompt that blocks the server.

        Args:
            file_path: Full path for the copy, extension included (.par, .asm).
            overwrite: Permit replacing an existing file at ``file_path``.

        Returns:
            Dict with status and file info.
        """
        try:
            if not self.active_document:
                return {"error": "No active document"}

            err = guard_overwrite(file_path, overwrite)
            if err:
                return err

            self.active_document.SaveCopyAs(file_path)

            return {
                "status": "copy_saved",
                "path": file_path,
                "active_document": self.active_document.Name,
            }
        except Exception as e:
            return error_result(e)

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
            return error_result(e)

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
            return error_result(e)

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
