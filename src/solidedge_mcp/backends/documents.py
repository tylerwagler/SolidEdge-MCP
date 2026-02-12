"""
Solid Edge Document Operations

Handles creating, opening, saving, and closing documents.
"""

from typing import Optional, Dict, Any, List
import os
import traceback
from .constants import DocumentTypeConstants


class DocumentManager:
    """Manages Solid Edge documents"""

    def __init__(self, connection):
        self.connection = connection
        self.active_document: Optional[Any] = None

    def create_part(self, template: Optional[str] = None) -> Dict[str, Any]:
        """Create a new part document"""
        try:
            app = self.connection.get_application()

            if template and os.path.exists(template):
                doc = app.Documents.Add(template)
            else:
                doc = app.Documents.Add("SolidEdge.PartDocument")

            self.active_document = doc

            return {
                "status": "created",
                "type": "Part",
                "name": doc.Name,
                "path": doc.FullName if doc.FullName else "untitled"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_assembly(self, template: Optional[str] = None) -> Dict[str, Any]:
        """Create a new assembly document"""
        try:
            app = self.connection.get_application()

            if template and os.path.exists(template):
                doc = app.Documents.Add(template)
            else:
                doc = app.Documents.Add("SolidEdge.AssemblyDocument")

            self.active_document = doc

            return {
                "status": "created",
                "type": "Assembly",
                "name": doc.Name,
                "path": doc.FullName if doc.FullName else "untitled"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_sheet_metal(self, template: Optional[str] = None) -> Dict[str, Any]:
        """Create a new sheet metal document"""
        try:
            app = self.connection.get_application()

            if template and os.path.exists(template):
                doc = app.Documents.Add(template)
            else:
                doc = app.Documents.Add("SolidEdge.SheetMetalDocument")

            self.active_document = doc

            return {
                "status": "created",
                "type": "SheetMetal",
                "name": doc.Name,
                "path": doc.FullName if doc.FullName else "untitled"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_draft(self, template: Optional[str] = None) -> Dict[str, Any]:
        """Create a new draft document"""
        try:
            app = self.connection.get_application()

            if template and os.path.exists(template):
                doc = app.Documents.Add(template)
            else:
                doc = app.Documents.Add("SolidEdge.DraftDocument")

            self.active_document = doc

            return {
                "status": "created",
                "type": "Draft",
                "name": doc.Name,
                "path": doc.FullName if doc.FullName else "untitled"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def open_document(self, file_path: str) -> Dict[str, Any]:
        """Open an existing document"""
        try:
            if not os.path.exists(file_path):
                return {"error": f"File not found: {file_path}"}

            app = self.connection.get_application()
            doc = app.Documents.Open(file_path)
            self.active_document = doc

            return {
                "status": "opened",
                "path": file_path,
                "name": doc.Name,
                "type": self._get_document_type(doc)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def save_document(self, file_path: Optional[str] = None) -> Dict[str, Any]:
        """Save the active document"""
        try:
            if not self.active_document:
                return {"error": "No active document"}

            if file_path:
                self.active_document.SaveAs(file_path)
                return {
                    "status": "saved",
                    "path": file_path,
                    "name": self.active_document.Name
                }
            else:
                self.active_document.Save()
                return {
                    "status": "saved",
                    "path": self.active_document.FullName,
                    "name": self.active_document.Name
                }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def close_document(self, save: bool = True) -> Dict[str, Any]:
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
                try:
                    app.DisplayAlerts = False
                except Exception:
                    pass
                try:
                    self.active_document.Saved = True
                except Exception:
                    pass  # Some binding modes can't set Saved property

            self.active_document.Close()
            self.active_document = None

            # Re-enable alerts
            if not save:
                try:
                    app.DisplayAlerts = True
                except Exception:
                    pass

            return {
                "status": "closed",
                "document": doc_name,
                "saved": save
            }
        except Exception as e:
            # Make sure alerts are re-enabled
            try:
                app = self.connection.get_application()
                app.DisplayAlerts = True
            except Exception:
                pass
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def list_documents(self) -> Dict[str, Any]:
        """List all open documents"""
        try:
            app = self.connection.get_application()
            documents = []

            for i in range(app.Documents.Count):
                doc = app.Documents.Item(i + 1)  # COM is 1-indexed
                documents.append({
                    "index": i,
                    "name": doc.Name,
                    "full_path": doc.FullName if doc.FullName else "untitled",
                    "type": self._get_document_type(doc),
                    "modified": not doc.Saved,
                    "read_only": doc.ReadOnly
                })

            return {
                "documents": documents,
                "count": len(documents)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def activate_document(self, name_or_index) -> Dict[str, Any]:
        """
        Activate a specific open document by name or index.

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
            self.active_document = doc

            return {
                "status": "activated",
                "name": doc.Name,
                "path": doc.FullName if hasattr(doc, 'FullName') else "untitled",
                "type": self._get_document_type(doc)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def undo(self) -> Dict[str, Any]:
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
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def redo(self) -> Dict[str, Any]:
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
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_active_document(self):
        """Get the active document object"""
        if not self.active_document:
            # Try to get active document from application
            try:
                app = self.connection.get_application()
                self.active_document = app.ActiveDocument
            except:
                raise Exception("No active document")
        return self.active_document

    def get_active_document_type(self) -> Dict[str, Any]:
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
                "name": doc.Name if hasattr(doc, 'Name') else "Unknown",
                "path": doc.FullName if hasattr(doc, 'FullName') else "untitled"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def _get_document_type(self, doc) -> str:
        """Determine document type"""
        try:
            doc_type = doc.Type
            type_map = {
                DocumentTypeConstants.igPartDocument: "Part",
                DocumentTypeConstants.igAssemblyDocument: "Assembly",
                DocumentTypeConstants.igDraftDocument: "Draft",
                DocumentTypeConstants.igSheetMetalDocument: "SheetMetal",
                DocumentTypeConstants.igWeldmentDocument: "Weldment",
                DocumentTypeConstants.igWeldmentAssemblyDocument: "WeldmentAssembly"
            }
            return type_map.get(doc_type, f"Unknown({doc_type})")
        except:
            return "Unknown"
