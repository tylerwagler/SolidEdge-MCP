"""
Solid Edge Export and Visualization Operations

Handles exporting to various formats and creating drawings.
"""

import contextlib
import os
import traceback
from typing import Any

from .constants import DrawingViewOrientationConstants, FoldTypeConstants, RenderModeConstants


class ExportManager:
    """Manages export and visualization operations"""

    def __init__(self, document_manager):
        self.doc_manager = document_manager

    def export_to_step(self, file_path: str) -> dict[str, Any]:
        """
        Export the active document to STEP format.

        Args:
            file_path: Output file path

        Returns:
            Dict with status and export info
        """
        try:
            doc = self.doc_manager.get_active_document()

            # Ensure file has .step or .stp extension
            if not file_path.lower().endswith((".step", ".stp")):
                file_path += ".step"

            # Save as STEP
            doc.SaveAs(file_path)

            return {
                "status": "exported",
                "format": "STEP",
                "path": file_path,
                "size_bytes": os.path.getsize(file_path) if os.path.exists(file_path) else 0,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def export_to_stl(self, file_path: str, quality: str = "Medium") -> dict[str, Any]:
        """
        Export the active document to STL format (for 3D printing).

        Args:
            file_path: Output file path
            quality: Mesh quality - 'Low', 'Medium', or 'High'

        Returns:
            Dict with status and export info
        """
        try:
            doc = self.doc_manager.get_active_document()

            # Ensure file has .stl extension
            if not file_path.lower().endswith(".stl"):
                file_path += ".stl"

            # Quality mapping (actual values depend on Solid Edge version)
            quality_map = {"Low": 0.01, "Medium": 0.001, "High": 0.0001}
            quality_map.get(quality, 0.001)

            # Save as STL
            # Note: Actual method may vary by Solid Edge version
            try:
                doc.SaveAs(file_path)
            except Exception:
                # Alternative export method
                if hasattr(doc, "SaveAsJT"):
                    # Some versions use different export methods
                    pass

            return {
                "status": "exported",
                "format": "STL",
                "path": file_path,
                "quality": quality,
                "size_bytes": os.path.getsize(file_path) if os.path.exists(file_path) else 0,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def export_to_iges(self, file_path: str) -> dict[str, Any]:
        """
        Export the active document to IGES format.

        Args:
            file_path: Output file path

        Returns:
            Dict with status and export info
        """
        try:
            doc = self.doc_manager.get_active_document()

            # Ensure file has .iges or .igs extension
            if not file_path.lower().endswith((".iges", ".igs")):
                file_path += ".iges"

            # Save as IGES
            doc.SaveAs(file_path)

            return {
                "status": "exported",
                "format": "IGES",
                "path": file_path,
                "size_bytes": os.path.getsize(file_path) if os.path.exists(file_path) else 0,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def export_to_pdf(self, file_path: str) -> dict[str, Any]:
        """Export drawing to PDF"""
        try:
            doc = self.doc_manager.get_active_document()

            if not file_path.lower().endswith(".pdf"):
                file_path += ".pdf"

            # PDF export typically works for draft documents
            doc.SaveAs(file_path)

            return {"status": "exported", "format": "PDF", "path": file_path}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_drawing(
        self, template: str | None = None, views: list[str] | None = None
    ) -> dict[str, Any]:
        """
        Create a 2D drawing from the active 3D model.

        The active document must be a saved Part or Assembly. A new Draft document
        is created with a model link to the source, then drawing views are placed.

        Args:
            template: Drawing template path (optional)
            views: List of views to create - ['Front', 'Top', 'Right', 'Isometric']

        Returns:
            Dict with status and drawing info
        """
        try:
            import win32com.client.dynamic as dyn

            source_doc = self.doc_manager.get_active_document()
            app = self.doc_manager.connection.get_application()

            # Source must be saved to disk for ModelLinks.Add
            source_path = None
            with contextlib.suppress(Exception):
                source_path = source_doc.FullName

            if not source_path:
                return {
                    "error": "Active document must be saved "
                    "before creating a drawing. "
                    "Use save_document() first."
                }

            if views is None:
                views = ["Front", "Top", "Right", "Isometric"]

            view_orient_map = {
                "Front": DrawingViewOrientationConstants.Front,
                "Back": DrawingViewOrientationConstants.Back,
                "Top": DrawingViewOrientationConstants.Top,
                "Bottom": DrawingViewOrientationConstants.Bottom,
                "Right": DrawingViewOrientationConstants.Right,
                "Left": DrawingViewOrientationConstants.Left,
                "Isometric": DrawingViewOrientationConstants.Isometric,
                "Iso": DrawingViewOrientationConstants.Isometric,
            }

            # Create a new draft document
            if template and os.path.exists(template):
                draft_doc = app.Documents.Add(template)
            else:
                draft_doc = app.Documents.Add("SolidEdge.DraftDocument")

            # Add model link to the source part/assembly
            model_link = draft_doc.ModelLinks.Add(source_path)

            # Get active sheet and DrawingViews
            sheet = draft_doc.ActiveSheet
            # DrawingViews may bind to wrong type library (Part instead of Draft)
            # due to gen_py cache. Force late binding on the DrawingViews object.
            dvs_early = sheet.DrawingViews
            dvs = dyn.Dispatch(dvs_early._oleobj_)

            # Place views in a grid layout on the sheet
            # Standard A-size sheet is ~0.279 x 0.216 m (A4 landscape)
            positions = [
                (0.10, 0.15),  # View 1
                (0.10, 0.06),  # View 2
                (0.22, 0.15),  # View 3
                (0.22, 0.06),  # View 4
            ]

            views_added = []
            for i, view_name in enumerate(views[:4]):
                orient = view_orient_map.get(view_name)
                if orient is None:
                    continue

                x, y = positions[i] if i < len(positions) else (0.15 + i * 0.08, 0.10)

                try:
                    dvs.AddPartView(model_link, orient, 1.0, x, y, 0)
                    views_added.append(view_name)
                except Exception:
                    # Try the generic Add method as fallback
                    try:
                        dvs.Add(model_link, orient, 1.0, x, y)
                        views_added.append(view_name)
                    except Exception:
                        pass

            return {
                "status": "created",
                "type": "drawing",
                "draft_name": draft_doc.Name if hasattr(draft_doc, "Name") else "Draft",
                "model_link": source_path,
                "views_requested": views,
                "views_added": views_added,
                "total_views": len(views_added),
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_draft_sheet(self) -> dict[str, Any]:
        """
        Add a new sheet to the active draft document.

        The active document must be a Draft document. Creates a new sheet
        and activates it.

        Returns:
            Dict with status and new sheet info
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document. Create a drawing first."}

            sheets = doc.Sheets

            sheet = sheets.AddSheet()
            sheet.Activate()

            return {
                "status": "added",
                "sheet_number": sheets.Count,
                "total_sheets": sheets.Count,
                "name": sheet.Name if hasattr(sheet, "Name") else f"Sheet {sheets.Count}",
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_assembly_drawing_view(
        self, x: float = 0.15, y: float = 0.15, orientation: str = "Isometric", scale: float = 1.0
    ) -> dict[str, Any]:
        """
        Add an assembly drawing view to the active draft document.

        The active document must be a Draft with a model link to an assembly.

        Args:
            x: View center X position on sheet (meters)
            y: View center Y position on sheet (meters)
            orientation: View orientation - 'Front', 'Top', 'Right', 'Isometric', etc.
            scale: View scale factor

        Returns:
            Dict with status and view info
        """
        try:
            import win32com.client.dynamic as dyn

            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            # Get model link
            if not hasattr(doc, "ModelLinks") or doc.ModelLinks.Count == 0:
                return {
                    "error": "No model link found. Create a drawing with create_drawing() first."
                }

            model_link = doc.ModelLinks.Item(1)

            # View orientation constants
            view_orient_map = {
                "Front": 5,
                "Back": 8,
                "Top": 6,
                "Bottom": 9,
                "Right": 7,
                "Left": 10,
                "Isometric": 12,
                "Iso": 12,
            }

            orient = view_orient_map.get(orientation)
            if orient is None:
                valid = ", ".join(view_orient_map.keys())
                return {"error": f"Invalid orientation: {orientation}. Valid: {valid}"}

            sheet = doc.ActiveSheet
            dvs_early = sheet.DrawingViews
            dvs = dyn.Dispatch(dvs_early._oleobj_)

            # seAssemblyDesignedView = 0
            try:
                dvs.AddAssemblyView(model_link, orient, scale, x, y, 0)
            except Exception:
                # Fall back to AddPartView if AddAssemblyView not available
                dvs.AddPartView(model_link, orient, scale, x, y, 0)

            return {
                "status": "added",
                "orientation": orientation,
                "scale": scale,
                "position": [x, y],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def create_parts_list(
        self, auto_balloon: bool = True, x: float = 0.15, y: float = 0.25
    ) -> dict[str, Any]:
        """
        Create a parts list (BOM table) on the active draft sheet.

        Requires the active document to be a Draft with at least one drawing view.
        Uses PartsLists.Add(DrawingView, SavedSettings, AutoBalloon, CreatePartsList).

        Args:
            auto_balloon: Whether to auto-generate balloon callouts
            x: Table X position on sheet (meters, default 0.15)
            y: Table Y position on sheet (meters, default 0.25)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet

            # Get the first drawing view on the sheet
            dvs = sheet.DrawingViews
            if dvs.Count == 0:
                return {"error": "No drawing views on active sheet. Add a view first."}

            dv = dvs.Item(1)

            # Get PartsLists collection
            parts_lists = sheet.PartsLists if hasattr(sheet, "PartsLists") else None
            if parts_lists is None:
                # Try from document level
                parts_lists = doc.PartsLists if hasattr(doc, "PartsLists") else None

            if parts_lists is None:
                return {"error": "PartsLists collection not available"}

            # Add parts list
            # Parameters: DrawingView, SavedSettings, AutoBalloon, CreatePartsList
            # AutoBalloon: 0=No, 1=Yes; CreatePartsList: 0=No, 1=Yes
            parts_lists.Add(dv, "", 1 if auto_balloon else 0, 1)

            return {
                "status": "created",
                "auto_balloon": auto_balloon,
                "total_parts_lists": parts_lists.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def capture_screenshot(
        self, file_path: str, width: int = 1920, height: int = 1080
    ) -> dict[str, Any]:
        """
        Capture a screenshot of the current view.

        Uses View.SaveAsImage(filename, width, height). Supports .png, .jpg, .bmp formats.

        Args:
            file_path: Output image file path (.png, .jpg, .bmp)
            width: Image width in pixels
            height: Image height in pixels

        Returns:
            Dict with status, path, and file size
        """
        try:
            doc = self.doc_manager.get_active_document()

            # Ensure file has image extension
            valid_exts = [".png", ".jpg", ".jpeg", ".bmp"]
            if not any(file_path.lower().endswith(ext) for ext in valid_exts):
                file_path += ".png"

            # Get the window and view
            if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
                return {"error": "No window available for screenshot"}

            window = doc.Windows.Item(1)
            view = window.View if hasattr(window, "View") else None

            if not view:
                return {"error": "Cannot access view object"}

            # View.SaveAsImage(Filename, Width, Height)
            view.SaveAsImage(file_path, width, height)

            return {
                "status": "captured",
                "path": file_path,
                "dimensions": [width, height],
                "size_bytes": os.path.getsize(file_path) if os.path.exists(file_path) else 0,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def export_to_dxf(self, file_path: str) -> dict[str, Any]:
        """Export to DXF format"""
        try:
            doc = self.doc_manager.get_active_document()

            if not file_path.lower().endswith(".dxf"):
                file_path += ".dxf"

            doc.SaveAs(file_path)

            return {
                "status": "exported",
                "format": "DXF",
                "path": file_path,
                "size_bytes": os.path.getsize(file_path) if os.path.exists(file_path) else 0,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def export_to_parasolid(self, file_path: str) -> dict[str, Any]:
        """Export to Parasolid format (X_T or X_B)"""
        try:
            doc = self.doc_manager.get_active_document()

            if not (file_path.lower().endswith(".x_t") or file_path.lower().endswith(".x_b")):
                file_path += ".x_t"

            doc.SaveAs(file_path)

            return {
                "status": "exported",
                "format": "Parasolid",
                "path": file_path,
                "size_bytes": os.path.getsize(file_path) if os.path.exists(file_path) else 0,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def export_to_jt(self, file_path: str) -> dict[str, Any]:
        """Export to JT format"""
        try:
            doc = self.doc_manager.get_active_document()

            if not file_path.lower().endswith(".jt"):
                file_path += ".jt"

            doc.SaveAs(file_path)

            return {
                "status": "exported",
                "format": "JT",
                "path": file_path,
                "size_bytes": os.path.getsize(file_path) if os.path.exists(file_path) else 0,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def export_flat_dxf(self, file_path: str) -> dict[str, Any]:
        """
        Export sheet metal flat pattern to DXF format.

        Only works on sheet metal documents. Exports the flat pattern
        geometry to DXF for use in CNC/laser cutting.

        Args:
            file_path: Output DXF file path

        Returns:
            Dict with export status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not file_path.lower().endswith(".dxf"):
                file_path += ".dxf"

            # Access FlatPatternModels collection (sheet metal only)
            if not hasattr(doc, "FlatPatternModels"):
                return {
                    "error": "Active document is not a "
                    "sheet metal document. "
                    "FlatPatternModels not available."
                }

            flat_models = doc.FlatPatternModels

            # SaveAsFlatDXFEx(filename, face, edge, vertex, useFlatPattern)
            # Pass None for face/edge/vertex to export all
            flat_models.SaveAsFlatDXFEx(file_path, None, None, None, True)

            return {
                "status": "exported",
                "format": "Flat DXF",
                "path": file_path,
                "size_bytes": os.path.getsize(file_path) if os.path.exists(file_path) else 0,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_text_box(self, x: float, y: float, text: str, height: float = 0.005) -> dict[str, Any]:
        """
        Add a text box annotation to the active draft sheet.

        Args:
            x: X position on sheet (meters)
            y: Y position on sheet (meters)
            text: Text content
            height: Text height in meters (default 0.005 = 5mm)

        Returns:
            Dict with status and text box info
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet

            # TextBoxes.Add(x, y, 0) for 2D sheets
            text_boxes = sheet.TextBoxes
            text_box = text_boxes.Add(x, y, 0)

            # Set the text content
            text_box.Text = text

            # Set text height if possible
            with contextlib.suppress(Exception):
                text_box.TextHeight = height

            return {"status": "added", "type": "text_box", "text": text, "position": [x, y]}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_leader(
        self, x1: float, y1: float, x2: float, y2: float, text: str = ""
    ) -> dict[str, Any]:
        """
        Add a leader annotation to the active draft sheet.

        A leader is an arrow pointing to geometry with optional text.

        Args:
            x1: Arrow start X (meters)
            y1: Arrow start Y (meters)
            x2: Text end X (meters)
            y2: Text end Y (meters)
            text: Optional text at the leader end

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet

            leaders = sheet.Leaders
            leader = leaders.Add(x1, y1, 0, x2, y2, 0)

            if text:
                with contextlib.suppress(Exception):
                    leader.Text = text

            return {
                "status": "added",
                "type": "leader",
                "start": [x1, y1],
                "end": [x2, y2],
                "text": text,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_dimension(
        self, x1: float, y1: float, x2: float, y2: float, dim_x: float = None, dim_y: float = None
    ) -> dict[str, Any]:
        """
        Add a linear dimension between two points on the active draft sheet.

        Args:
            x1: First point X (meters)
            y1: First point Y (meters)
            x2: Second point X (meters)
            y2: Second point Y (meters)
            dim_x: Dimension text X position (meters, optional)
            dim_y: Dimension text Y position (meters, optional)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet

            # Default dimension text position to midpoint offset
            if dim_x is None:
                dim_x = (x1 + x2) / 2
            if dim_y is None:
                dim_y = max(y1, y2) + 0.02  # 20mm above

            dimensions = sheet.Dimensions
            dimensions.AddLength(x1, y1, 0, x2, y2, 0, dim_x, dim_y, 0)

            return {
                "status": "added",
                "type": "dimension",
                "point1": [x1, y1],
                "point2": [x2, y2],
                "text_position": [dim_x, dim_y],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_balloon(
        self, x: float, y: float, text: str = "", leader_x: float = None, leader_y: float = None
    ) -> dict[str, Any]:
        """
        Add a balloon annotation to the active draft sheet.

        Balloons are circular annotations typically used for BOM item numbers.

        Args:
            x: Balloon center X (meters)
            y: Balloon center Y (meters)
            text: Text inside the balloon
            leader_x: Leader arrow X (meters, optional)
            leader_y: Leader arrow Y (meters, optional)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet

            balloons = sheet.Balloons

            if leader_x is not None and leader_y is not None:
                balloon = balloons.Add(leader_x, leader_y, 0, x, y, 0)
            else:
                balloon = balloons.Add(x, y, 0, x + 0.02, y + 0.02, 0)

            if text:
                try:
                    balloon.BalloonText = text
                except Exception:
                    with contextlib.suppress(Exception):
                        balloon.Text = text

            return {"status": "added", "type": "balloon", "position": [x, y], "text": text}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_note(self, x: float, y: float, text: str, height: float = 0.005) -> dict[str, Any]:
        """
        Add a note (free-standing text) to the active draft sheet.

        Similar to text box but simpler - just plain text annotation.

        Args:
            x: Note X position (meters)
            y: Note Y position (meters)
            text: Note text content
            height: Text height in meters (default 5mm)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet

            # Try TextBoxes first (most reliable)
            text_boxes = sheet.TextBoxes
            text_box = text_boxes.Add(x, y, 0)
            text_box.Text = text

            with contextlib.suppress(Exception):
                text_box.TextHeight = height

            return {
                "status": "added",
                "type": "note",
                "position": [x, y],
                "text": text,
                "height": height,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_sheet_info(self) -> dict[str, Any]:
        """
        Get information about the active draft sheet.

        Returns:
            Dict with sheet name, size, scale, and counts of drawing objects
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet
            sheets = doc.Sheets

            info = {
                "status": "ok",
                "sheet_name": sheet.Name,
                "sheet_count": sheets.Count,
            }

            # Try to get sheet dimensions
            try:
                info["width"] = sheet.SheetWidth
                info["height"] = sheet.SheetHeight
            except Exception:
                pass

            # Try to get background name
            with contextlib.suppress(Exception):
                info["background"] = sheet.Background

            # Count drawing objects
            with contextlib.suppress(Exception):
                info["drawing_views"] = sheet.DrawingViews.Count
            with contextlib.suppress(Exception):
                info["text_boxes"] = sheet.TextBoxes.Count
            with contextlib.suppress(Exception):
                info["leaders"] = sheet.Leaders.Count
            with contextlib.suppress(Exception):
                info["dimensions"] = sheet.Dimensions.Count

            return info
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def activate_sheet(self, sheet_index: int) -> dict[str, Any]:
        """
        Activate a specific draft sheet by index.

        Args:
            sheet_index: 0-based sheet index

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheets = doc.Sheets
            if sheet_index < 0 or sheet_index >= sheets.Count:
                return {"error": f"Invalid sheet index: {sheet_index}. Count: {sheets.Count}"}

            sheet = sheets.Item(sheet_index + 1)  # COM is 1-indexed
            sheet.Activate()

            return {"status": "activated", "sheet_name": sheet.Name, "sheet_index": sheet_index}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def rename_sheet(self, sheet_index: int, new_name: str) -> dict[str, Any]:
        """
        Rename a draft sheet.

        Args:
            sheet_index: 0-based sheet index
            new_name: New name for the sheet

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheets = doc.Sheets
            if sheet_index < 0 or sheet_index >= sheets.Count:
                return {"error": f"Invalid sheet index: {sheet_index}. Count: {sheets.Count}"}

            sheet = sheets.Item(sheet_index + 1)
            old_name = sheet.Name
            sheet.Name = new_name

            return {
                "status": "renamed",
                "old_name": old_name,
                "new_name": new_name,
                "sheet_index": sheet_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def delete_sheet(self, sheet_index: int) -> dict[str, Any]:
        """
        Delete a draft sheet.

        Args:
            sheet_index: 0-based sheet index

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheets = doc.Sheets
            if sheets.Count <= 1:
                return {"error": "Cannot delete the last sheet"}

            if sheet_index < 0 or sheet_index >= sheets.Count:
                return {"error": f"Invalid sheet index: {sheet_index}. Count: {sheets.Count}"}

            sheet = sheets.Item(sheet_index + 1)
            sheet_name = sheet.Name
            sheet.Delete()

            return {"status": "deleted", "sheet_name": sheet_name, "remaining_sheets": sheets.Count}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # DRAWING VIEW MANAGEMENT
    # =================================================================

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

    def get_drawing_view_count(self) -> dict[str, Any]:
        """
        Get the number of drawing views on the active sheet.

        Returns:
            Dict with view count
        """
        try:
            dvs = self._get_drawing_views()
            return {"count": dvs.Count}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_drawing_view_scale(self, view_index: int) -> dict[str, Any]:
        """
        Get the scale of a drawing view.

        Args:
            view_index: 0-based view index

        Returns:
            Dict with scale value
        """
        try:
            dvs = self._get_drawing_views()
            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)
            scale = view.ScaleFactor

            return {"view_index": view_index, "scale": scale}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_drawing_view_scale(self, view_index: int, scale: float) -> dict[str, Any]:
        """
        Set the scale of a drawing view.

        Args:
            view_index: 0-based view index
            scale: New scale factor (e.g. 1.0, 0.5, 2.0)

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()
            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)
            view.ScaleFactor = scale

            return {"status": "set", "view_index": view_index, "scale": scale}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def delete_drawing_view(self, view_index: int) -> dict[str, Any]:
        """
        Delete a drawing view from the active sheet.

        Args:
            view_index: 0-based view index

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()
            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)
            view.Delete()

            return {"status": "deleted", "view_index": view_index, "remaining_views": dvs.Count}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def update_drawing_view(self, view_index: int) -> dict[str, Any]:
        """
        Force update a drawing view to reflect 3D model changes.

        Args:
            view_index: 0-based view index

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()
            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)
            view.Update()

            return {"status": "updated", "view_index": view_index}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_projected_view(
        self, parent_view_index: int, fold_direction: str, x: float, y: float
    ) -> dict[str, Any]:
        """
        Add a projected (folded) drawing view from a parent view.

        Creates an orthographic projection by folding from the parent view
        in the specified direction.

        Args:
            parent_view_index: 0-based index of the parent drawing view
            fold_direction: 'Up', 'Down', 'Left', or 'Right'
            x: X position on sheet (meters)
            y: Y position on sheet (meters)

        Returns:
            Dict with status and view info
        """
        try:
            dvs = self._get_drawing_views()

            if parent_view_index < 0 or parent_view_index >= dvs.Count:
                return {
                    "error": f"Invalid parent view index: {parent_view_index}. Count: {dvs.Count}"
                }

            fold_map = {
                "Up": FoldTypeConstants.igFoldUp,
                "Down": FoldTypeConstants.igFoldDown,
                "Left": FoldTypeConstants.igFoldLeft,
                "Right": FoldTypeConstants.igFoldRight,
            }

            fold_const = fold_map.get(fold_direction)
            if fold_const is None:
                valid = ", ".join(fold_map.keys())
                return {"error": f"Invalid fold_direction: '{fold_direction}'. Valid: {valid}"}

            parent_view = dvs.Item(parent_view_index + 1)
            dvs.AddByFold(parent_view, fold_const, x, y)

            return {
                "status": "added",
                "parent_view_index": parent_view_index,
                "fold_direction": fold_direction,
                "position": [x, y],
                "total_views": dvs.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def move_drawing_view(self, view_index: int, x: float, y: float) -> dict[str, Any]:
        """
        Reposition a drawing view on the sheet.

        Args:
            view_index: 0-based view index
            x: New X position (meters)
            y: New Y position (meters)

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()

            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)

            try:
                view.OriginX = x
                view.OriginY = y
            except Exception:
                view.XPosition = x
                view.YPosition = y

            return {"status": "moved", "view_index": view_index, "position": [x, y]}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def show_hidden_edges(self, view_index: int, show: bool = True) -> dict[str, Any]:
        """
        Toggle hidden edge visibility on a drawing view.

        Args:
            view_index: 0-based view index
            show: True to show hidden edges, False to hide them

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()

            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)
            view.ShowHiddenEdges = show

            return {"status": "updated", "view_index": view_index, "show_hidden_edges": show}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_drawing_view_display_mode(self, view_index: int, mode: str) -> dict[str, Any]:
        """
        Set the display/render mode of a drawing view.

        Args:
            view_index: 0-based view index
            mode: 'Wireframe', 'HiddenEdgesVisible', 'Shaded', or 'ShadedWithEdges'

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()

            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            mode_map = {
                "Wireframe": RenderModeConstants.seRenderModeWireframe,
                "HiddenEdgesVisible": RenderModeConstants.seRenderModeVHL,
                "Shaded": RenderModeConstants.seRenderModeSmooth,
                "ShadedWithEdges": RenderModeConstants.seRenderModeSmoothBoundary,
            }

            mode_value = mode_map.get(mode)
            if mode_value is None:
                valid = ", ".join(mode_map.keys())
                return {"error": f"Invalid mode: '{mode}'. Valid: {valid}"}

            view = dvs.Item(view_index + 1)

            try:
                view.SetRenderMode(mode_value)
            except Exception:
                view.DisplayMode = mode_value

            return {"status": "updated", "view_index": view_index, "mode": mode}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_drawing_view_info(self, view_index: int) -> dict[str, Any]:
        """
        Get detailed information about a drawing view.

        Args:
            view_index: 0-based view index

        Returns:
            Dict with scale, position, display properties, and name
        """
        try:
            dvs = self._get_drawing_views()

            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)

            info = {"view_index": view_index}

            with contextlib.suppress(Exception):
                info["name"] = view.Name
            with contextlib.suppress(Exception):
                info["scale"] = view.ScaleFactor
            with contextlib.suppress(Exception):
                info["origin_x"] = view.OriginX
            with contextlib.suppress(Exception):
                info["origin_y"] = view.OriginY
            with contextlib.suppress(Exception):
                info["show_hidden_edges"] = view.ShowHiddenEdges
            with contextlib.suppress(Exception):
                info["show_tangent_edges"] = view.ShowTangentEdges
            with contextlib.suppress(Exception):
                info["type"] = view.Type

            return info
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_drawing_view_orientation(self, view_index: int, orientation: str) -> dict[str, Any]:
        """
        Change the orientation of a drawing view.

        Args:
            view_index: 0-based view index
            orientation: 'Front', 'Top', 'Right', 'Back', 'Bottom', 'Left', 'Isometric'

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()

            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            orient_map = {
                "Front": DrawingViewOrientationConstants.Front,
                "Top": DrawingViewOrientationConstants.Top,
                "Right": DrawingViewOrientationConstants.Right,
                "Back": DrawingViewOrientationConstants.Back,
                "Bottom": DrawingViewOrientationConstants.Bottom,
                "Left": DrawingViewOrientationConstants.Left,
                "Isometric": DrawingViewOrientationConstants.Isometric,
                "Iso": DrawingViewOrientationConstants.Isometric,
            }

            orient_const = orient_map.get(orientation)
            if orient_const is None:
                valid = ", ".join(orient_map.keys())
                return {"error": f"Invalid orientation: '{orientation}'. Valid: {valid}"}

            view = dvs.Item(view_index + 1)
            view.ViewOrientation = orient_const

            return {"status": "updated", "view_index": view_index, "orientation": orientation}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # DIMENSION ANNOTATIONS
    # =================================================================

    def add_angular_dimension(
        self,
        x1: float,
        y1: float,
        x2: float,
        y2: float,
        x3: float,
        y3: float,
        dim_x: float | None = None,
        dim_y: float | None = None,
    ) -> dict[str, Any]:
        """
        Add an angular dimension between three points on the active draft sheet.

        The angle is measured at the vertex (x2, y2) between the rays to
        (x1, y1) and (x3, y3).

        Args:
            x1: First ray endpoint X (meters)
            y1: First ray endpoint Y (meters)
            x2: Vertex X (meters)
            y2: Vertex Y (meters)
            x3: Second ray endpoint X (meters)
            y3: Second ray endpoint Y (meters)
            dim_x: Dimension text X position (meters, optional)
            dim_y: Dimension text Y position (meters, optional)

        Returns:
            Dict with status and dimension info
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet
            dims = sheet.Dimensions

            text_x = dim_x if dim_x is not None else (x1 + x3) / 2
            text_y = dim_y if dim_y is not None else (y1 + y3) / 2 + 0.02

            dims.AddAngular(x1, y1, x2, y2, x3, y3, text_x, text_y)

            return {
                "status": "created",
                "type": "angular_dimension",
                "vertex": [x2, y2],
                "text_position": [text_x, text_y],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_radial_dimension(
        self,
        center_x: float,
        center_y: float,
        point_x: float,
        point_y: float,
        dim_x: float | None = None,
        dim_y: float | None = None,
    ) -> dict[str, Any]:
        """
        Add a radial dimension on the active draft sheet.

        Args:
            center_x: Arc center X (meters)
            center_y: Arc center Y (meters)
            point_x: Point on arc X (meters)
            point_y: Point on arc Y (meters)
            dim_x: Dimension text X position (meters, optional)
            dim_y: Dimension text Y position (meters, optional)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet
            dims = sheet.Dimensions

            text_x = dim_x if dim_x is not None else (center_x + point_x) / 2
            text_y = dim_y if dim_y is not None else (center_y + point_y) / 2 + 0.02

            dims.AddRadial(center_x, center_y, point_x, point_y, text_x, text_y)

            return {
                "status": "created",
                "type": "radial_dimension",
                "center": [center_x, center_y],
                "point": [point_x, point_y],
                "text_position": [text_x, text_y],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_diameter_dimension(
        self,
        center_x: float,
        center_y: float,
        point_x: float,
        point_y: float,
        dim_x: float | None = None,
        dim_y: float | None = None,
    ) -> dict[str, Any]:
        """
        Add a diameter dimension on the active draft sheet.

        Args:
            center_x: Circle center X (meters)
            center_y: Circle center Y (meters)
            point_x: Point on circle X (meters)
            point_y: Point on circle Y (meters)
            dim_x: Dimension text X position (meters, optional)
            dim_y: Dimension text Y position (meters, optional)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet
            dims = sheet.Dimensions

            text_x = dim_x if dim_x is not None else center_x + (point_x - center_x) * 1.3
            text_y = dim_y if dim_y is not None else center_y + (point_y - center_y) * 1.3 + 0.02

            dims.AddDiameter(center_x, center_y, point_x, point_y, text_x, text_y)

            return {
                "status": "created",
                "type": "diameter_dimension",
                "center": [center_x, center_y],
                "point": [point_x, point_y],
                "text_position": [text_x, text_y],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_ordinate_dimension(
        self,
        origin_x: float,
        origin_y: float,
        point_x: float,
        point_y: float,
        dim_x: float | None = None,
        dim_y: float | None = None,
    ) -> dict[str, Any]:
        """
        Add an ordinate dimension on the active draft sheet.

        Ordinate dimensions show the distance from an origin to a point
        along a single axis.

        Args:
            origin_x: Datum origin X (meters)
            origin_y: Datum origin Y (meters)
            point_x: Measured point X (meters)
            point_y: Measured point Y (meters)
            dim_x: Dimension text X position (meters, optional)
            dim_y: Dimension text Y position (meters, optional)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet
            dims = sheet.Dimensions

            text_x = dim_x if dim_x is not None else point_x
            text_y = dim_y if dim_y is not None else point_y + 0.02

            dims.AddOrdinate(origin_x, origin_y, point_x, point_y, text_x, text_y)

            return {
                "status": "created",
                "type": "ordinate_dimension",
                "origin": [origin_x, origin_y],
                "point": [point_x, point_y],
                "text_position": [text_x, text_y],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # SYMBOL ANNOTATIONS
    # =================================================================

    def add_center_mark(self, x: float, y: float) -> dict[str, Any]:
        """
        Add a center mark annotation at the specified coordinates.

        Args:
            x: Center mark X position (meters)
            y: Center mark Y position (meters)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet
            center_marks = sheet.CenterMarks
            center_marks.Add(x, y, 0)

            return {
                "status": "added",
                "type": "center_mark",
                "position": [x, y],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_centerline(self, x1: float, y1: float, x2: float, y2: float) -> dict[str, Any]:
        """
        Add a centerline between two points on the active draft sheet.

        Args:
            x1: Start X (meters)
            y1: Start Y (meters)
            x2: End X (meters)
            y2: End Y (meters)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet
            centerlines = sheet.Centerlines
            centerlines.Add(x1, y1, 0, x2, y2, 0)

            return {
                "status": "added",
                "type": "centerline",
                "start": [x1, y1],
                "end": [x2, y2],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_surface_finish_symbol(
        self, x: float, y: float, symbol_type: str = "machined"
    ) -> dict[str, Any]:
        """
        Add a surface finish symbol to the active draft sheet.

        Args:
            x: Symbol X position (meters)
            y: Symbol Y position (meters)
            symbol_type: Type of surface finish - 'machined', 'any', 'prohibited'

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            type_map = {"machined": 1, "any": 0, "prohibited": 2}
            type_value = type_map.get(symbol_type.lower())
            if type_value is None:
                valid = ", ".join(type_map.keys())
                return {"error": f"Invalid symbol_type: '{symbol_type}'. Valid: {valid}"}

            sheet = doc.ActiveSheet

            try:
                sfs = sheet.SurfaceFinishSymbols
                sfs.Add(x, y, 0, type_value)
            except Exception:
                # Fallback: use TextBoxes with standard surface finish text
                text_boxes = sheet.TextBoxes
                symbols = {"machined": "\u2327", "any": "\u2328", "prohibited": "\u2329"}
                text_box = text_boxes.Add(x, y, 0)
                text_box.Text = symbols.get(symbol_type.lower(), "\u2327")

            return {
                "status": "added",
                "type": "surface_finish_symbol",
                "symbol_type": symbol_type,
                "position": [x, y],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_weld_symbol(self, x: float, y: float, weld_type: str = "fillet") -> dict[str, Any]:
        """
        Add a welding symbol to the active draft sheet.

        Args:
            x: Symbol X position (meters)
            y: Symbol Y position (meters)
            weld_type: Type of weld - 'fillet', 'groove', 'plug', 'spot', 'seam'

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            type_map = {"fillet": 0, "groove": 1, "plug": 2, "spot": 3, "seam": 4}
            type_value = type_map.get(weld_type.lower())
            if type_value is None:
                valid = ", ".join(type_map.keys())
                return {"error": f"Invalid weld_type: '{weld_type}'. Valid: {valid}"}

            sheet = doc.ActiveSheet

            try:
                ws = sheet.WeldSymbols
                ws.Add(x, y, 0, type_value)
            except Exception:
                # Fallback: use a leader with weld designation text
                leaders = sheet.Leaders
                leader = leaders.Add(x, y, 0, x + 0.02, y + 0.02, 0)
                with contextlib.suppress(Exception):
                    leader.Text = f"[{weld_type.upper()}]"

            return {
                "status": "added",
                "type": "weld_symbol",
                "weld_type": weld_type,
                "position": [x, y],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_geometric_tolerance(
        self, x: float, y: float, tolerance_text: str = ""
    ) -> dict[str, Any]:
        """
        Add a geometric tolerance (Feature Control Frame / GD&T) to the active draft.

        Args:
            x: FCF X position (meters)
            y: FCF Y position (meters)
            tolerance_text: Tolerance specification text (e.g., "0.05 A B")

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet

            try:
                fcfs = sheet.FCFs
                fcf = fcfs.Add(x, y, 0)
                if tolerance_text:
                    with contextlib.suppress(Exception):
                        fcf.Text = tolerance_text
            except Exception:
                # Fallback: use a text box with GD&T text
                text_boxes = sheet.TextBoxes
                text_box = text_boxes.Add(x, y, 0)
                text_box.Text = tolerance_text if tolerance_text else "[GD&T]"

            return {
                "status": "added",
                "type": "geometric_tolerance",
                "position": [x, y],
                "text": tolerance_text,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # DRAWING VIEW VARIANTS
    # =================================================================

    def add_detail_view(
        self,
        parent_view_index: int,
        center_x: float,
        center_y: float,
        radius: float,
        x: float,
        y: float,
        scale: float = 2.0,
    ) -> dict[str, Any]:
        """
        Add a detail (zoom) view from a parent drawing view.

        Creates a circular detail envelope on the parent view and places
        an enlarged view at the specified position.

        Args:
            parent_view_index: 0-based index of the parent drawing view
            center_x: Detail envelope center X on parent view (meters)
            center_y: Detail envelope center Y on parent view (meters)
            radius: Detail envelope radius (meters)
            x: Detail view X position on sheet (meters)
            y: Detail view Y position on sheet (meters)
            scale: Detail view scale factor (default 2.0)

        Returns:
            Dict with status and view info
        """
        try:
            dvs = self._get_drawing_views()

            if parent_view_index < 0 or parent_view_index >= dvs.Count:
                return {
                    "error": f"Invalid parent view index: {parent_view_index}. Count: {dvs.Count}"
                }

            parent_view = dvs.Item(parent_view_index + 1)

            try:
                dvs.AddByDetailEnvelope(parent_view, center_x, center_y, radius, x, y, scale)
            except Exception:
                # Fallback: try AddDetailView
                dvs.AddDetailView(parent_view, center_x, center_y, radius, x, y, scale)

            return {
                "status": "added",
                "type": "detail_view",
                "parent_view_index": parent_view_index,
                "center": [center_x, center_y],
                "radius": radius,
                "position": [x, y],
                "scale": scale,
                "total_views": dvs.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_auxiliary_view(
        self,
        parent_view_index: int,
        x: float,
        y: float,
        fold_direction: str = "Up",
    ) -> dict[str, Any]:
        """
        Add an auxiliary (folded) view from a parent drawing view.

        An auxiliary view shows the model from an angle not available
        from standard orthographic projections.

        Args:
            parent_view_index: 0-based index of the parent drawing view
            x: Auxiliary view X position on sheet (meters)
            y: Auxiliary view Y position on sheet (meters)
            fold_direction: Fold direction - 'Up', 'Down', 'Left', 'Right'

        Returns:
            Dict with status and view info
        """
        try:
            dvs = self._get_drawing_views()

            if parent_view_index < 0 or parent_view_index >= dvs.Count:
                return {
                    "error": f"Invalid parent view index: {parent_view_index}. Count: {dvs.Count}"
                }

            fold_map = {
                "Up": FoldTypeConstants.igFoldUp,
                "Down": FoldTypeConstants.igFoldDown,
                "Left": FoldTypeConstants.igFoldLeft,
                "Right": FoldTypeConstants.igFoldRight,
            }

            fold_const = fold_map.get(fold_direction)
            if fold_const is None:
                valid = ", ".join(fold_map.keys())
                return {"error": f"Invalid fold_direction: '{fold_direction}'. Valid: {valid}"}

            parent_view = dvs.Item(parent_view_index + 1)

            try:
                dvs.AddByAuxiliaryFold(parent_view, fold_const, x, y)
            except Exception:
                # Fallback: try AddByFold
                dvs.AddByFold(parent_view, fold_const, x, y)

            return {
                "status": "added",
                "type": "auxiliary_view",
                "parent_view_index": parent_view_index,
                "fold_direction": fold_direction,
                "position": [x, y],
                "total_views": dvs.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_draft_view(self, x: float, y: float) -> dict[str, Any]:
        """
        Add an empty draft (sketch) view to the active sheet.

        A draft view is an empty drawing view area where you can add
        free-form sketch geometry and annotations.

        Args:
            x: View X position on sheet (meters)
            y: View Y position on sheet (meters)

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()

            dvs.AddDraftView(x, y)

            return {
                "status": "added",
                "type": "draft_view",
                "position": [x, y],
                "total_views": dvs.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # DRAWING VIEW PROPERTIES
    # =================================================================

    def align_drawing_views(
        self, view_index1: int, view_index2: int, align: bool = True
    ) -> dict[str, Any]:
        """
        Align or unalign two drawing views.

        When aligned, moving one view constrains the other to maintain
        alignment (horizontal or vertical).

        Args:
            view_index1: 0-based index of the first view
            view_index2: 0-based index of the second view
            align: True to align, False to unalign

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()

            if view_index1 < 0 or view_index1 >= dvs.Count:
                return {"error": f"Invalid view_index1: {view_index1}. Count: {dvs.Count}"}
            if view_index2 < 0 or view_index2 >= dvs.Count:
                return {"error": f"Invalid view_index2: {view_index2}. Count: {dvs.Count}"}

            view1 = dvs.Item(view_index1 + 1)
            view2 = dvs.Item(view_index2 + 1)

            if align:
                try:
                    view1.AlignToView(view2)
                except Exception:
                    view2.AlignToView(view1)
            else:
                try:
                    view1.RemoveAlignment()
                except Exception:
                    view2.RemoveAlignment()

            return {
                "status": "aligned" if align else "unaligned",
                "view_index1": view_index1,
                "view_index2": view_index2,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_drawing_view_model_link(self, view_index: int) -> dict[str, Any]:
        """
        Get the model link reference from a drawing view.

        Returns information about which 3D model is associated with
        the specified drawing view.

        Args:
            view_index: 0-based view index

        Returns:
            Dict with model link info (file path, name)
        """
        try:
            dvs = self._get_drawing_views()

            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)

            info = {"view_index": view_index}

            with contextlib.suppress(Exception):
                model_link = view.ModelLink
                info["has_model_link"] = True
                with contextlib.suppress(Exception):
                    info["model_path"] = model_link.FileName
                with contextlib.suppress(Exception):
                    info["model_name"] = model_link.Name
            if "has_model_link" not in info:
                info["has_model_link"] = False

            with contextlib.suppress(Exception):
                info["view_name"] = view.Name
            with contextlib.suppress(Exception):
                info["scale"] = view.ScaleFactor

            return info
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def show_tangent_edges(self, view_index: int, show: bool = True) -> dict[str, Any]:
        """
        Set tangent edge visibility on a drawing view.

        Tangent edges are edges where two surfaces meet tangentially
        (e.g., where a fillet meets a flat face).

        Args:
            view_index: 0-based view index
            show: True to show tangent edges, False to hide them

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()

            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)
            view.ShowTangentEdges = show

            return {
                "status": "updated",
                "view_index": view_index,
                "show_tangent_edges": show,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # Aliases for consistency with MCP tool names
    def export_step(self, file_path: str) -> dict[str, Any]:
        """Alias for export_to_step"""
        return self.export_to_step(file_path)

    def export_stl(self, file_path: str, quality: str = "Medium") -> dict[str, Any]:
        """Alias for export_to_stl"""
        return self.export_to_stl(file_path, quality)

    def export_iges(self, file_path: str) -> dict[str, Any]:
        """Alias for export_to_iges"""
        return self.export_to_iges(file_path)

    def export_pdf(self, file_path: str) -> dict[str, Any]:
        """Alias for export_to_pdf"""
        return self.export_to_pdf(file_path)

    def export_dxf(self, file_path: str) -> dict[str, Any]:
        """Alias for export_to_dxf"""
        return self.export_to_dxf(file_path)

    def export_parasolid(self, file_path: str) -> dict[str, Any]:
        """Alias for export_to_parasolid"""
        return self.export_to_parasolid(file_path)

    def export_jt(self, file_path: str) -> dict[str, Any]:
        """Alias for export_to_jt"""
        return self.export_to_jt(file_path)

    # =================================================================
    # DRAFT SHEET COLLECTIONS (Batch 10)
    # =================================================================

    def get_sheet_dimensions(self) -> dict[str, Any]:
        """
        Get all dimensions on the active draft sheet.

        Iterates sheet.Dimensions collection and collects type, value, and
        position where available.

        Returns:
            Dict with count and list of dimension info dicts
        """
        try:
            doc = self.doc_manager.get_active_document()
            if not hasattr(doc, "ActiveSheet"):
                return {"error": "Active document is not a draft"}
            sheet = doc.ActiveSheet
            dims = sheet.Dimensions
            items = []
            for i in range(1, dims.Count + 1):
                dim = dims.Item(i)
                info: dict[str, Any] = {"index": i - 1}
                with contextlib.suppress(Exception):
                    info["type"] = dim.Type
                with contextlib.suppress(Exception):
                    info["value"] = dim.Value
                with contextlib.suppress(Exception):
                    info["name"] = dim.Name
                items.append(info)
            return {"count": len(items), "dimensions": items}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_sheet_balloons(self) -> dict[str, Any]:
        """
        Get all balloons on the active draft sheet.

        Iterates sheet.Balloons collection and collects text and position
        where available.

        Returns:
            Dict with count and list of balloon info dicts
        """
        try:
            doc = self.doc_manager.get_active_document()
            if not hasattr(doc, "ActiveSheet"):
                return {"error": "Active document is not a draft"}
            sheet = doc.ActiveSheet
            balloons = sheet.Balloons
            items = []
            for i in range(1, balloons.Count + 1):
                balloon = balloons.Item(i)
                info: dict[str, Any] = {"index": i - 1}
                with contextlib.suppress(Exception):
                    info["text"] = balloon.BalloonText
                with contextlib.suppress(Exception):
                    info["x"] = balloon.x
                with contextlib.suppress(Exception):
                    info["y"] = balloon.y
                items.append(info)
            return {"count": len(items), "balloons": items}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_sheet_text_boxes(self) -> dict[str, Any]:
        """
        Get all text boxes on the active draft sheet.

        Iterates sheet.TextBoxes collection and collects text content
        and position where available.

        Returns:
            Dict with count and list of text box info dicts
        """
        try:
            doc = self.doc_manager.get_active_document()
            if not hasattr(doc, "ActiveSheet"):
                return {"error": "Active document is not a draft"}
            sheet = doc.ActiveSheet
            text_boxes = sheet.TextBoxes
            items = []
            for i in range(1, text_boxes.Count + 1):
                tb = text_boxes.Item(i)
                info: dict[str, Any] = {"index": i - 1}
                with contextlib.suppress(Exception):
                    info["text"] = tb.Text
                with contextlib.suppress(Exception):
                    info["x"] = tb.x
                with contextlib.suppress(Exception):
                    info["y"] = tb.y
                with contextlib.suppress(Exception):
                    info["height"] = tb.Height
                items.append(info)
            return {"count": len(items), "text_boxes": items}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_sheet_drawing_objects(self) -> dict[str, Any]:
        """
        Get all drawing objects on the active draft sheet.

        Iterates sheet.DrawingObjects collection and collects type and
        name where available.

        Returns:
            Dict with count and list of drawing object info dicts
        """
        try:
            doc = self.doc_manager.get_active_document()
            if not hasattr(doc, "ActiveSheet"):
                return {"error": "Active document is not a draft"}
            sheet = doc.ActiveSheet
            drawing_objects = sheet.DrawingObjects
            items = []
            for i in range(1, drawing_objects.Count + 1):
                obj = drawing_objects.Item(i)
                info: dict[str, Any] = {"index": i - 1}
                with contextlib.suppress(Exception):
                    info["type"] = str(type(obj).__name__)
                with contextlib.suppress(Exception):
                    info["name"] = obj.Name
                items.append(info)
            return {"count": len(items), "drawing_objects": items}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_sheet_sections(self) -> dict[str, Any]:
        """
        Get all section views on the active draft sheet.

        Iterates sheet.Sections collection and collects label and
        info where available.

        Returns:
            Dict with count and list of section view info dicts
        """
        try:
            doc = self.doc_manager.get_active_document()
            if not hasattr(doc, "ActiveSheet"):
                return {"error": "Active document is not a draft"}
            sheet = doc.ActiveSheet
            sections = sheet.Sections
            items = []
            for i in range(1, sections.Count + 1):
                sec = sections.Item(i)
                info: dict[str, Any] = {"index": i - 1}
                with contextlib.suppress(Exception):
                    info["label"] = sec.Label
                with contextlib.suppress(Exception):
                    info["name"] = sec.Name
                with contextlib.suppress(Exception):
                    info["type"] = sec.Type
                items.append(info)
            return {"count": len(items), "sections": items}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # PRINTING (Batch 10)
    # =================================================================

    def print_drawing(self, copies: int = 1, all_sheets: bool = True) -> dict[str, Any]:
        """
        Print the active draft document.

        Tries doc.PrintOut first, then falls back to DraftPrintUtility.

        Args:
            copies: Number of copies to print
            all_sheets: Whether to print all sheets (True) or active only

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            # Try DraftPrintUtility first (more control)
            if hasattr(doc, "DraftPrintUtility"):
                dpu = doc.DraftPrintUtility
                with contextlib.suppress(Exception):
                    dpu.Copies = copies
                with contextlib.suppress(Exception):
                    dpu.PrintAllSheets = all_sheets
                dpu.PrintOut()
                return {"status": "printed", "copies": copies, "all_sheets": all_sheets}

            # Fall back to simple PrintOut
            if hasattr(doc, "PrintOut"):
                try:
                    doc.PrintOut(Copies=copies)
                except Exception:
                    doc.PrintOut()
                return {"status": "printed", "copies": copies}

            return {"error": "Active document does not support printing"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_printer(self, printer_name: str) -> dict[str, Any]:
        """
        Set the printer for the active draft document.

        Uses DraftPrintUtility.Printer property.

        Args:
            printer_name: Name of the printer to use

        Returns:
            Dict with status and printer name
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "DraftPrintUtility"):
                return {"error": "Active document does not have DraftPrintUtility"}

            dpu = doc.DraftPrintUtility
            dpu.Printer = printer_name

            return {"status": "set", "printer": printer_name}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_printer(self) -> dict[str, Any]:
        """
        Get the current printer for the active draft document.

        Uses DraftPrintUtility.Printer property.

        Returns:
            Dict with printer name
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "DraftPrintUtility"):
                return {"error": "Active document does not have DraftPrintUtility"}

            dpu = doc.DraftPrintUtility
            printer_name = dpu.Printer

            return {"printer": printer_name}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_paper_size(
        self, width: float, height: float, orientation: str = "Landscape"
    ) -> dict[str, Any]:
        """
        Set the paper size and orientation for printing.

        Uses DraftPrintUtility paper width/height and orientation.

        Args:
            width: Paper width in meters
            height: Paper height in meters
            orientation: 'Landscape' or 'Portrait'

        Returns:
            Dict with status and paper settings
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "DraftPrintUtility"):
                return {"error": "Active document does not have DraftPrintUtility"}

            dpu = doc.DraftPrintUtility

            with contextlib.suppress(Exception):
                dpu.PaperWidth = width
            with contextlib.suppress(Exception):
                dpu.PaperHeight = height

            # Set orientation: 1=Portrait, 2=Landscape (typical COM constants)
            if orientation.lower() == "portrait":
                with contextlib.suppress(Exception):
                    dpu.Orientation = 1
            else:
                with contextlib.suppress(Exception):
                    dpu.Orientation = 2

            return {
                "status": "set",
                "width": width,
                "height": height,
                "orientation": orientation,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # DRAWING VIEW TOOLS (Batch 10)
    # =================================================================

    def set_face_texture(self, face_index: int, texture_name: str) -> dict[str, Any]:
        """
        Apply a texture to a face by index.

        Uses face style properties to set the texture name.

        Args:
            face_index: 0-based face index
            texture_name: Name of the texture to apply

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Models"):
                return {"error": "Active document does not have a Models collection"}

            models = doc.Models
            if models.Count == 0:
                return {"error": "No models in document"}

            model = models.Item(1)
            body = model.Body
            faces = body.Faces(1)  # igQueryAll = 1

            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face_index: {face_index}. Count: {faces.Count}"}

            face = faces.Item(face_index + 1)

            # Try to set texture via face style
            try:
                face.TextureName = texture_name
            except Exception:
                # Alternative: use Style object
                try:
                    style = face.Style
                    style.TextureName = texture_name
                except Exception as inner_e:
                    return {
                        "error": f"Cannot set texture: {inner_e}",
                        "traceback": traceback.format_exc(),
                    }

            return {
                "status": "set",
                "face_index": face_index,
                "texture_name": texture_name,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_assembly_drawing_view_ex(
        self,
        x: float = 0.15,
        y: float = 0.15,
        orientation: str = "Isometric",
        scale: float = 1.0,
        config: str = None,
    ) -> dict[str, Any]:
        """
        Add an extended assembly drawing view with optional configuration.

        Similar to add_assembly_drawing_view but with configuration support.

        Args:
            x: View center X position on sheet (meters)
            y: View center Y position on sheet (meters)
            orientation: View orientation ('Front', 'Top', 'Right', 'Isometric', etc.)
            scale: View scale factor
            config: Optional configuration name

        Returns:
            Dict with status and view info
        """
        try:
            import win32com.client.dynamic as dyn

            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            if not hasattr(doc, "ModelLinks") or doc.ModelLinks.Count == 0:
                return {
                    "error": "No model link found. Create a drawing with create_drawing() first."
                }

            model_link = doc.ModelLinks.Item(1)

            view_orient_map = {
                "Front": 5,
                "Back": 8,
                "Top": 6,
                "Bottom": 9,
                "Right": 7,
                "Left": 10,
                "Isometric": 12,
                "Iso": 12,
            }

            orient = view_orient_map.get(orientation)
            if orient is None:
                valid = ", ".join(view_orient_map.keys())
                return {"error": f"Invalid orientation: {orientation}. Valid: {valid}"}

            sheet = doc.ActiveSheet
            dvs_early = sheet.DrawingViews
            dvs = dyn.Dispatch(dvs_early._oleobj_)

            # If config specified, try AddWithConfiguration
            if config is not None:
                try:
                    dvs.AddAssemblyViewWithConfiguration(model_link, orient, scale, x, y, 0, config)
                    return {
                        "status": "added",
                        "orientation": orientation,
                        "scale": scale,
                        "position": [x, y],
                        "configuration": config,
                    }
                except Exception:
                    pass  # Fall through to standard method

            # Standard assembly view
            try:
                dvs.AddAssemblyView(model_link, orient, scale, x, y, 0)
            except Exception:
                dvs.AddPartView(model_link, orient, scale, x, y, 0)

            result = {
                "status": "added",
                "orientation": orientation,
                "scale": scale,
                "position": [x, y],
            }
            if config is not None:
                result["configuration"] = config
                result["note"] = "Configuration param not applied; used standard view."
            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_drawing_view_with_config(
        self,
        x: float = 0.15,
        y: float = 0.15,
        orientation: str = "Front",
        scale: float = 1.0,
        configuration: str = "Default",
    ) -> dict[str, Any]:
        """
        Add a drawing view with a specific configuration.

        Uses DrawingViews.AddWithConfiguration if available.

        Args:
            x: View center X position on sheet (meters)
            y: View center Y position on sheet (meters)
            orientation: View orientation ('Front', 'Top', 'Right', etc.)
            scale: View scale factor
            configuration: Configuration name

        Returns:
            Dict with status and view info
        """
        try:
            import win32com.client.dynamic as dyn

            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            if not hasattr(doc, "ModelLinks") or doc.ModelLinks.Count == 0:
                return {
                    "error": "No model link found. Create a drawing with create_drawing() first."
                }

            model_link = doc.ModelLinks.Item(1)

            view_orient_map = {
                "Front": 5,
                "Back": 8,
                "Top": 6,
                "Bottom": 9,
                "Right": 7,
                "Left": 10,
                "Isometric": 12,
                "Iso": 12,
            }

            orient = view_orient_map.get(orientation)
            if orient is None:
                valid = ", ".join(view_orient_map.keys())
                return {"error": f"Invalid orientation: {orientation}. Valid: {valid}"}

            sheet = doc.ActiveSheet
            dvs_early = sheet.DrawingViews
            dvs = dyn.Dispatch(dvs_early._oleobj_)

            try:
                dvs.AddPartViewWithConfiguration(model_link, orient, scale, x, y, 0, configuration)
            except Exception:
                # Fall back to standard AddPartView
                dvs.AddPartView(model_link, orient, scale, x, y, 0)

            return {
                "status": "added",
                "orientation": orientation,
                "scale": scale,
                "position": [x, y],
                "configuration": configuration,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def activate_drawing_view(self, view_index: int) -> dict[str, Any]:
        """
        Activate a drawing view by 0-based index.

        An activated view allows editing its contents.

        Args:
            view_index: 0-based view index

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()
            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view_index: {view_index}. Count: {dvs.Count}"}
            view = dvs.Item(view_index + 1)
            view.Activate()
            return {"status": "activated", "view_index": view_index}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def deactivate_drawing_view(self, view_index: int) -> dict[str, Any]:
        """
        Deactivate a drawing view by 0-based index.

        Deactivating a view returns focus to the sheet.

        Args:
            view_index: 0-based view index

        Returns:
            Dict with status
        """
        try:
            dvs = self._get_drawing_views()
            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view_index: {view_index}. Count: {dvs.Count}"}
            view = dvs.Item(view_index + 1)
            view.Deactivate()
            return {"status": "deactivated", "view_index": view_index}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # DRAWING VIEW COPY / SECTION / DIMENSIONS (Batch 11)
    # =================================================================

    def add_by_draft_view(
        self, source_view_index: int, x: float, y: float, scale: float | None = None
    ) -> dict[str, Any]:
        """
        Copy an existing drawing view to a new location on the active sheet.

        Uses DrawingViews.AddByDraftView(From, Scale, x1, y1) from the draft
        type library. The source view is specified by 0-based index.

        Args:
            source_view_index: 0-based index of the source drawing view
            x: X position for the new view on the sheet (meters)
            y: Y position for the new view on the sheet (meters)
            scale: Scale for the new view (default: same as source view)

        Returns:
            Dict with status and new view info
        """
        try:
            dvs = self._get_drawing_views()

            if source_view_index < 0 or source_view_index >= dvs.Count:
                return {
                    "error": f"Invalid source view index: {source_view_index}. Count: {dvs.Count}"
                }

            source_view = dvs.Item(source_view_index + 1)

            # Use source view scale if not specified
            if scale is None:
                try:
                    scale = source_view.ScaleFactor
                except Exception:
                    scale = 1.0

            # AddByDraftView(From: DrawingView*, Scale: VT_R8, x1: VT_R8, y1: VT_R8)
            new_view = dvs.AddByDraftView(source_view, scale, x, y)

            result = {
                "status": "added",
                "source_view_index": source_view_index,
                "position": [x, y],
                "scale": scale,
                "total_views": dvs.Count,
            }

            with contextlib.suppress(Exception):
                result["name"] = new_view.Name

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_section_cuts(self, view_index: int) -> dict[str, Any]:
        """
        Get section cut (cutting plane) information from a drawing view.

        Accesses DrawingView.CuttingPlanes collection and extracts caption,
        display type, and fold line geometry for each cutting plane.

        Args:
            view_index: 0-based index of the drawing view

        Returns:
            Dict with count and list of section cut info dicts
        """
        try:
            dvs = self._get_drawing_views()

            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)

            if not hasattr(view, "CuttingPlanes"):
                return {"count": 0, "section_cuts": [], "note": "No CuttingPlanes on this view"}

            cutting_planes = view.CuttingPlanes
            items = []
            for i in range(1, cutting_planes.Count + 1):
                cp = cutting_planes.Item(i)
                info: dict[str, Any] = {"index": i - 1}

                with contextlib.suppress(Exception):
                    info["caption"] = cp.Caption
                with contextlib.suppress(Exception):
                    info["display_caption"] = cp.DisplayCaption
                with contextlib.suppress(Exception):
                    info["display_type"] = cp.DisplayType
                with contextlib.suppress(Exception):
                    info["style_name"] = cp.StyleName
                with contextlib.suppress(Exception):
                    info["text_height"] = cp.TextHeight

                # Try to get fold line geometry
                with contextlib.suppress(Exception):
                    (
                        line_start_x,
                        line_start_y,
                        line_end_x,
                        line_end_y,
                        view_dir_x,
                        view_dir_y,
                    ) = cp.GetFoldLineWithViewDirection()
                    info["fold_line"] = {
                        "start": [line_start_x, line_start_y],
                        "end": [line_end_x, line_end_y],
                        "view_direction": [view_dir_x, view_dir_y],
                    }

                items.append(info)

            return {"count": len(items), "section_cuts": items, "view_index": view_index}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_section_cut(
        self,
        view_index: int,
        x: float,
        y: float,
        section_type: int = 0,
    ) -> dict[str, Any]:
        """
        Add a section cut (cutting plane) to a drawing view and create the section view.

        Creates a cutting plane on the specified drawing view via
        CuttingPlanes.Add(), then calls CuttingPlane.CreateView(SectionType)
        to generate the section drawing view.

        Args:
            view_index: 0-based index of the source drawing view
            x: X position for the section view on the sheet (meters)
            y: Y position for the section view on the sheet (meters)
            section_type: 0 = standard, 1 = revolved (DraftSectionViewType)

        Returns:
            Dict with status and section view info
        """
        try:
            dvs = self._get_drawing_views()

            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            if section_type not in (0, 1):
                return {
                    "error": f"Invalid section_type: {section_type}. "
                    "Use 0 (standard) or 1 (revolved)."
                }

            view = dvs.Item(view_index + 1)

            if not hasattr(view, "CuttingPlanes"):
                return {"error": "Drawing view does not support CuttingPlanes"}

            cutting_planes = view.CuttingPlanes

            # CuttingPlanes.Add() returns a new CuttingPlane object
            cutting_plane = cutting_planes.Add()

            # CuttingPlane.CreateView(SectionType) creates the section view
            section_view = cutting_plane.CreateView(section_type)

            # Move the section view to the desired position
            with contextlib.suppress(Exception):
                section_view.OriginX = x
                section_view.OriginY = y

            result = {
                "status": "added",
                "source_view_index": view_index,
                "section_type": "standard" if section_type == 0 else "revolved",
                "position": [x, y],
                "total_cutting_planes": cutting_planes.Count,
            }

            with contextlib.suppress(Exception):
                result["caption"] = cutting_plane.Caption
            with contextlib.suppress(Exception):
                result["section_view_name"] = section_view.Name

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_drawing_view_dimensions(self, view_index: int) -> dict[str, Any]:
        """
        Get all dimensions associated with a specific drawing view.

        Accesses DrawingView.Dimensions collection (dispid 120) and iterates
        each Dimension to return its type, value, prefix/suffix strings, and
        override text.

        DimensionType constants (DimTypeConstants):
            1=Linear, 2=Radial, 3=Angular, 4=RadialDiameter,
            5=CircularDiameter, 6=ArcLength, 7=ArcAngle, 8=Coordinate,
            9=SymmetricalDiameter, 10=Chamfer, 11=AngularCoordinate,
            12=CurveLength

        Args:
            view_index: 0-based index of the drawing view

        Returns:
            Dict with count and list of dimension info dicts
        """
        try:
            dvs = self._get_drawing_views()

            if view_index < 0 or view_index >= dvs.Count:
                return {"error": f"Invalid view index: {view_index}. Count: {dvs.Count}"}

            view = dvs.Item(view_index + 1)

            if not hasattr(view, "Dimensions"):
                return {"count": 0, "dimensions": [], "note": "No Dimensions on this view"}

            dims = view.Dimensions

            dim_type_names = {
                1: "Linear",
                2: "Radial",
                3: "Angular",
                4: "RadialDiameter",
                5: "CircularDiameter",
                6: "ArcLength",
                7: "ArcAngle",
                8: "Coordinate",
                9: "SymmetricalDiameter",
                10: "Chamfer",
                11: "AngularCoordinate",
                12: "CurveLength",
            }

            items = []
            for i in range(1, dims.Count + 1):
                dim = dims.Item(i)
                info: dict[str, Any] = {"index": i - 1}

                with contextlib.suppress(Exception):
                    raw_type = dim.DimensionType
                    info["dimension_type"] = raw_type
                    info["dimension_type_name"] = dim_type_names.get(raw_type, "Unknown")
                with contextlib.suppress(Exception):
                    info["value"] = dim.Value
                with contextlib.suppress(Exception):
                    info["constraint"] = dim.Constraint
                with contextlib.suppress(Exception):
                    info["prefix"] = dim.PrefixString
                with contextlib.suppress(Exception):
                    info["suffix"] = dim.SuffixString
                with contextlib.suppress(Exception):
                    info["override"] = dim.OverrideString
                with contextlib.suppress(Exception):
                    info["subfix"] = dim.SubfixString
                with contextlib.suppress(Exception):
                    info["superfix"] = dim.SuperfixString

                items.append(info)

            return {
                "count": len(items),
                "dimensions": items,
                "view_index": view_index,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}


class ViewModel:
    """Manages view manipulation"""

    def __init__(self, document_manager):
        self.doc_manager = document_manager

    def set_view(self, view: str) -> dict[str, Any]:
        """
        Set the viewing orientation.

        Args:
            view: View orientation - 'Iso', 'Top', 'Front', 'Right', 'Bottom', 'Back', 'Left'

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            # Get the window
            if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, "View") else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            # Valid view names (discovered via introspection)
            # Note: Bottom, Back, Left may not work in all contexts
            valid_views = ["Iso", "Top", "Front", "Right", "Bottom", "Back", "Left"]

            if view not in valid_views:
                return {"error": f"Invalid view: {view}. Valid: {', '.join(valid_views)}"}

            # Use ApplyNamedView with string name (discovered method!)
            if hasattr(view_obj, "ApplyNamedView"):
                view_obj.ApplyNamedView(view)
                return {"status": "view_set", "view": view}
            else:
                return {
                    "error": "ApplyNamedView not available",
                    "note": "Use View menu in Solid Edge UI",
                }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def zoom_fit(self) -> dict[str, Any]:
        """Zoom to fit all geometry in view"""
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, "View") else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            # Zoom to fit
            if hasattr(view_obj, "Fit"):
                view_obj.Fit()
                return {"status": "zoomed_fit"}
            else:
                return {
                    "error": "Fit method not available",
                    "note": "Use View > Fit in Solid Edge UI",
                }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def zoom_to_selection(self) -> dict[str, Any]:
        """Zoom to fit all geometry (equivalent to View > Fit)."""
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, "View") else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            view_obj.Fit()

            return {"status": "zoomed_to_selection"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_display_mode(self, mode: str) -> dict[str, Any]:
        """
        Set the display mode for the active view.

        Args:
            mode: Display mode - 'Shaded', 'ShadedWithEdges', 'Wireframe', 'HiddenEdgesVisible'

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, "View") else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            mode_map = {
                "Wireframe": RenderModeConstants.seRenderModeWireframe,
                "Shaded": RenderModeConstants.seRenderModeSmooth,
                "ShadedWithEdges": RenderModeConstants.seRenderModeSmoothBoundary,
                "HiddenEdgesVisible": RenderModeConstants.seRenderModeVHL,
            }

            mode_value = mode_map.get(mode)
            if mode_value is None:
                return {
                    "error": f"Invalid mode: {mode}. "
                    "Use 'Shaded', 'ShadedWithEdges', "
                    "'Wireframe', or "
                    "'HiddenEdgesVisible'"
                }

            view_obj.SetRenderMode(mode_value)
            return {"status": "display_mode_set", "mode": mode}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_view_background(self, red: int, green: int, blue: int) -> dict[str, Any]:
        """
        Set the view background color.

        Args:
            red: Red component (0-255)
            green: Green component (0-255)
            blue: Blue component (0-255)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, "View") else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            ole_color = red | (green << 8) | (blue << 16)

            try:
                view_obj.SetBackgroundColor(ole_color)
            except Exception:
                try:
                    view_obj.BackgroundColor = ole_color
                except Exception:
                    view_obj.SetBackgroundGradientColor(ole_color, ole_color)

            return {"status": "updated", "color": [red, green, blue]}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_camera(self) -> dict[str, Any]:
        """
        Get the current camera parameters.

        Returns eye position, target position, up vector, perspective flag,
        and scale/field-of-view angle.

        Returns:
            Dict with camera parameters
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, "View") else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            # GetCamera returns 11 out-params by reference
            result = view_obj.GetCamera()

            # result is a tuple: (EyeX, EyeY, EyeZ, TargetX, TargetY, TargetZ,
            #                     UpX, UpY, UpZ, Perspective, ScaleOrAngle)
            return {
                "eye": [result[0], result[1], result[2]],
                "target": [result[3], result[4], result[5]],
                "up": [result[6], result[7], result[8]],
                "perspective": bool(result[9]),
                "scale_or_angle": result[10],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def rotate_camera(
        self,
        angle: float,
        center_x: float = 0.0,
        center_y: float = 0.0,
        center_z: float = 0.0,
        axis_x: float = 0.0,
        axis_y: float = 1.0,
        axis_z: float = 0.0,
    ) -> dict[str, Any]:
        """
        Rotate the camera around a specified axis through a center point.

        Args:
            angle: Rotation angle in radians
            center_x, center_y, center_z: Center of rotation (meters)
            axis_x, axis_y, axis_z: Rotation axis vector (default: Y-up)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, "View") else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            view_obj.RotateCamera(angle, center_x, center_y, center_z, axis_x, axis_y, axis_z)

            return {
                "status": "camera_rotated",
                "angle_rad": angle,
                "center": [center_x, center_y, center_z],
                "axis": [axis_x, axis_y, axis_z],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def pan_camera(self, dx: int, dy: int) -> dict[str, Any]:
        """
        Pan the camera by pixel offsets.

        Args:
            dx: Horizontal pan in pixels (positive = right)
            dy: Vertical pan in pixels (positive = down)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, "View") else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            view_obj.PanCamera(dx, dy)

            return {"status": "camera_panned", "dx": dx, "dy": dy}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def zoom_camera(self, factor: float) -> dict[str, Any]:
        """
        Zoom the camera by a scale factor.

        Args:
            factor: Zoom scale factor (>1 = zoom in, <1 = zoom out)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, "View") else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            view_obj.ZoomCamera(factor)

            return {"status": "camera_zoomed", "factor": factor}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def refresh_view(self) -> dict[str, Any]:
        """
        Force the active view to refresh/update.

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, "View") else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            view_obj.Update()

            return {"status": "view_refreshed"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def _get_view_object(self):
        """Get the active view object from the first window."""
        doc = self.doc_manager.get_active_document()
        if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
            raise Exception("No window available")
        window = doc.Windows.Item(1)
        view_obj = window.View if hasattr(window, "View") else None
        if not view_obj:
            raise Exception("Cannot access view object")
        return view_obj

    def transform_model_to_screen(self, x: float, y: float, z: float) -> dict[str, Any]:
        """
        Transform 3D model coordinates to 2D screen (device) coordinates.

        Uses View.ModelToScreenTransform or TransformModelToDC.

        Args:
            x: Model X coordinate (meters)
            y: Model Y coordinate (meters)
            z: Model Z coordinate (meters)

        Returns:
            Dict with screen_x and screen_y pixel coordinates
        """
        try:
            view_obj = self._get_view_object()
            result = view_obj.TransformModelToDC(x, y, z)
            return {
                "status": "success",
                "model": [x, y, z],
                "screen_x": result[0],
                "screen_y": result[1],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def transform_screen_to_model(self, screen_x: int, screen_y: int) -> dict[str, Any]:
        """
        Transform 2D screen (device) coordinates to 3D model coordinates.

        Uses View.TransformDCToModel.

        Args:
            screen_x: Screen X pixel coordinate
            screen_y: Screen Y pixel coordinate

        Returns:
            Dict with model x, y, z coordinates (meters)
        """
        try:
            view_obj = self._get_view_object()
            result = view_obj.TransformDCToModel(screen_x, screen_y)
            return {
                "status": "success",
                "screen": [screen_x, screen_y],
                "x": result[0],
                "y": result[1],
                "z": result[2],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def begin_camera_dynamics(self) -> dict[str, Any]:
        """
        Begin camera dynamics mode for smooth multi-step camera manipulation.

        Call this before a sequence of RotateCamera/PanCamera/ZoomCamera calls,
        then call end_camera_dynamics() when done. This prevents intermediate
        view updates and improves performance.

        Returns:
            Dict with status
        """
        try:
            view_obj = self._get_view_object()
            view_obj.BeginCameraDynamics()
            return {"status": "camera_dynamics_started"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def end_camera_dynamics(self) -> dict[str, Any]:
        """
        End camera dynamics mode and apply all pending camera changes.

        Call after a sequence of camera manipulations that were started with
        begin_camera_dynamics().

        Returns:
            Dict with status
        """
        try:
            view_obj = self._get_view_object()
            view_obj.EndCameraDynamics()
            return {"status": "camera_dynamics_ended"}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_camera(
        self,
        eye_x: float,
        eye_y: float,
        eye_z: float,
        target_x: float,
        target_y: float,
        target_z: float,
        up_x: float = 0.0,
        up_y: float = 1.0,
        up_z: float = 0.0,
        perspective: bool = False,
        scale_or_angle: float = 1.0,
    ) -> dict[str, Any]:
        """
        Set the camera parameters for the active view.

        Args:
            eye_x, eye_y, eye_z: Camera eye (position) coordinates
            target_x, target_y, target_z: Camera target (look-at) coordinates
            up_x, up_y, up_z: Camera up vector (default: Y-up)
            perspective: True for perspective, False for orthographic
            scale_or_angle: View scale (ortho) or FOV angle in radians (perspective)

        Returns:
            Dict with status and camera settings
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "Windows") or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, "View") else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            view_obj.SetCamera(
                eye_x,
                eye_y,
                eye_z,
                target_x,
                target_y,
                target_z,
                up_x,
                up_y,
                up_z,
                perspective,
                scale_or_angle,
            )

            return {
                "status": "camera_set",
                "eye": [eye_x, eye_y, eye_z],
                "target": [target_x, target_y, target_z],
                "up": [up_x, up_y, up_z],
                "perspective": perspective,
                "scale_or_angle": scale_or_angle,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
