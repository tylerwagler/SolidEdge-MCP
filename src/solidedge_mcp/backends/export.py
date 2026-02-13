"""
Solid Edge Export and Visualization Operations

Handles exporting to various formats and creating drawings.
"""

from typing import Dict, Any, Optional, List
import os
import traceback
from .constants import DrawingViewOrientationConstants, RenderModeConstants


class ExportManager:
    """Manages export and visualization operations"""

    def __init__(self, document_manager):
        self.doc_manager = document_manager

    def export_to_step(self, file_path: str) -> Dict[str, Any]:
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
            if not file_path.lower().endswith(('.step', '.stp')):
                file_path += '.step'

            # Save as STEP
            doc.SaveAs(file_path)

            return {
                "status": "exported",
                "format": "STEP",
                "path": file_path,
                "size_bytes": os.path.getsize(file_path) if os.path.exists(file_path) else 0
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def export_to_stl(self, file_path: str, quality: str = "Medium") -> Dict[str, Any]:
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
            if not file_path.lower().endswith('.stl'):
                file_path += '.stl'

            # Quality mapping (actual values depend on Solid Edge version)
            quality_map = {
                "Low": 0.01,
                "Medium": 0.001,
                "High": 0.0001
            }
            tolerance = quality_map.get(quality, 0.001)

            # Save as STL
            # Note: Actual method may vary by Solid Edge version
            try:
                doc.SaveAs(file_path)
            except:
                # Alternative export method
                if hasattr(doc, 'SaveAsJT'):
                    # Some versions use different export methods
                    pass

            return {
                "status": "exported",
                "format": "STL",
                "path": file_path,
                "quality": quality,
                "size_bytes": os.path.getsize(file_path) if os.path.exists(file_path) else 0
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def export_to_iges(self, file_path: str) -> Dict[str, Any]:
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
            if not file_path.lower().endswith(('.iges', '.igs')):
                file_path += '.iges'

            # Save as IGES
            doc.SaveAs(file_path)

            return {
                "status": "exported",
                "format": "IGES",
                "path": file_path,
                "size_bytes": os.path.getsize(file_path) if os.path.exists(file_path) else 0
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def export_to_pdf(self, file_path: str) -> Dict[str, Any]:
        """Export drawing to PDF"""
        try:
            doc = self.doc_manager.get_active_document()

            if not file_path.lower().endswith('.pdf'):
                file_path += '.pdf'

            # PDF export typically works for draft documents
            doc.SaveAs(file_path)

            return {
                "status": "exported",
                "format": "PDF",
                "path": file_path
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_drawing(self, template: Optional[str] = None,
                      views: Optional[List[str]] = None) -> Dict[str, Any]:
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
            try:
                source_path = source_doc.FullName
            except Exception:
                pass

            if not source_path:
                return {"error": "Active document must be saved before creating a drawing. Use save_document() first."}

            if views is None:
                views = ['Front', 'Top', 'Right', 'Isometric']

            view_orient_map = {
                'Front': DrawingViewOrientationConstants.Front,
                'Back': DrawingViewOrientationConstants.Back,
                'Top': DrawingViewOrientationConstants.Top,
                'Bottom': DrawingViewOrientationConstants.Bottom,
                'Right': DrawingViewOrientationConstants.Right,
                'Left': DrawingViewOrientationConstants.Left,
                'Isometric': DrawingViewOrientationConstants.Isometric,
                'Iso': DrawingViewOrientationConstants.Isometric,
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
                (0.10, 0.15),   # View 1
                (0.10, 0.06),   # View 2
                (0.22, 0.15),   # View 3
                (0.22, 0.06),   # View 4
            ]

            views_added = []
            for i, view_name in enumerate(views[:4]):
                orient = view_orient_map.get(view_name)
                if orient is None:
                    continue

                x, y = positions[i] if i < len(positions) else (0.15 + i * 0.08, 0.10)

                try:
                    dv = dvs.AddPartView(model_link, orient, 1.0, x, y, 0)
                    views_added.append(view_name)
                except Exception:
                    # Try the generic Add method as fallback
                    try:
                        dv = dvs.Add(model_link, orient, 1.0, x, y)
                        views_added.append(view_name)
                    except Exception:
                        pass

            return {
                "status": "created",
                "type": "drawing",
                "draft_name": draft_doc.Name if hasattr(draft_doc, 'Name') else "Draft",
                "model_link": source_path,
                "views_requested": views,
                "views_added": views_added,
                "total_views": len(views_added)
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def add_draft_sheet(self) -> Dict[str, Any]:
        """
        Add a new sheet to the active draft document.

        The active document must be a Draft document. Creates a new sheet
        and activates it.

        Returns:
            Dict with status and new sheet info
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, 'Sheets'):
                return {"error": "Active document is not a draft document. Create a drawing first."}

            sheets = doc.Sheets
            old_count = sheets.Count

            sheet = sheets.AddSheet()
            sheet.Activate()

            return {
                "status": "added",
                "sheet_number": sheets.Count,
                "total_sheets": sheets.Count,
                "name": sheet.Name if hasattr(sheet, 'Name') else f"Sheet {sheets.Count}"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def add_assembly_drawing_view(
        self,
        x: float = 0.15,
        y: float = 0.15,
        orientation: str = "Isometric",
        scale: float = 1.0
    ) -> Dict[str, Any]:
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

            if not hasattr(doc, 'Sheets'):
                return {"error": "Active document is not a draft document"}

            # Get model link
            if not hasattr(doc, 'ModelLinks') or doc.ModelLinks.Count == 0:
                return {"error": "No model link found. Create a drawing with create_drawing() first."}

            model_link = doc.ModelLinks.Item(1)

            # View orientation constants
            view_orient_map = {
                'Front': 5, 'Back': 8, 'Top': 6, 'Bottom': 9,
                'Right': 7, 'Left': 10, 'Isometric': 12, 'Iso': 12,
            }

            orient = view_orient_map.get(orientation)
            if orient is None:
                return {"error": f"Invalid orientation: {orientation}. Valid: {', '.join(view_orient_map.keys())}"}

            sheet = doc.ActiveSheet
            dvs_early = sheet.DrawingViews
            dvs = dyn.Dispatch(dvs_early._oleobj_)

            # seAssemblyDesignedView = 0
            try:
                dv = dvs.AddAssemblyView(model_link, orient, scale, x, y, 0)
            except Exception:
                # Fall back to AddPartView if AddAssemblyView not available
                dv = dvs.AddPartView(model_link, orient, scale, x, y, 0)

            return {
                "status": "added",
                "orientation": orientation,
                "scale": scale,
                "position": [x, y]
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def create_parts_list(self, auto_balloon: bool = True,
                          x: float = 0.15, y: float = 0.25) -> Dict[str, Any]:
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

            if not hasattr(doc, 'Sheets'):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet

            # Get the first drawing view on the sheet
            dvs = sheet.DrawingViews
            if dvs.Count == 0:
                return {"error": "No drawing views on active sheet. Add a view first."}

            dv = dvs.Item(1)

            # Get PartsLists collection
            parts_lists = sheet.PartsLists if hasattr(sheet, 'PartsLists') else None
            if parts_lists is None:
                # Try from document level
                parts_lists = doc.PartsLists if hasattr(doc, 'PartsLists') else None

            if parts_lists is None:
                return {"error": "PartsLists collection not available"}

            # Add parts list
            # Parameters: DrawingView, SavedSettings, AutoBalloon, CreatePartsList
            # AutoBalloon: 0=No, 1=Yes; CreatePartsList: 0=No, 1=Yes
            parts_list = parts_lists.Add(dv, "", 1 if auto_balloon else 0, 1)

            return {
                "status": "created",
                "auto_balloon": auto_balloon,
                "total_parts_lists": parts_lists.Count
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def capture_screenshot(self, file_path: str, width: int = 1920, height: int = 1080) -> Dict[str, Any]:
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
            if not any(file_path.lower().endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.bmp']):
                file_path += '.png'

            # Get the window and view
            if not hasattr(doc, 'Windows') or doc.Windows.Count == 0:
                return {"error": "No window available for screenshot"}

            window = doc.Windows.Item(1)
            view = window.View if hasattr(window, 'View') else None

            if not view:
                return {"error": "Cannot access view object"}

            # View.SaveAsImage(Filename, Width, Height)
            view.SaveAsImage(file_path, width, height)

            return {
                "status": "captured",
                "path": file_path,
                "dimensions": [width, height],
                "size_bytes": os.path.getsize(file_path) if os.path.exists(file_path) else 0
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }


    def export_to_dxf(self, file_path: str) -> Dict[str, Any]:
        """Export to DXF format"""
        try:
            doc = self.doc_manager.get_active_document()

            if not file_path.lower().endswith('.dxf'):
                file_path += '.dxf'

            doc.SaveAs(file_path)

            return {
                "status": "exported",
                "format": "DXF",
                "path": file_path,
                "size_bytes": os.path.getsize(file_path) if os.path.exists(file_path) else 0
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def export_to_parasolid(self, file_path: str) -> Dict[str, Any]:
        """Export to Parasolid format (X_T or X_B)"""
        try:
            doc = self.doc_manager.get_active_document()

            if not (file_path.lower().endswith('.x_t') or file_path.lower().endswith('.x_b')):
                file_path += '.x_t'

            doc.SaveAs(file_path)

            return {
                "status": "exported",
                "format": "Parasolid",
                "path": file_path,
                "size_bytes": os.path.getsize(file_path) if os.path.exists(file_path) else 0
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def export_to_jt(self, file_path: str) -> Dict[str, Any]:
        """Export to JT format"""
        try:
            doc = self.doc_manager.get_active_document()

            if not file_path.lower().endswith('.jt'):
                file_path += '.jt'

            doc.SaveAs(file_path)

            return {
                "status": "exported",
                "format": "JT",
                "path": file_path,
                "size_bytes": os.path.getsize(file_path) if os.path.exists(file_path) else 0
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def export_flat_dxf(self, file_path: str) -> Dict[str, Any]:
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

            if not file_path.lower().endswith('.dxf'):
                file_path += '.dxf'

            # Access FlatPatternModels collection (sheet metal only)
            if not hasattr(doc, 'FlatPatternModels'):
                return {"error": "Active document is not a sheet metal document. FlatPatternModels not available."}

            flat_models = doc.FlatPatternModels

            # SaveAsFlatDXFEx(filename, face, edge, vertex, useFlatPattern)
            # Pass None for face/edge/vertex to export all
            flat_models.SaveAsFlatDXFEx(file_path, None, None, None, True)

            return {
                "status": "exported",
                "format": "Flat DXF",
                "path": file_path,
                "size_bytes": os.path.getsize(file_path) if os.path.exists(file_path) else 0
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def add_text_box(self, x: float, y: float, text: str, height: float = 0.005) -> Dict[str, Any]:
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

            if not hasattr(doc, 'Sheets'):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet

            # TextBoxes.Add(x, y, 0) for 2D sheets
            text_boxes = sheet.TextBoxes
            text_box = text_boxes.Add(x, y, 0)

            # Set the text content
            text_box.Text = text

            # Set text height if possible
            try:
                text_box.TextHeight = height
            except Exception:
                pass

            return {
                "status": "added",
                "type": "text_box",
                "text": text,
                "position": [x, y]
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def add_leader(self, x1: float, y1: float, x2: float, y2: float, text: str = "") -> Dict[str, Any]:
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

            if not hasattr(doc, 'Sheets'):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet

            leaders = sheet.Leaders
            leader = leaders.Add(x1, y1, 0, x2, y2, 0)

            if text:
                try:
                    leader.Text = text
                except Exception:
                    pass

            return {
                "status": "added",
                "type": "leader",
                "start": [x1, y1],
                "end": [x2, y2],
                "text": text
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def add_dimension(self, x1: float, y1: float, x2: float, y2: float,
                      dim_x: float = None, dim_y: float = None) -> Dict[str, Any]:
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

            if not hasattr(doc, 'Sheets'):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet

            # Default dimension text position to midpoint offset
            if dim_x is None:
                dim_x = (x1 + x2) / 2
            if dim_y is None:
                dim_y = max(y1, y2) + 0.02  # 20mm above

            dimensions = sheet.Dimensions
            dim = dimensions.AddLength(x1, y1, 0, x2, y2, 0, dim_x, dim_y, 0)

            return {
                "status": "added",
                "type": "dimension",
                "point1": [x1, y1],
                "point2": [x2, y2],
                "text_position": [dim_x, dim_y]
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def add_balloon(self, x: float, y: float, text: str = "",
                    leader_x: float = None, leader_y: float = None) -> Dict[str, Any]:
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

            if not hasattr(doc, 'Sheets'):
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
                    try:
                        balloon.Text = text
                    except Exception:
                        pass

            return {
                "status": "added",
                "type": "balloon",
                "position": [x, y],
                "text": text
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def add_note(self, x: float, y: float, text: str, height: float = 0.005) -> Dict[str, Any]:
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

            if not hasattr(doc, 'Sheets'):
                return {"error": "Active document is not a draft document"}

            sheet = doc.ActiveSheet

            # Try TextBoxes first (most reliable)
            text_boxes = sheet.TextBoxes
            text_box = text_boxes.Add(x, y, 0)
            text_box.Text = text

            try:
                text_box.TextHeight = height
            except Exception:
                pass

            return {
                "status": "added",
                "type": "note",
                "position": [x, y],
                "text": text,
                "height": height
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_sheet_info(self) -> Dict[str, Any]:
        """
        Get information about the active draft sheet.

        Returns:
            Dict with sheet name, size, scale, and counts of drawing objects
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, 'Sheets'):
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
            try:
                info["background"] = sheet.Background
            except Exception:
                pass

            # Count drawing objects
            try:
                info["drawing_views"] = sheet.DrawingViews.Count
            except Exception:
                pass
            try:
                info["text_boxes"] = sheet.TextBoxes.Count
            except Exception:
                pass
            try:
                info["leaders"] = sheet.Leaders.Count
            except Exception:
                pass
            try:
                info["dimensions"] = sheet.Dimensions.Count
            except Exception:
                pass

            return info
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def activate_sheet(self, sheet_index: int) -> Dict[str, Any]:
        """
        Activate a specific draft sheet by index.

        Args:
            sheet_index: 0-based sheet index

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, 'Sheets'):
                return {"error": "Active document is not a draft document"}

            sheets = doc.Sheets
            if sheet_index < 0 or sheet_index >= sheets.Count:
                return {"error": f"Invalid sheet index: {sheet_index}. Count: {sheets.Count}"}

            sheet = sheets.Item(sheet_index + 1)  # COM is 1-indexed
            sheet.Activate()

            return {
                "status": "activated",
                "sheet_name": sheet.Name,
                "sheet_index": sheet_index
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def rename_sheet(self, sheet_index: int, new_name: str) -> Dict[str, Any]:
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

            if not hasattr(doc, 'Sheets'):
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
                "sheet_index": sheet_index
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def delete_sheet(self, sheet_index: int) -> Dict[str, Any]:
        """
        Delete a draft sheet.

        Args:
            sheet_index: 0-based sheet index

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, 'Sheets'):
                return {"error": "Active document is not a draft document"}

            sheets = doc.Sheets
            if sheets.Count <= 1:
                return {"error": "Cannot delete the last sheet"}

            if sheet_index < 0 or sheet_index >= sheets.Count:
                return {"error": f"Invalid sheet index: {sheet_index}. Count: {sheets.Count}"}

            sheet = sheets.Item(sheet_index + 1)
            sheet_name = sheet.Name
            sheet.Delete()

            return {
                "status": "deleted",
                "sheet_name": sheet_name,
                "remaining_sheets": sheets.Count
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    # Aliases for consistency with MCP tool names
    def export_step(self, file_path: str) -> Dict[str, Any]:
        """Alias for export_to_step"""
        return self.export_to_step(file_path)

    def export_stl(self, file_path: str, quality: str = "Medium") -> Dict[str, Any]:
        """Alias for export_to_stl"""
        return self.export_to_stl(file_path, quality)

    def export_iges(self, file_path: str) -> Dict[str, Any]:
        """Alias for export_to_iges"""
        return self.export_to_iges(file_path)

    def export_pdf(self, file_path: str) -> Dict[str, Any]:
        """Alias for export_to_pdf"""
        return self.export_to_pdf(file_path)

    def export_dxf(self, file_path: str) -> Dict[str, Any]:
        """Alias for export_to_dxf"""
        return self.export_to_dxf(file_path)

    def export_parasolid(self, file_path: str) -> Dict[str, Any]:
        """Alias for export_to_parasolid"""
        return self.export_to_parasolid(file_path)

    def export_jt(self, file_path: str) -> Dict[str, Any]:
        """Alias for export_to_jt"""
        return self.export_to_jt(file_path)


class ViewModel:
    """Manages view manipulation"""

    def __init__(self, document_manager):
        self.doc_manager = document_manager

    def set_view(self, view: str) -> Dict[str, Any]:
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
            if not hasattr(doc, 'Windows') or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, 'View') else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            # Valid view names (discovered via introspection)
            # Note: Bottom, Back, Left may not work in all contexts
            valid_views = ["Iso", "Top", "Front", "Right", "Bottom", "Back", "Left"]

            if view not in valid_views:
                return {"error": f"Invalid view: {view}. Valid: {', '.join(valid_views)}"}

            # Use ApplyNamedView with string name (discovered method!)
            if hasattr(view_obj, 'ApplyNamedView'):
                view_obj.ApplyNamedView(view)
                return {
                    "status": "view_set",
                    "view": view
                }
            else:
                return {
                    "error": "ApplyNamedView not available",
                    "note": "Use View menu in Solid Edge UI"
                }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def zoom_fit(self) -> Dict[str, Any]:
        """Zoom to fit all geometry in view"""
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, 'Windows') or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, 'View') else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            # Zoom to fit
            if hasattr(view_obj, 'Fit'):
                view_obj.Fit()
                return {"status": "zoomed_fit"}
            else:
                return {
                    "error": "Fit method not available",
                    "note": "Use View > Fit in Solid Edge UI"
                }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def zoom_to_selection(self) -> Dict[str, Any]:
        """Zoom to fit all geometry (equivalent to View > Fit)."""
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, 'Windows') or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, 'View') else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            view_obj.Fit()

            return {"status": "zoomed_to_selection"}
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def set_display_mode(self, mode: str) -> Dict[str, Any]:
        """
        Set the display mode for the active view.

        Args:
            mode: Display mode - 'Shaded', 'ShadedWithEdges', 'Wireframe', 'HiddenEdgesVisible'

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, 'Windows') or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, 'View') else None

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
                return {"error": f"Invalid mode: {mode}. Use 'Shaded', 'ShadedWithEdges', 'Wireframe', or 'HiddenEdgesVisible'"}

            view_obj.SetRenderMode(mode_value)
            return {
                "status": "display_mode_set",
                "mode": mode
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def set_view_background(self, red: int, green: int, blue: int) -> Dict[str, Any]:
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

            if not hasattr(doc, 'Windows') or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, 'View') else None

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

            return {
                "status": "updated",
                "color": [red, green, blue]
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def get_camera(self) -> Dict[str, Any]:
        """
        Get the current camera parameters.

        Returns eye position, target position, up vector, perspective flag,
        and scale/field-of-view angle.

        Returns:
            Dict with camera parameters
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, 'Windows') or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, 'View') else None

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
                "scale_or_angle": result[10]
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def set_camera(self, eye_x: float, eye_y: float, eye_z: float,
                   target_x: float, target_y: float, target_z: float,
                   up_x: float = 0.0, up_y: float = 1.0, up_z: float = 0.0,
                   perspective: bool = False,
                   scale_or_angle: float = 1.0) -> Dict[str, Any]:
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

            if not hasattr(doc, 'Windows') or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, 'View') else None

            if not view_obj:
                return {"error": "Cannot access view object"}

            view_obj.SetCamera(
                eye_x, eye_y, eye_z,
                target_x, target_y, target_z,
                up_x, up_y, up_z,
                perspective, scale_or_angle
            )

            return {
                "status": "camera_set",
                "eye": [eye_x, eye_y, eye_z],
                "target": [target_x, target_y, target_z],
                "up": [up_x, up_y, up_z],
                "perspective": perspective,
                "scale_or_angle": scale_or_angle
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }
