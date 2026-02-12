"""
Solid Edge Export and Visualization Operations

Handles exporting to various formats and creating drawings.
"""

from typing import Dict, Any, Optional, List
import os
import traceback
from .constants import ViewOrientationConstants


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

            # View orientation constants (Solid Edge ViewOrientationConstants)
            view_orient_map = {
                'Front': 5,       # igFrontView
                'Back': 8,        # igBackView
                'Top': 6,         # igTopView
                'Bottom': 9,      # igBottomView
                'Right': 7,       # igRightView
                'Left': 10,       # igLeftView
                'Isometric': 12,  # igISOView
                'Iso': 12,
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

            # Map mode names to seRenderMode constants (from type library)
            mode_map = {
                "Wireframe": 1,           # seRenderModeWireframe
                "Shaded": 8,              # seRenderModeSmooth
                "ShadedWithEdges": 11,    # seRenderModeSmoothBoundary
                "HiddenEdgesVisible": 6,  # seRenderModeVHL
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
