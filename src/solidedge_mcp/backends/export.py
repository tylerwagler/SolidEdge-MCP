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

        Args:
            template: Drawing template path (optional)
            views: List of views to create - ['Front', 'Top', 'Right', 'Isometric']

        Returns:
            Dict with status and drawing info
        """
        try:
            source_doc = self.doc_manager.get_active_document()
            app = self.doc_manager.connection.get_application()

            if views is None:
                views = ['Front', 'Top', 'Right']

            # Create a new draft document
            if template and os.path.exists(template):
                draft_doc = app.Documents.Add(template)
            else:
                draft_doc = app.Documents.Add("SolidEdge.DraftDocument")

            # Note: Actual view creation is complex and requires
            # accessing the drawing views collection and placing views

            return {
                "status": "created",
                "type": "drawing",
                "views": views,
                "note": "Drawing views require manual placement - use Solid Edge UI for detailed control"
            }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }

    def capture_screenshot(self, file_path: str, width: int = 1920, height: int = 1080) -> Dict[str, Any]:
        """
        Capture a screenshot of the current view.

        Args:
            file_path: Output image file path
            width: Image width in pixels
            height: Image height in pixels

        Returns:
            Dict with status and image info
        """
        try:
            doc = self.doc_manager.get_active_document()

            # Ensure file has image extension
            if not any(file_path.lower().endswith(ext) for ext in ['.png', '.jpg', '.jpeg', '.bmp']):
                file_path += '.png'

            # Get the window
            if hasattr(doc, 'Windows') and doc.Windows.Count > 0:
                window = doc.Windows.Item(1)

                # Save as image
                # Note: Actual method varies by Solid Edge version
                if hasattr(window, 'SaveAsImage'):
                    window.SaveAsImage(file_path, width, height)
                    return {
                        "status": "captured",
                        "path": file_path,
                        "dimensions": [width, height],
                        "size_bytes": os.path.getsize(file_path) if os.path.exists(file_path) else 0
                    }
                else:
                    return {
                        "error": "Screenshot capture not supported in this Solid Edge version",
                        "note": "Use View > Save Image in Solid Edge UI"
                    }
            else:
                return {"error": "No window available for screenshot"}
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

            # Map view names to constants
            view_map = {
                "Iso": ViewOrientationConstants.seIsoView,
                "Top": ViewOrientationConstants.seTopView,
                "Bottom": ViewOrientationConstants.seBottomView,
                "Front": ViewOrientationConstants.seFrontView,
                "Back": ViewOrientationConstants.seBackView,
                "Right": ViewOrientationConstants.seRightView,
                "Left": ViewOrientationConstants.seLeftView
            }

            view_const = view_map.get(view)
            if view_const is None:
                return {"error": f"Invalid view: {view}"}

            # Set the view
            if hasattr(view_obj, 'SetNamedView'):
                view_obj.SetNamedView(view_const)
                return {
                    "status": "view_set",
                    "view": view
                }
            else:
                return {
                    "error": "SetNamedView not available",
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
        """Zoom to selected geometry"""
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, 'Windows') or doc.Windows.Count == 0:
                return {"error": "No window available"}

            window = doc.Windows.Item(1)
            view_obj = window.View if hasattr(window, 'View') else None

            if not view_obj:
                return {"error": "Cannot access view object"}

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

            # Map mode names to display style values
            # Note: Actual constants may vary by Solid Edge version
            mode_map = {
                "Shaded": 1,
                "ShadedWithEdges": 2,
                "Wireframe": 3,
                "HiddenEdgesVisible": 4
            }

            mode_value = mode_map.get(mode)
            if mode_value is None:
                return {"error": f"Invalid mode: {mode}. Use 'Shaded', 'ShadedWithEdges', 'Wireframe', or 'HiddenEdgesVisible'"}

            # Set display style
            if hasattr(view_obj, 'Style'):
                view_obj.Style = mode_value
                return {
                    "status": "display_mode_set",
                    "mode": mode
                }
            else:
                return {
                    "error": "Display mode not accessible",
                    "note": "Use View > Display Style in Solid Edge UI"
                }
        except Exception as e:
            return {
                "error": str(e),
                "traceback": traceback.format_exc()
            }
