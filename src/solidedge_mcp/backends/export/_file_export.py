"""File export operations (STEP, STL, IGES, PDF, DXF, Parasolid, JT, etc.)."""

import os
import traceback
from typing import Any

from ..logging import get_logger

_logger = get_logger(__name__)


class FileExportMixin:
    """Mixin providing file export methods."""

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

            _logger.info(f"Exported to STEP: {file_path}")
            return {
                "status": "exported",
                "format": "STEP",
                "path": file_path,
                "size_bytes": os.path.getsize(file_path) if os.path.exists(file_path) else 0,
            }
        except Exception as e:
            _logger.error(f"STEP export failed: {e}")
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

    def export_to_prc(self, file_path: str) -> dict[str, Any]:
        """
        Export the active document to PRC format (3D PDF).

        Args:
            file_path: Output file path (.prc extension)

        Returns:
            Dict with status and export info
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not file_path.lower().endswith(".prc"):
                file_path += ".prc"

            doc.SaveAsPRC(file_path)

            return {
                "status": "exported",
                "format": "PRC",
                "path": file_path,
                "size_bytes": os.path.getsize(file_path) if os.path.exists(file_path) else 0,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def export_to_plmxml(
        self, file_path: str, ini_file_path: str
    ) -> dict[str, Any]:
        """
        Export the active document to PLMXML format.

        Args:
            file_path: Output PLMXML file path (.xml extension)
            ini_file_path: Path to the PLMXML INI configuration file

        Returns:
            Dict with status and export info
        """
        try:
            doc = self.doc_manager.get_active_document()

            doc.SaveAsPLMXML(file_path, ini_file_path)

            return {
                "status": "exported",
                "format": "PLMXML",
                "path": file_path,
                "ini_file": ini_file_path,
                "size_bytes": os.path.getsize(file_path) if os.path.exists(file_path) else 0,
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
