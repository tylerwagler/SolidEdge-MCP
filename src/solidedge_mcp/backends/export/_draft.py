"""Draft-specific operations (smart frames, symbols, PMI, printing, etc.)."""

import contextlib
import traceback
from typing import Any


class DraftMixin:
    """Mixin providing draft document operations."""

    # =================================================================
    # SMART FRAMES
    # =================================================================

    def add_smart_frame(
        self, style_name: str, x1: float, y1: float, x2: float, y2: float
    ) -> dict[str, Any]:
        """
        Add a smart frame (title block / border) to the active drawing sheet.

        Uses sheet.SmartFrames2d.AddBy2Points to place a bordered frame
        defined by two corner points.

        Args:
            style_name: Name of the smart frame style (e.g. 'A4', 'A3')
            x1: Lower-left corner X (meters)
            y1: Lower-left corner Y (meters)
            x2: Upper-right corner X (meters)
            y2: Upper-right corner Y (meters)

        Returns:
            Dict with status and style info
        """
        try:
            doc = self.doc_manager.get_active_document()
            if not hasattr(doc, "ActiveSheet"):
                return {"error": "Active document is not a draft"}
            sheet = doc.ActiveSheet

            smart_frames = sheet.SmartFrames2d
            smart_frames.AddBy2Points(style_name, x1, y1, x2, y2)

            return {
                "status": "added",
                "type": "smart_frame",
                "style": style_name,
                "corner1": [x1, y1],
                "corner2": [x2, y2],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def add_smart_frame_by_origin(
        self,
        style_name: str,
        x: float,
        y: float,
        top: float,
        bottom: float,
        left: float,
        right: float,
    ) -> dict[str, Any]:
        """
        Add a smart frame by origin point and margin extents.

        Uses sheet.SmartFrames2d.AddByOrigin to place a bordered frame
        defined by an origin and directional margins.

        Args:
            style_name: Name of the smart frame style
            x: Origin X (meters)
            y: Origin Y (meters)
            top: Top margin extent (meters)
            bottom: Bottom margin extent (meters)
            left: Left margin extent (meters)
            right: Right margin extent (meters)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            if not hasattr(doc, "ActiveSheet"):
                return {"error": "Active document is not a draft"}
            sheet = doc.ActiveSheet

            smart_frames = sheet.SmartFrames2d
            smart_frames.AddByOrigin(style_name, x, y, top, bottom, left, right)

            return {
                "status": "added",
                "type": "smart_frame",
                "style": style_name,
                "origin": [x, y],
                "margins": {
                    "top": top,
                    "bottom": bottom,
                    "left": left,
                    "right": right,
                },
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # SYMBOLS
    # =================================================================

    def add_symbol(
        self, file_path: str, x: float, y: float, insertion_type: int = 0
    ) -> dict[str, Any]:
        """
        Place a symbol from a symbol file onto the active drawing sheet.

        Uses sheet.Symbols.Add to insert a pre-defined symbol at the
        given coordinates.

        Args:
            file_path: Path to the symbol file (.sym)
            x: Placement X coordinate (meters)
            y: Placement Y coordinate (meters)
            insertion_type: Symbol insertion type constant (default 0)

        Returns:
            Dict with status and placement info
        """
        try:
            doc = self.doc_manager.get_active_document()
            if not hasattr(doc, "ActiveSheet"):
                return {"error": "Active document is not a draft"}
            sheet = doc.ActiveSheet

            symbols = sheet.Symbols
            symbols.Add(insertion_type, file_path, x, y)

            return {
                "status": "placed",
                "type": "symbol",
                "file": file_path,
                "position": [x, y],
                "insertion_type": insertion_type,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def get_symbols(self) -> dict[str, Any]:
        """
        List all symbols on the active draft sheet.

        Iterates sheet.Symbols collection and collects name and position
        information for each symbol.

        Returns:
            Dict with count and list of symbol info dicts
        """
        try:
            doc = self.doc_manager.get_active_document()
            if not hasattr(doc, "ActiveSheet"):
                return {"error": "Active document is not a draft"}
            sheet = doc.ActiveSheet

            symbols = sheet.Symbols
            items = []
            for i in range(1, symbols.Count + 1):
                sym = symbols.Item(i)
                info: dict[str, Any] = {"index": i - 1}
                with contextlib.suppress(Exception):
                    info["name"] = sym.Name
                with contextlib.suppress(Exception):
                    info["x"] = sym.OriginX
                with contextlib.suppress(Exception):
                    info["y"] = sym.OriginY
                items.append(info)
            return {"count": len(items), "symbols": items}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # PMI (Product Manufacturing Information)
    # =================================================================

    def get_pmi_info(self) -> dict[str, Any]:
        """
        Get PMI annotations summary for the active part document.

        Accesses doc.PMI and enumerates sub-collections to provide counts
        of each annotation type (dimensions, balloons, datum frames, etc.).

        Returns:
            Dict with has_pmi flag and counts per annotation type
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "PMI"):
                return {
                    "has_pmi": False,
                    "error": "PMI not available on this document",
                }

            pmi = doc.PMI

            result: dict[str, Any] = {"has_pmi": True}

            # Enumerate known PMI sub-collections
            pmi_collections = {
                "dimensions": "Dimensions",
                "balloons": "Balloons",
                "datum_frames": "DatumFrames",
                "feature_control_frames": "FeatureControlFrames",
                "surface_finish_symbols": "SurfaceFinishSymbols",
                "weld_symbols": "WeldSymbols",
                "center_marks": "CenterMarks",
                "center_lines": "CenterLines",
                "text_boxes": "TextBoxes",
            }

            for key, attr_name in pmi_collections.items():
                with contextlib.suppress(Exception):
                    coll = getattr(pmi, attr_name, None)
                    if coll is not None:
                        result[key] = coll.Count
                    else:
                        result[key] = 0

            return result
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_pmi_visibility(
        self,
        show: bool = True,
        show_dimensions: bool = True,
        show_annotations: bool = True,
    ) -> dict[str, Any]:
        """
        Show or hide PMI annotations on the active part document.

        Controls the overall visibility of PMI data as well as sub-categories
        for dimensions and annotations.

        Args:
            show: Master PMI visibility toggle
            show_dimensions: Show/hide dimension PMI annotations
            show_annotations: Show/hide non-dimension PMI annotations

        Returns:
            Dict with updated visibility settings
        """
        try:
            doc = self.doc_manager.get_active_document()

            if not hasattr(doc, "PMI"):
                return {"error": "PMI not available on this document"}

            pmi = doc.PMI

            with contextlib.suppress(Exception):
                pmi.Show = show
            with contextlib.suppress(Exception):
                pmi.ShowDimensions = show_dimensions
            with contextlib.suppress(Exception):
                pmi.ShowAnnotations = show_annotations

            return {
                "status": "updated",
                "show": show,
                "show_dimensions": show_dimensions,
                "show_annotations": show_annotations,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # DRAFT GLOBAL PARAMETERS
    # =================================================================

    def get_draft_global_parameter(self, parameter: int) -> dict[str, Any]:
        """
        Get a draft document global parameter.

        Uses DraftDocument.GetGlobalParameter(DraftGlobalConstants).

        Args:
            parameter: Draft global parameter ID (from DraftGlobalConstants)

        Returns:
            Dict with parameter value
        """
        try:
            doc = self.doc_manager.get_active_document()
            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            value = doc.GetGlobalParameter(parameter)
            return {"status": "success", "parameter": parameter, "value": value}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_draft_global_parameter(self, parameter: int, value) -> dict[str, Any]:
        """
        Set a draft document global parameter.

        Uses DraftDocument.SetGlobalParameter(DraftGlobalConstants, value).

        Args:
            parameter: Draft global parameter ID (from DraftGlobalConstants)
            value: New value for the parameter

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            doc.SetGlobalParameter(parameter, value)
            return {"status": "set", "parameter": parameter, "value": value}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # SYMBOL FILE ORIGIN
    # =================================================================

    def get_symbol_file_origin(self) -> dict[str, Any]:
        """
        Get the symbol file origin of the active draft document.

        Returns the origin point used when this draft is used as a symbol.

        Returns:
            Dict with x, y origin coordinates
        """
        try:
            doc = self.doc_manager.get_active_document()
            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            result = doc.GetSymbolFileOrigin()
            return {
                "status": "success",
                "x": result[0],
                "y": result[1],
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    def set_symbol_file_origin(self, x: float, y: float) -> dict[str, Any]:
        """
        Set the symbol file origin of the active draft document.

        Controls the origin point used when this draft is used as a symbol.

        Args:
            x: Origin X coordinate (meters)
            y: Origin Y coordinate (meters)

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()
            if not hasattr(doc, "Sheets"):
                return {"error": "Active document is not a draft document"}

            doc.SetSymbolFileOrigin(x, y)
            return {"status": "set", "x": x, "y": y}
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # FACE TEXTURE
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

    # =================================================================
    # BEND TABLE
    # =================================================================

    def create_bend_table(
        self,
        view_index: int = 0,
        saved_settings: str = "",
        auto_balloon: bool = True,
    ) -> dict[str, Any]:
        """
        Create a bend table on the active draft sheet (sheet metal).

        Args:
            view_index: 0-based index of the drawing view to attach to
            saved_settings: Name of saved settings to use (empty for default)
            auto_balloon: Whether to auto-create balloons

        Returns:
            Dict with status and table info
        """
        try:
            doc = self.doc_manager.get_active_document()
            sheet = doc.ActiveSheet
            dvs = self._get_drawing_views()
            if dvs is None:
                return {"error": "No drawing views available"}

            dv = dvs.Item(view_index + 1)

            bend_tables = sheet.DraftBendTables
            bend_tables.Add(
                dv,
                saved_settings,
                1 if auto_balloon else 0,
                1,  # CreateDraftBendTable = True
            )
            return {
                "status": "created",
                "type": "bend_table",
                "view_index": view_index,
                "count": bend_tables.Count,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}

    # =================================================================
    # PRINTING
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

    def print_document(
        self,
        printer: str | None = None,
        num_copies: int = 1,
        orientation: int | None = None,
        paper_size: int | None = None,
        scale: float | None = None,
        print_to_file: bool = False,
        output_file_name: str | None = None,
        print_range: int | None = None,
        sheets: str | None = None,
        color_as_black: bool = False,
        collate: bool = True,
    ) -> dict[str, Any]:
        """
        Print the active document.

        All parameters are optional and use system defaults when not specified.

        Args:
            printer: Printer name (None = default printer)
            num_copies: Number of copies
            orientation: Paper orientation constant
            paper_size: Paper size constant
            scale: Print scale (1.0 = 100%)
            print_to_file: Print to file instead of printer
            output_file_name: Output file path when print_to_file is True
            print_range: Print range constant
            sheets: Sheet specification string
            color_as_black: Print colors as black
            collate: Collate multiple copies

        Returns:
            Dict with status and print info
        """
        try:
            doc = self.doc_manager.get_active_document()

            # Build keyword args, omitting None values (COM optional params)
            kwargs = {}
            if printer is not None:
                kwargs["Printer"] = printer
            if num_copies != 1:
                kwargs["NumCopies"] = num_copies
            if orientation is not None:
                kwargs["Orientation"] = orientation
            if paper_size is not None:
                kwargs["PaperSize"] = paper_size
            if scale is not None:
                kwargs["Scale"] = scale
            if print_to_file:
                kwargs["PrintToFile"] = print_to_file
            if output_file_name is not None:
                kwargs["OutputFileName"] = output_file_name
            if print_range is not None:
                kwargs["PrintRange"] = print_range
            if sheets is not None:
                kwargs["Sheets"] = sheets
            if color_as_black:
                kwargs["ColorAsBlack"] = color_as_black
            if not collate:
                kwargs["Collate"] = collate

            doc.PrintOut(**kwargs)

            return {
                "status": "printed",
                "document": doc.Name,
                "printer": printer or "default",
                "copies": num_copies,
            }
        except Exception as e:
            return {"error": str(e), "traceback": traceback.format_exc()}
