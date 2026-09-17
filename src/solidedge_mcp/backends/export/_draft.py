"""Draft-specific operations (smart frames, symbols, PMI, printing, etc.)."""

import contextlib
import math
from typing import Any

import pythoncom

from solidedge_mcp.backends.errors import com_hresult, error_result

from ..comutil import owned_style_for
from ..constants import DraftPrintOrientationConstants
from ..features._base import verifies_collection_growth
from ..logging import get_logger
from ..query._base import BodyNotReachableError, all_faces, body_of
from ._base import NOT_A_DRAFT, com_get

_logger = get_logger(__name__)


class DraftMixin:
    """Mixin providing draft document operations."""

    # =================================================================
    # SMART FRAMES
    # =================================================================

    def _print_utility(self) -> Any:
        """The draft print utility, or None.

        ``Document.DraftPrintUtility`` does not exist on any Solid Edge
        document, so reading it there returned None every time and each of
        these four tools took its "not available" branch.
        ``Application.GetDraftPrintUtility()`` is the real route, verified on
        Solid Edge 2026: the object it returns answers Copies, Printer,
        PaperWidth, RemoveAllDocuments and AddSheet.
        """
        try:
            app = self.doc_manager.connection.get_application()
        except Exception:
            return None
        try:
            return app.GetDraftPrintUtility()
        except Exception:
            return None

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
            err = self._require_draft(doc, NOT_A_DRAFT)
            if err:
                return err
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
            return error_result(e)

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
            err = self._require_draft(doc, NOT_A_DRAFT)
            if err:
                return err
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
            return error_result(e)

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
            err = self._require_draft(doc, NOT_A_DRAFT)
            if err:
                return err
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
            return error_result(e)

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
            err = self._require_draft(doc, NOT_A_DRAFT)
            if err:
                return err
            sheet = doc.ActiveSheet

            symbols = sheet.Symbols
            items = []
            for i in range(1, symbols.Count + 1):
                sym = symbols.Item(i)
                info: dict[str, Any] = {"index": i - 1}
                with contextlib.suppress(Exception):
                    info["name"] = sym.Name
                # Symbol2d has no OriginX/OriginY, so both keys were always
                # missing. Its position comes from its first keypoint.
                with contextlib.suppress(Exception):
                    keypoint = sym.GetKeyPoint(0)
                    info["x"] = float(keypoint[0])
                    info["y"] = float(keypoint[1])
                with contextlib.suppress(Exception):
                    info["scale"] = sym.ScaleFactor
                # Symbol2d.Angle is radians, like every other Solid Edge
                # Angle property; degrees is the unit at this boundary.
                with contextlib.suppress(Exception):
                    info["angle_degrees"] = math.degrees(sym.Angle)
                items.append(info)
            return {"count": len(items), "symbols": items}
        except Exception as e:
            return error_result(e)

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

            pmi = com_get(doc, "PMI")
            if pmi is None:
                return {
                    "has_pmi": False,
                    "error": "PMI not available on this document",
                }

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
            return error_result(e)

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

        PMI belongs to a part, sheet metal or assembly document, not to a
        draft. Solid Edge does not always accept ``Show``: on a part with no
        PMI content it stays False however it is written, so the result
        reports what the document holds afterwards rather than what was asked
        for. The three writes used to sit inside suppresses, which hid both
        that and any real failure.

        Args:
            show: Master PMI visibility toggle
            show_dimensions: Show/hide dimension PMI annotations
            show_annotations: Show/hide non-dimension PMI annotations

        Returns:
            Dict with updated visibility settings
        """
        try:
            doc = self.doc_manager.get_active_document()

            pmi = com_get(doc, "PMI")
            if pmi is None:
                return {"error": "PMI not available on this document"}

            pmi.Show = show
            pmi.ShowDimensions = show_dimensions
            pmi.ShowAnnotations = show_annotations

            return {
                "status": "updated",
                "requested": {
                    "show": show,
                    "show_dimensions": show_dimensions,
                    "show_annotations": show_annotations,
                },
                "show": com_get(pmi, "Show"),
                "show_dimensions": com_get(pmi, "ShowDimensions"),
                "show_annotations": com_get(pmi, "ShowAnnotations"),
            }
        except Exception as e:
            return error_result(e)

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
            err = self._require_draft(doc)
            if err:
                return err

            value = doc.GetGlobalParameter(parameter)
            return {"status": "success", "parameter": parameter, "value": value}
        except Exception as e:
            return error_result(e)

    def set_draft_global_parameter(self, parameter: int, value: Any) -> dict[str, Any]:
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
            err = self._require_draft(doc)
            if err:
                return err

            doc.SetGlobalParameter(parameter, value)
            return {"status": "set", "parameter": parameter, "value": value}
        except Exception as e:
            return error_result(e)

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
            err = self._require_draft(doc)
            if err:
                return err

            # GetSymbolFileOrigin(pxOrigin as VT_R8*, pyOrigin as VT_R8*). The
            # type library declares both [in], so the ordinary late-bound call
            # returns None and the values are lost. Invoking with the parameters
            # declared [in, out] by reference hands them back (verified on
            # Solid Edge 2026: (0.1, 0.2) after SetSymbolFileOrigin(0.1, 0.2)).
            ole = doc._oleobj_
            dispid = ole.GetIDsOfNames(0, "GetSymbolFileOrigin")
            byref_r8 = pythoncom.VT_BYREF | pythoncom.VT_R8
            try:
                x, y = ole.InvokeTypes(
                    dispid,
                    0,
                    pythoncom.DISPATCH_METHOD,
                    (pythoncom.VT_VOID, 0),
                    ((byref_r8, 3), (byref_r8, 3)),
                    0.0,
                    0.0,
                )
            except pythoncom.com_error as e:
                # DISP_E_BADINDEX is what a draft with no origin answers; it
                # arrives inside excepinfo under DISP_E_EXCEPTION.
                if com_hresult(e) == 0x8002000B:
                    return {
                        "error": (
                            "This draft has no symbol file origin. Set one with set_origin first."
                        )
                    }
                raise
            return {
                "status": "success",
                "x": x,
                "y": y,
            }
        except Exception as e:
            return error_result(e)

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
            err = self._require_draft(doc)
            if err:
                return err

            doc.SetSymbolFileOrigin(x, y)
            return {"status": "set", "x": x, "y": y}
        except Exception as e:
            return error_result(e)

    # =================================================================
    # FACE TEXTURE
    # =================================================================

    def set_face_texture(self, face_index: int, texture_name: str) -> dict[str, Any]:
        """Apply a texture to one face.

        ``Face.TextureFileName`` is on no Solid Edge interface, so the first
        attempt always raised and only the fallback ever ran. ``TextureFileName``
        belongs to ``FaceStyle``, which is what ``Face.Style`` holds, and that
        style may be a stock one shared across the document -- writing the
        texture onto it would texture every face using it. The face gets a
        style of its own instead, the same way its colour does.

        Args:
            face_index: 0-based face index.
            texture_name: Texture file name to apply.

        Returns:
            Dict with status and the texture Solid Edge reports afterwards.
        """
        try:
            doc = self.doc_manager.get_active_document()

            models = com_get(doc, "Models")
            if models is None:
                return {"error": "Active document does not have a Models collection"}
            if models.Count == 0:
                return {"error": "No models in document"}

            model = models.Item(1)
            body = body_of(model)
            faces = all_faces(body, model)

            if face_index < 0 or face_index >= faces.Count:
                return {"error": f"Invalid face_index: {face_index}. Count: {faces.Count}"}

            face = faces.Item(face_index + 1)
            style, err = owned_style_for(doc, face, f"Face {face_index}")
            if err:
                return err
            style.TextureFileName = texture_name

            return {
                "status": "set",
                "face_index": face_index,
                "texture_name": texture_name,
                "reads_back": com_get(com_get(face, "Style"), "TextureFileName"),
            }
        except BodyNotReachableError as e:
            return {"error": str(e)}
        except Exception as e:
            return error_result(e)

    # =================================================================
    # BEND TABLE
    # =================================================================

    @verifies_collection_growth("DraftBendTables")
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
            dvs = self._get_drawing_views()
            if dvs is None:
                return {"error": "No drawing views available"}

            dv = dvs.Item(view_index + 1)

            # DraftBendTables is on DraftDocument, not on Sheet.
            bend_tables = doc.DraftBendTables
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
            return error_result(e)

    # =================================================================
    # PRINTING
    # =================================================================

    def print_drawing(self, copies: int = 1, all_sheets: bool = True) -> dict[str, Any]:
        """
        Print the active draft document.

        Prefers DraftPrintUtility, which gives more control, and falls back to
        Document.PrintOut(Printer, NumCopies, ...).

        Args:
            copies: Number of copies to print
            all_sheets: Whether to print all sheets (True) or active only

        Returns:
            Dict with status
        """
        try:
            doc = self.doc_manager.get_active_document()

            # DraftPrintUtility gives more control than Document.PrintOut.
            dpu = self._print_utility()
            if dpu is not None:
                # Verified on Solid Edge 2026: Copies is get/put and takes.
                dpu.Copies = copies
                # There is no PrintAllSheets property; setting it did nothing
                # and every print silently used whatever was already queued.
                # AddDocument queues the whole document, AddSheet just one.
                with contextlib.suppress(Exception):
                    dpu.RemoveAllDocuments()
                if all_sheets:
                    dpu.AddDocument(doc)
                else:
                    dpu.AddSheet(doc.ActiveSheet)
                dpu.PrintOut()
                return {
                    "status": "printed",
                    "copies": com_get(dpu, "Copies", copies),
                    "all_sheets": all_sheets,
                }

            # Fall back to Document.PrintOut(Printer, NumCopies, ...). The
            # keyword was Copies, which is not a parameter of anything, so the
            # call raised and the bare retry below printed a single copy while
            # the result still claimed the number asked for.
            doc.PrintOut(NumCopies=copies)
            return {"status": "printed", "copies": copies}
        except Exception as e:
            return error_result(e)

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
            self.doc_manager.get_active_document()  # refuse with no document

            dpu = self._print_utility()
            if dpu is None:
                return {
                    "error": (
                        "Solid Edge did not hand out a print utility. It comes "
                        "from Application.GetDraftPrintUtility()."
                    )
                }
            dpu.Printer = printer_name

            return {"status": "set", "printer": printer_name}
        except Exception as e:
            return error_result(e)

    def get_printer(self) -> dict[str, Any]:
        """
        Get the current printer for the active draft document.

        Uses DraftPrintUtility.Printer property.

        Returns:
            Dict with printer name
        """
        try:
            self.doc_manager.get_active_document()  # refuse with no document

            dpu = self._print_utility()
            if dpu is None:
                return {
                    "error": (
                        "Solid Edge did not hand out a print utility. It comes "
                        "from Application.GetDraftPrintUtility()."
                    )
                }
            printer_name = dpu.Printer

            return {"printer": printer_name}
        except Exception as e:
            return error_result(e)

    def set_paper_size(
        self, width: float, height: float, orientation: str = "Landscape"
    ) -> dict[str, Any]:
        """Set the paper size and orientation for printing.

        ``DraftPrintUtility.PaperWidth`` and ``PaperHeight`` are in
        **millimetres**, and this passed meters: 0.42 was read back as the
        untouched default, so the size never changed while the result reported
        the requested one. The tool boundary stays meters and the conversion
        happens here.

        Orientation was passed as 1 for Portrait and 2 for Landscape, called
        "typical COM constants" in a comment.
        ``DraftPrintOrientationConstants`` has Portrait at 0 and Landscape at
        1, so "Portrait" selected landscape and 2 is not a member -- Solid Edge
        2026 rejects it and leaves the orientation alone.

        The printer driver constrains what it will accept: asking for A3 on a
        letter-size printer comes back as something else entirely. The result
        therefore reports what Solid Edge holds afterwards, in meters, not what
        was asked for.

        Args:
            width: Paper width in meters.
            height: Paper height in meters.
            orientation: 'Landscape' or 'Portrait'.

        Returns:
            Dict with status and the paper settings Solid Edge kept.
        """
        try:
            self.doc_manager.get_active_document()  # refuse with no document

            dpu = self._print_utility()
            if dpu is None:
                return {
                    "error": (
                        "Solid Edge did not hand out a print utility. It comes "
                        "from Application.GetDraftPrintUtility()."
                    )
                }

            if orientation.lower() == "portrait":
                orient = DraftPrintOrientationConstants.igDraftPrintPortrait
            else:
                orient = DraftPrintOrientationConstants.igDraftPrintLandscape
            dpu.Orientation = orient

            dpu.PaperWidth = width * 1000.0
            dpu.PaperHeight = height * 1000.0

            kept_width = com_get(dpu, "PaperWidth")
            kept_height = com_get(dpu, "PaperHeight")
            kept_orient = com_get(dpu, "Orientation")

            result: dict[str, Any] = {
                "status": "set",
                "requested": {"width": width, "height": height, "orientation": orientation},
                "orientation": (
                    "Portrait"
                    if kept_orient == DraftPrintOrientationConstants.igDraftPrintPortrait
                    else "Landscape"
                ),
            }
            if kept_width is not None:
                result["width"] = kept_width / 1000.0
            if kept_height is not None:
                result["height"] = kept_height / 1000.0
            return result
        except Exception as e:
            return error_result(e)

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
            kwargs: dict[str, Any] = {}
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
            return error_result(e)
