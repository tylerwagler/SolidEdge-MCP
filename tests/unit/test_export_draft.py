"""
Unit tests for ExportManager draft backend methods.

Tests printing, paper size, face textures, bend tables,
smart frames, symbols, PMI, draft global parameters, and symbol file origins.
Uses unittest.mock to simulate COM objects.
"""

from unittest.mock import MagicMock

import pytest


@pytest.fixture
def export_mgr():
    """Create ExportManager with mocked dependencies."""
    from solidedge_mcp.backends.export import ExportManager

    dm = MagicMock()
    doc = MagicMock()
    dm.get_active_document.return_value = doc
    return ExportManager(dm), doc


# ============================================================================
# BATCH 10: PRINTING
# ============================================================================


class TestPrintDrawing:
    def test_with_draft_print_utility(self, export_mgr):
        em, doc = export_mgr
        dpu = MagicMock()
        doc.DraftPrintUtility = dpu

        result = em.print_drawing(copies=2, all_sheets=False)
        assert result["status"] == "printed"
        assert result["copies"] == 2
        dpu.PrintOut.assert_called_once()

    def test_fallback_printout(self, export_mgr):
        em, doc = export_mgr
        del doc.DraftPrintUtility

        result = em.print_drawing()
        assert result["status"] == "printed"

    def test_no_print_support(self, export_mgr):
        em, doc = export_mgr
        del doc.DraftPrintUtility
        del doc.PrintOut

        result = em.print_drawing()
        assert "error" in result


class TestSetPrinter:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        dpu = MagicMock()
        doc.DraftPrintUtility = dpu

        result = em.set_printer("HP LaserJet")
        assert result["status"] == "set"
        assert result["printer"] == "HP LaserJet"
        assert dpu.Printer == "HP LaserJet"

    def test_no_dpu(self, export_mgr):
        em, doc = export_mgr
        del doc.DraftPrintUtility

        result = em.set_printer("HP LaserJet")
        assert "error" in result

    def test_different_printer(self, export_mgr):
        em, doc = export_mgr
        dpu = MagicMock()
        doc.DraftPrintUtility = dpu

        result = em.set_printer("PDF Printer")
        assert result["printer"] == "PDF Printer"


class TestGetPrinter:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        dpu = MagicMock()
        dpu.Printer = "HP LaserJet"
        doc.DraftPrintUtility = dpu

        result = em.get_printer()
        assert result["printer"] == "HP LaserJet"

    def test_no_dpu(self, export_mgr):
        em, doc = export_mgr
        del doc.DraftPrintUtility

        result = em.get_printer()
        assert "error" in result

    def test_different_printer(self, export_mgr):
        em, doc = export_mgr
        dpu = MagicMock()
        dpu.Printer = "PDF Printer"
        doc.DraftPrintUtility = dpu

        result = em.get_printer()
        assert result["printer"] == "PDF Printer"


class TestSetPaperSize:
    def test_landscape(self, export_mgr):
        em, doc = export_mgr
        dpu = MagicMock()
        doc.DraftPrintUtility = dpu

        result = em.set_paper_size(0.297, 0.210, "Landscape")
        assert result["status"] == "set"
        assert result["orientation"] == "Landscape"
        assert result["width"] == 0.297

    def test_portrait(self, export_mgr):
        em, doc = export_mgr
        dpu = MagicMock()
        doc.DraftPrintUtility = dpu

        result = em.set_paper_size(0.210, 0.297, "Portrait")
        assert result["status"] == "set"
        assert result["orientation"] == "Portrait"

    def test_no_dpu(self, export_mgr):
        em, doc = export_mgr
        del doc.DraftPrintUtility

        result = em.set_paper_size(0.297, 0.210)
        assert "error" in result


class TestPrintDocument:
    def test_default_print(self, export_mgr):
        em, doc = export_mgr
        doc.Name = "part1.par"

        result = em.print_document()
        assert result["status"] == "printed"
        assert result["printer"] == "default"
        assert result["copies"] == 1
        doc.PrintOut.assert_called_once_with()

    def test_with_printer_and_copies(self, export_mgr):
        em, doc = export_mgr
        doc.Name = "part1.par"

        result = em.print_document(printer="HP LaserJet", num_copies=3)
        assert result["status"] == "printed"
        assert result["printer"] == "HP LaserJet"
        assert result["copies"] == 3
        doc.PrintOut.assert_called_once_with(Printer="HP LaserJet", NumCopies=3)

    def test_print_to_file(self, export_mgr):
        em, doc = export_mgr
        doc.Name = "part1.par"

        result = em.print_document(print_to_file=True, output_file_name="C:/temp/out.pdf")
        assert result["status"] == "printed"
        doc.PrintOut.assert_called_once_with(
            PrintToFile=True, OutputFileName="C:/temp/out.pdf"
        )

    def test_color_as_black(self, export_mgr):
        em, doc = export_mgr
        doc.Name = "part1.par"

        result = em.print_document(color_as_black=True)
        assert result["status"] == "printed"
        doc.PrintOut.assert_called_once_with(ColorAsBlack=True)

    def test_com_error(self, export_mgr):
        em, doc = export_mgr
        doc.PrintOut.side_effect = Exception("Printer offline")

        result = em.print_document()
        assert "error" in result


# ============================================================================
# BATCH 10: DRAWING VIEW TOOLS
# ============================================================================


class TestSetFaceTexture:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        face = MagicMock()
        faces = MagicMock()
        faces.Count = 3
        faces.Item.return_value = face

        body = MagicMock()
        body.Faces.return_value = faces

        model = MagicMock()
        model.Body = body

        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        result = em.set_face_texture(1, "Wood")
        assert result["status"] == "set"
        assert result["texture_name"] == "Wood"

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        faces = MagicMock()
        faces.Count = 2

        body = MagicMock()
        body.Faces.return_value = faces

        model = MagicMock()
        model.Body = body

        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        result = em.set_face_texture(5, "Wood")
        assert "error" in result

    def test_no_models(self, export_mgr):
        em, doc = export_mgr
        del doc.Models

        result = em.set_face_texture(0, "Wood")
        assert "error" in result


class TestCreateBendTable:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        dv = MagicMock()
        dvs = MagicMock()
        dvs.Item.return_value = dv
        # Mock _get_drawing_views
        em._get_drawing_views = MagicMock(return_value=dvs)
        sheet = MagicMock()
        bend_tables = MagicMock()
        bend_tables.Count = 1
        sheet.DraftBendTables = bend_tables
        doc.ActiveSheet = sheet

        result = em.create_bend_table(view_index=0)
        assert result["status"] == "created"
        assert result["type"] == "bend_table"
        bend_tables.Add.assert_called_once()

    def test_no_drawing_views(self, export_mgr):
        em, doc = export_mgr
        em._get_drawing_views = MagicMock(return_value=None)

        result = em.create_bend_table()
        assert "error" in result
        assert "No drawing views" in result["error"]

    def test_com_error(self, export_mgr):
        em, doc = export_mgr
        dvs = MagicMock()
        dvs.Item.side_effect = Exception("fail")
        em._get_drawing_views = MagicMock(return_value=dvs)

        result = em.create_bend_table()
        assert "error" in result


# ============================================================================
# SMART FRAMES
# ============================================================================


class TestAddSmartFrame:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        smart_frames = MagicMock()
        sheet.SmartFrames2d = smart_frames
        doc.ActiveSheet = sheet

        result = em.add_smart_frame("A4", 0.0, 0.0, 0.297, 0.21)
        assert result["status"] == "added"
        assert result["style"] == "A4"
        assert result["corner1"] == [0.0, 0.0]
        assert result["corner2"] == [0.297, 0.21]
        smart_frames.AddBy2Points.assert_called_once_with("A4", 0.0, 0.0, 0.297, 0.21)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.ActiveSheet

        result = em.add_smart_frame("A4", 0.0, 0.0, 0.297, 0.21)
        assert "error" in result

    def test_com_error(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        smart_frames = MagicMock()
        smart_frames.AddBy2Points.side_effect = Exception("Style not found")
        sheet.SmartFrames2d = smart_frames
        doc.ActiveSheet = sheet

        result = em.add_smart_frame("InvalidStyle", 0.0, 0.0, 0.297, 0.21)
        assert "error" in result


class TestAddSmartFrameByOrigin:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        smart_frames = MagicMock()
        sheet.SmartFrames2d = smart_frames
        doc.ActiveSheet = sheet

        result = em.add_smart_frame_by_origin("A3", 0.01, 0.01, 0.28, 0.01, 0.01, 0.40)
        assert result["status"] == "added"
        assert result["style"] == "A3"
        assert result["origin"] == [0.01, 0.01]
        smart_frames.AddByOrigin.assert_called_once_with("A3", 0.01, 0.01, 0.28, 0.01, 0.01, 0.40)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.ActiveSheet

        result = em.add_smart_frame_by_origin("A3", 0.01, 0.01, 0.28, 0.01, 0.01, 0.40)
        assert "error" in result

    def test_com_error(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        smart_frames = MagicMock()
        smart_frames.AddByOrigin.side_effect = Exception("COM error")
        sheet.SmartFrames2d = smart_frames
        doc.ActiveSheet = sheet

        result = em.add_smart_frame_by_origin("A3", 0.01, 0.01, 0.28, 0.01, 0.01, 0.40)
        assert "error" in result


# ============================================================================
# SYMBOLS
# ============================================================================


class TestAddSymbol:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        symbols = MagicMock()
        sheet.Symbols = symbols
        doc.ActiveSheet = sheet

        result = em.add_symbol("C:/symbols/arrow.sym", 0.1, 0.1, 0)
        assert result["status"] == "placed"
        assert result["file"] == "C:/symbols/arrow.sym"
        assert result["position"] == [0.1, 0.1]
        symbols.Add.assert_called_once_with(0, "C:/symbols/arrow.sym", 0.1, 0.1)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.ActiveSheet

        result = em.add_symbol("C:/symbols/arrow.sym", 0.1, 0.1)
        assert "error" in result

    def test_com_error(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        symbols = MagicMock()
        symbols.Add.side_effect = Exception("File not found")
        sheet.Symbols = symbols
        doc.ActiveSheet = sheet

        result = em.add_symbol("C:/symbols/nonexistent.sym", 0.1, 0.1)
        assert "error" in result


class TestGetSymbols:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sym1 = MagicMock()
        sym1.Name = "Arrow"
        sym1.OriginX = 0.1
        sym1.OriginY = 0.1

        sym2 = MagicMock()
        sym2.Name = "Star"
        sym2.OriginX = 0.2
        sym2.OriginY = 0.2

        symbols = MagicMock()
        symbols.Count = 2
        symbols.Item.side_effect = lambda i: {1: sym1, 2: sym2}[i]
        sheet.Symbols = symbols
        doc.ActiveSheet = sheet

        result = em.get_symbols()
        assert result["count"] == 2
        assert result["symbols"][0]["name"] == "Arrow"
        assert result["symbols"][0]["x"] == 0.1
        assert result["symbols"][1]["name"] == "Star"
        assert result["symbols"][1]["index"] == 1

    def test_empty(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        symbols = MagicMock()
        symbols.Count = 0
        sheet.Symbols = symbols
        doc.ActiveSheet = sheet

        result = em.get_symbols()
        assert result["count"] == 0
        assert result["symbols"] == []

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.ActiveSheet

        result = em.get_symbols()
        assert "error" in result


# ============================================================================
# PMI (Product Manufacturing Information)
# ============================================================================


class TestGetPmiInfo:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        pmi = MagicMock()

        pmi_dims = MagicMock()
        pmi_dims.Count = 5
        pmi.Dimensions = pmi_dims

        pmi_balloons = MagicMock()
        pmi_balloons.Count = 3
        pmi.Balloons = pmi_balloons

        pmi_datum = MagicMock()
        pmi_datum.Count = 2
        pmi.DatumFrames = pmi_datum

        pmi_fcf = MagicMock()
        pmi_fcf.Count = 1
        pmi.FeatureControlFrames = pmi_fcf

        pmi_sf = MagicMock()
        pmi_sf.Count = 0
        pmi.SurfaceFinishSymbols = pmi_sf

        pmi_ws = MagicMock()
        pmi_ws.Count = 0
        pmi.WeldSymbols = pmi_ws

        pmi_cm = MagicMock()
        pmi_cm.Count = 4
        pmi.CenterMarks = pmi_cm

        pmi_cl = MagicMock()
        pmi_cl.Count = 2
        pmi.CenterLines = pmi_cl

        pmi_tb = MagicMock()
        pmi_tb.Count = 1
        pmi.TextBoxes = pmi_tb

        doc.PMI = pmi

        result = em.get_pmi_info()
        assert result["has_pmi"] is True
        assert result["dimensions"] == 5
        assert result["balloons"] == 3
        assert result["datum_frames"] == 2
        assert result["feature_control_frames"] == 1
        assert result["center_marks"] == 4
        assert result["text_boxes"] == 1

    def test_no_pmi(self, export_mgr):
        em, doc = export_mgr
        del doc.PMI

        result = em.get_pmi_info()
        assert result["has_pmi"] is False

    def test_partial_collections(self, export_mgr):
        em, doc = export_mgr
        pmi = MagicMock(spec=["Dimensions"])
        pmi_dims = MagicMock()
        pmi_dims.Count = 3
        pmi.Dimensions = pmi_dims
        doc.PMI = pmi

        result = em.get_pmi_info()
        assert result["has_pmi"] is True
        assert result["dimensions"] == 3


class TestSetPmiVisibility:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        pmi = MagicMock()
        doc.PMI = pmi

        result = em.set_pmi_visibility(True, False, True)
        assert result["status"] == "updated"
        assert result["show"] is True
        assert result["show_dimensions"] is False
        assert result["show_annotations"] is True

    def test_no_pmi(self, export_mgr):
        em, doc = export_mgr
        del doc.PMI

        result = em.set_pmi_visibility()
        assert "error" in result

    def test_defaults(self, export_mgr):
        em, doc = export_mgr
        pmi = MagicMock()
        doc.PMI = pmi

        result = em.set_pmi_visibility()
        assert result["status"] == "updated"
        assert result["show"] is True
        assert result["show_dimensions"] is True
        assert result["show_annotations"] is True


# ============================================================================
# DRAFT: GET DRAFT GLOBAL PARAMETER
# ============================================================================


class TestGetDraftGlobalParameter:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        doc.Sheets = MagicMock()
        doc.GetGlobalParameter.return_value = 0.003

        result = em.get_draft_global_parameter(5)
        assert result["status"] == "success"
        assert result["parameter"] == 5
        assert result["value"] == 0.003
        doc.GetGlobalParameter.assert_called_once_with(5)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.get_draft_global_parameter(5)
        assert "error" in result

    def test_com_error(self, export_mgr):
        em, doc = export_mgr
        doc.Sheets = MagicMock()
        doc.GetGlobalParameter.side_effect = Exception("Invalid parameter")

        result = em.get_draft_global_parameter(999)
        assert "error" in result


# ============================================================================
# DRAFT: SET DRAFT GLOBAL PARAMETER
# ============================================================================


class TestSetDraftGlobalParameter:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        doc.Sheets = MagicMock()

        result = em.set_draft_global_parameter(5, 0.005)
        assert result["status"] == "set"
        assert result["parameter"] == 5
        assert result["value"] == 0.005
        doc.SetGlobalParameter.assert_called_once_with(5, 0.005)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.set_draft_global_parameter(5, 0.005)
        assert "error" in result

    def test_com_error(self, export_mgr):
        em, doc = export_mgr
        doc.Sheets = MagicMock()
        doc.SetGlobalParameter.side_effect = Exception("Read-only parameter")

        result = em.set_draft_global_parameter(5, 0.005)
        assert "error" in result


# ============================================================================
# DRAFT: GET SYMBOL FILE ORIGIN
# ============================================================================


class TestGetSymbolFileOrigin:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        doc.Sheets = MagicMock()
        doc.GetSymbolFileOrigin.return_value = (0.05, 0.10)

        result = em.get_symbol_file_origin()
        assert result["status"] == "success"
        assert result["x"] == 0.05
        assert result["y"] == 0.10
        doc.GetSymbolFileOrigin.assert_called_once()

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.get_symbol_file_origin()
        assert "error" in result

    def test_com_error(self, export_mgr):
        em, doc = export_mgr
        doc.Sheets = MagicMock()
        doc.GetSymbolFileOrigin.side_effect = Exception("No origin set")

        result = em.get_symbol_file_origin()
        assert "error" in result


# ============================================================================
# DRAFT: SET SYMBOL FILE ORIGIN
# ============================================================================


class TestSetSymbolFileOrigin:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        doc.Sheets = MagicMock()

        result = em.set_symbol_file_origin(0.05, 0.10)
        assert result["status"] == "set"
        assert result["x"] == 0.05
        assert result["y"] == 0.10
        doc.SetSymbolFileOrigin.assert_called_once_with(0.05, 0.10)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.set_symbol_file_origin(0.05, 0.10)
        assert "error" in result

    def test_com_error(self, export_mgr):
        em, doc = export_mgr
        doc.Sheets = MagicMock()
        doc.SetSymbolFileOrigin.side_effect = Exception("COM error")

        result = em.set_symbol_file_origin(0.05, 0.10)
        assert "error" in result
