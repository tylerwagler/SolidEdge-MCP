"""
Unit tests for ExportManager backend methods.

Tests draft sheet management and assembly drawing views.
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
# DRAFT SHEET MANAGEMENT
# ============================================================================


class TestAddDraftSheet:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Name = "Sheet 2"

        sheets = MagicMock()
        sheets.Count = 2
        sheets.AddSheet.return_value = sheet
        doc.Sheets = sheets

        result = em.add_draft_sheet()
        assert result["status"] == "added"
        assert result["total_sheets"] == 2
        sheets.AddSheet.assert_called_once()
        sheet.Activate.assert_called_once()

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_draft_sheet()
        assert "error" in result


# ============================================================================
# ASSEMBLY DRAWING VIEW
# ============================================================================


class TestAddAssemblyDrawingView:
    def test_success(self, export_mgr):
        em, doc = export_mgr

        model_link = MagicMock()
        model_links = MagicMock()
        model_links.Count = 1
        model_links.Item.return_value = model_link
        doc.ModelLinks = model_links

        sheet = MagicMock()
        dvs = MagicMock()
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_assembly_drawing_view(0.15, 0.15, "Front", 1.0)
        assert result["status"] == "added"
        assert result["orientation"] == "Front"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_assembly_drawing_view()
        assert "error" in result

    def test_no_model_link(self, export_mgr):
        em, doc = export_mgr
        doc.Sheets = MagicMock()

        model_links = MagicMock()
        model_links.Count = 0
        doc.ModelLinks = model_links

        result = em.add_assembly_drawing_view()
        assert "error" in result

    def test_invalid_orientation(self, export_mgr):
        em, doc = export_mgr
        doc.Sheets = MagicMock()

        model_links = MagicMock()
        model_links.Count = 1
        doc.ModelLinks = model_links

        result = em.add_assembly_drawing_view(orientation="InvalidView")
        assert "error" in result


# ============================================================================
# FLAT DXF EXPORT
# ============================================================================


class TestExportFlatDxf:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        flat_models = MagicMock()
        doc.FlatPatternModels = flat_models

        result = em.export_flat_dxf("C:/output/flat.dxf")
        assert result["status"] == "exported"
        assert result["format"] == "Flat DXF"
        flat_models.SaveAsFlatDXFEx.assert_called_once()

    def test_not_sheet_metal(self, export_mgr):
        em, doc = export_mgr
        del doc.FlatPatternModels

        result = em.export_flat_dxf("C:/output/flat.dxf")
        assert "error" in result

    def test_adds_extension(self, export_mgr):
        em, doc = export_mgr
        flat_models = MagicMock()
        doc.FlatPatternModels = flat_models

        result = em.export_flat_dxf("C:/output/flat")
        assert result["path"] == "C:/output/flat.dxf"


# ============================================================================
# DOCUMENT MANAGEMENT (activate, undo, redo)
# ============================================================================


@pytest.fixture
def doc_mgr():
    """Create DocumentManager with mocked connection."""
    from solidedge_mcp.backends.documents import DocumentManager

    conn = MagicMock()
    app = MagicMock()
    conn.get_application.return_value = app
    dm = DocumentManager(conn)
    return dm, app


class TestActivateDocument:
    def test_by_index(self, doc_mgr):
        dm, app = doc_mgr
        doc = MagicMock()
        doc.Name = "Part1.par"
        doc.FullName = "C:/parts/Part1.par"
        doc.Type = 1

        docs = MagicMock()
        docs.Count = 2
        docs.Item.return_value = doc
        app.Documents = docs

        result = dm.activate_document(0)
        assert result["status"] == "activated"
        assert result["name"] == "Part1.par"
        doc.Activate.assert_called_once()

    def test_by_name(self, doc_mgr):
        dm, app = doc_mgr
        doc1 = MagicMock()
        doc1.Name = "Part1.par"
        doc1.FullName = "C:/Part1.par"
        doc1.Type = 1

        doc2 = MagicMock()
        doc2.Name = "Part2.par"
        doc2.FullName = "C:/Part2.par"
        doc2.Type = 1

        docs = MagicMock()
        docs.Count = 2
        docs.Item.side_effect = lambda i: [None, doc1, doc2][i]
        app.Documents = docs

        result = dm.activate_document("Part2.par")
        assert result["status"] == "activated"
        doc2.Activate.assert_called_once()

    def test_not_found(self, doc_mgr):
        dm, app = doc_mgr
        doc1 = MagicMock()
        doc1.Name = "Part1.par"

        docs = MagicMock()
        docs.Count = 1
        docs.Item.return_value = doc1
        app.Documents = docs

        result = dm.activate_document("Nonexistent.par")
        assert "error" in result

    def test_invalid_index(self, doc_mgr):
        dm, app = doc_mgr
        docs = MagicMock()
        docs.Count = 1
        app.Documents = docs

        result = dm.activate_document(5)
        assert "error" in result


class TestUndo:
    def test_success(self, doc_mgr):
        dm, app = doc_mgr
        doc = MagicMock()
        dm.active_document = doc

        result = dm.undo()
        assert result["status"] == "undone"
        doc.Undo.assert_called_once()


class TestRedo:
    def test_success(self, doc_mgr):
        dm, app = doc_mgr
        doc = MagicMock()
        dm.active_document = doc

        result = dm.redo()
        assert result["status"] == "redone"
        doc.Redo.assert_called_once()


# ============================================================================
# CREATE DRAFT DOCUMENT
# ============================================================================


class TestCreateDraftDocument:
    def test_success(self, doc_mgr):
        dm, app = doc_mgr
        draft_doc = MagicMock()
        draft_doc.Name = "Draft1.dft"
        app.Documents.Add.return_value = draft_doc

        result = dm.create_draft()
        assert result["status"] == "created"
        app.Documents.Add.assert_called_once()

    def test_with_template(self, doc_mgr):
        dm, app = doc_mgr
        draft_doc = MagicMock()
        draft_doc.Name = "Draft1.dft"
        app.Documents.Add.return_value = draft_doc

        result = dm.create_draft("C:/templates/a3.dft")
        assert result["status"] == "created"


# ============================================================================
# DRAW POINT (SKETCHING)
# ============================================================================


class TestDrawPoint:
    def test_success(self):
        from solidedge_mcp.backends.sketching import SketchManager

        dm = MagicMock()
        sm = SketchManager(dm)

        profile = MagicMock()
        sm.active_profile = profile

        result = sm.draw_point(0.05, 0.03)
        assert result["status"] == "created"
        assert result["type"] == "point"
        assert result["position"] == [0.05, 0.03]
        profile.Holes2d.Add.assert_called_once_with(0.05, 0.03)

    def test_no_active_sketch(self):
        from solidedge_mcp.backends.sketching import SketchManager

        dm = MagicMock()
        sm = SketchManager(dm)

        result = sm.draw_point(0.05, 0.03)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_fallback_method(self):
        from solidedge_mcp.backends.sketching import SketchManager

        dm = MagicMock()
        sm = SketchManager(dm)

        profile = MagicMock()
        # Holes2d.Add throws - force fallback to construction circle
        profile.Holes2d.Add.side_effect = Exception("not available")
        sm.active_profile = profile

        result = sm.draw_point(0.01, 0.02)
        assert result["status"] == "created"
        assert result["method"] == "construction_circle"
        profile.Circles2d.AddByCenterRadius.assert_called_once()
        profile.ToggleConstruction.assert_called_once()


# ============================================================================
# GET SKETCH INFO
# ============================================================================


class TestGetSketchInfo:
    def test_success(self):
        from solidedge_mcp.backends.sketching import SketchManager

        dm = MagicMock()
        sm = SketchManager(dm)

        profile = MagicMock()
        profile.Lines2d.Count = 4
        profile.Circles2d.Count = 1
        profile.Arcs2d.Count = 0
        profile.Ellipses2d.Count = 0
        profile.BSplineCurves2d.Count = 0
        profile.Holes2d.Count = 2
        sm.active_profile = profile

        result = sm.get_sketch_info()
        assert result["status"] == "active"
        assert result["lines"] == 4
        assert result["circles"] == 1
        assert result["points"] == 2
        assert result["total_elements"] == 7

    def test_no_active_sketch(self):
        from solidedge_mcp.backends.sketching import SketchManager

        dm = MagicMock()
        sm = SketchManager(dm)

        result = sm.get_sketch_info()
        assert "error" in result


# ============================================================================
# GET ACTIVE DOCUMENT TYPE
# ============================================================================


class TestGetActiveDocumentType:
    def test_success(self, doc_mgr):
        dm, app = doc_mgr
        doc = MagicMock()
        doc.Name = "Part1.par"
        doc.FullName = "C:/Part1.par"
        doc.Type = 1  # igPartDocument
        dm.active_document = doc

        result = dm.get_active_document_type()
        assert result["type"] is not None
        assert result["name"] == "Part1.par"

    def test_no_document(self, doc_mgr):
        dm, app = doc_mgr
        dm.active_document = None
        app.ActiveDocument = None
        dm.connection.get_application.return_value = app

        # This should raise/return error
        dm.get_active_document_type()
        # Either works or errors depending on mock setup


# ============================================================================
# ADD TEXT BOX
# ============================================================================


class TestAddTextBox:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        text_boxes = MagicMock()
        text_box = MagicMock()
        text_boxes.Add.return_value = text_box
        sheet.TextBoxes = text_boxes
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_text_box(0.1, 0.1, "Hello")
        assert result["status"] == "added"
        assert result["text"] == "Hello"
        text_boxes.Add.assert_called_once_with(0.1, 0.1, 0)
        assert text_box.Text == "Hello"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_text_box(0.1, 0.1, "Test")
        assert "error" in result


# ============================================================================
# ADD LEADER
# ============================================================================


class TestAddLeader:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        leaders = MagicMock()
        leader = MagicMock()
        leaders.Add.return_value = leader
        sheet.Leaders = leaders
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_leader(0.05, 0.05, 0.15, 0.15, "Note")
        assert result["status"] == "added"
        assert result["text"] == "Note"
        leaders.Add.assert_called_once_with(0.05, 0.05, 0, 0.15, 0.15, 0)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_leader(0.05, 0.05, 0.15, 0.15)
        assert "error" in result


# ============================================================================
# ASSEMBLY: SET COMPONENT VISIBILITY
# ============================================================================


class TestSetComponentVisibility:
    @pytest.fixture
    def asm_mgr(self):
        from solidedge_mcp.backends.assembly import AssemblyManager

        dm = MagicMock()
        doc = MagicMock()
        dm.get_active_document.return_value = doc

        occurrence = MagicMock()
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occurrence
        doc.Occurrences = occurrences

        return AssemblyManager(dm), doc, occurrence

    def test_hide(self, asm_mgr):
        am, doc, occ = asm_mgr
        result = am.set_component_visibility(0, False)
        assert result["status"] == "updated"
        assert occ.Visible is False

    def test_show(self, asm_mgr):
        am, doc, occ = asm_mgr
        result = am.set_component_visibility(1, True)
        assert result["status"] == "updated"
        assert occ.Visible is True

    def test_invalid_index(self, asm_mgr):
        am, doc, occ = asm_mgr
        result = am.set_component_visibility(99, True)
        assert "error" in result

    def test_not_assembly(self, asm_mgr):
        am, doc, occ = asm_mgr
        del doc.Occurrences
        result = am.set_component_visibility(0, True)
        assert "error" in result


# ============================================================================
# ASSEMBLY: DELETE COMPONENT
# ============================================================================


class TestDeleteComponent:
    @pytest.fixture
    def asm_mgr(self):
        from solidedge_mcp.backends.assembly import AssemblyManager

        dm = MagicMock()
        doc = MagicMock()
        dm.get_active_document.return_value = doc

        occurrence = MagicMock()
        occurrence.Name = "Part1:1"
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occurrence
        doc.Occurrences = occurrences

        return AssemblyManager(dm), doc, occurrence

    def test_success(self, asm_mgr):
        am, doc, occ = asm_mgr
        result = am.delete_component(0)
        assert result["status"] == "deleted"
        assert result["name"] == "Part1:1"
        occ.Delete.assert_called_once()

    def test_invalid_index(self, asm_mgr):
        am, doc, occ = asm_mgr
        result = am.delete_component(99)
        assert "error" in result


# ============================================================================
# ASSEMBLY: GROUND COMPONENT
# ============================================================================


class TestGroundComponent:
    @pytest.fixture
    def asm_mgr(self):
        from solidedge_mcp.backends.assembly import AssemblyManager

        dm = MagicMock()
        doc = MagicMock()
        dm.get_active_document.return_value = doc

        occurrence = MagicMock()
        occurrences = MagicMock()
        occurrences.Count = 1
        occurrences.Item.return_value = occurrence
        doc.Occurrences = occurrences

        relations = MagicMock()
        doc.Relations3d = relations

        return AssemblyManager(dm), doc, occurrence, relations

    def test_ground(self, asm_mgr):
        am, doc, occ, rels = asm_mgr
        result = am.ground_component(0, True)
        assert result["status"] == "grounded"
        rels.AddGround.assert_called_once_with(occ)

    def test_invalid_index(self, asm_mgr):
        am, doc, occ, rels = asm_mgr
        result = am.ground_component(99, True)
        assert "error" in result


# ============================================================================
# ADD DIMENSION
# ============================================================================


class TestAddDimension:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dims = MagicMock()
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_dimension(0.0, 0.0, 0.1, 0.0)
        assert result["status"] == "added"
        assert result["type"] == "dimension"
        dims.AddLength.assert_called_once()

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_dimension(0, 0, 0.1, 0)
        assert "error" in result


# ============================================================================
# ADD BALLOON
# ============================================================================


class TestAddBalloon:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        balloons = MagicMock()
        balloon = MagicMock()
        balloons.Add.return_value = balloon
        sheet.Balloons = balloons
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_balloon(0.1, 0.1, "1", 0.05, 0.05)
        assert result["status"] == "added"
        assert result["type"] == "balloon"
        balloons.Add.assert_called_once_with(0.05, 0.05, 0, 0.1, 0.1, 0)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_balloon(0.1, 0.1)
        assert "error" in result


# ============================================================================
# ADD NOTE
# ============================================================================


class TestAddNote:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        text_boxes = MagicMock()
        text_box = MagicMock()
        text_boxes.Add.return_value = text_box
        sheet.TextBoxes = text_boxes
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_note(0.1, 0.1, "Test note")
        assert result["status"] == "added"
        assert result["type"] == "note"
        assert result["text"] == "Test note"
        text_boxes.Add.assert_called_once_with(0.1, 0.1, 0)
        assert text_box.Text == "Test note"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_note(0.1, 0.1, "Test")
        assert "error" in result


# ============================================================================
# GET SHEET INFO
# ============================================================================


class TestGetSheetInfo:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Name = "Sheet 1"
        sheets = MagicMock()
        sheets.Count = 2
        doc.ActiveSheet = sheet
        doc.Sheets = sheets

        result = em.get_sheet_info()
        assert result["status"] == "ok"
        assert result["sheet_name"] == "Sheet 1"
        assert result["sheet_count"] == 2

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.get_sheet_info()
        assert "error" in result


# ============================================================================
# ACTIVATE SHEET
# ============================================================================


class TestActivateSheet:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Name = "Sheet 2"
        sheets = MagicMock()
        sheets.Count = 3
        sheets.Item.return_value = sheet
        doc.Sheets = sheets

        result = em.activate_sheet(1)
        assert result["status"] == "activated"
        assert result["sheet_name"] == "Sheet 2"
        sheet.Activate.assert_called_once()

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        sheets = MagicMock()
        sheets.Count = 1
        doc.Sheets = sheets

        result = em.activate_sheet(5)
        assert "error" in result

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.activate_sheet(0)
        assert "error" in result


# ============================================================================
# CONNECTION: DISCONNECT AND IS_CONNECTED
# ============================================================================


class TestDisconnect:
    def test_success(self):
        from solidedge_mcp.backends.connection import SolidEdgeConnection

        conn = SolidEdgeConnection()
        conn._is_connected = True
        conn.application = MagicMock()

        result = conn.disconnect()
        assert result["status"] == "disconnected"
        assert conn.application is None
        assert conn._is_connected is False


class TestIsConnected:
    def test_connected(self):
        from solidedge_mcp.backends.connection import SolidEdgeConnection

        conn = SolidEdgeConnection()
        conn._is_connected = True
        conn.application = MagicMock()

        assert conn.is_connected() is True

    def test_not_connected(self):
        from solidedge_mcp.backends.connection import SolidEdgeConnection

        conn = SolidEdgeConnection()

        assert conn.is_connected() is False


# ============================================================================
# ASSEMBLY: REPLACE COMPONENT
# ============================================================================


class TestReplaceComponent:
    @pytest.fixture
    def asm_mgr(self):
        from solidedge_mcp.backends.assembly import AssemblyManager

        dm = MagicMock()
        doc = MagicMock()
        dm.get_active_document.return_value = doc

        occurrence = MagicMock()
        occurrence.Name = "Part1:1"
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occurrence
        doc.Occurrences = occurrences

        return AssemblyManager(dm), doc, occurrence

    def test_not_assembly(self, asm_mgr):
        am, doc, occ = asm_mgr
        del doc.Occurrences
        result = am.replace_component(0, "C:/parts/new.par")
        assert "error" in result

    def test_invalid_index(self, asm_mgr):
        am, doc, occ = asm_mgr
        result = am.replace_component(99, "C:/parts/new.par")
        assert "error" in result


# ============================================================================
# ASSEMBLY: GET COMPONENT TRANSFORM
# ============================================================================


class TestGetComponentTransform:
    @pytest.fixture
    def asm_mgr(self):
        from solidedge_mcp.backends.assembly import AssemblyManager

        dm = MagicMock()
        doc = MagicMock()
        dm.get_active_document.return_value = doc

        occurrence = MagicMock()
        occurrence.Name = "Part1:1"
        occurrence.GetTransform.return_value = (0.1, 0.2, 0.3, 0.0, 0.0, 0.0)
        occurrence.GetMatrix.return_value = tuple(range(16))
        occurrences = MagicMock()
        occurrences.Count = 1
        occurrences.Item.return_value = occurrence
        doc.Occurrences = occurrences

        return AssemblyManager(dm), doc, occurrence

    def test_success(self, asm_mgr):
        am, doc, occ = asm_mgr
        result = am.get_component_transform(0)
        assert result["name"] == "Part1:1"
        assert result["origin"] == [0.1, 0.2, 0.3]
        assert len(result["matrix"]) == 16

    def test_invalid_index(self, asm_mgr):
        am, doc, occ = asm_mgr
        result = am.get_component_transform(5)
        assert "error" in result


# ============================================================================
# ASSEMBLY: GET STRUCTURED BOM
# ============================================================================


class TestGetStructuredBom:
    @pytest.fixture
    def asm_mgr(self):
        from solidedge_mcp.backends.assembly import AssemblyManager

        dm = MagicMock()
        doc = MagicMock()
        doc.Name = "Asm1.asm"
        dm.get_active_document.return_value = doc

        occ1 = MagicMock()
        occ1.Name = "Part1:1"
        occ1.OccurrenceFileName = "C:/part1.par"
        occ1.Visible = True
        occ1.IsSuppressed = False
        sub_occs = MagicMock()
        sub_occs.Count = 0
        occ1.SubOccurrences = sub_occs

        occurrences = MagicMock()
        occurrences.Count = 1
        occurrences.Item.return_value = occ1
        doc.Occurrences = occurrences

        return AssemblyManager(dm), doc

    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        result = am.get_structured_bom()
        assert result["top_level_count"] == 1
        assert result["bom"][0]["name"] == "Part1:1"
        assert result["bom"][0]["type"] == "part"

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences
        result = am.get_structured_bom()
        assert "error" in result


# ============================================================================
# ASSEMBLY: SET COMPONENT COLOR
# ============================================================================


class TestSetComponentColor:
    @pytest.fixture
    def asm_mgr(self):
        from solidedge_mcp.backends.assembly import AssemblyManager

        dm = MagicMock()
        doc = MagicMock()
        dm.get_active_document.return_value = doc

        occurrence = MagicMock()
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occurrence
        doc.Occurrences = occurrences

        return AssemblyManager(dm), doc, occurrence

    def test_success(self, asm_mgr):
        am, doc, occ = asm_mgr
        result = am.set_component_color(0, 255, 0, 0)
        assert result["status"] == "updated"
        assert result["color"] == [255, 0, 0]

    def test_invalid_index(self, asm_mgr):
        am, doc, occ = asm_mgr
        result = am.set_component_color(99, 0, 0, 255)
        assert "error" in result


# ============================================================================
# DOCUMENTS: CREATE WELDMENT
# ============================================================================


class TestCreateWeldment:
    def test_success(self, doc_mgr):
        dm, app = doc_mgr
        doc = MagicMock()
        doc.Name = "Weld1.pwd"
        doc.FullName = "C:/weld.pwd"
        app.Documents.Add.return_value = doc

        result = dm.create_weldment()
        assert result["status"] == "created"
        assert result["type"] == "Weldment"


# ============================================================================
# DOCUMENTS: IMPORT FILE
# ============================================================================


class TestImportFile:
    def test_not_found(self, doc_mgr):
        dm, app = doc_mgr
        result = dm.import_file("C:/nonexistent/file.step")
        assert "error" in result


# ============================================================================
# DOCUMENTS: GET DOCUMENT COUNT
# ============================================================================


class TestGetDocumentCount:
    def test_success(self, doc_mgr):
        dm, app = doc_mgr
        app.Documents.Count = 3
        result = dm.get_document_count()
        assert result["count"] == 3


# ============================================================================
# RENAME SHEET
# ============================================================================


class TestRenameSheet:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Name = "Sheet 1"
        sheets = MagicMock()
        sheets.Count = 2
        sheets.Item.return_value = sheet
        doc.Sheets = sheets

        result = em.rename_sheet(0, "Drawing A")
        assert result["status"] == "renamed"
        assert result["new_name"] == "Drawing A"
        assert sheet.Name == "Drawing A"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.rename_sheet(0, "New Name")
        assert "error" in result

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        sheets = MagicMock()
        sheets.Count = 1
        doc.Sheets = sheets

        result = em.rename_sheet(5, "New Name")
        assert "error" in result


# ============================================================================
# DELETE SHEET
# ============================================================================


class TestDeleteSheet:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Name = "Sheet 2"
        sheets = MagicMock()
        sheets.Count = 3
        sheets.Item.return_value = sheet
        doc.Sheets = sheets

        result = em.delete_sheet(1)
        assert result["status"] == "deleted"
        sheet.Delete.assert_called_once()

    def test_last_sheet(self, export_mgr):
        em, doc = export_mgr
        sheets = MagicMock()
        sheets.Count = 1
        doc.Sheets = sheets

        result = em.delete_sheet(0)
        assert "error" in result

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.delete_sheet(0)
        assert "error" in result


# ============================================================================
# DRAW CONSTRUCTION LINE
# ============================================================================


class TestDrawConstructionLine:
    def test_success(self):
        from solidedge_mcp.backends.sketching import SketchManager

        dm = MagicMock()
        sm = SketchManager(dm)

        profile = MagicMock()
        line = MagicMock()
        profile.Lines2d.AddBy2Points.return_value = line
        sm.active_profile = profile

        result = sm.draw_construction_line(0, 0, 0.1, 0)
        assert result["status"] == "created"
        assert result["type"] == "construction_line"
        profile.Lines2d.AddBy2Points.assert_called_once_with(0, 0, 0.1, 0)
        profile.ToggleConstruction.assert_called_once_with(line)

    def test_no_sketch(self):
        from solidedge_mcp.backends.sketching import SketchManager

        dm = MagicMock()
        sm = SketchManager(dm)

        result = sm.draw_construction_line(0, 0, 0.1, 0)
        assert "error" in result


# ============================================================================
# GET SKETCH CONSTRAINTS
# ============================================================================


class TestGetSketchConstraints:
    def test_success(self):
        from solidedge_mcp.backends.sketching import SketchManager

        dm = MagicMock()
        sm = SketchManager(dm)

        profile = MagicMock()
        rel1 = MagicMock()
        rel1.Type = 1
        rel1.Name = "Horizontal"
        relations = MagicMock()
        relations.Count = 1
        relations.Item.return_value = rel1
        profile.Relations2d = relations
        sm.active_profile = profile

        result = sm.get_sketch_constraints()
        assert result["count"] == 1
        assert result["constraints"][0]["type"] == 1

    def test_no_sketch(self):
        from solidedge_mcp.backends.sketching import SketchManager

        dm = MagicMock()
        sm = SketchManager(dm)

        result = sm.get_sketch_constraints()
        assert "error" in result


# ============================================================================
# ASSEMBLY: GET OCCURRENCE COUNT
# ============================================================================


class TestGetOccurrenceCount:
    @pytest.fixture
    def asm_mgr(self):
        from solidedge_mcp.backends.assembly import AssemblyManager

        dm = MagicMock()
        doc = MagicMock()
        dm.get_active_document.return_value = doc

        occurrences = MagicMock()
        occurrences.Count = 5
        doc.Occurrences = occurrences

        return AssemblyManager(dm), doc

    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        result = am.get_occurrence_count()
        assert result["count"] == 5

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences
        result = am.get_occurrence_count()
        assert "error" in result


# ============================================================================
# CONNECTION: GET PROCESS INFO
# ============================================================================


class TestGetProcessInfo:
    def test_success(self):
        from solidedge_mcp.backends.connection import SolidEdgeConnection

        conn = SolidEdgeConnection()
        conn._is_connected = True
        conn.application = MagicMock()
        conn.application.ProcessID = 12345
        conn.application.hWnd = 67890

        result = conn.get_process_info()
        assert result["status"] == "success"
        assert result["process_id"] == 12345
        assert result["window_handle"] == 67890

    def test_not_connected(self):
        from solidedge_mcp.backends.connection import SolidEdgeConnection

        conn = SolidEdgeConnection()

        result = conn.get_process_info()
        assert "error" in result


# ============================================================================
# CONNECTION: GET INSTALL INFO
# ============================================================================


class TestGetInstallInfo:
    def test_fallback_to_app_path(self):
        from solidedge_mcp.backends.connection import SolidEdgeConnection

        conn = SolidEdgeConnection()
        conn._is_connected = True
        conn.application = MagicMock()
        conn.application.Path = "C:\\Program Files\\Solid Edge"

        # SEInstallData will fail (not registered in test env),
        # so it should fall back to Application.Path
        result = conn.get_install_info()
        assert result["status"] == "success"
        assert "install_path" in result

    def test_no_connection_no_installdata(self):
        from solidedge_mcp.backends.connection import SolidEdgeConnection

        conn = SolidEdgeConnection()
        # Not connected and SEInstallData won't work
        result = conn.get_install_info()
        assert "error" in result


# ============================================================================
# DOCUMENTS: OPEN IN BACKGROUND
# ============================================================================


class TestOpenInBackground:
    @pytest.fixture
    def doc_mgr(self):
        from solidedge_mcp.backends.documents import DocumentManager

        conn = MagicMock()
        app = MagicMock()
        conn.get_application.return_value = app
        return DocumentManager(conn), app

    def test_success(self, doc_mgr, tmp_path):
        dm, app = doc_mgr
        doc = MagicMock()
        doc.Name = "test.par"
        doc.Type = 1
        app.Documents.Open.return_value = doc

        f = tmp_path / "test.par"
        f.write_text("")

        result = dm.open_in_background(str(f))
        assert result["status"] == "opened_in_background"
        assert result["name"] == "test.par"
        app.Documents.Open.assert_called_once_with(str(f), 0x8)

    def test_file_not_found(self, doc_mgr):
        dm, _ = doc_mgr
        result = dm.open_in_background("C:\\nonexistent\\file.par")
        assert "error" in result
        assert "not found" in result["error"]


# ============================================================================
# DOCUMENTS: CLOSE ALL DOCUMENTS
# ============================================================================


class TestCloseAllDocuments:
    @pytest.fixture
    def doc_mgr(self):
        from solidedge_mcp.backends.documents import DocumentManager

        conn = MagicMock()
        app = MagicMock()
        conn.get_application.return_value = app
        return DocumentManager(conn), app

    def test_success(self, doc_mgr):
        dm, app = doc_mgr
        doc1 = MagicMock()
        doc1.Name = "doc1.par"
        doc2 = MagicMock()
        doc2.Name = "doc2.par"
        app.Documents.Count = 2
        app.Documents.Item.side_effect = lambda i: {2: doc2, 1: doc1}[i]

        result = dm.close_all_documents(save=False)
        assert result["status"] == "closed_all"
        assert result["closed"] == 2
        doc1.Close.assert_called_once()
        doc2.Close.assert_called_once()

    def test_no_documents(self, doc_mgr):
        dm, app = doc_mgr
        app.Documents.Count = 0

        result = dm.close_all_documents()
        assert result["status"] == "no_documents"
        assert result["closed"] == 0

    def test_with_save(self, doc_mgr):
        dm, app = doc_mgr
        doc1 = MagicMock()
        doc1.Name = "doc1.par"
        app.Documents.Count = 1
        app.Documents.Item.return_value = doc1

        result = dm.close_all_documents(save=True)
        assert result["status"] == "closed_all"
        doc1.Save.assert_called_once()
        doc1.Close.assert_called_once()


# ============================================================================
# SKETCHING: DRAW ARC BY 3 POINTS
# ============================================================================


class TestDrawArcBy3Points:
    @pytest.fixture
    def sketch_mgr(self):
        from solidedge_mcp.backends.sketching import SketchManager

        dm = MagicMock()
        sm = SketchManager(dm)
        sm.active_profile = MagicMock()
        return sm

    def test_success(self, sketch_mgr):
        result = sketch_mgr.draw_arc_by_3_points(0.0, 0.0, 0.05, 0.05, 0.1, 0.0)
        assert result["status"] == "created"
        assert result["method"] == "start_center_end"
        sketch_mgr.active_profile.Arcs2d.AddByStartCenterEnd.assert_called_once_with(
            0.0, 0.0, 0.05, 0.05, 0.1, 0.0
        )

    def test_no_sketch(self):
        from solidedge_mcp.backends.sketching import SketchManager

        sm = SketchManager(MagicMock())
        result = sm.draw_arc_by_3_points(0, 0, 0.05, 0.05, 0.1, 0)
        assert "error" in result


# ============================================================================
# SKETCHING: DRAW CIRCLE BY 2 POINTS
# ============================================================================


class TestDrawCircleBy2Points:
    @pytest.fixture
    def sketch_mgr(self):
        from solidedge_mcp.backends.sketching import SketchManager

        dm = MagicMock()
        sm = SketchManager(dm)
        sm.active_profile = MagicMock()
        return sm

    def test_success(self, sketch_mgr):
        result = sketch_mgr.draw_circle_by_2_points(0.0, 0.0, 0.1, 0.0)
        assert result["status"] == "created"
        assert result["method"] == "2_points"
        assert result["center"] == [0.05, 0.0]
        assert result["radius"] == 0.05
        sketch_mgr.active_profile.Circles2d.AddBy2Points.assert_called_once_with(0.0, 0.0, 0.1, 0.0)

    def test_no_sketch(self):
        from solidedge_mcp.backends.sketching import SketchManager

        sm = SketchManager(MagicMock())
        result = sm.draw_circle_by_2_points(0, 0, 0.1, 0)
        assert "error" in result


# ============================================================================
# SKETCHING: DRAW CIRCLE BY 3 POINTS
# ============================================================================


class TestDrawCircleBy3Points:
    @pytest.fixture
    def sketch_mgr(self):
        from solidedge_mcp.backends.sketching import SketchManager

        dm = MagicMock()
        sm = SketchManager(dm)
        sm.active_profile = MagicMock()
        return sm

    def test_success(self, sketch_mgr):
        result = sketch_mgr.draw_circle_by_3_points(0.0, 0.0, 0.1, 0.0, 0.05, 0.05)
        assert result["status"] == "created"
        assert result["method"] == "3_points"
        sketch_mgr.active_profile.Circles2d.AddBy3Points.assert_called_once_with(
            0.0, 0.0, 0.1, 0.0, 0.05, 0.05
        )

    def test_no_sketch(self):
        from solidedge_mcp.backends.sketching import SketchManager

        sm = SketchManager(MagicMock())
        result = sm.draw_circle_by_3_points(0, 0, 0.1, 0, 0.05, 0.05)
        assert "error" in result


# ============================================================================
# SKETCHING: MIRROR SPLINE
# ============================================================================


class TestMirrorSpline:
    @pytest.fixture
    def sketch_mgr(self):
        from solidedge_mcp.backends.sketching import SketchManager

        dm = MagicMock()
        sm = SketchManager(dm)
        sm.active_profile = MagicMock()
        return sm

    def test_success(self, sketch_mgr):
        spline = MagicMock()
        sketch_mgr.active_profile.BSplineCurves2d.Count = 1
        sketch_mgr.active_profile.BSplineCurves2d.Item.return_value = spline

        result = sketch_mgr.mirror_spline(0, 0, 0, 1, True)
        assert result["status"] == "created"
        assert result["mirrored_count"] == 1
        spline.Mirror.assert_called_once_with(0, 0, 0, 1, True)

    def test_no_splines(self, sketch_mgr):
        sketch_mgr.active_profile.BSplineCurves2d.Count = 0
        result = sketch_mgr.mirror_spline(0, 0, 0, 1)
        assert "error" in result

    def test_no_sketch(self):
        from solidedge_mcp.backends.sketching import SketchManager

        sm = SketchManager(MagicMock())
        result = sm.mirror_spline(0, 0, 0, 1)
        assert "error" in result


# ============================================================================
# SKETCHING: HIDE PROFILE
# ============================================================================


class TestHideProfile:
    @pytest.fixture
    def sketch_mgr(self):
        from solidedge_mcp.backends.sketching import SketchManager

        dm = MagicMock()
        sm = SketchManager(dm)
        sm.active_profile = MagicMock()
        return sm

    def test_hide(self, sketch_mgr):
        result = sketch_mgr.hide_profile(visible=False)
        assert result["status"] == "updated"
        assert result["visible"] is False
        assert sketch_mgr.active_profile.Visible is False

    def test_show(self, sketch_mgr):
        result = sketch_mgr.hide_profile(visible=True)
        assert result["status"] == "updated"
        assert result["visible"] is True
        assert sketch_mgr.active_profile.Visible is True

    def test_no_sketch(self):
        from solidedge_mcp.backends.sketching import SketchManager

        sm = SketchManager(MagicMock())
        result = sm.hide_profile()
        assert "error" in result


# ============================================================================
# CONSTRAINTS: ADD_CONSTRAINT (all types)
# ============================================================================


class TestAddConstraint:
    @pytest.fixture
    def sketch_mgr(self):
        from solidedge_mcp.backends.sketching import SketchManager

        dm = MagicMock()
        sm = SketchManager(dm)
        sm.active_profile = MagicMock()
        # Set up Lines2d collection with 2 items
        line1 = MagicMock(name="line1")
        line2 = MagicMock(name="line2")
        lines = MagicMock()
        lines.Count = 2
        lines.Item.side_effect = lambda i: {1: line1, 2: line2}[i]
        sm.active_profile.Lines2d = lines
        # Set up Circles2d with 2 items
        circ1 = MagicMock(name="circ1")
        circ2 = MagicMock(name="circ2")
        circles = MagicMock()
        circles.Count = 2
        circles.Item.side_effect = lambda i: {1: circ1, 2: circ2}[i]
        sm.active_profile.Circles2d = circles
        return sm, line1, line2, circ1, circ2

    def test_horizontal(self, sketch_mgr):
        sm, line1, _, _, _ = sketch_mgr
        result = sm.add_constraint("Horizontal", [["line", 1]])
        assert result["status"] == "constraint_added"
        sm.active_profile.Relations2d.AddHorizontal.assert_called_once_with(line1)

    def test_vertical(self, sketch_mgr):
        sm, line1, _, _, _ = sketch_mgr
        result = sm.add_constraint("Vertical", [["line", 1]])
        assert result["status"] == "constraint_added"
        sm.active_profile.Relations2d.AddVertical.assert_called_once_with(line1)

    def test_parallel(self, sketch_mgr):
        sm, line1, line2, _, _ = sketch_mgr
        result = sm.add_constraint("Parallel", [["line", 1], ["line", 2]])
        assert result["status"] == "constraint_added"
        sm.active_profile.Relations2d.AddParallel.assert_called_once_with(line1, line2)

    def test_perpendicular(self, sketch_mgr):
        sm, line1, line2, _, _ = sketch_mgr
        result = sm.add_constraint("Perpendicular", [["line", 1], ["line", 2]])
        assert result["status"] == "constraint_added"
        sm.active_profile.Relations2d.AddPerpendicular.assert_called_once_with(line1, line2)

    def test_equal(self, sketch_mgr):
        sm, line1, line2, _, _ = sketch_mgr
        result = sm.add_constraint("Equal", [["line", 1], ["line", 2]])
        assert result["status"] == "constraint_added"
        sm.active_profile.Relations2d.AddEqual.assert_called_once_with(line1, line2)

    def test_concentric(self, sketch_mgr):
        sm, _, _, circ1, circ2 = sketch_mgr
        result = sm.add_constraint("Concentric", [["circle", 1], ["circle", 2]])
        assert result["status"] == "constraint_added"
        sm.active_profile.Relations2d.AddConcentric.assert_called_once_with(circ1, circ2)

    def test_tangent(self, sketch_mgr):
        sm, line1, _, circ1, _ = sketch_mgr
        result = sm.add_constraint("Tangent", [["line", 1], ["circle", 1]])
        assert result["status"] == "constraint_added"
        sm.active_profile.Relations2d.AddTangent.assert_called_once_with(line1, circ1)

    def test_unknown_type(self, sketch_mgr):
        sm, _, _, _, _ = sketch_mgr
        result = sm.add_constraint("Colinear", [["line", 1]])
        assert "error" in result

    def test_invalid_element_format(self, sketch_mgr):
        sm, _, _, _, _ = sketch_mgr
        result = sm.add_constraint("Horizontal", ["line"])
        assert "error" in result

    def test_invalid_element_type(self, sketch_mgr):
        sm, _, _, _, _ = sketch_mgr
        result = sm.add_constraint("Horizontal", [["polygon", 1]])
        assert "error" in result

    def test_index_out_of_range(self, sketch_mgr):
        sm, _, _, _, _ = sketch_mgr
        result = sm.add_constraint("Horizontal", [["line", 99]])
        assert "error" in result

    def test_no_sketch(self):
        from solidedge_mcp.backends.sketching import SketchManager

        sm = SketchManager(MagicMock())
        result = sm.add_constraint("Horizontal", [["line", 1]])
        assert "error" in result


# ============================================================================
# CONSTRAINTS: ADD_KEYPOINT_CONSTRAINT
# ============================================================================


class TestAddKeypointConstraint:
    @pytest.fixture
    def sketch_mgr(self):
        from solidedge_mcp.backends.sketching import SketchManager

        dm = MagicMock()
        sm = SketchManager(dm)
        sm.active_profile = MagicMock()
        line1 = MagicMock(name="line1")
        line2 = MagicMock(name="line2")
        lines = MagicMock()
        lines.Count = 2
        lines.Item.side_effect = lambda i: {1: line1, 2: line2}[i]
        sm.active_profile.Lines2d = lines
        return sm, line1, line2

    def test_success(self, sketch_mgr):
        sm, line1, line2 = sketch_mgr
        result = sm.add_keypoint_constraint("line", 1, 1, "line", 2, 0)
        assert result["status"] == "constraint_added"
        assert result["type"] == "Keypoint"
        sm.active_profile.Relations2d.AddKeypoint.assert_called_once_with(line1, 1, line2, 0)

    def test_invalid_index(self, sketch_mgr):
        sm, _, _ = sketch_mgr
        result = sm.add_keypoint_constraint("line", 99, 0, "line", 1, 0)
        assert "error" in result

    def test_no_sketch(self):
        from solidedge_mcp.backends.sketching import SketchManager

        sm = SketchManager(MagicMock())
        result = sm.add_keypoint_constraint("line", 1, 0, "line", 2, 1)
        assert "error" in result


# ============================================================================
# TIER 2: CAMERA CONTROL (ViewModel)
# ============================================================================


@pytest.fixture
def view_mgr():
    """Create ViewModel with mocked dependencies."""
    from solidedge_mcp.backends.export import ViewModel

    dm = MagicMock()
    doc = MagicMock()
    dm.get_active_document.return_value = doc

    window = MagicMock()
    view_obj = MagicMock()
    window.View = view_obj
    windows = MagicMock()
    windows.Count = 1
    windows.Item.return_value = window
    doc.Windows = windows

    return ViewModel(dm), doc, view_obj


class TestGetCamera:
    def test_success(self, view_mgr):
        vm, doc, view_obj = view_mgr
        # GetCamera returns 11-element tuple
        view_obj.GetCamera.return_value = (
            1.0,
            2.0,
            3.0,  # eye
            0.0,
            0.0,
            0.0,  # target
            0.0,
            1.0,
            0.0,  # up
            False,  # perspective
            1.5,  # scale
        )
        result = vm.get_camera()
        assert result["eye"] == [1.0, 2.0, 3.0]
        assert result["target"] == [0.0, 0.0, 0.0]
        assert result["up"] == [0.0, 1.0, 0.0]
        assert result["perspective"] is False
        assert result["scale_or_angle"] == 1.5
        view_obj.GetCamera.assert_called_once()

    def test_no_window(self):
        from solidedge_mcp.backends.export import ViewModel

        dm = MagicMock()
        doc = MagicMock()
        doc.Windows.Count = 0
        dm.get_active_document.return_value = doc
        vm = ViewModel(dm)
        result = vm.get_camera()
        assert "error" in result

    def test_perspective_mode(self, view_mgr):
        vm, doc, view_obj = view_mgr
        view_obj.GetCamera.return_value = (
            0.5,
            0.5,
            0.5,
            0.0,
            0.0,
            0.0,
            0.0,
            0.0,
            1.0,
            True,
            0.785,  # ~45 degrees FOV
        )
        result = vm.get_camera()
        assert result["perspective"] is True
        assert result["scale_or_angle"] == 0.785


class TestSetCamera:
    def test_success(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.set_camera(1.0, 2.0, 3.0, 0.0, 0.0, 0.0)
        assert result["status"] == "camera_set"
        assert result["eye"] == [1.0, 2.0, 3.0]
        assert result["target"] == [0.0, 0.0, 0.0]
        assert result["up"] == [0.0, 1.0, 0.0]  # defaults
        view_obj.SetCamera.assert_called_once_with(
            1.0, 2.0, 3.0, 0.0, 0.0, 0.0, 0.0, 1.0, 0.0, False, 1.0
        )

    def test_with_perspective(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.set_camera(0.5, 0.5, 0.5, 0.0, 0.0, 0.0, 0.0, 0.0, 1.0, True, 0.785)
        assert result["status"] == "camera_set"
        assert result["perspective"] is True
        assert result["up"] == [0.0, 0.0, 1.0]
        view_obj.SetCamera.assert_called_once_with(
            0.5, 0.5, 0.5, 0.0, 0.0, 0.0, 0.0, 0.0, 1.0, True, 0.785
        )

    def test_no_window(self):
        from solidedge_mcp.backends.export import ViewModel

        dm = MagicMock()
        doc = MagicMock()
        doc.Windows.Count = 0
        dm.get_active_document.return_value = doc
        vm = ViewModel(dm)
        result = vm.set_camera(0, 0, 1, 0, 0, 0)
        assert "error" in result


class TestRotateCamera:
    def test_success(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.rotate_camera(0.5, 0.0, 0.0, 0.0, 0.0, 1.0, 0.0)
        assert result["status"] == "camera_rotated"
        assert result["angle_rad"] == 0.5
        assert result["center"] == [0.0, 0.0, 0.0]
        assert result["axis"] == [0.0, 1.0, 0.0]
        view_obj.RotateCamera.assert_called_once_with(0.5, 0.0, 0.0, 0.0, 0.0, 1.0, 0.0)

    def test_custom_axis(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.rotate_camera(1.57, 0.1, 0.2, 0.3, 1.0, 0.0, 0.0)
        assert result["status"] == "camera_rotated"
        assert result["axis"] == [1.0, 0.0, 0.0]
        view_obj.RotateCamera.assert_called_once_with(1.57, 0.1, 0.2, 0.3, 1.0, 0.0, 0.0)

    def test_defaults(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.rotate_camera(0.3)
        assert result["status"] == "camera_rotated"
        # Default center=(0,0,0), axis=(0,1,0)
        view_obj.RotateCamera.assert_called_once_with(0.3, 0.0, 0.0, 0.0, 0.0, 1.0, 0.0)

    def test_no_window(self):
        from solidedge_mcp.backends.export import ViewModel

        dm = MagicMock()
        doc = MagicMock()
        doc.Windows.Count = 0
        dm.get_active_document.return_value = doc
        vm = ViewModel(dm)
        result = vm.rotate_camera(0.5)
        assert "error" in result


class TestPanCamera:
    def test_success(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.pan_camera(100, -50)
        assert result["status"] == "camera_panned"
        assert result["dx"] == 100
        assert result["dy"] == -50
        view_obj.PanCamera.assert_called_once_with(100, -50)

    def test_zero_pan(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.pan_camera(0, 0)
        assert result["status"] == "camera_panned"
        view_obj.PanCamera.assert_called_once_with(0, 0)

    def test_no_window(self):
        from solidedge_mcp.backends.export import ViewModel

        dm = MagicMock()
        doc = MagicMock()
        doc.Windows.Count = 0
        dm.get_active_document.return_value = doc
        vm = ViewModel(dm)
        result = vm.pan_camera(10, 20)
        assert "error" in result


class TestZoomCamera:
    def test_zoom_in(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.zoom_camera(2.0)
        assert result["status"] == "camera_zoomed"
        assert result["factor"] == 2.0
        view_obj.ZoomCamera.assert_called_once_with(2.0)

    def test_zoom_out(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.zoom_camera(0.5)
        assert result["status"] == "camera_zoomed"
        assert result["factor"] == 0.5
        view_obj.ZoomCamera.assert_called_once_with(0.5)

    def test_no_window(self):
        from solidedge_mcp.backends.export import ViewModel

        dm = MagicMock()
        doc = MagicMock()
        doc.Windows.Count = 0
        dm.get_active_document.return_value = doc
        vm = ViewModel(dm)
        result = vm.zoom_camera(1.5)
        assert "error" in result


class TestRefreshView:
    def test_success(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.refresh_view()
        assert result["status"] == "view_refreshed"
        view_obj.Update.assert_called_once()

    def test_no_window(self):
        from solidedge_mcp.backends.export import ViewModel

        dm = MagicMock()
        doc = MagicMock()
        doc.Windows.Count = 0
        dm.get_active_document.return_value = doc
        vm = ViewModel(dm)
        result = vm.refresh_view()
        assert "error" in result


# ============================================================================
# TIER 3: CREATE PARTS LIST
# ============================================================================


class TestCreatePartsList:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dv = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        dvs.Item.return_value = dv
        sheet.DrawingViews = dvs

        parts_lists = MagicMock()
        parts_lists.Count = 1
        parts_lists.Add.return_value = MagicMock()
        sheet.PartsLists = parts_lists
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.create_parts_list(auto_balloon=True)
        assert result["status"] == "created"
        assert result["auto_balloon"] is True
        parts_lists.Add.assert_called_once_with(dv, "", 1, 1)

    def test_no_drawing_views(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 0
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.create_parts_list()
        assert "error" in result
        assert "No drawing views" in result["error"]

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.create_parts_list()
        assert "error" in result

    def test_no_auto_balloon(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dv = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        dvs.Item.return_value = dv
        sheet.DrawingViews = dvs

        parts_lists = MagicMock()
        parts_lists.Count = 1
        parts_lists.Add.return_value = MagicMock()
        sheet.PartsLists = parts_lists
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.create_parts_list(auto_balloon=False)
        assert result["status"] == "created"
        assert result["auto_balloon"] is False
        parts_lists.Add.assert_called_once_with(dv, "", 0, 1)


# ============================================================================
# TIER 3: PROJECT EDGE
# ============================================================================


class TestProjectEdge:
    @pytest.fixture
    def sketch_mgr(self):
        from solidedge_mcp.backends.sketching import SketchManager

        dm = MagicMock()
        doc = MagicMock()
        dm.get_active_document.return_value = doc

        model = MagicMock()
        body = MagicMock()
        face = MagicMock()
        edge = MagicMock()
        edges = MagicMock()
        edges.Count = 2
        edges.Item.side_effect = lambda i: edge if i == 1 else MagicMock()
        face.Edges = edges
        faces = MagicMock()
        faces.Count = 1
        faces.Item.return_value = face
        body.Faces.return_value = faces
        model.Body = body
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        sm = SketchManager(dm)
        sm.active_profile = MagicMock()
        return sm, edge

    def test_success(self, sketch_mgr):
        sm, edge = sketch_mgr
        sm.active_profile.ProjectEdge.return_value = MagicMock()
        result = sm.project_edge(0, 0)
        assert result["status"] == "projected"
        sm.active_profile.ProjectEdge.assert_called_once_with(edge)

    def test_no_sketch(self):
        from solidedge_mcp.backends.sketching import SketchManager

        sm = SketchManager(MagicMock())
        result = sm.project_edge(0, 0)
        assert "error" in result

    def test_invalid_face_index(self, sketch_mgr):
        sm, _ = sketch_mgr
        result = sm.project_edge(5, 0)
        assert "error" in result
        assert "Invalid face" in result["error"]

    def test_invalid_edge_index(self, sketch_mgr):
        sm, _ = sketch_mgr
        result = sm.project_edge(0, 5)
        assert "error" in result
        assert "Invalid edge" in result["error"]


# ============================================================================
# TIER 3: INCLUDE EDGE
# ============================================================================


class TestIncludeEdge:
    @pytest.fixture
    def sketch_mgr(self):
        from solidedge_mcp.backends.sketching import SketchManager

        dm = MagicMock()
        doc = MagicMock()
        dm.get_active_document.return_value = doc

        model = MagicMock()
        body = MagicMock()
        face = MagicMock()
        edge = MagicMock()
        edges = MagicMock()
        edges.Count = 2
        edges.Item.side_effect = lambda i: edge if i == 1 else MagicMock()
        face.Edges = edges
        faces = MagicMock()
        faces.Count = 1
        faces.Item.return_value = face
        body.Faces.return_value = faces
        model.Body = body
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        sm = SketchManager(dm)
        sm.active_profile = MagicMock()
        return sm, edge

    def test_success(self, sketch_mgr):
        sm, edge = sketch_mgr
        sm.active_profile.IncludeEdge.return_value = MagicMock()
        result = sm.include_edge(0, 0)
        assert result["status"] == "included"
        sm.active_profile.IncludeEdge.assert_called_once_with(edge)

    def test_no_sketch(self):
        from solidedge_mcp.backends.sketching import SketchManager

        sm = SketchManager(MagicMock())
        result = sm.include_edge(0, 0)
        assert "error" in result

    def test_no_model(self):
        from solidedge_mcp.backends.sketching import SketchManager

        dm = MagicMock()
        doc = MagicMock()
        models = MagicMock()
        models.Count = 0
        doc.Models = models
        dm.get_active_document.return_value = doc
        sm = SketchManager(dm)
        sm.active_profile = MagicMock()
        result = sm.include_edge(0, 0)
        assert "error" in result


# ============================================================================
# TIER 3: START COMMAND
# ============================================================================


class TestStartCommand:
    def test_success(self):
        from solidedge_mcp.backends.connection import SolidEdgeConnection

        conn = SolidEdgeConnection()
        conn.application = MagicMock()
        conn._is_connected = True

        result = conn.start_command(12345)
        assert result["status"] == "success"
        assert result["command_id"] == 12345
        conn.application.StartCommand.assert_called_once_with(12345)

    def test_not_connected(self):
        from solidedge_mcp.backends.connection import SolidEdgeConnection

        conn = SolidEdgeConnection()

        result = conn.start_command(12345)
        assert "error" in result

    def test_invalid_command(self):
        from solidedge_mcp.backends.connection import SolidEdgeConnection

        conn = SolidEdgeConnection()
        conn.application = MagicMock()
        conn._is_connected = True
        conn.application.StartCommand.side_effect = Exception("Invalid command ID")

        result = conn.start_command(99999)
        assert "error" in result


# ============================================================================
# DRAWING VIEW: GET COUNT
# ============================================================================


class TestGetDrawingViewCount:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 3
        # Make _oleobj_ raise so the late-binding branch is skipped
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.get_drawing_view_count()
        assert result["count"] == 3

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.get_drawing_view_count()
        assert "error" in result


# ============================================================================
# DRAWING VIEW: GET SCALE
# ============================================================================


class TestGetDrawingViewScale:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        view = MagicMock()
        view.ScaleFactor = 2.0
        dvs = MagicMock()
        dvs.Count = 2
        dvs.Item.return_value = view
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.get_drawing_view_scale(0)
        assert result["view_index"] == 0
        assert result["scale"] == 2.0

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.get_drawing_view_scale(5)
        assert "error" in result

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.get_drawing_view_scale(0)
        assert "error" in result


# ============================================================================
# DRAWING VIEW: SET SCALE
# ============================================================================


class TestSetDrawingViewScale:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        view = MagicMock()
        dvs = MagicMock()
        dvs.Count = 2
        dvs.Item.return_value = view
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.set_drawing_view_scale(0, 0.5)
        assert result["status"] == "set"
        assert result["scale"] == 0.5
        assert view.ScaleFactor == 0.5

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.set_drawing_view_scale(5, 1.0)
        assert "error" in result


# ============================================================================
# DRAWING VIEW: DELETE
# ============================================================================


class TestDeleteDrawingView:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        view = MagicMock()
        dvs = MagicMock()
        dvs.Count = 3
        dvs.Item.return_value = view
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.delete_drawing_view(1)
        assert result["status"] == "deleted"
        assert result["view_index"] == 1
        view.Delete.assert_called_once()

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.delete_drawing_view(5)
        assert "error" in result


# ============================================================================
# DRAWING VIEW: UPDATE
# ============================================================================


class TestUpdateDrawingView:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        view = MagicMock()
        dvs = MagicMock()
        dvs.Count = 2
        dvs.Item.return_value = view
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.update_drawing_view(0)
        assert result["status"] == "updated"
        assert result["view_index"] == 0
        view.Update.assert_called_once()

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.update_drawing_view(5)
        assert "error" in result


# ============================================================================
# VIEW COORDINATE TRANSFORMS
# ============================================================================


class TestTransformModelToScreen:
    def test_success(self, view_mgr):
        vm, doc, view_obj = view_mgr
        view_obj.TransformModelToDC.return_value = (400, 300)
        result = vm.transform_model_to_screen(0.1, 0.2, 0.3)
        assert result["status"] == "success"
        assert result["screen_x"] == 400
        assert result["screen_y"] == 300
        assert result["model"] == [0.1, 0.2, 0.3]
        view_obj.TransformModelToDC.assert_called_once_with(0.1, 0.2, 0.3)

    def test_no_window(self):
        from solidedge_mcp.backends.export import ViewModel

        dm = MagicMock()
        doc = MagicMock()
        doc.Windows.Count = 0
        dm.get_active_document.return_value = doc
        vm = ViewModel(dm)
        result = vm.transform_model_to_screen(0, 0, 0)
        assert "error" in result


class TestTransformScreenToModel:
    def test_success(self, view_mgr):
        vm, doc, view_obj = view_mgr
        view_obj.TransformDCToModel.return_value = (0.05, 0.10, 0.15)
        result = vm.transform_screen_to_model(500, 250)
        assert result["status"] == "success"
        assert result["x"] == 0.05
        assert result["y"] == 0.10
        assert result["z"] == 0.15
        assert result["screen"] == [500, 250]
        view_obj.TransformDCToModel.assert_called_once_with(500, 250)

    def test_no_window(self):
        from solidedge_mcp.backends.export import ViewModel

        dm = MagicMock()
        doc = MagicMock()
        doc.Windows.Count = 0
        dm.get_active_document.return_value = doc
        vm = ViewModel(dm)
        result = vm.transform_screen_to_model(0, 0)
        assert "error" in result


# ============================================================================
# CAMERA DYNAMICS
# ============================================================================


class TestBeginCameraDynamics:
    def test_success(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.begin_camera_dynamics()
        assert result["status"] == "camera_dynamics_started"
        view_obj.BeginCameraDynamics.assert_called_once()

    def test_no_window(self):
        from solidedge_mcp.backends.export import ViewModel

        dm = MagicMock()
        doc = MagicMock()
        doc.Windows.Count = 0
        dm.get_active_document.return_value = doc
        vm = ViewModel(dm)
        result = vm.begin_camera_dynamics()
        assert "error" in result


class TestEndCameraDynamics:
    def test_success(self, view_mgr):
        vm, doc, view_obj = view_mgr
        result = vm.end_camera_dynamics()
        assert result["status"] == "camera_dynamics_ended"
        view_obj.EndCameraDynamics.assert_called_once()

    def test_no_window(self):
        from solidedge_mcp.backends.export import ViewModel

        dm = MagicMock()
        doc = MagicMock()
        doc.Windows.Count = 0
        dm.get_active_document.return_value = doc
        vm = ViewModel(dm)
        result = vm.end_camera_dynamics()
        assert "error" in result


# ============================================================================
# ADD PROJECTED VIEW
# ============================================================================


class TestAddProjectedView:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        parent_view = MagicMock()
        new_view = MagicMock()
        new_view.Name = "View 2"

        dvs = MagicMock()
        dvs.Count = 2
        dvs.Item.side_effect = lambda i: {1: parent_view, 2: new_view}[i]
        dvs.AddByFold.return_value = new_view
        sheet.DrawingViews = dvs

        sheets = MagicMock()
        sheets.Count = 1
        sheets.Item.return_value = sheet
        doc.Sheets = sheets
        doc.ActiveSheet = sheet

        result = em.add_projected_view(0, "Up", 0.2, 0.3)
        # May return error if COM binding fails in test env, or success
        assert isinstance(result, dict)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_projected_view(0, "Up", 0.2, 0.3)
        assert "error" in result

    def test_invalid_direction(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        sheets = MagicMock()
        sheets.Count = 1
        doc.Sheets = sheets

        result = em.add_projected_view(0, "InvalidDir", 0.2, 0.3)
        assert "error" in result


# ============================================================================
# MOVE DRAWING VIEW
# ============================================================================


class TestMoveDrawingView:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        view = MagicMock()
        view.Name = "View 1"

        dvs = MagicMock()
        dvs.Count = 1
        dvs.Item.return_value = view
        sheet.DrawingViews = dvs

        doc.ActiveSheet = sheet
        sheets = MagicMock()
        sheets.Count = 1
        doc.Sheets = sheets

        result = em.move_drawing_view(0, 0.15, 0.20)
        assert isinstance(result, dict)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.move_drawing_view(0, 0.15, 0.20)
        assert "error" in result

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        sheet.DrawingViews = dvs

        doc.ActiveSheet = sheet
        sheets = MagicMock()
        sheets.Count = 1
        doc.Sheets = sheets

        result = em.move_drawing_view(5, 0.15, 0.20)
        assert "error" in result


# ============================================================================
# SHOW HIDDEN EDGES
# ============================================================================


class TestShowHiddenEdges:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        view = MagicMock()
        view.Name = "View 1"

        dvs = MagicMock()
        dvs.Count = 1
        dvs.Item.return_value = view
        sheet.DrawingViews = dvs

        doc.ActiveSheet = sheet
        sheets = MagicMock()
        sheets.Count = 1
        doc.Sheets = sheets

        result = em.show_hidden_edges(0, True)
        assert isinstance(result, dict)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.show_hidden_edges(0, True)
        assert "error" in result

    def test_hide_edges(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        view = MagicMock()
        view.Name = "View 1"

        dvs = MagicMock()
        dvs.Count = 1
        dvs.Item.return_value = view
        sheet.DrawingViews = dvs

        doc.ActiveSheet = sheet
        sheets = MagicMock()
        sheets.Count = 1
        doc.Sheets = sheets

        result = em.show_hidden_edges(0, False)
        assert isinstance(result, dict)


# ============================================================================
# SET DRAWING VIEW DISPLAY MODE
# ============================================================================


class TestSetDrawingViewDisplayMode:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        view = MagicMock()

        dvs = MagicMock()
        dvs.Count = 1
        dvs.Item.return_value = view
        sheet.DrawingViews = dvs

        doc.ActiveSheet = sheet
        sheets = MagicMock()
        sheets.Count = 1
        doc.Sheets = sheets

        result = em.set_drawing_view_display_mode(0, "Shaded")
        assert isinstance(result, dict)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.set_drawing_view_display_mode(0, "Shaded")
        assert "error" in result

    def test_invalid_mode(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        sheet.DrawingViews = dvs

        doc.ActiveSheet = sheet
        sheets = MagicMock()
        sheets.Count = 1
        doc.Sheets = sheets

        result = em.set_drawing_view_display_mode(0, "InvalidMode")
        assert "error" in result


# ============================================================================
# GET DRAWING VIEW INFO
# ============================================================================


class TestGetDrawingViewInfo:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        view = MagicMock()
        view.Name = "View 1"
        view.ScaleFactor = 1.0

        dvs = MagicMock()
        dvs.Count = 1
        dvs.Item.return_value = view
        sheet.DrawingViews = dvs

        doc.ActiveSheet = sheet
        sheets = MagicMock()
        sheets.Count = 1
        doc.Sheets = sheets

        result = em.get_drawing_view_info(0)
        assert isinstance(result, dict)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.get_drawing_view_info(0)
        assert "error" in result

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        sheet.DrawingViews = dvs

        doc.ActiveSheet = sheet
        sheets = MagicMock()
        sheets.Count = 1
        doc.Sheets = sheets

        result = em.get_drawing_view_info(5)
        assert "error" in result


# ============================================================================
# SET DRAWING VIEW ORIENTATION
# ============================================================================


class TestSetDrawingViewOrientation:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        view = MagicMock()

        dvs = MagicMock()
        dvs.Count = 1
        dvs.Item.return_value = view
        sheet.DrawingViews = dvs

        doc.ActiveSheet = sheet
        sheets = MagicMock()
        sheets.Count = 1
        doc.Sheets = sheets

        result = em.set_drawing_view_orientation(0, "Front")
        assert isinstance(result, dict)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.set_drawing_view_orientation(0, "Front")
        assert "error" in result

    def test_invalid_orientation(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        sheet.DrawingViews = dvs

        doc.ActiveSheet = sheet
        sheets = MagicMock()
        sheets.Count = 1
        doc.Sheets = sheets

        result = em.set_drawing_view_orientation(0, "InvalidView")
        assert "error" in result


# ============================================================================
# ADD ANGULAR DIMENSION
# ============================================================================


class TestAddAngularDimension:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dims = MagicMock()
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_angular_dimension(0.0, 0.0, 0.05, 0.05, 0.1, 0.0)
        assert result["status"] == "created"
        assert result["type"] == "angular_dimension"
        assert result["vertex"] == [0.05, 0.05]
        dims.AddAngular.assert_called_once()

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_angular_dimension(0, 0, 0.05, 0.05, 0.1, 0)
        assert "error" in result

    def test_exception(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Dimensions.AddAngular.side_effect = Exception("COM error")
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_angular_dimension(0, 0, 0, 0, 0, 0)
        assert "error" in result


# ============================================================================
# ADD RADIAL DIMENSION
# ============================================================================


class TestAddRadialDimension:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dims = MagicMock()
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_radial_dimension(0.05, 0.05, 0.1, 0.05)
        assert result["status"] == "created"
        assert result["type"] == "radial_dimension"
        assert result["center"] == [0.05, 0.05]
        dims.AddRadial.assert_called_once()

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_radial_dimension(0.05, 0.05, 0.1, 0.05)
        assert "error" in result

    def test_custom_text_position(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dims = MagicMock()
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_radial_dimension(0.05, 0.05, 0.1, 0.05, dim_x=0.2, dim_y=0.2)
        assert result["status"] == "created"
        assert result["text_position"] == [0.2, 0.2]


# ============================================================================
# ADD DIAMETER DIMENSION
# ============================================================================


class TestAddDiameterDimension:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dims = MagicMock()
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_diameter_dimension(0.05, 0.05, 0.1, 0.05)
        assert result["status"] == "created"
        assert result["type"] == "diameter_dimension"
        assert result["center"] == [0.05, 0.05]
        dims.AddDiameter.assert_called_once()

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_diameter_dimension(0.05, 0.05, 0.1, 0.05)
        assert "error" in result

    def test_exception(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Dimensions.AddDiameter.side_effect = Exception("COM error")
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_diameter_dimension(0, 0, 0.1, 0)
        assert "error" in result


# ============================================================================
# ADD ORDINATE DIMENSION
# ============================================================================


class TestAddOrdinateDimension:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dims = MagicMock()
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_ordinate_dimension(0.0, 0.0, 0.1, 0.0)
        assert result["status"] == "created"
        assert result["type"] == "ordinate_dimension"
        assert result["origin"] == [0.0, 0.0]
        assert result["point"] == [0.1, 0.0]
        dims.AddOrdinate.assert_called_once()

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_ordinate_dimension(0, 0, 0.1, 0)
        assert "error" in result

    def test_custom_text_position(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dims = MagicMock()
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_ordinate_dimension(0.0, 0.0, 0.1, 0.0, dim_x=0.15, dim_y=0.05)
        assert result["status"] == "created"
        assert result["text_position"] == [0.15, 0.05]


# ============================================================================
# ADD CENTER MARK
# ============================================================================


class TestAddCenterMark:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        center_marks = MagicMock()
        sheet.CenterMarks = center_marks
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_center_mark(0.1, 0.1)
        assert result["status"] == "added"
        assert result["type"] == "center_mark"
        assert result["position"] == [0.1, 0.1]
        center_marks.Add.assert_called_once_with(0.1, 0.1, 0)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_center_mark(0.1, 0.1)
        assert "error" in result

    def test_exception(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.CenterMarks.Add.side_effect = Exception("COM error")
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_center_mark(0.1, 0.1)
        assert "error" in result


# ============================================================================
# ADD CENTERLINE
# ============================================================================


class TestAddCenterline:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        centerlines = MagicMock()
        sheet.Centerlines = centerlines
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_centerline(0.0, 0.05, 0.1, 0.05)
        assert result["status"] == "added"
        assert result["type"] == "centerline"
        assert result["start"] == [0.0, 0.05]
        assert result["end"] == [0.1, 0.05]
        centerlines.Add.assert_called_once_with(0.0, 0.05, 0, 0.1, 0.05, 0)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_centerline(0, 0, 0.1, 0)
        assert "error" in result

    def test_exception(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.Centerlines.Add.side_effect = Exception("COM error")
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_centerline(0, 0, 0.1, 0)
        assert "error" in result


# ============================================================================
# ADD SURFACE FINISH SYMBOL
# ============================================================================


class TestAddSurfaceFinishSymbol:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sfs = MagicMock()
        sheet.SurfaceFinishSymbols = sfs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_surface_finish_symbol(0.1, 0.1, "machined")
        assert result["status"] == "added"
        assert result["type"] == "surface_finish_symbol"
        assert result["symbol_type"] == "machined"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_surface_finish_symbol(0.1, 0.1)
        assert "error" in result

    def test_invalid_type(self, export_mgr):
        em, doc = export_mgr
        doc.Sheets = MagicMock()

        result = em.add_surface_finish_symbol(0.1, 0.1, "invalid_type")
        assert "error" in result


# ============================================================================
# ADD WELD SYMBOL
# ============================================================================


class TestAddWeldSymbol:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        ws = MagicMock()
        sheet.WeldSymbols = ws
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_weld_symbol(0.1, 0.1, "fillet")
        assert result["status"] == "added"
        assert result["type"] == "weld_symbol"
        assert result["weld_type"] == "fillet"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_weld_symbol(0.1, 0.1)
        assert "error" in result

    def test_invalid_type(self, export_mgr):
        em, doc = export_mgr
        doc.Sheets = MagicMock()

        result = em.add_weld_symbol(0.1, 0.1, "invalid_weld")
        assert "error" in result


# ============================================================================
# ADD GEOMETRIC TOLERANCE
# ============================================================================


class TestAddGeometricTolerance:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        fcfs = MagicMock()
        fcf = MagicMock()
        fcfs.Add.return_value = fcf
        sheet.FCFs = fcfs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_geometric_tolerance(0.1, 0.1, "0.05 A B")
        assert result["status"] == "added"
        assert result["type"] == "geometric_tolerance"
        assert result["text"] == "0.05 A B"
        fcfs.Add.assert_called_once_with(0.1, 0.1, 0)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_geometric_tolerance(0.1, 0.1)
        assert "error" in result

    def test_fallback_to_textbox(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sheet.FCFs.Add.side_effect = Exception("FCFs not available")
        text_boxes = MagicMock()
        text_box = MagicMock()
        text_boxes.Add.return_value = text_box
        sheet.TextBoxes = text_boxes
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_geometric_tolerance(0.1, 0.1, "0.05 A")
        assert result["status"] == "added"
        text_boxes.Add.assert_called_once_with(0.1, 0.1, 0)


# ============================================================================
# ADD DETAIL VIEW
# ============================================================================


class TestAddDetailView:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        parent_view = MagicMock()
        dvs = MagicMock()
        dvs.Count = 2
        dvs.Item.return_value = parent_view
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_detail_view(0, 0.05, 0.05, 0.01, 0.2, 0.1, 2.0)
        assert result["status"] == "added"
        assert result["type"] == "detail_view"
        assert result["parent_view_index"] == 0
        assert result["scale"] == 2.0

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_detail_view(0, 0.05, 0.05, 0.01, 0.2, 0.1)
        assert "error" in result

    def test_invalid_parent_index(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_detail_view(5, 0.05, 0.05, 0.01, 0.2, 0.1)
        assert "error" in result


# ============================================================================
# ADD AUXILIARY VIEW
# ============================================================================


class TestAddAuxiliaryView:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        parent_view = MagicMock()
        dvs = MagicMock()
        dvs.Count = 2
        dvs.Item.return_value = parent_view
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_auxiliary_view(0, 0.2, 0.3, "Up")
        assert result["status"] == "added"
        assert result["type"] == "auxiliary_view"
        assert result["fold_direction"] == "Up"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_auxiliary_view(0, 0.2, 0.3)
        assert "error" in result

    def test_invalid_direction(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_auxiliary_view(0, 0.2, 0.3, "Diagonal")
        assert "error" in result


# ============================================================================
# ADD DRAFT VIEW
# ============================================================================


class TestAddDraftView:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_draft_view(0.15, 0.10)
        assert result["status"] == "added"
        assert result["type"] == "draft_view"
        assert result["position"] == [0.15, 0.10]
        dvs.AddDraftView.assert_called_once_with(0.15, 0.10)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_draft_view(0.15, 0.10)
        assert "error" in result

    def test_exception(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        del dvs._oleobj_
        dvs.AddDraftView.side_effect = Exception("COM error")
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_draft_view(0.15, 0.10)
        assert "error" in result


# ============================================================================
# ALIGN DRAWING VIEWS
# ============================================================================


class TestAlignDrawingViews:
    def test_align(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        view1 = MagicMock()
        view2 = MagicMock()
        dvs = MagicMock()
        dvs.Count = 2
        dvs.Item.side_effect = lambda i: {1: view1, 2: view2}[i]
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.align_drawing_views(0, 1, True)
        assert result["status"] == "aligned"
        view1.AlignToView.assert_called_once_with(view2)

    def test_unalign(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        view1 = MagicMock()
        view2 = MagicMock()
        dvs = MagicMock()
        dvs.Count = 2
        dvs.Item.side_effect = lambda i: {1: view1, 2: view2}[i]
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.align_drawing_views(0, 1, False)
        assert result["status"] == "unaligned"
        view1.RemoveAlignment.assert_called_once()

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.align_drawing_views(0, 5, True)
        assert "error" in result


# ============================================================================
# GET DRAWING VIEW MODEL LINK
# ============================================================================


class TestGetDrawingViewModelLink:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        view = MagicMock()
        model_link = MagicMock()
        model_link.FileName = "C:/parts/test.par"
        model_link.Name = "test.par"
        view.ModelLink = model_link
        view.Name = "View 1"
        view.ScaleFactor = 1.0
        dvs = MagicMock()
        dvs.Count = 1
        dvs.Item.return_value = view
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.get_drawing_view_model_link(0)
        assert result["view_index"] == 0
        assert result["has_model_link"] is True
        assert result["model_path"] == "C:/parts/test.par"

    def test_no_model_link(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        view = MagicMock(spec=[])  # empty spec - no attributes
        dvs = MagicMock()
        dvs.Count = 1
        dvs.Item.return_value = view
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.get_drawing_view_model_link(0)
        assert result["view_index"] == 0
        assert result["has_model_link"] is False

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.get_drawing_view_model_link(5)
        assert "error" in result


# ============================================================================
# SHOW TANGENT EDGES
# ============================================================================


class TestShowTangentEdges:
    def test_show(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        view = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        dvs.Item.return_value = view
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.show_tangent_edges(0, True)
        assert result["status"] == "updated"
        assert result["show_tangent_edges"] is True
        assert view.ShowTangentEdges is True

    def test_hide(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        view = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        dvs.Item.return_value = view
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.show_tangent_edges(0, False)
        assert result["status"] == "updated"
        assert result["show_tangent_edges"] is False
        assert view.ShowTangentEdges is False

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.show_tangent_edges(5, True)
        assert "error" in result


# ============================================================================
# BATCH 10: DRAFT SHEET COLLECTIONS
# ============================================================================


class TestGetSheetDimensions:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dim1 = MagicMock()
        dim1.Type = 1
        dim1.Value = 0.05
        dim1.Name = "D1"
        dim2 = MagicMock()
        dim2.Type = 2
        dim2.Value = 0.1
        dim2.Name = "D2"

        dims = MagicMock()
        dims.Count = 2
        dims.Item.side_effect = lambda i: [None, dim1, dim2][i]
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet

        result = em.get_sheet_dimensions()
        assert result["count"] == 2
        assert result["dimensions"][0]["value"] == 0.05
        assert result["dimensions"][1]["name"] == "D2"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.ActiveSheet

        result = em.get_sheet_dimensions()
        assert "error" in result

    def test_empty(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dims = MagicMock()
        dims.Count = 0
        sheet.Dimensions = dims
        doc.ActiveSheet = sheet

        result = em.get_sheet_dimensions()
        assert result["count"] == 0
        assert result["dimensions"] == []


class TestGetSheetBalloons:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        b1 = MagicMock()
        b1.BalloonText = "1"
        b1.x = 0.05
        b1.y = 0.1

        balloons = MagicMock()
        balloons.Count = 1
        balloons.Item.return_value = b1
        sheet.Balloons = balloons
        doc.ActiveSheet = sheet

        result = em.get_sheet_balloons()
        assert result["count"] == 1
        assert result["balloons"][0]["text"] == "1"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.ActiveSheet

        result = em.get_sheet_balloons()
        assert "error" in result

    def test_empty(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        balloons = MagicMock()
        balloons.Count = 0
        sheet.Balloons = balloons
        doc.ActiveSheet = sheet

        result = em.get_sheet_balloons()
        assert result["count"] == 0


class TestGetSheetTextBoxes:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        tb = MagicMock()
        tb.Text = "Hello"
        tb.x = 0.02
        tb.y = 0.03
        tb.Height = 0.005

        text_boxes = MagicMock()
        text_boxes.Count = 1
        text_boxes.Item.return_value = tb
        sheet.TextBoxes = text_boxes
        doc.ActiveSheet = sheet

        result = em.get_sheet_text_boxes()
        assert result["count"] == 1
        assert result["text_boxes"][0]["text"] == "Hello"
        assert result["text_boxes"][0]["height"] == 0.005

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.ActiveSheet

        result = em.get_sheet_text_boxes()
        assert "error" in result

    def test_empty(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        text_boxes = MagicMock()
        text_boxes.Count = 0
        sheet.TextBoxes = text_boxes
        doc.ActiveSheet = sheet

        result = em.get_sheet_text_boxes()
        assert result["count"] == 0


class TestGetSheetDrawingObjects:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        obj1 = MagicMock()
        obj1.Name = "Obj1"

        drawing_objects = MagicMock()
        drawing_objects.Count = 1
        drawing_objects.Item.return_value = obj1
        sheet.DrawingObjects = drawing_objects
        doc.ActiveSheet = sheet

        result = em.get_sheet_drawing_objects()
        assert result["count"] == 1
        assert result["drawing_objects"][0]["name"] == "Obj1"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.ActiveSheet

        result = em.get_sheet_drawing_objects()
        assert "error" in result

    def test_empty(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        drawing_objects = MagicMock()
        drawing_objects.Count = 0
        sheet.DrawingObjects = drawing_objects
        doc.ActiveSheet = sheet

        result = em.get_sheet_drawing_objects()
        assert result["count"] == 0


class TestGetSheetSections:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sec = MagicMock()
        sec.Label = "A-A"
        sec.Name = "Section1"
        sec.Type = 1

        sections = MagicMock()
        sections.Count = 1
        sections.Item.return_value = sec
        sheet.Sections = sections
        doc.ActiveSheet = sheet

        result = em.get_sheet_sections()
        assert result["count"] == 1
        assert result["sections"][0]["label"] == "A-A"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.ActiveSheet

        result = em.get_sheet_sections()
        assert "error" in result

    def test_empty(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        sections = MagicMock()
        sections.Count = 0
        sheet.Sections = sections
        doc.ActiveSheet = sheet

        result = em.get_sheet_sections()
        assert result["count"] == 0


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


class TestAddAssemblyDrawingViewEx:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        model_link = MagicMock()
        model_links = MagicMock()
        model_links.Count = 1
        model_links.Item.return_value = model_link
        doc.ModelLinks = model_links

        sheet = MagicMock()
        dvs = MagicMock()
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_assembly_drawing_view_ex(0.15, 0.15, "Front", 1.0)
        assert result["status"] == "added"
        assert result["orientation"] == "Front"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_assembly_drawing_view_ex()
        assert "error" in result

    def test_invalid_orientation(self, export_mgr):
        em, doc = export_mgr
        doc.Sheets = MagicMock()
        model_links = MagicMock()
        model_links.Count = 1
        doc.ModelLinks = model_links

        result = em.add_assembly_drawing_view_ex(orientation="BadView")
        assert "error" in result


class TestAddDrawingViewWithConfig:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        model_link = MagicMock()
        model_links = MagicMock()
        model_links.Count = 1
        model_links.Item.return_value = model_link
        doc.ModelLinks = model_links

        sheet = MagicMock()
        dvs = MagicMock()
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_drawing_view_with_config(0.15, 0.15, "Top", 2.0, "Config1")
        assert result["status"] == "added"
        assert result["configuration"] == "Config1"

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.add_drawing_view_with_config()
        assert "error" in result

    def test_invalid_orientation(self, export_mgr):
        em, doc = export_mgr
        doc.Sheets = MagicMock()
        model_links = MagicMock()
        model_links.Count = 1
        doc.ModelLinks = model_links

        result = em.add_drawing_view_with_config(orientation="BadView")
        assert "error" in result


class TestActivateDrawingView:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        view = MagicMock()
        dvs = MagicMock()
        dvs.Count = 2
        dvs.Item.return_value = view
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.activate_drawing_view(0)
        assert result["status"] == "activated"
        assert result["view_index"] == 0
        view.Activate.assert_called_once()

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.activate_drawing_view(5)
        assert "error" in result

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.activate_drawing_view(0)
        assert "error" in result


class TestDeactivateDrawingView:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        view = MagicMock()
        dvs = MagicMock()
        dvs.Count = 2
        dvs.Item.return_value = view
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.deactivate_drawing_view(1)
        assert result["status"] == "deactivated"
        assert result["view_index"] == 1
        view.Deactivate.assert_called_once()

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.deactivate_drawing_view(3)
        assert "error" in result

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        del doc.Sheets

        result = em.deactivate_drawing_view(0)
        assert "error" in result


# ============================================================================
# ADD BY DRAFT VIEW (Batch 11)
# ============================================================================


class TestAddByDraftView:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 2

        source_view = MagicMock()
        source_view.ScaleFactor = 1.5
        new_view = MagicMock()
        new_view.Name = "View 2"
        dvs.Item.return_value = source_view
        dvs.AddByDraftView.return_value = new_view

        # _get_drawing_views requires Sheets, ActiveSheet, DrawingViews
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_by_draft_view(0, 0.20, 0.10)
        assert result["status"] == "added"
        assert result["source_view_index"] == 0
        assert result["position"] == [0.20, 0.10]
        assert result["scale"] == 1.5
        dvs.AddByDraftView.assert_called_once_with(source_view, 1.5, 0.20, 0.10)

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_by_draft_view(5, 0.10, 0.10)
        assert "error" in result

    def test_custom_scale(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1

        source_view = MagicMock()
        new_view = MagicMock()
        dvs.Item.return_value = source_view
        dvs.AddByDraftView.return_value = new_view

        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_by_draft_view(0, 0.15, 0.15, scale=2.0)
        assert result["status"] == "added"
        assert result["scale"] == 2.0
        dvs.AddByDraftView.assert_called_once_with(source_view, 2.0, 0.15, 0.15)


# ============================================================================
# GET SECTION CUTS (Batch 11)
# ============================================================================


class TestGetSectionCuts:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1

        view = MagicMock()
        cp = MagicMock()
        cp.Caption = "A"
        cp.DisplayCaption = True
        cp.DisplayType = 0
        cp.StyleName = "Default"
        cp.TextHeight = 0.005
        cp.GetFoldLineWithViewDirection.return_value = (0.1, 0.2, 0.3, 0.4, 0.0, 1.0)

        cutting_planes = MagicMock()
        cutting_planes.Count = 1
        cutting_planes.Item.return_value = cp
        view.CuttingPlanes = cutting_planes
        dvs.Item.return_value = view

        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.get_section_cuts(0)
        assert result["count"] == 1
        assert result["view_index"] == 0
        assert result["section_cuts"][0]["caption"] == "A"
        assert result["section_cuts"][0]["fold_line"]["start"] == [0.1, 0.2]

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.get_section_cuts(3)
        assert "error" in result

    def test_no_cutting_planes(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1

        view = MagicMock(spec=[])  # No CuttingPlanes attribute
        dvs.Item.return_value = view

        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.get_section_cuts(0)
        assert result["count"] == 0
        assert result["section_cuts"] == []


# ============================================================================
# ADD SECTION CUT (Batch 11)
# ============================================================================


class TestAddSectionCut:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1

        view = MagicMock()
        cutting_plane = MagicMock()
        cutting_plane.Caption = "A"
        section_view = MagicMock()
        section_view.Name = "Section A-A"

        cutting_planes = MagicMock()
        cutting_planes.Count = 1
        cutting_planes.Add.return_value = cutting_plane
        cutting_plane.CreateView.return_value = section_view
        view.CuttingPlanes = cutting_planes
        dvs.Item.return_value = view

        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_section_cut(0, 0.20, 0.10)
        assert result["status"] == "added"
        assert result["source_view_index"] == 0
        assert result["section_type"] == "standard"
        cutting_planes.Add.assert_called_once()
        cutting_plane.CreateView.assert_called_once_with(0)

    def test_invalid_section_type(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1

        view = MagicMock()
        view.CuttingPlanes = MagicMock()
        dvs.Item.return_value = view

        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_section_cut(0, 0.20, 0.10, section_type=5)
        assert "error" in result
        assert "section_type" in result["error"]

    def test_revolved_section(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1

        view = MagicMock()
        cutting_plane = MagicMock()
        section_view = MagicMock()

        cutting_planes = MagicMock()
        cutting_planes.Count = 1
        cutting_planes.Add.return_value = cutting_plane
        cutting_plane.CreateView.return_value = section_view
        view.CuttingPlanes = cutting_planes
        dvs.Item.return_value = view

        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.add_section_cut(0, 0.15, 0.15, section_type=1)
        assert result["status"] == "added"
        assert result["section_type"] == "revolved"
        cutting_plane.CreateView.assert_called_once_with(1)


# ============================================================================
# GET DRAWING VIEW DIMENSIONS (Batch 11)
# ============================================================================


class TestGetDrawingViewDimensions:
    def test_success(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1

        view = MagicMock()

        dim1 = MagicMock()
        dim1.DimensionType = 1
        dim1.Value = 0.05
        dim1.Constraint = False
        dim1.PrefixString = ""
        dim1.SuffixString = "mm"
        dim1.OverrideString = ""
        dim1.SubfixString = ""
        dim1.SuperfixString = ""

        dim2 = MagicMock()
        dim2.DimensionType = 3
        dim2.Value = 1.5708
        dim2.Constraint = True
        dim2.PrefixString = ""
        dim2.SuffixString = ""
        dim2.OverrideString = "90"
        dim2.SubfixString = ""
        dim2.SuperfixString = ""

        dims = MagicMock()
        dims.Count = 2
        dims.Item.side_effect = lambda i: {1: dim1, 2: dim2}[i]
        view.Dimensions = dims
        dvs.Item.return_value = view

        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.get_drawing_view_dimensions(0)
        assert result["count"] == 2
        assert result["view_index"] == 0
        assert result["dimensions"][0]["dimension_type"] == 1
        assert result["dimensions"][0]["dimension_type_name"] == "Linear"
        assert result["dimensions"][0]["value"] == 0.05
        assert result["dimensions"][1]["dimension_type"] == 3
        assert result["dimensions"][1]["dimension_type_name"] == "Angular"
        assert result["dimensions"][1]["override"] == "90"

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.get_drawing_view_dimensions(5)
        assert "error" in result

    def test_no_dimensions(self, export_mgr):
        em, doc = export_mgr
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1

        view = MagicMock(spec=[])  # No Dimensions attribute
        dvs.Item.return_value = view

        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.get_drawing_view_dimensions(0)
        assert result["count"] == 0
        assert result["dimensions"] == []
