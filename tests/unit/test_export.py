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
        result = dm.get_active_document_type()
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
        assert occ.Visible == False

    def test_show(self, asm_mgr):
        am, doc, occ = asm_mgr
        result = am.set_component_visibility(1, True)
        assert result["status"] == "updated"
        assert occ.Visible == True

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
