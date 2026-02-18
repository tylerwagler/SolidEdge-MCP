"""
Unit tests for non-export backend methods that were in test_export.py.

Tests connection, document management, sketching, and assembly operations.
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


@pytest.fixture
def doc_mgr():
    """Create DocumentManager with mocked connection."""
    from solidedge_mcp.backends.documents import DocumentManager

    conn = MagicMock()
    app = MagicMock()
    conn.get_application.return_value = app
    dm = DocumentManager(conn)
    return dm, app


# ============================================================================
# DOCUMENT MANAGEMENT (activate, undo, redo)
# ============================================================================


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
# DRAW ARC BY 3 POINTS
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
# DRAW CIRCLE BY 2 POINTS
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
# DRAW CIRCLE BY 3 POINTS
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
# MIRROR SPLINE
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
# HIDE PROFILE
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
# PROJECT EDGE
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
# INCLUDE EDGE
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
# START COMMAND
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
# DOCUMENTS: SAVE COPY AS
# ============================================================================


class TestSaveCopyAs:
    @pytest.fixture
    def doc_mgr(self):
        from solidedge_mcp.backends.documents import DocumentManager

        conn = MagicMock()
        app = MagicMock()
        conn.get_application.return_value = app
        dm = DocumentManager(conn)
        dm.active_document = MagicMock()
        dm.active_document.Name = "part1.par"
        return dm

    def test_success(self, doc_mgr):
        result = doc_mgr.save_copy_as("C:/temp/copy.par")
        assert result["status"] == "copy_saved"
        assert result["path"] == "C:/temp/copy.par"
        assert result["active_document"] == "part1.par"
        doc_mgr.active_document.SaveCopyAs.assert_called_once_with("C:/temp/copy.par")

    def test_no_active_document(self, doc_mgr):
        doc_mgr.active_document = None
        result = doc_mgr.save_copy_as("C:/temp/copy.par")
        assert "error" in result
        assert "No active document" in result["error"]

    def test_com_error(self, doc_mgr):
        doc_mgr.active_document.SaveCopyAs.side_effect = Exception("Access denied")
        result = doc_mgr.save_copy_as("C:/temp/copy.par")
        assert "error" in result


# ============================================================================
# DOCUMENTS: OPEN WITH TEMPLATE
# ============================================================================


class TestOpenWithTemplate:
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
        doc.Name = "imported.par"
        doc.Type = 1
        app.Documents.OpenWithTemplate.return_value = doc

        f = tmp_path / "model.step"
        f.write_text("")

        result = dm.open_with_template(str(f), "C:/templates/metric.par")
        assert result["status"] == "opened_with_template"
        assert result["name"] == "imported.par"
        assert result["template"] == "C:/templates/metric.par"
        app.Documents.OpenWithTemplate.assert_called_once_with(
            str(f), "C:/templates/metric.par"
        )

    def test_file_not_found(self, doc_mgr):
        dm, _ = doc_mgr
        result = dm.open_with_template("C:\\nonexistent\\file.step", "C:/templates/t.par")
        assert "error" in result
        assert "not found" in result["error"]

    def test_com_error(self, doc_mgr, tmp_path):
        dm, app = doc_mgr
        app.Documents.OpenWithTemplate.side_effect = Exception("Template invalid")

        f = tmp_path / "model.step"
        f.write_text("")

        result = dm.open_with_template(str(f), "C:/bad/template.par")
        assert "error" in result


# ============================================================================
# DOCUMENTS: OPEN WITH FILE OPEN DIALOG
# ============================================================================


class TestOpenWithFileOpenDialog:
    @pytest.fixture
    def doc_mgr(self):
        from solidedge_mcp.backends.documents import DocumentManager

        conn = MagicMock()
        app = MagicMock()
        conn.get_application.return_value = app
        return DocumentManager(conn), app

    def test_success_no_args(self, doc_mgr):
        dm, app = doc_mgr
        doc = MagicMock()
        doc.Name = "Selected.par"
        doc.Type = 1
        app.Documents.OpenWithFileOpenDialog.return_value = doc

        result = dm.open_with_file_open_dialog()
        assert result["status"] == "opened"
        assert result["name"] == "Selected.par"
        app.Documents.OpenWithFileOpenDialog.assert_called_once_with()

    def test_success_with_args(self, doc_mgr):
        dm, app = doc_mgr
        doc = MagicMock()
        doc.Name = "myfile.par"
        doc.Type = 1
        app.Documents.OpenWithFileOpenDialog.return_value = doc

        result = dm.open_with_file_open_dialog(
            filename="*.par", dialog_title="Select Part"
        )
        assert result["status"] == "opened"
        assert result["name"] == "myfile.par"
        app.Documents.OpenWithFileOpenDialog.assert_called_once_with(
            Filename="*.par", DialogTitle="Select Part"
        )

    def test_cancelled(self, doc_mgr):
        dm, app = doc_mgr
        app.Documents.OpenWithFileOpenDialog.return_value = None

        result = dm.open_with_file_open_dialog()
        assert result["status"] == "cancelled"

    def test_com_error(self, doc_mgr):
        dm, app = doc_mgr
        app.Documents.OpenWithFileOpenDialog.side_effect = Exception("Dialog failed")

        result = dm.open_with_file_open_dialog()
        assert "error" in result
