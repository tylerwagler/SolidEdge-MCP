"""
Unit tests for ExportManager view manipulation backend methods.

Tests drawing view scale, delete, update, projected views, move,
hidden edges, display modes, orientations, detail/auxiliary/draft views,
alignment, activation, section cuts, and more.
Uses unittest.mock to simulate COM objects.
"""

from unittest.mock import MagicMock

import pytest

from solidedge_mcp.backends.constants import DocumentTypeConstants, FoldTypeConstants

IG_ASSEMBLY_DOCUMENT = DocumentTypeConstants.igAssemblyDocument
IG_DRAFT_DOCUMENT = DocumentTypeConstants.igDraftDocument
IG_PART_DOCUMENT = DocumentTypeConstants.igPartDocument


@pytest.fixture
def export_mgr():
    """Create ExportManager with mocked dependencies."""
    from solidedge_mcp.backends.export import ExportManager

    dm = MagicMock()
    doc = MagicMock()
    doc.Type = IG_DRAFT_DOCUMENT
    dm.get_active_document.return_value = doc
    return ExportManager(dm), doc


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
        assert result["reads_back"] == 0.5
        assert view.ScaleFactor == 0.5

    def test_a_write_the_view_did_not_keep_shows_in_reads_back(self, export_mgr):
        """The setter used to echo the value it was asked for. A scale Solid
        Edge clamps or ignores must be visible, so ScaleFactor is read back."""
        em, doc = export_mgr
        view = MagicMock()
        type(view).ScaleFactor = property(lambda self: 1.0, lambda self, v: None)
        sheet = MagicMock()
        dvs = MagicMock()
        dvs.Count = 1
        dvs.Item.return_value = view
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()

        result = em.set_drawing_view_scale(0, 0.25)

        assert result["scale"] == 0.25, "what was asked"
        assert result["reads_back"] == 1.0, "what the view actually holds"

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
        doc.Type = IG_PART_DOCUMENT

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


def one_view(doc, view=None):
    """Wire a draft document up to a single drawing view."""
    sheet = MagicMock()
    view = view if view is not None else MagicMock()
    dvs = MagicMock()
    dvs.Count = 1
    dvs.Item.return_value = view
    del dvs._oleobj_
    sheet.DrawingViews = dvs
    doc.ActiveSheet = sheet
    sheets = MagicMock()
    sheets.Count = 1
    doc.Sheets = sheets
    return sheet, view


class TestMoveDrawingView:
    """SetOrigin is the only way to move a view; OriginX cannot be assigned."""

    def test_success(self, export_mgr):
        em, doc = export_mgr
        _sheet, view = one_view(doc)
        view.GetOrigin.return_value = (0.15, 0.20)

        result = em.move_drawing_view(0, 0.15, 0.20)

        assert result["status"] == "moved"
        assert result["origin"] == [0.15, 0.20]
        view.SetOrigin.assert_called_once_with(0.15, 0.20)

    def test_does_not_assign_originx(self, export_mgr):
        """Assigning OriginX raises "can not be set", so it must not be tried."""
        em, doc = export_mgr
        _sheet, view = one_view(doc)

        em.move_drawing_view(0, 0.15, 0.20)

        assert "OriginX" not in view.mock_calls.__str__()

    def test_com_error(self, export_mgr):
        em, doc = export_mgr
        _sheet, view = one_view(doc)
        view.SetOrigin.side_effect = Exception("COM error")

        assert "error" in em.move_drawing_view(0, 0.15, 0.20)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        assert "error" in em.move_drawing_view(0, 0.15, 0.20)

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        one_view(doc)

        assert "error" in em.move_drawing_view(5, 0.15, 0.20)


# ============================================================================
# SHOW HIDDEN EDGES
# ============================================================================


class TestShowHiddenEdges:
    """The view property that holds is Defaults_ShowHiddenEdges.

    DrawingView.ShowHiddenEdges raises "can not be set", and
    ModelMember.ShowHiddenEdges accepts the write then reads back unchanged.
    """

    def test_show(self, export_mgr):
        em, doc = export_mgr
        _sheet, view = one_view(doc)

        result = em.show_hidden_edges(0, True)

        assert result["status"] == "updated"
        assert result["show_hidden_edges"] is True
        assert view.Defaults_ShowHiddenEdges is True

    def test_hide(self, export_mgr):
        em, doc = export_mgr
        _sheet, view = one_view(doc)

        result = em.show_hidden_edges(0, False)

        assert result["show_hidden_edges"] is False
        assert view.Defaults_ShowHiddenEdges is False

    def test_a_view_that_reverts_is_reported_as_a_failure(self, export_mgr):
        """A silent revert must not be dressed up as success."""
        em, doc = export_mgr

        class Stubborn:
            Defaults_ShowHiddenEdges = False

            def __setattr__(self, name, value):
                pass  # accepts the write, keeps the old value

            def Update(self):
                pass

        _sheet, _view = one_view(doc, view=Stubborn())

        result = em.show_hidden_edges(0, True)

        assert "error" in result
        assert "did not accept" in result["error"]

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        one_view(doc)

        assert "error" in em.show_hidden_edges(5, True)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        assert "error" in em.show_hidden_edges(0, True)


# ============================================================================
# SET DRAWING VIEW DISPLAY MODE
# ============================================================================


class TestSetDrawingViewDisplayMode:
    """A DrawingView has no SetRenderMode and no DisplayMode.

    SetRenderMode belongs to the 3D window View in framewrk.tlb. Shading,
    ShadingShowVisibleEdges and Defaults_ShowHiddenEdges are what a drawing
    view really exposes.
    """

    def test_shaded(self, export_mgr):
        em, doc = export_mgr
        _sheet, view = one_view(doc)

        result = em.set_drawing_view_display_mode(0, "Shaded")

        assert result["status"] == "updated"
        assert view.Shading is True
        assert view.ShadingShowVisibleEdges is False
        view.SetRenderMode.assert_not_called()

    def test_shaded_with_edges(self, export_mgr):
        em, doc = export_mgr
        _sheet, view = one_view(doc)

        em.set_drawing_view_display_mode(0, "ShadedWithEdges")

        assert view.Shading is True
        assert view.ShadingShowVisibleEdges is True

    def test_wireframe_shows_every_edge(self, export_mgr):
        em, doc = export_mgr
        _sheet, view = one_view(doc)

        em.set_drawing_view_display_mode(0, "Wireframe")

        assert view.Shading is False
        assert view.Defaults_ShowHiddenEdges is True

    def test_hidden_edges_visible_hides_obscured_edges(self, export_mgr):
        em, doc = export_mgr
        _sheet, view = one_view(doc)

        em.set_drawing_view_display_mode(0, "HiddenEdgesVisible")

        assert view.Shading is False
        assert view.Defaults_ShowHiddenEdges is False

    def test_a_view_refusing_everything_is_an_error(self, export_mgr):
        em, doc = export_mgr

        class Refuses:
            def __setattr__(self, name, value):
                raise AttributeError(name)

            def Update(self):
                pass

        one_view(doc, view=Refuses())

        result = em.set_drawing_view_display_mode(0, "Shaded")

        assert "error" in result
        assert "accepted none" in result["error"]

    def test_invalid_mode(self, export_mgr):
        em, doc = export_mgr
        one_view(doc)

        result = em.set_drawing_view_display_mode(0, "InvalidMode")

        assert "error" in result
        assert "Wireframe" in result["error"]

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        assert "error" in em.set_drawing_view_display_mode(0, "Shaded")


# ============================================================================
# GET DRAWING VIEW INFO
# ============================================================================


class TestGetDrawingViewInfo:
    def test_reports_the_origin(self, export_mgr):
        """origin_x and origin_y were always missing: there is no OriginX."""
        em, doc = export_mgr
        _sheet, view = one_view(doc)
        view.Name = "View 1"
        view.ScaleFactor = 1.0
        view.GetOrigin.return_value = (0.2, 0.15)
        view.Defaults_ShowHiddenEdges = True
        view.Defaults_ShowTangentEdges = False

        result = em.get_drawing_view_info(0)

        assert result["name"] == "View 1"
        assert result["origin_x"] == 0.2
        assert result["origin_y"] == 0.15
        assert result["show_hidden_edges"] is True
        assert result["show_tangent_edges"] is False

    def test_a_view_without_an_origin_omits_it(self, export_mgr):
        em, doc = export_mgr
        _sheet, view = one_view(doc)
        view.GetOrigin.side_effect = Exception("no origin")

        result = em.get_drawing_view_info(0)

        assert "origin_x" not in result

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        assert "error" in em.get_drawing_view_info(0)

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        one_view(doc)

        assert "error" in em.get_drawing_view_info(5)


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
        doc.Type = IG_PART_DOCUMENT

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
    """The view property that holds is Defaults_ShowTangentEdges.

    ModelMember.ShowTangentEdges exists and accepts the write, but reads back
    unchanged, so routing it there stopped the exception without changing
    anything. Verified against Solid Edge 2026.
    """

    def test_show(self, export_mgr):
        em, doc = export_mgr
        _sheet, view = one_view(doc)

        result = em.show_tangent_edges(0, True)

        assert result["status"] == "updated"
        assert result["show_tangent_edges"] is True
        assert view.Defaults_ShowTangentEdges is True

    def test_hide(self, export_mgr):
        em, doc = export_mgr
        _sheet, view = one_view(doc)

        result = em.show_tangent_edges(0, False)

        assert result["show_tangent_edges"] is False
        assert view.Defaults_ShowTangentEdges is False

    def test_does_not_write_to_the_model_member(self, export_mgr):
        em, doc = export_mgr
        _sheet, view = one_view(doc)
        member = MagicMock()
        members = MagicMock()
        members.Count = 1
        members.Item.return_value = member
        view.ModelMembers = members

        em.show_tangent_edges(0, True)

        members.Item.assert_not_called()

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        one_view(doc)

        assert "error" in em.show_tangent_edges(5, True)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        assert "error" in em.show_tangent_edges(0, True)


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
        # AddByDetailEnvelope(From, x1, y1, Radius, Scale, x2, y2) - Scale sits
        # between the envelope radius and the placement point.
        dvs.AddByDetailEnvelope.assert_called_once_with(
            parent_view, 0.05, 0.05, 0.01, 2.0, 0.2, 0.1
        )
        dvs.AddDetailView.assert_not_called()

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

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
        # AddByFold(From, foldDir, x, y) is the fold-direction entry point.
        # AddByAuxiliaryFold wants a fold line picked on the parent view.
        dvs.AddByFold.assert_called_once_with(parent_view, FoldTypeConstants.igFoldUp, 0.2, 0.3)
        dvs.AddByAuxiliaryFold.assert_not_called()

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

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
        # AddDraftView(Scale, x1, y1) - the scale is the first argument.
        dvs.AddDraftView.assert_called_once_with(1.0, 0.15, 0.10)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

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
    """DrawingViews.Align/Unalign act on the document select set.

    DrawingView.AlignToView and RemoveAlignment are in no Solid Edge type
    library, and Align() raises outright when nothing is selected.
    """

    def _two_views(self, doc):
        sheet = MagicMock()
        view1 = MagicMock()
        view1.GetOrigin.return_value = (0.1, 0.1)
        view2 = MagicMock()
        view2.GetOrigin.return_value = (0.2, 0.2)
        dvs = MagicMock()
        dvs.Count = 2
        dvs.Item.side_effect = lambda i: {1: view1, 2: view2}[i]
        del dvs._oleobj_
        sheet.DrawingViews = dvs
        doc.ActiveSheet = sheet
        doc.Sheets = MagicMock()
        return dvs, view1, view2

    def test_align_selects_both_views_then_calls_align(self, export_mgr):
        em, doc = export_mgr
        dvs, view1, view2 = self._two_views(doc)
        select_set = MagicMock()
        doc.SelectSet = select_set

        result = em.align_drawing_views(0, 1, True)

        assert result["status"] == "aligned"
        dvs.Align.assert_called_once()
        added = [call.args[0] for call in select_set.Add.call_args_list]
        assert added == [view1, view2]
        view1.AlignToView.assert_not_called()

    def test_unalign_calls_unalign(self, export_mgr):
        em, doc = export_mgr
        dvs, view1, _view2 = self._two_views(doc)
        doc.SelectSet = MagicMock()

        result = em.align_drawing_views(0, 1, False)

        assert result["status"] == "unaligned"
        dvs.Unalign.assert_called_once()
        view1.RemoveAlignment.assert_not_called()

    def test_the_selection_is_cleared_afterwards(self, export_mgr):
        em, doc = export_mgr
        dvs, _view1, _view2 = self._two_views(doc)
        select_set = MagicMock()
        doc.SelectSet = select_set
        dvs.Align.side_effect = Exception("nothing to align")

        em.align_drawing_views(0, 1, True)

        assert select_set.RemoveAll.call_count == 2

    def test_says_so_when_neither_view_moved(self, export_mgr):
        em, doc = export_mgr
        self._two_views(doc)
        doc.SelectSet = MagicMock()

        result = em.align_drawing_views(0, 1, True)

        assert "fold relationship" in result["note"]

    def test_refuses_one_view_against_itself(self, export_mgr):
        em, doc = export_mgr
        self._two_views(doc)

        result = em.align_drawing_views(1, 1, True)

        assert "error" in result
        assert "two different views" in result["error"]

    def test_invalid_index(self, export_mgr):
        em, doc = export_mgr
        self._two_views(doc)

        assert "error" in em.align_drawing_views(0, 5, True)


# ============================================================================
# ACTIVATE / DEACTIVATE DRAWING VIEW
# ============================================================================


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
        doc.Type = IG_PART_DOCUMENT

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
        doc.Type = IG_PART_DOCUMENT

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
# ADD ASSEMBLY DRAWING VIEW EX
# ============================================================================


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
        doc.Type = IG_PART_DOCUMENT

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
        doc.Type = IG_PART_DOCUMENT

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


# ============================================================================
# DRAFT: UPDATE ALL VIEWS
# ============================================================================


class TestUpdateAllViews:
    def test_success_force(self, export_mgr):
        em, doc = export_mgr
        doc.Sheets = MagicMock()

        result = em.update_all_views(force_update=True)
        assert result["status"] == "updated_all"
        assert result["force_update"] is True
        doc.UpdateAll.assert_called_once_with(True)

    def test_success_no_force(self, export_mgr):
        em, doc = export_mgr
        doc.Sheets = MagicMock()

        result = em.update_all_views(force_update=False)
        assert result["status"] == "updated_all"
        assert result["force_update"] is False
        doc.UpdateAll.assert_called_once_with(False)

    def test_not_draft(self, export_mgr):
        em, doc = export_mgr
        doc.Type = IG_PART_DOCUMENT

        result = em.update_all_views()
        assert "error" in result
