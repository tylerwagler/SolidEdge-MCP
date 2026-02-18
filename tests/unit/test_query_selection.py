"""
Unit tests for QueryManager backend methods (_selection.py mixin).

Uses unittest.mock to simulate COM objects.
"""

from unittest.mock import MagicMock

import pytest


@pytest.fixture
def doc_mgr():
    """Create mock doc_manager."""
    dm = MagicMock()
    doc = MagicMock()
    dm.get_active_document.return_value = doc
    return dm, doc


@pytest.fixture
def query_mgr(doc_mgr):
    """Create QueryManager with mocked dependencies."""
    from solidedge_mcp.backends.query import QueryManager

    dm, doc = doc_mgr
    return QueryManager(dm), doc


# ============================================================================
# SELECT SET
# ============================================================================


class TestGetSelectSet:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        item1 = MagicMock()
        item1.Name = "Face_1"

        item2 = MagicMock()
        item2.Name = "Edge_1"

        select_set = MagicMock()
        select_set.Count = 2
        select_set.Item.side_effect = lambda i: [None, item1, item2][i]
        doc.SelectSet = select_set

        result = qm.get_select_set()
        assert result["count"] == 2
        assert result["items"][0]["name"] == "Face_1"
        assert result["items"][1]["name"] == "Edge_1"

    def test_empty(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 0
        doc.SelectSet = select_set

        result = qm.get_select_set()
        assert result["count"] == 0
        assert result["items"] == []


class TestClearSelectSet:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 3
        doc.SelectSet = select_set

        result = qm.clear_select_set()
        assert result["status"] == "cleared"
        assert result["items_removed"] == 3
        select_set.RemoveAll.assert_called_once()

    def test_already_empty(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 0
        doc.SelectSet = select_set

        result = qm.clear_select_set()
        assert result["status"] == "cleared"
        assert result["items_removed"] == 0


# ============================================================================
# TIER 1: SELECT ADD
# ============================================================================


class TestSelectAdd:
    def test_add_feature(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 1
        doc.SelectSet = select_set

        features = MagicMock()
        features.Count = 3
        feat = MagicMock()
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.select_add("feature", 0)
        assert result["status"] == "added"
        assert result["object_type"] == "feature"
        assert result["selection_count"] == 1
        select_set.Add.assert_called_once_with(feat)

    def test_add_face(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 1
        doc.SelectSet = select_set

        models = MagicMock()
        models.Count = 1
        model = MagicMock()
        models.Item.return_value = model
        doc.Models = models

        body = MagicMock()
        model.Body = body
        face = MagicMock()
        faces = MagicMock()
        faces.Count = 5
        faces.Item.return_value = face
        body.Faces.return_value = faces

        result = qm.select_add("face", 2)
        assert result["status"] == "added"
        assert result["object_type"] == "face"
        select_set.Add.assert_called_once_with(face)

    def test_add_plane(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 1
        doc.SelectSet = select_set

        ref_planes = MagicMock()
        ref_planes.Count = 3
        plane = MagicMock()
        ref_planes.Item.return_value = plane
        doc.RefPlanes = ref_planes

        result = qm.select_add("plane", 0)
        assert result["status"] == "added"
        assert result["object_type"] == "plane"
        select_set.Add.assert_called_once_with(plane)

    def test_invalid_type(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        doc.SelectSet = select_set

        result = qm.select_add("invalid", 0)
        assert "error" in result
        assert "Unsupported object type" in result["error"]

    def test_invalid_feature_index(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        doc.SelectSet = select_set

        features = MagicMock()
        features.Count = 2
        doc.DesignEdgebarFeatures = features

        result = qm.select_add("feature", 5)
        assert "error" in result
        assert "Invalid feature index" in result["error"]

    def test_invalid_face_index(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        doc.SelectSet = select_set

        models = MagicMock()
        models.Count = 1
        model = MagicMock()
        models.Item.return_value = model
        doc.Models = models

        body = MagicMock()
        model.Body = body
        faces = MagicMock()
        faces.Count = 3
        body.Faces.return_value = faces

        result = qm.select_add("face", 10)
        assert "error" in result
        assert "Invalid face index" in result["error"]


# ============================================================================
# SELECT SET: REMOVE
# ============================================================================


class TestSelectRemove:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 3
        doc.SelectSet = select_set

        result = qm.select_remove(1)
        assert result["status"] == "removed"
        assert result["index"] == 1
        select_set.Remove.assert_called_once_with(2)  # 0-based -> 1-based

    def test_invalid_index(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 2
        doc.SelectSet = select_set

        result = qm.select_remove(5)
        assert "error" in result


# ============================================================================
# SELECT SET: ALL
# ============================================================================


class TestSelectAll:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 10
        doc.SelectSet = select_set

        result = qm.select_all()
        assert result["status"] == "selected_all"
        assert result["selection_count"] == 10
        select_set.AddAll.assert_called_once()


# ============================================================================
# SELECT SET: COPY / CUT / DELETE
# ============================================================================


class TestSelectCopy:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 2
        doc.SelectSet = select_set

        result = qm.select_copy()
        assert result["status"] == "copied"
        assert result["items_copied"] == 2
        select_set.Copy.assert_called_once()

    def test_empty_selection(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 0
        doc.SelectSet = select_set

        result = qm.select_copy()
        assert "error" in result


class TestSelectCut:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 3
        doc.SelectSet = select_set

        result = qm.select_cut()
        assert result["status"] == "cut"
        assert result["items_cut"] == 3
        select_set.Cut.assert_called_once()

    def test_empty_selection(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 0
        doc.SelectSet = select_set

        result = qm.select_cut()
        assert "error" in result


class TestSelectDelete:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 1
        doc.SelectSet = select_set

        result = qm.select_delete()
        assert result["status"] == "deleted"
        assert result["items_deleted"] == 1
        select_set.Delete.assert_called_once()

    def test_empty_selection(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 0
        doc.SelectSet = select_set

        result = qm.select_delete()
        assert "error" in result


# ============================================================================
# SELECT SET: DISPLAY CONTROL
# ============================================================================


class TestSelectSuspendDisplay:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        doc.SelectSet = select_set

        result = qm.select_suspend_display()
        assert result["status"] == "display_suspended"
        select_set.SuspendDisplay.assert_called_once()


class TestSelectResumeDisplay:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        doc.SelectSet = select_set

        result = qm.select_resume_display()
        assert result["status"] == "display_resumed"
        select_set.ResumeDisplay.assert_called_once()


class TestSelectRefreshDisplay:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        doc.SelectSet = select_set

        result = qm.select_refresh_display()
        assert result["status"] == "display_refreshed"
        select_set.RefreshDisplay.assert_called_once()
