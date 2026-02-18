"""
Unit tests for QueryManager backend methods (_document.py mixin).

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
# REFERENCE PLANES
# ============================================================================


class TestGetRefPlanes:
    def test_default_planes(self, query_mgr):
        qm, doc = query_mgr

        plane1 = MagicMock()
        plane1.Name = "Top"
        plane1.Visible = True

        plane2 = MagicMock()
        plane2.Name = "Front"
        plane2.Visible = True

        plane3 = MagicMock()
        plane3.Name = "Right"
        plane3.Visible = True

        ref_planes = MagicMock()
        ref_planes.Count = 3
        ref_planes.Item.side_effect = lambda i: [None, plane1, plane2, plane3][i]
        doc.RefPlanes = ref_planes

        result = qm.get_ref_planes()
        assert result["count"] == 3
        assert result["planes"][0]["is_default"] is True
        assert result["planes"][0]["name"] == "Top"

    def test_with_offset_planes(self, query_mgr):
        qm, doc = query_mgr

        planes = []
        for i in range(5):
            p = MagicMock()
            p.Name = f"Plane_{i + 1}"
            p.Visible = True
            planes.append(p)

        ref_planes = MagicMock()
        ref_planes.Count = 5
        ref_planes.Item.side_effect = lambda i: planes[i - 1]
        doc.RefPlanes = ref_planes

        result = qm.get_ref_planes()
        assert result["count"] == 5
        assert result["planes"][2]["is_default"] is True
        assert result["planes"][3]["is_default"] is False


# ============================================================================
# RECOMPUTE
# ============================================================================


class TestRecompute:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        result = qm.recompute()
        assert result["status"] == "recomputed"
        model.Recompute.assert_called_once()


# ============================================================================
# DOCUMENT-LEVEL RECOMPUTE
# ============================================================================


class TestRecomputeDocument:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        result = qm.recompute_document()
        assert result["status"] == "recomputed_document"
        doc.Recompute.assert_called_once()

    def test_com_error(self, query_mgr):
        qm, doc = query_mgr
        doc.Recompute.side_effect = Exception("Recompute failed")

        result = qm.recompute_document()
        assert "error" in result


# ============================================================================
# MODELING MODE
# ============================================================================


class TestGetModelingMode:
    def test_synchronous(self, query_mgr):
        qm, doc = query_mgr
        doc.ModelingMode = 1  # seModelingModeSynchronous = 1

        result = qm.get_modeling_mode()
        assert result["mode"] == "synchronous"

    def test_ordered(self, query_mgr):
        qm, doc = query_mgr
        doc.ModelingMode = 2  # seModelingModeOrdered = 2

        result = qm.get_modeling_mode()
        assert result["mode"] == "ordered"


class TestSetModelingMode:
    def test_to_synchronous(self, query_mgr):
        qm, doc = query_mgr
        doc.ModelingMode = 1

        result = qm.set_modeling_mode("synchronous")
        assert result["status"] == "changed"

    def test_invalid_mode(self, query_mgr):
        qm, doc = query_mgr

        result = qm.set_modeling_mode("invalid")
        assert "error" in result


# ============================================================================
# SET DOCUMENT PROPERTY
# ============================================================================


class TestSetDocumentProperty:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        summary = MagicMock()
        doc.SummaryInfo = summary

        result = qm.set_document_property("Title", "My Part")
        assert result["status"] == "set"
        assert summary.Title == "My Part"

    def test_invalid_property(self, query_mgr):
        qm, doc = query_mgr
        summary = MagicMock()
        doc.SummaryInfo = summary

        result = qm.set_document_property("InvalidProp", "value")
        assert "error" in result

    def test_no_summary_info(self, query_mgr):
        qm, doc = query_mgr
        del doc.SummaryInfo

        result = qm.set_document_property("Title", "Test")
        assert "error" in result


# ============================================================================
# SOLID BODIES
# ============================================================================


class TestGetSolidBodies:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        model = MagicMock()
        model.Name = "Model_1"
        body = MagicMock()
        body.IsSolid = True
        body.Volume = 0.001
        shells = MagicMock()
        shells.Count = 1
        body.Shells = shells
        model.Body = body

        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        # No constructions
        doc.Constructions = MagicMock()
        doc.Constructions.Count = 0

        result = qm.get_solid_bodies()
        assert result["total_bodies"] == 1
        assert result["bodies"][0]["type"] == "design"
        assert result["bodies"][0]["is_solid"] is True


# ============================================================================
# QUIT APPLICATION
# ============================================================================


class TestQuitApplication:
    def test_success(self):
        from solidedge_mcp.backends.connection import SolidEdgeConnection

        conn = SolidEdgeConnection()
        conn.application = MagicMock()
        conn._is_connected = True

        result = conn.quit_application()
        assert result["status"] == "quit"
        assert conn.application is None
        assert conn._is_connected is False

    def test_not_connected(self):
        from solidedge_mcp.backends.connection import SolidEdgeConnection

        conn = SolidEdgeConnection()

        result = conn.quit_application()
        assert "error" in result


# ============================================================================
# PERFORMANCE FLAGS
# ============================================================================


class TestSetPerformanceMode:
    def test_success(self):
        from solidedge_mcp.backends.connection import SolidEdgeConnection

        conn = SolidEdgeConnection()
        conn.application = MagicMock()
        conn._is_connected = True

        result = conn.set_performance_mode(
            delay_compute=True, screen_updating=False, display_alerts=False
        )
        assert result["status"] == "updated"
        assert result["settings"]["delay_compute"] is True
        assert result["settings"]["screen_updating"] is False
        assert result["settings"]["display_alerts"] is False

    def test_partial_update(self):
        from solidedge_mcp.backends.connection import SolidEdgeConnection

        conn = SolidEdgeConnection()
        conn.application = MagicMock()
        conn._is_connected = True

        result = conn.set_performance_mode(screen_updating=False)
        assert result["status"] == "updated"
        assert "screen_updating" in result["settings"]
        assert "delay_compute" not in result["settings"]
