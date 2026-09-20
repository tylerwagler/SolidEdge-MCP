"""Unit tests for QueryManager.get_spatial_context (_spatial.py mixin)."""

from unittest.mock import MagicMock

import pytest

from solidedge_mcp.backends.query import QueryManager
from solidedge_mcp.backends.query._spatial import ORIGIN_TOLERANCE, PLANE_AXIS_MAP


@pytest.fixture
def doc_mgr():
    dm = MagicMock()
    doc = MagicMock()
    dm.get_active_document.return_value = doc
    return dm, doc


@pytest.fixture
def query_mgr(doc_mgr):
    dm, doc = doc_mgr
    return QueryManager(dm), doc


@pytest.fixture
def no_sketch(monkeypatch):
    """No sketch open: get_active_plane_index() returns None."""
    from solidedge_mcp import managers

    stub = MagicMock()
    stub.get_active_plane_index.return_value = None
    monkeypatch.setattr(managers, "sketch_manager", stub)
    return stub


def _body(doc, minimum, maximum, body_count=1):
    model = MagicMock()
    models = MagicMock()
    models.Count = body_count
    models.Item.return_value = model
    doc.Models = models
    model.Body.GetRange.return_value = (minimum, maximum)
    return model


# === Plane map ===


class TestPlaneAxisMap:
    def test_canonical_mapping(self):
        assert PLANE_AXIS_MAP["1"]["name"] == "Top"
        assert PLANE_AXIS_MAP["1"]["plane"] == "XY"
        assert PLANE_AXIS_MAP["1"]["normal"] == [0.0, 0.0, 1.0]
        assert PLANE_AXIS_MAP["2"]["name"] == "Right"
        assert PLANE_AXIS_MAP["2"]["plane"] == "YZ"
        assert PLANE_AXIS_MAP["2"]["normal"] == [1.0, 0.0, 0.0]
        assert PLANE_AXIS_MAP["3"]["name"] == "Front"
        assert PLANE_AXIS_MAP["3"]["plane"] == "XZ"
        assert PLANE_AXIS_MAP["3"]["normal"] == [0.0, 1.0, 0.0]

    def test_map_is_always_included(self, query_mgr, no_sketch):
        qm, doc = query_mgr
        _body(doc, (0.0, 0.0, 0.0), (0.1, 0.1, 0.1))
        assert qm.get_spatial_context()["plane_axis_map"] == PLANE_AXIS_MAP


# === No bodies ===


class TestNoBodies:
    def test_reports_zero_and_nulls(self, query_mgr, no_sketch):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.get_spatial_context()
        assert result["body_count"] == 0
        assert result["bounding_box"] is None
        assert result["centered_on_origin"] is None
        assert result["active_sketch"] is None
        assert any("No bodies" in w for w in result["warnings"])

    def test_bounding_box_is_not_attempted(self, query_mgr, no_sketch):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models
        qm.get_bounding_box = MagicMock()

        qm.get_spatial_context()
        qm.get_bounding_box.assert_not_called()


# === With a body ===


class TestWithBody:
    def test_bounding_box_center_and_dimensions(self, query_mgr, no_sketch):
        qm, doc = query_mgr
        _body(doc, (0.0, 0.0, 0.0), (0.1, 0.2, 0.3))

        result = qm.get_spatial_context()
        assert result["body_count"] == 1
        box = result["bounding_box"]
        assert box["min"] == [0.0, 0.0, 0.0]
        assert box["max"] == [0.1, 0.2, 0.3]
        assert box["center"] == pytest.approx([0.05, 0.1, 0.15])
        assert box["dimensions"] == pytest.approx({"x": 0.1, "y": 0.2, "z": 0.3})
        assert result["units"] == "meters"

    def test_body_centered_on_origin(self, query_mgr, no_sketch):
        qm, doc = query_mgr
        _body(doc, (-0.05, -0.05, -0.05), (0.05, 0.05, 0.05))

        result = qm.get_spatial_context()
        assert result["centered_on_origin"] is True

    def test_body_off_origin(self, query_mgr, no_sketch):
        qm, doc = query_mgr
        _body(doc, (1.0, 0.0, 0.0), (2.0, 0.1, 0.1))

        result = qm.get_spatial_context()
        assert result["centered_on_origin"] is False

    def test_tolerance_boundary(self, query_mgr, no_sketch):
        qm, doc = query_mgr
        # centre is exactly the tolerance on X -> not "on the origin"
        edge = ORIGIN_TOLERANCE * 2
        _body(doc, (0.0, 0.0, 0.0), (edge, 0.0, 0.0))

        assert qm.get_spatial_context()["centered_on_origin"] is False

    def test_just_inside_tolerance(self, query_mgr, no_sketch):
        qm, doc = query_mgr
        edge = ORIGIN_TOLERANCE
        _body(doc, (0.0, 0.0, 0.0), (edge, 0.0, 0.0))

        assert qm.get_spatial_context()["centered_on_origin"] is True

    def test_multiple_bodies_are_counted(self, query_mgr, no_sketch):
        qm, doc = query_mgr
        _body(doc, (0.0, 0.0, 0.0), (0.1, 0.1, 0.1), body_count=3)

        assert qm.get_spatial_context()["body_count"] == 3


# === Active sketch ===


class TestActiveSketch:
    def _sketch(self, monkeypatch, plane_index):
        from solidedge_mcp import managers

        stub = MagicMock()
        stub.get_active_plane_index.return_value = plane_index
        monkeypatch.setattr(managers, "sketch_manager", stub)
        return stub

    def test_open_sketch_on_a_default_plane(self, query_mgr, monkeypatch):
        qm, doc = query_mgr
        _body(doc, (0.0, 0.0, 0.0), (0.1, 0.1, 0.1))
        self._sketch(monkeypatch, 2)

        sketch = qm.get_spatial_context()["active_sketch"]
        assert sketch == {
            "plane_index": 2,
            "plane_name": "Right",
            "plane": "YZ",
            "normal": [1.0, 0.0, 0.0],
        }

    def test_user_plane_falls_back_to_the_com_name(self, query_mgr, monkeypatch):
        qm, doc = query_mgr
        _body(doc, (0.0, 0.0, 0.0), (0.1, 0.1, 0.1))
        self._sketch(monkeypatch, 7)
        doc.RefPlanes.Item.return_value.Name = "Offset Plane 4"

        sketch = qm.get_spatial_context()["active_sketch"]
        assert sketch == {"plane_index": 7, "plane_name": "Offset Plane 4"}

    def test_no_sketch_open(self, query_mgr, no_sketch):
        qm, doc = query_mgr
        _body(doc, (0.0, 0.0, 0.0), (0.1, 0.1, 0.1))

        assert qm.get_spatial_context()["active_sketch"] is None


# === COM failures ===


class TestFailuresAreNeverRaised:
    def test_document_access_failure(self, query_mgr, no_sketch):
        qm, doc = query_mgr
        qm.doc_manager.get_active_document.side_effect = Exception("not connected")

        result = qm.get_spatial_context()
        assert result["body_count"] is None
        assert result["bounding_box"] is None
        assert result["centered_on_origin"] is None
        assert result["plane_axis_map"] == PLANE_AXIS_MAP
        assert any("Body count unavailable" in w for w in result["warnings"])

    def test_bounding_box_failure_keeps_the_body_count(self, query_mgr, no_sketch):
        qm, doc = query_mgr
        model = _body(doc, (0.0, 0.0, 0.0), (0.1, 0.1, 0.1))
        model.Body.GetRange.side_effect = Exception("0x8002000F Parameter not optional")

        result = qm.get_spatial_context()
        assert result["body_count"] == 1
        assert result["bounding_box"] is None
        assert result["centered_on_origin"] is None
        assert any("Bounding box unavailable" in w for w in result["warnings"])

    def test_malformed_bounding_box(self, query_mgr, no_sketch):
        qm, doc = query_mgr
        _body(doc, (0.0, 0.0, 0.0), (0.1, 0.1, 0.1))
        qm.get_bounding_box = MagicMock(return_value={"min": "nope", "max": "nope"})

        result = qm.get_spatial_context()
        assert result["bounding_box"] is None
        assert any("malformed" in w for w in result["warnings"])

    def test_sketch_manager_failure(self, query_mgr, monkeypatch):
        from solidedge_mcp import managers

        qm, doc = query_mgr
        _body(doc, (0.0, 0.0, 0.0), (0.1, 0.1, 0.1))
        stub = MagicMock()
        stub.get_active_plane_index.side_effect = Exception("boom")
        monkeypatch.setattr(managers, "sketch_manager", stub)

        result = qm.get_spatial_context()
        assert result["active_sketch"] is None
        assert result["bounding_box"] is not None
        assert any("Active sketch plane unavailable" in w for w in result["warnings"])
