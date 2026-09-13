"""Tests for tools/resources.py — the 53 read-only solidedge:// endpoints."""

import asyncio
import json
from unittest.mock import MagicMock

import pytest
from fastmcp import Client, FastMCP

from solidedge_mcp.tools import resources

EXPECTED_TOTAL = 53
EXPECTED_STATIC = 38
EXPECTED_TEMPLATES = 15

MANAGER_NAMES = (
    "connection",
    "doc_manager",
    "export_manager",
    "feature_manager",
    "query_manager",
    "sketch_manager",
    "view_manager",
)


def _run(coro):
    return asyncio.run(coro)


@pytest.fixture(scope="module")
def mcp():
    server = FastMCP("t")
    resources.register(server)
    return server


@pytest.fixture
def managers(monkeypatch):
    """Replace every backend manager on the resources module with a MagicMock."""
    mocks = {}
    for name in MANAGER_NAMES:
        mock = MagicMock()
        monkeypatch.setattr(resources, name, mock)
        mocks[name] = mock
    return mocks


# === Registration ===


class TestRegistration:
    def test_registers_all_endpoints(self, mcp):
        static = _run(mcp.get_resources())
        templates = _run(mcp.get_resource_templates())
        assert len(static) == EXPECTED_STATIC
        assert len(templates) == EXPECTED_TEMPLATES
        assert len(static) + len(templates) == EXPECTED_TOTAL

    def test_registration_table_matches_the_server(self, mcp):
        static = _run(mcp.get_resources())
        templates = _run(mcp.get_resource_templates())
        assert set(static) | set(templates) == {uri for uri, _ in resources.RESOURCES}

    def test_uris_are_unique(self):
        uris = [uri for uri, _ in resources.RESOURCES]
        assert len(uris) == len(set(uris)) == EXPECTED_TOTAL

    def test_every_uri_uses_the_solidedge_scheme(self, mcp):
        static = _run(mcp.get_resources())
        templates = _run(mcp.get_resource_templates())
        for uri in list(static) + list(templates):
            assert uri.startswith("solidedge://"), uri

    def test_every_mime_type_is_json(self, mcp):
        static = _run(mcp.get_resources())
        templates = _run(mcp.get_resource_templates())
        for uri, res in list(static.items()) + list(templates.items()):
            assert res.mime_type == "application/json", uri

    def test_templates_are_exactly_the_parameterised_uris(self, mcp):
        templates = _run(mcp.get_resource_templates())
        assert set(templates) == {uri for uri, _ in resources.RESOURCES if "{" in uri}

    def test_every_resource_is_tagged_query(self, mcp):
        static = _run(mcp.get_resources())
        templates = _run(mcp.get_resource_templates())
        for uri, res in list(static.items()) + list(templates.items()):
            assert res.tags == {"query"}, uri

    def test_resources_run_on_the_com_thread(self, mcp):
        static = _run(mcp.get_resources())
        for uri, res in static.items():
            assert getattr(res.fn, "__wrapped__", None) is not None, uri


# === End-to-end reads through a FastMCP client ===


def _read(mcp, uri: str) -> object:
    async def go():
        async with Client(mcp) as client:
            out = await client.read_resource(uri)
            return json.loads(out[0].text)

    return _run(go())


class TestResourceReads:
    def test_static_resource_returns_manager_json(self, mcp, managers):
        managers["connection"].get_info.return_value = {"version": "226", "documents": 2}
        assert _read(mcp, "solidedge://app/info") == {"version": "226", "documents": 2}
        managers["connection"].get_info.assert_called_once_with()

    def test_static_resource_wraps_a_bare_value(self, mcp, managers):
        managers["connection"].is_connected.return_value = True
        assert _read(mcp, "solidedge://app/connection-status") == {"connected": True}

    def test_feature_list_resource(self, mcp, managers):
        managers["feature_manager"].list_features.return_value = {
            "features": [{"name": "Protrusion 1"}]
        }
        assert _read(mcp, "solidedge://model/features") == {"features": [{"name": "Protrusion 1"}]}

    def test_template_resource_passes_int_index(self, mcp, managers):
        managers["query_manager"].get_face_area.return_value = {"area": 0.01}
        assert _read(mcp, "solidedge://geometry/face/3/area") == {"area": 0.01}
        managers["query_manager"].get_face_area.assert_called_once_with(3)

    def test_two_parameter_template_passes_both(self, mcp, managers):
        managers["query_manager"].get_edge_info.return_value = {"length": 0.05}
        assert _read(mcp, "solidedge://geometry/face/2/edge/1") == {"length": 0.05}
        managers["query_manager"].get_edge_info.assert_called_once_with(2, 1)

    def test_float_template_passes_density(self, mcp, managers):
        managers["query_manager"].get_mass_properties.return_value = {"mass": 1.5}
        assert _read(mcp, "solidedge://geometry/mass-properties/7850.0") == {"mass": 1.5}
        managers["query_manager"].get_mass_properties.assert_called_once_with(7850.0)

    def test_named_template_passes_string(self, mcp, managers):
        managers["query_manager"].get_variable.return_value = {"value": 0.1}
        assert _read(mcp, "solidedge://model/variable/Width") == {"value": 0.1}
        managers["query_manager"].get_variable.assert_called_once_with("Width")

    def test_export_manager_backed_resource(self, mcp, managers):
        managers["export_manager"].get_drawing_view_count.return_value = {"count": 4}
        assert _read(mcp, "solidedge://drawing/view-count") == {"count": 4}


# === Direct handler calls ===


class _JsonStub:
    """Stands in for a manager: any method call returns a JSON-serialisable dict."""

    def __getattr__(self, name: str):
        def call(*args, **kwargs):
            return {"called": name, "args": list(args)}

        return call


@pytest.fixture
def json_managers(monkeypatch):
    for name in MANAGER_NAMES:
        monkeypatch.setattr(resources, name, _JsonStub())


class TestHandlers:
    def test_every_handler_returns_a_json_string(self, json_managers):
        for uri, fn in resources.RESOURCES:
            args = {
                "index": 0,
                "face": 0,
                "edge": 0,
                "name": "x",
                "density": 1.0,
            }
            kwargs = {k: v for k, v in args.items() if "{" + k + "}" in uri}
            out = fn(**kwargs)
            assert isinstance(out, str), uri
            json.loads(out)


class TestSpatialContextResource:
    def test_registered_at_the_documented_uri(self, mcp):
        static = _run(mcp.get_resources())
        assert "solidedge://spatial-context" in static
        assert static["solidedge://spatial-context"].mime_type == "application/json"
        assert static["solidedge://spatial-context"].tags == {"query"}

    def test_reads_from_the_query_manager(self, mcp, managers):
        payload = {
            "body_count": 1,
            "bounding_box": {"min": [0, 0, 0], "max": [1, 1, 1], "center": [0.5, 0.5, 0.5]},
            "centered_on_origin": False,
            "active_sketch": None,
            "plane_axis_map": {"1": {"name": "Top"}},
        }
        managers["query_manager"].get_spatial_context.return_value = payload
        assert _read(mcp, "solidedge://spatial-context") == payload
        managers["query_manager"].get_spatial_context.assert_called_once_with()


class TestPagedGeometryResources:
    """The paged geometry resources keep their URI and expose the page flags."""

    def test_faces_resource_forwards_the_default_page(self, mcp, managers):
        managers["query_manager"].get_body_faces.return_value = {
            "total": 900,
            "offset": 0,
            "limit": 200,
            "items": [{"index": 0}],
            "truncated": True,
        }
        payload = _read(mcp, "solidedge://geometry/faces")
        assert payload["total"] == 900
        assert payload["truncated"] is True
        managers["query_manager"].get_body_faces.assert_called_once_with()

    def test_edges_resource_forwards_the_default_page(self, mcp, managers):
        managers["query_manager"].get_body_edges.return_value = {
            "total": 4,
            "offset": 0,
            "limit": 200,
            "items": [],
            "truncated": False,
        }
        payload = _read(mcp, "solidedge://geometry/edges")
        assert payload["truncated"] is False
        managers["query_manager"].get_body_edges.assert_called_once_with()
