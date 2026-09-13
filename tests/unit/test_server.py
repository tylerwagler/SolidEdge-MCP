"""Smoke tests for the assembled server: registration, annotations, tags, prompts."""

from __future__ import annotations

import asyncio
import json
from unittest.mock import MagicMock

import pytest
from fastmcp import Client

from solidedge_mcp import server as server_module
from solidedge_mcp.server import create_server

EXPECTED_TAGS = {
    "app",
    "document",
    "sketch",
    "part",
    "assembly",
    "query",
    "export",
    "draft",
    "sheet_metal",
    "diagnostics",
}


@pytest.fixture(scope="module")
def mcp():
    return create_server()


def _run(coro):
    return asyncio.run(coro)


class TestRegistration:
    def test_module_level_server_exists(self):
        assert server_module.mcp.name == "Solid Edge MCP Server"

    def test_instructions_carry_the_core_contract(self, mcp):
        text = mcp.instructions or ""
        for needle in ("METERS", "DEGREES", "1=Top", "0-based", "close_sketch", "solidedge://"):
            assert needle in text

    def test_every_tool_has_annotations_and_tags(self, mcp):
        tools = _run(mcp.get_tools())
        assert len(tools) >= 100
        missing_annotations = [n for n, t in tools.items() if t.annotations is None]
        missing_tags = [n for n, t in tools.items() if not t.tags]
        assert missing_annotations == []
        assert missing_tags == []
        unknown_tags = {tag for t in tools.values() for tag in t.tags} - EXPECTED_TAGS
        assert unknown_tags == set()

    def test_read_only_and_destructive_hints_are_populated(self, mcp):
        tools = _run(mcp.get_tools())
        read_only = [n for n, t in tools.items() if t.annotations and t.annotations.readOnlyHint]
        destructive = [
            n for n, t in tools.items() if t.annotations and t.annotations.destructiveHint
        ]
        assert len(read_only) >= 5, read_only
        assert len(destructive) >= 3, destructive
        assert not (set(read_only) & set(destructive))

    def test_tools_run_on_com_thread(self, mcp):
        tools = _run(mcp.get_tools())
        for tool in tools.values():
            fn = tool.fn
            assert getattr(fn, "__wrapped__", None) is not None, tool.name

    def test_resources_and_guides(self, mcp):
        resources = _run(mcp.get_resources())
        templates = _run(mcp.get_resource_templates())
        assert len(resources) + len(templates) == 54  # 52 data + 2 guides
        assert "solidedge://guide/workflows" in resources
        assert "solidedge://guide/conventions" in resources
        for uri, res in resources.items():
            if uri.startswith("solidedge://guide/"):
                assert res.mime_type == "text/markdown"
            else:
                assert res.mime_type == "application/json", uri

    def test_prompts_registered(self, mcp):
        prompts = _run(mcp.get_prompts())
        assert {
            "new_part",
            "design_review",
            "manufacturability_check",
            "troubleshoot_feature",
        } <= set(prompts)


class TestEndToEnd:
    def test_guide_resource_readable(self, mcp):
        async def go():
            async with Client(mcp) as c:
                out = await c.read_resource("solidedge://guide/workflows")
                return out[0].text

        assert "create_sketch" in _run(go())

    def test_prompt_renders(self, mcp):
        async def go():
            async with Client(mcp) as c:
                out = await c.get_prompt("new_part", {"description": "a 10 mm cube"})
                return out.messages[0].content.text

        assert "a 10 mm cube" in _run(go())

    def test_tool_call_reaches_manager(self, mcp, monkeypatch):
        import solidedge_mcp.tools.connection as conn_tools

        fake = MagicMock()
        fake.get_info.return_value = {"version": "226"}
        fake.connect.return_value = {"status": "connected", "version": "226"}
        monkeypatch.setattr(conn_tools, "connection", fake)

        async def go():
            async with Client(mcp) as c:
                out = await c.call_tool("manage_connection", {"action": "connect"})
                return out.data if out.data is not None else json.loads(out.content[0].text)

        result = _run(go())
        assert result["status"] == "connected"
        fake.connect.assert_called_once()
