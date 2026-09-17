"""The counts the docs publish must be the counts the server registers.

Every number in ``reference/VERIFICATION_STATUS.md`` and the ``Surface:`` line
of ``CLAUDE.md`` is meant to be measured, not estimated -- and they drift the
moment nobody re-measures. Both were wrong on the day this test was written:
the status document said 119 tools and CLAUDE.md said 117, because one count
had included ``def register_tool`` itself and the other had never been updated.
The real number was 118.

This makes the documents a ratchet: register a tool and the published count
fails until the document is re-measured alongside it.
"""

from __future__ import annotations

import asyncio
import pathlib
import re

import pytest

from solidedge_mcp.server import create_server

ROOT = pathlib.Path(__file__).resolve().parent.parent.parent
STATUS = ROOT / "reference" / "VERIFICATION_STATUS.md"
CLAUDE = ROOT / "CLAUDE.md"


@pytest.fixture(scope="module")
def registered():
    mcp = create_server()
    tools = asyncio.run(mcp.get_tools())
    resources = asyncio.run(mcp.get_resources())
    templates = asyncio.run(mcp.get_resource_templates())
    guides = [k for k in resources if "/guide/" in str(k)]
    return {
        "tools": len(tools),
        "data_resources": len(resources) + len(templates) - len(guides),
        "guides": len(guides),
    }


def _status_row(label: str) -> str:
    text = STATUS.read_text(encoding="utf-8")
    match = re.search(rf"^\|\s*{re.escape(label)}\s*\|\s*([^|]+?)\s*\|", text, re.MULTILINE)
    assert match, f"no '| {label} |' row in {STATUS.name}"
    return match.group(1)


class TestVerificationStatusIsMeasured:
    def test_the_tool_count(self, registered):
        published = int(_status_row("Tools"))
        assert published == registered["tools"], (
            f"{STATUS.name} says {published} tools; the server registers "
            f"{registered['tools']}. Re-measure the document."
        )

    def test_the_resource_count(self, registered):
        cell = _status_row("Resources")  # e.g. "53 + 2 guides"
        numbers = [int(n) for n in re.findall(r"\d+", cell)]
        assert numbers[:2] == [registered["data_resources"], registered["guides"]], (
            f"{STATUS.name} says {cell!r}; the server registers "
            f"{registered['data_resources']} data resources + {registered['guides']} guides."
        )


class TestClaudeMdSurfaceLine:
    def test_matches_the_server(self, registered):
        text = CLAUDE.read_text(encoding="utf-8")
        match = re.search(
            r"^Surface:\s*(\d+) tools,\s*(\d+) data resources \+ (\d+) guide resources",
            text,
            re.MULTILINE,
        )
        assert match, (
            "CLAUDE.md has no 'Surface: N tools, M data resources + K guide resources' line"
        )
        tools, data, guides = (int(g) for g in match.groups())
        assert (tools, data, guides) == (
            registered["tools"],
            registered["data_resources"],
            registered["guides"],
        ), (
            f"CLAUDE.md says {tools}/{data}/{guides}; the server registers "
            f"{registered['tools']}/{registered['data_resources']}/{registered['guides']}."
        )
