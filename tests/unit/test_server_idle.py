"""Regression: the server must not start FastMCP's docket worker.

The in-memory docket worker busy-polls (fakeredis) and burns a full CPU core
while the server is idle on Windows. This server uses no background tasks.
"""

import asyncio
from unittest.mock import patch

from fastmcp import Client

from solidedge_mcp.server import create_server


def _run(coro):
    return asyncio.run(coro)


def test_docket_lifespan_is_noop():
    server = create_server()
    assert server._docket_lifespan.__name__ == "_no_docket"


def test_lifespan_does_not_start_docket_or_worker():
    server = create_server()

    async def go():
        with (
            patch("fastmcp.server.server.Docket") as docket_cls,
            patch("fastmcp.server.server.Worker") as worker_cls,
        ):
            async with server._lifespan_manager():
                assert server._docket is None
                assert server._worker is None
            docket_cls.assert_not_called()
            worker_cls.assert_not_called()

    _run(go())


def test_list_tools_works_with_in_memory_client():
    server = create_server()

    async def go():
        async with Client(server) as client:
            return await client.list_tools()

    tools = _run(go())
    assert len(tools) > 0
