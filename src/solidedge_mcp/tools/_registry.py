"""Helpers for registering tools and resources with FastMCP.

Every registration goes through here so that:

* the callable is marshalled onto the single COM worker thread
  (:mod:`solidedge_mcp.backends.com_thread`);
* MCP tool annotations (``readOnlyHint`` / ``destructiveHint`` /
  ``idempotentHint``) are always set, so clients can gate destructive tools;
* tags are applied consistently, so a client can include or exclude whole
  categories (``part``, ``assembly``, ``draft``, ``sheet_metal`` ...).
"""

from __future__ import annotations

import functools
from collections.abc import Callable
from typing import Any

from solidedge_mcp.backends.com_thread import on_com_thread
from solidedge_mcp.backends.errors import error_result

JSON_MIME = "application/json"


def _never_raises(fn: Callable[..., Any]) -> Callable[..., Any]:
    """Turn an escaped exception into the error dict every tool promises.

    Backend methods catch their own COM errors, but a tool that reaches COM
    itself can still raise -- diagnose_api did, and FastMCP turned that into a
    protocol-level failure with a traceback instead of the
    ``{"error": ...}`` shape callers handle. One tool breaking that contract
    is one too many, so the guarantee is enforced here rather than trusted to
    every author.
    """

    @functools.wraps(fn)
    def wrapper(*args: Any, **kwargs: Any) -> Any:
        try:
            return fn(*args, **kwargs)
        except Exception as exc:  # noqa: BLE001 - the whole point is to catch all
            return error_result(exc, context=f"{fn.__name__} failed")

    return wrapper


def tool_annotations(
    *,
    read_only: bool = False,
    destructive: bool = False,
    idempotent: bool = False,
) -> dict[str, Any]:
    """Build an MCP ToolAnnotations dict. Solid Edge is a closed, local world."""
    return {
        "readOnlyHint": read_only,
        "destructiveHint": destructive,
        "idempotentHint": idempotent,
        "openWorldHint": False,
    }


def register_tool(
    mcp: Any,
    fn: Callable[..., Any],
    *,
    tags: set[str] | None = None,
    read_only: bool = False,
    destructive: bool = False,
    idempotent: bool = False,
) -> Any:
    """Register ``fn`` as an MCP tool, run on the COM thread, with annotations."""
    return mcp.tool(
        on_com_thread(_never_raises(fn)),
        tags=tags,
        annotations=tool_annotations(
            read_only=read_only, destructive=destructive, idempotent=idempotent
        ),
    )


def register_resource(
    mcp: Any,
    uri: str,
    fn: Callable[..., Any],
    *,
    tags: set[str] | None = None,
    mime_type: str = JSON_MIME,
) -> Any:
    """Register ``fn`` as an MCP resource (or template), run on the COM thread."""
    return mcp.resource(uri, mime_type=mime_type, tags=tags)(on_com_thread(fn))
