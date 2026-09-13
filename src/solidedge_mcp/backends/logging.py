"""Logging for the Solid Edge MCP server.

All output goes to **stderr**. On the stdio transport, stdout is the JSON-RPC
channel, so anything written there corrupts the protocol stream.

The level is read from the ``SOLIDEDGE_MCP_LOG_LEVEL`` environment variable
(DEBUG, INFO, WARNING, ERROR); it defaults to WARNING so a normal session is
quiet. Set ``SOLIDEDGE_MCP_DEBUG=1`` to force DEBUG and to include Python
tracebacks in error results (see :mod:`solidedge_mcp.backends.errors`).
"""

from __future__ import annotations

import logging
import os
import sys

ROOT_LOGGER_NAME = "solidedge_mcp"


def _level_from_env() -> int:
    if os.environ.get("SOLIDEDGE_MCP_DEBUG", "").strip().lower() in {"1", "true", "yes", "on"}:
        return logging.DEBUG
    name = os.environ.get("SOLIDEDGE_MCP_LOG_LEVEL", "WARNING").strip().upper()
    return logging.getLevelNamesMapping().get(name, logging.WARNING)


def configure_logging(level: int | None = None) -> logging.Logger:
    """Attach a single stderr handler to the package logger (idempotent)."""
    root = logging.getLogger(ROOT_LOGGER_NAME)
    if not any(getattr(h, "_solidedge_mcp", False) for h in root.handlers):
        handler = logging.StreamHandler(sys.stderr)
        handler.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s"))
        handler._solidedge_mcp = True  # type: ignore[attr-defined]
        root.addHandler(handler)
    root.setLevel(level if level is not None else _level_from_env())
    # Never bubble up into a host application's root logger config.
    root.propagate = False
    return root


logger: logging.Logger = configure_logging()


def get_logger(name: str) -> logging.Logger:
    """Return a child logger, e.g. ``get_logger(__name__)``."""
    if name.startswith(ROOT_LOGGER_NAME):
        return logging.getLogger(name)
    return logging.getLogger(f"{ROOT_LOGGER_NAME}.{name}")
