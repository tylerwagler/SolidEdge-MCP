"""Static guide resources: workflows and conventions for the LLM caller."""

from __future__ import annotations

from typing import Any

from solidedge_mcp.prompts import CONVENTIONS_GUIDE, WORKFLOWS_GUIDE


def register(mcp: Any) -> None:
    """Register the two guide resources. Pure text; no COM, no thread hop."""

    @mcp.resource("solidedge://guide/workflows", mime_type="text/markdown", tags={"guide"})
    def guide_workflows() -> str:
        """Step-by-step tool sequences for common Solid Edge tasks."""
        return WORKFLOWS_GUIDE

    @mcp.resource("solidedge://guide/conventions", mime_type="text/markdown", tags={"guide"})
    def guide_conventions() -> str:
        """Units, index bases, sketch lifecycle, error shape, and known COM limits."""
        return CONVENTIONS_GUIDE
