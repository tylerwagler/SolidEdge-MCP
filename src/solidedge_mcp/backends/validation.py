"""Shared argument validation helpers for tool-layer entry points."""

from __future__ import annotations

from math import isfinite
from numbers import Real
from pathlib import Path
from typing import Any

ErrorResult = dict[str, str]


def validate_numerics(**values: Any) -> ErrorResult | None:
    """Validate numeric MCP arguments before dispatching to Solid Edge."""
    for name, value in values.items():
        if isinstance(value, bool) or not isinstance(value, Real):
            return {"error": f"{name} must be a finite number"}
        if not isfinite(float(value)):
            return {"error": f"{name} must be a finite number"}
    return None


def validate_path(path: str, *, must_exist: bool) -> tuple[str, ErrorResult | None]:
    """Normalize a file path and optionally require that it already exists."""
    if not isinstance(path, str) or not path.strip():
        return path, {"error": "file path is required"}

    normalized = str(Path(path).expanduser())
    file_path = Path(normalized)

    if must_exist and not file_path.exists():
        return normalized, {"error": f"file path does not exist: {normalized}"}

    if not must_exist:
        parent = file_path.parent
        if str(parent) not in ("", ".") and not parent.exists():
            return normalized, {"error": f"parent directory does not exist: {parent}"}

    return normalized, None
