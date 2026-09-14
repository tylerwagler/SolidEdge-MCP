"""Input validation helpers for MCP tools.

Provides two small helpers used by the tool layer:

- ``validate_numerics`` — guards numeric parameters against non-numbers,
  NaN, and infinities before they are handed to COM calls.
- ``validate_path`` — normalizes a filesystem path and optionally checks
  that it exists.

Both return error payloads in the same ``{"error": ...}`` shape that the
rest of the tool layer uses, so callers can simply ``return err``.
"""

from __future__ import annotations

import math
import os
from pathlib import Path
from typing import Any


def validate_numerics(**values: Any) -> dict[str, Any] | None:
    """Validate that keyword values are finite real numbers.

    ``None`` values are skipped (treated as "not provided") so optional
    parameters can be forwarded directly. Booleans are rejected because they
    are almost never an intended numeric input here.

    Returns an error dict on the first invalid value, or ``None`` if all
    values are valid.
    """
    for name, value in values.items():
        if value is None:
            continue
        if isinstance(value, bool) or not isinstance(value, (int, float)):
            return {"error": (f"Parameter '{name}' must be a number, got {type(value).__name__}")}
        if not math.isfinite(float(value)):
            return {"error": f"Parameter '{name}' must be finite, got {value!r}"}
    return None


def validate_path(path: str, must_exist: bool = False) -> tuple[str, dict[str, Any] | None]:
    """Normalize a filesystem path and optionally verify it exists.

    Returns a ``(normalized_path, error)`` tuple. ``error`` is ``None`` on
    success. When ``must_exist`` is true and the path is absent, an error
    dict is returned alongside the normalized path.
    """
    if not isinstance(path, str) or not path.strip():
        return path, {"error": "A non-empty file path is required"}

    normalized = os.path.normpath(os.path.expanduser(os.path.expandvars(path)))

    if must_exist and not os.path.exists(normalized):
        return normalized, {"error": f"Path does not exist: {normalized}"}

    return normalized, None


def guard_overwrite(file_path: str, overwrite: bool) -> dict[str, Any] | None:
    """Refuse to write over an existing file, or clear the way for it.

    Every Solid Edge call that writes a path answers an existing one with a
    modal "This file exists. Do you want to overwrite it?" prompt.
    Application.DisplayAlerts does not suppress it, and Solid Edge has a
    single UI thread, so the COM call never returns and every later call
    queues behind it. An automation client sees the whole server stop
    answering. Reproduced and photographed on Solid Edge 2026.

    Returns None when the write may go ahead, or the error to return.
    """
    target = Path(file_path)
    if not target.exists():
        return None
    if not overwrite:
        return {
            "error": (
                f"{target} already exists. Solid Edge would raise a modal overwrite "
                f"prompt, which blocks the server until somebody clicks it. Pass "
                f"overwrite=true to replace the file, or choose another path."
            ),
            "path": str(target),
            "exists": True,
        }
    try:
        target.unlink()
    except OSError as exc:
        return {
            "error": (
                f"{target} exists and could not be removed, so the write would stop "
                f"on an overwrite prompt: {exc}"
            ),
            "path": str(target),
        }
    return None
