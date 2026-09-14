"""Uniform error results for backend operations.

Every backend method returns ``{"error": ...}`` on failure. This module turns a
caught exception into that dict, decoding ``pywintypes.com_error`` into a
readable message (hex HRESULT, symbolic name where known, and the description
Solid Edge attached to the exception) instead of the raw tuple ``str()`` form.

Tracebacks are only included when the ``SOLIDEDGE_MCP_DEBUG`` environment
variable is truthy, so a failed tool call does not bill the caller for a
multi-frame Python traceback.
"""

from __future__ import annotations

import os
import traceback
from typing import Any

try:  # pywintypes only exists on Windows
    from pywintypes import com_error as _ComError
except ImportError:  # pragma: no cover - non-Windows test environments

    class _ComError(Exception):  # type: ignore[no-redef]
        """Placeholder so isinstance checks work off-Windows."""


# Values from winerror.h / oleauto. Add entries as they show up in practice.
HRESULT_NAMES: dict[int, str] = {
    0x80004001: "E_NOTIMPL",
    0x80004002: "E_NOINTERFACE",
    0x80004003: "E_POINTER",
    0x80004004: "E_ABORT",
    0x80004005: "E_FAIL",
    0x80020003: "DISP_E_MEMBERNOTFOUND",
    0x80020004: "DISP_E_PARAMNOTFOUND",
    0x80020005: "DISP_E_TYPEMISMATCH",
    0x80020006: "DISP_E_UNKNOWNNAME",
    0x80020009: "DISP_E_EXCEPTION",
    0x8002000E: "DISP_E_BADPARAMCOUNT",
    0x80070005: "E_ACCESSDENIED",
    0x8007000E: "E_OUTOFMEMORY",
    0x80070057: "E_INVALIDARG",
    0x800401E3: "MK_E_UNAVAILABLE",
    0x800706BA: "RPC_S_SERVER_UNAVAILABLE",
    0x800706BE: "RPC_S_CALL_FAILED",
    0x80010108: "RPC_E_DISCONNECTED",
    0x8001010A: "RPC_E_SERVERCALL_RETRYLATER",
    0x8001010E: "RPC_E_WRONG_THREAD",
    0x800401F0: "CO_E_NOTINITIALIZED",
}

# HRESULTs that mean the COM server (Solid Edge) is gone or unreachable.
DISCONNECTED_HRESULTS: frozenset[int] = frozenset(
    {0x800706BA, 0x800706BE, 0x80010108, 0x800401E3, 0x8001010A}
)


def debug_enabled() -> bool:
    """True when SOLIDEDGE_MCP_DEBUG asks for tracebacks in error results."""
    return os.environ.get("SOLIDEDGE_MCP_DEBUG", "").strip().lower() in {"1", "true", "yes", "on"}


def _unsigned(hresult: int) -> int:
    return hresult & 0xFFFFFFFF


def com_hresult(exc: BaseException) -> int | None:
    """Return the unsigned HRESULT of a com_error, else None."""
    if not isinstance(exc, _ComError):
        return None
    hresult = getattr(exc, "hresult", None)
    if not isinstance(hresult, int):
        args = getattr(exc, "args", ())
        if args and isinstance(args[0], int):
            hresult = args[0]
        else:
            return None
    # DISP_E_EXCEPTION carries the real code inside excepinfo[5]
    excepinfo = getattr(exc, "excepinfo", None)
    if _unsigned(hresult) == 0x80020009 and excepinfo and len(excepinfo) >= 6:
        inner = excepinfo[5]
        if isinstance(inner, int) and inner != 0:
            return _unsigned(inner)
    return _unsigned(hresult)


def is_disconnected_error(exc: BaseException) -> bool:
    """True when the exception means Solid Edge is no longer reachable."""
    hresult = com_hresult(exc)
    return hresult is not None and hresult in DISCONNECTED_HRESULTS


def describe_exception(exc: BaseException) -> str:
    """Human-readable, single-line description of any exception."""
    if not isinstance(exc, _ComError):
        return str(exc)

    hresult = com_hresult(exc)
    parts: list[str] = []
    if hresult is not None:
        name = HRESULT_NAMES.get(hresult)
        parts.append(f"COM error 0x{hresult:08X}" + (f" ({name})" if name else ""))
    else:
        parts.append("COM error")

    description = None
    excepinfo = getattr(exc, "excepinfo", None)
    if excepinfo and len(excepinfo) >= 3 and excepinfo[2]:
        description = str(excepinfo[2]).strip()
    if not description:
        strerror = getattr(exc, "strerror", None)
        if strerror:
            description = str(strerror).strip()
    if description:
        parts.append(description)
    return ": ".join(parts)


def error_result(exc: BaseException, **extra: Any) -> dict[str, Any]:
    """Build the standard ``{"error": ...}`` dict for a caught exception.

    Extra keyword arguments are merged into the result so callers can attach
    context (``feature="extrude"``) without rebuilding the dict by hand.

    A ``context`` string is treated specially: it leads the error message
    rather than sitting in a key beside it. "COM error 0x80004005" on its own
    tells a caller nothing, and the sentence explaining it is the whole point
    of passing one, so it must be in the field everybody reads.
    """
    described = describe_exception(exc)
    context = extra.pop("context", None)
    message = f"{context} ({described})" if context else described

    result: dict[str, Any] = {"error": message}
    hresult = com_hresult(exc)
    if hresult is not None:
        result["hresult"] = f"0x{hresult:08X}"
        if hresult in DISCONNECTED_HRESULTS:
            result["disconnected"] = True
    if context:
        result["com_error"] = described
    if debug_enabled():
        result["traceback"] = traceback.format_exc()
    result.update(extra)
    return result
