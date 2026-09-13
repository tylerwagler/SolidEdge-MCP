"""A single, long-lived COM worker thread.

COM interface pointers obtained on one apartment must not be used from another
thread without marshaling. FastMCP dispatches synchronous tool functions onto
an anyio worker-thread pool, so without intervention each tool call may run on
a different thread than the one that created the cached ``Application``,
``Document`` and ``Profile`` pointers held by the managers. That produces
``CO_E_NOTINITIALIZED`` on first use and ``RPC_E_WRONG_THREAD`` afterwards.

Every registered tool and resource is wrapped with :func:`on_com_thread`, which
marshals the call onto one dedicated STA thread that calls ``CoInitializeEx``
exactly once. Calls made while already on that thread run inline, so backend
code may call other backend code freely.
"""

from __future__ import annotations

import functools
import threading
from collections.abc import Callable
from concurrent.futures import ThreadPoolExecutor
from typing import ParamSpec, TypeVar

from solidedge_mcp.backends.logging import get_logger

_logger = get_logger(__name__)

P = ParamSpec("P")
R = TypeVar("R")

try:
    import pythoncom
except ImportError:  # pragma: no cover - non-Windows
    pythoncom = None


class ComThread:
    """Owns the one thread on which all COM traffic happens."""

    def __init__(self) -> None:
        self._executor: ThreadPoolExecutor | None = None
        self._thread_ident: int | None = None
        self._lock = threading.Lock()

    def _initializer(self) -> None:
        self._thread_ident = threading.get_ident()
        if pythoncom is not None:
            try:
                pythoncom.CoInitializeEx(pythoncom.COINIT_APARTMENTTHREADED)
                _logger.debug("COM apartment initialised on worker thread")
            except Exception as exc:  # already initialised, or non-COM host
                _logger.debug(f"CoInitializeEx skipped: {exc}")

    def _get_executor(self) -> ThreadPoolExecutor:
        with self._lock:
            if self._executor is None:
                self._executor = ThreadPoolExecutor(
                    max_workers=1,
                    thread_name_prefix="solidedge-com",
                    initializer=self._initializer,
                )
            return self._executor

    def is_current(self) -> bool:
        """True when called from the COM worker thread itself."""
        return self._thread_ident is not None and threading.get_ident() == self._thread_ident

    def run(self, fn: Callable[P, R], *args: P.args, **kwargs: P.kwargs) -> R:
        """Run ``fn`` on the COM thread and return its result (or re-raise)."""
        if self.is_current():
            return fn(*args, **kwargs)
        return self._get_executor().submit(fn, *args, **kwargs).result()

    def shutdown(self) -> None:
        with self._lock:
            if self._executor is not None:
                self._executor.shutdown(wait=False)
                self._executor = None
                self._thread_ident = None


com_thread = ComThread()


def on_com_thread(fn: Callable[P, R]) -> Callable[P, R]:
    """Decorator: run ``fn`` on the COM worker thread.

    Preserves the signature and annotations (via ``functools.wraps`` and
    ``__wrapped__``) so FastMCP still derives the JSON schema from ``fn``.
    """

    @functools.wraps(fn)
    def wrapper(*args: P.args, **kwargs: P.kwargs) -> R:
        return com_thread.run(fn, *args, **kwargs)

    return wrapper
