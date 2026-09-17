"""Document creators must refuse before they can hang, and must not lie.

create_weldment without a template used to call
Documents.Add("SolidEdge.WeldmentDocument"). Solid Edge accepts the ProgID,
then looks for a default weldment template; on an install without the
weldment environment that file does not exist, and the answer is a modal
"Path not found" that DisplayAlerts does not suppress -- the one UI thread
blocks and the server hangs until someone clicks it. Every creator also used
to ignore a template path that was not on disk and fall through to the
default document.

These run with a modal watchdog so that, if a regression ever reintroduces
the dialog, the test fails and reports it instead of hanging the suite.

Run with: uv run pytest -m integration
"""

from __future__ import annotations

import ctypes
import threading
import time

import pytest

pytestmark = pytest.mark.integration


class _Watchdog:
    """Dismiss any #32770 dialog while active, and remember its title."""

    def __init__(self) -> None:
        self.titles: list[str] = []
        self._stop = threading.Event()
        self._thread = threading.Thread(target=self._run, daemon=True)

    def _run(self) -> None:
        user32 = ctypes.windll.user32
        proto = ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
        while not self._stop.is_set():

            def cb(hwnd, _):
                cls = ctypes.create_unicode_buffer(256)
                user32.GetClassNameW(hwnd, cls, 256)
                if cls.value == "#32770" and user32.IsWindowVisible(hwnd):
                    title = ctypes.create_unicode_buffer(512)
                    user32.GetWindowTextW(hwnd, title, 512)
                    self.titles.append(title.value)
                    btn = user32.GetDlgItem(hwnd, 1) or user32.GetDlgItem(hwnd, 2)
                    if btn:
                        user32.SendMessageW(btn, 0x00F5, 0, 0)
                return True

            user32.EnumWindows(proto(cb), 0)
            time.sleep(0.3)

    def __enter__(self) -> _Watchdog:
        self._thread.start()
        return self

    def __exit__(self, *_: object) -> None:
        self._stop.set()
        self._thread.join(timeout=2)


class TestCreateWeldment:
    def test_without_a_template_it_refuses_at_once_with_no_dialog(self, stack):
        with _Watchdog() as dog:
            started = time.time()
            result = stack.doc.create_weldment()
            elapsed = time.time() - started

        assert "error" in result, result
        assert result.get("unsupported") is True
        assert elapsed < 5, f"took {elapsed:.1f}s -- was it waiting on something?"
        assert dog.titles == [], f"a dialog appeared: {dog.titles}"
        assert stack.connection.get_application().Documents.Count == 0


class TestAMissingTemplateIsRefused:
    @pytest.mark.parametrize(
        "creator", ["create_part", "create_assembly", "create_sheet_metal", "create_draft"]
    )
    def test_before_any_document_is_made(self, stack, creator):
        app = stack.connection.get_application()
        before = app.Documents.Count

        with _Watchdog() as dog:
            result = getattr(stack.doc, creator)(template=r"C:\nowhere\template.xxx")

        assert "error" in result, (creator, result)
        assert "Template not found" in result["error"]
        assert app.Documents.Count == before, "nothing may have been created"
        assert dog.titles == [], dog.titles
