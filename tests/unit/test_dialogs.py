"""backends/dialogs.py: the scoped dismisser clicks only the dialog it was told about."""

from __future__ import annotations

import time

from solidedge_mcp.backends import dialogs


def _run(monkeypatch, windows, prefix):
    clicked: list[int] = []
    monkeypatch.setattr(dialogs, "_visible_dialogs", lambda: windows)
    monkeypatch.setattr(dialogs, "_click_ok", lambda hwnd: clicked.append(hwnd) or True)
    monkeypatch.setattr(dialogs, "_user32", lambda: object())
    with dialogs.dismiss_informational_dialog(prefix, poll_seconds=0.01) as d:
        time.sleep(0.08)
    return clicked, d


def test_the_named_dialog_is_clicked_and_counted(monkeypatch):
    windows = [(7, "Solid Edge", ["OK", "The Segments group of commands are replaced ..."])]
    clicked, d = _run(monkeypatch, windows, "The Segments group of commands")
    assert 7 in clicked
    assert d.dismissed >= 1


def test_other_dialogs_are_left_alone(monkeypatch):
    windows = [(8, "Solid Edge", ["OK", "This file exists. Do you want to overwrite it?"])]
    clicked, d = _run(monkeypatch, windows, "The Segments group of commands")
    assert clicked == []
    assert d.dismissed == 0


def test_without_user32_it_is_inert(monkeypatch):
    monkeypatch.setattr(dialogs, "_user32", lambda: None)
    with dialogs.dismiss_informational_dialog("x") as d:
        pass
    assert d.dismissed == 0
