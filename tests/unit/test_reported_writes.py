"""Fail when a COM write is swallowed and then reported as applied.

    with contextlib.suppress(Exception):
        obj.Member = value
    return {"status": "added", "thing": value}

If Solid Edge refuses the write, the suppress hides it and the result tells
the caller the value was applied. The call succeeds, the dict looks right, and
nothing downstream can notice.

Every bug this found was of exactly that shape, and every one had shipped:

    TextBox.TextHeight            not a member; text boxes kept the default
    Leader.Text                   not a member; the leader was a bare arrow
    FeatureControlFrame.Text      not a member; the frame was placed empty
    variable.DisplayName          read-only; rename reported "not found"
    DraftPrintUtility.PaperWidth  millimetres, given meters; size never changed
    PMI.Show                      refused on a part with no PMI content

A write that genuinely may not take goes in the script's ALLOWED, and its
result must report what Solid Edge holds rather than what was asked for.
Needs no type library, so it never skips.
"""

from __future__ import annotations

import ast
import importlib.util
import pathlib
import sys
from types import ModuleType

import pytest

REPO_ROOT = pathlib.Path(__file__).resolve().parents[2]
AUDIT_SCRIPT = REPO_ROOT / "scripts" / "audit_reported_writes.py"


def _load() -> ModuleType:
    spec = importlib.util.spec_from_file_location("_audit_reported_writes", AUDIT_SCRIPT)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


@pytest.fixture(scope="module")
def audit() -> ModuleType:
    return _load()


def _fn(source: str) -> ast.FunctionDef:
    node = ast.parse(source).body[0]
    assert isinstance(node, ast.FunctionDef)
    return node


def test_no_swallowed_write_is_reported_as_applied(audit):
    findings = audit.audit()
    assert not findings, (
        "swallowed COM writes reported as applied:\n  "
        + "\n  ".join(str(f) for f in findings)
        + "\n\nLet the write raise, report what Solid Edge holds afterwards, or "
        "add the pair to ALLOWED once it has been driven live."
    )


def test_a_suppressed_write_is_caught(audit):
    fn = _fn(
        "def f(self, text):\n"
        "    with contextlib.suppress(Exception):\n"
        "        leader.Text = text\n"
        '    return {"status": "added", "text": text}\n'
    )
    assert audit.swallowed_writes(fn) == [("Text", "text", 3)]
    assert "text" in audit.reported_names(fn)


def test_a_try_that_passes_counts_as_swallowing(audit):
    fn = _fn(
        "def f(self, value):\n"
        "    try:\n"
        "        thing.Member = value\n"
        "    except Exception:\n"
        "        pass\n"
        '    return {"member": value}\n'
    )
    assert audit.swallowed_writes(fn) == [("Member", "value", 3)]


def test_an_unguarded_write_is_not_flagged(audit):
    """Letting it raise is the fix, so it must not still be a finding."""
    fn = _fn(
        "def f(self, text):\n"
        "    frame.PrimaryFrame = text\n"
        '    return {"status": "added", "text": text}\n'
    )
    assert audit.swallowed_writes(fn) == []


def test_a_write_nobody_reports_is_not_flagged(audit):
    """Best-effort settings like DisplayAlerts are not the target."""
    fn = _fn(
        "def f(self, app):\n"
        "    with contextlib.suppress(Exception):\n"
        "        app.DisplayAlerts = quiet\n"
        '    return {"status": "closed"}\n'
    )
    findings = audit.swallowed_writes(fn)
    assert findings and not (set(v for _m, v, _l in findings) & audit.reported_names(fn))
