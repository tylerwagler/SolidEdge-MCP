"""Fail when a COM property write is one Solid Edge will refuse.

Reading a member that does not exist raises, and the name and receiver checks
already cover that. Writing is worse: Solid Edge answers "Property 'Item.X'
can not be set.", and a surrounding try/except turns that into
``{"status": "set"}``, so the feature is dead and nothing says so. Four
shipped that way and were only found by reading the value back off the live
model:

    body.FaceStyle.Opacity          FaceStyle is on no interface; Body.Style is
    face.Color = rgb                Face has Style, not Color or SetColor
    view.ViewOrientation = const    a method with seven out-parameters
    occurrence.OccurrenceFileName   read-only; Occurrence.Replace is the setter

``scripts/audit_com_writes.py`` resolves each write's receiver with the same
type inference the receiver audit uses, then asks the type library whether
that member can be written at all. This pins the result at zero.
"""

from __future__ import annotations

import ast
import importlib.util
import json
import pathlib
import sys
from types import ModuleType

import pytest

REPO_ROOT = pathlib.Path(__file__).resolve().parents[2]
AUDIT_SCRIPT = REPO_ROOT / "scripts" / "audit_com_writes.py"
DUMP = REPO_ROOT / "reference" / "typelib_dump.json"

pytestmark = pytest.mark.skipif(
    not AUDIT_SCRIPT.exists() or not DUMP.exists(),
    reason="needs scripts/audit_com_writes.py and reference/typelib_dump.json",
)


def _load() -> ModuleType:
    spec = importlib.util.spec_from_file_location("_audit_com_writes", AUDIT_SCRIPT)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


@pytest.fixture(scope="module")
def audit() -> ModuleType:
    return _load()


@pytest.fixture(scope="module")
def result(audit) -> tuple[list, int]:
    return audit.audit()


def _writes(audit, source: str) -> list[tuple[str, str, str, str]]:
    receivers = audit.load_receivers()
    typelib = receivers.TypeLib(json.loads(DUMP.read_text(encoding="utf-8")))
    visitor = audit.build_visitor(receivers.Visitor, typelib)(REPO_ROOT / "x.py", typelib)
    visitor.visit(ast.parse(source))
    return visitor.writes


def test_no_write_lands_on_a_method_or_a_read_only_property(result):
    findings, _checked = result
    assert not findings, (
        "writes Solid Edge will refuse:\n  "
        + "\n  ".join(f"{loc}  {iface}.{member} {why}" for loc, iface, member, why in findings)
        + "\n\nUse the setter the type library declares, or drop the write."
    )


def test_the_audit_resolves_a_useful_number_of_writes(result):
    """A silent pass from broken inference would be worse than a failure."""
    _findings, checked = result
    assert checked >= 10, (
        f"only {checked} property writes resolved to a known interface. The "
        f"inference has probably broken, so a clean run means nothing."
    )


def test_writing_to_a_method_is_caught(audit):
    """DrawingView.ViewOrientation is a method, and this was the real bug."""
    source = (
        "def f(self):\n"
        "    doc = self.doc_manager.get_active_document()\n"
        "    view = doc.ActiveSheet.DrawingViews.Item(1)\n"
        "    view.ViewOrientation = 4\n"
    )
    found = _writes(audit, source)
    assert [(iface, member) for _loc, iface, member, _why in found] == [
        ("DrawingView", "ViewOrientation")
    ]
    assert "is a method" in found[0][3]


def test_writing_to_a_read_only_property_is_caught(audit):
    source = (
        "def f(self):\n"
        "    doc = self.doc_manager.get_active_document()\n"
        "    occurrence = doc.Occurrences.Item(1)\n"
        "    occurrence.OccurrenceFileName = 'x.par'\n"
    )
    found = _writes(audit, source)
    assert [(iface, member) for _loc, iface, member, _why in found] == [
        ("Occurrence", "OccurrenceFileName")
    ]
    assert "read-only" in found[0][3]


def test_a_writable_property_passes(audit):
    """Occurrence.Visible is get/put, so assigning to it is fine."""
    source = (
        "def f(self):\n"
        "    doc = self.doc_manager.get_active_document()\n"
        "    occurrence = doc.Occurrences.Item(1)\n"
        "    occurrence.Visible = True\n"
    )
    assert _writes(audit, source) == []
