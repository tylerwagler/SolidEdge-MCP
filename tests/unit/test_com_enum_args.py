"""Fail when a constant is in no member of the enum it is passed to.

Solid Edge does not police its enums: hand it a value from the wrong one and
it usually accepts the call, does nothing, and reports no error.

This covers a narrow slice of that, and the limit is worth stating. It fires
only when the value is in *no* member of the declared enum, so it catches a
stray literal, a local bound to one, and a value taken from the wrong
constants class when that value falls outside the target enum. It cannot
catch the wrong *member* of the right enum -- AddPartView passed 5 is a real
igBottomView, which is exactly why every "Front" drawing view came out a
bottom view and why only live verification found it.

Bitmask enums are handled: where every member is a power of two, 0 and any
combination of members are legitimate. Profile.End(0) means "no validation
criteria" and cuts a hole perfectly well.
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
AUDIT_SCRIPT = REPO_ROOT / "scripts" / "audit_com_enum_args.py"
DUMP = REPO_ROOT / "reference" / "typelib_dump.json"

pytestmark = pytest.mark.skipif(
    not AUDIT_SCRIPT.exists() or not DUMP.exists(),
    reason="needs scripts/audit_com_enum_args.py and reference/typelib_dump.json",
)


def _load() -> ModuleType:
    spec = importlib.util.spec_from_file_location("_audit_com_enum_args", AUDIT_SCRIPT)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


@pytest.fixture(scope="module")
def audit() -> ModuleType:
    return _load()


@pytest.fixture(scope="module")
def visitor_cls(audit):
    data = json.loads(DUMP.read_text(encoding="utf-8"))
    receivers = audit.load_receivers()
    typelib = receivers.TypeLib(data)
    cls = audit.build_visitor(
        receivers,
        audit.load_enums(data),
        audit.signatures(data),
        audit.property_types(data),
        audit.load_our_constants(),
    )
    return cls, typelib


def _findings(visitor_cls, source: str):
    cls, typelib = visitor_cls
    visitor = cls(REPO_ROOT / "x.py", typelib)
    visitor.visit(ast.parse(source))
    return visitor.enum_findings


def test_no_argument_is_outside_its_enum(audit):
    findings, checked = audit.audit()
    assert checked > 200, f"only {checked} arguments resolved; the audit has stopped looking"
    assert not findings, "constants outside the enum their parameter declares:\n  " + "\n  ".join(
        str(f) for f in findings
    )


def test_a_stray_literal_is_caught(visitor_cls):
    found = _findings(
        visitor_cls,
        "def f(self):\n"
        "    doc = self.doc_manager.get_active_document()\n"
        "    dvs = doc.ActiveSheet.DrawingViews\n"
        "    dvs.AddPartView(link, 999, 1.0, 0.1, 0.15, 0)\n",
    )
    assert [(f.param, f.value) for f in found] == [("Orientation", 999)]


def test_a_local_bound_to_one_is_caught(visitor_cls):
    found = _findings(
        visitor_cls,
        "def f(self):\n"
        "    doc = self.doc_manager.get_active_document()\n"
        "    orient = 999\n"
        "    dvs = doc.ActiveSheet.DrawingViews\n"
        "    dvs.AddPartView(link, orient, 1.0, 0.1, 0.15, 0)\n",
    )
    assert [f.value for f in found] == [999]


def test_the_wrong_constants_class_is_caught_when_it_falls_outside(visitor_cls):
    """ExtentTypeConstants.igNone is 44; ViewOrientationConstants stops short."""
    found = _findings(
        visitor_cls,
        "def f(self):\n"
        "    doc = self.doc_manager.get_active_document()\n"
        "    dvs = doc.ActiveSheet.DrawingViews\n"
        "    dvs.AddPartView(link, ExtentTypeConstants.igNone, 1.0, 0.1, 0.15, 0)\n",
    )
    assert [(f.value, f.enum) for f in found] == [(44, "ViewOrientationConstants")]


def test_the_right_constant_passes(visitor_cls):
    found = _findings(
        visitor_cls,
        "def f(self):\n"
        "    doc = self.doc_manager.get_active_document()\n"
        "    dvs = doc.ActiveSheet.DrawingViews\n"
        "    dvs.AddPartView(link, ViewOrientationConstants.igFrontView, 1.0, 0.1, 0.15, 0)\n",
    )
    assert found == []


def test_the_wrong_member_of_the_right_enum_is_not_claimed(visitor_cls):
    """The documented limit. 5 is igBottomView: a real member, wrongly chosen.

    This is what made every "Front" view a bottom view, and it is invisible
    here by construction -- the test exists so nobody mistakes a clean run for
    proof that the constants are semantically right.
    """
    found = _findings(
        visitor_cls,
        "def f(self):\n"
        "    doc = self.doc_manager.get_active_document()\n"
        "    dvs = doc.ActiveSheet.DrawingViews\n"
        "    dvs.AddPartView(link, 5, 1.0, 0.1, 0.15, 0)\n",
    )
    assert found == []


def test_a_bitmask_enum_accepts_zero_and_combinations(audit):
    flags = {"igProfileClosed": 1, "igProfileSingle": 4, "igProfileRefAxisRequired": 16}
    assert audit.is_flags(flags)
    assert audit.accepts(flags, 0)
    assert audit.accepts(flags, 17)  # igProfileClosed | igProfileRefAxisRequired
    assert not audit.accepts(flags, 64)


def test_a_choice_enum_does_not_accept_zero(audit):
    choices = {"igTopView": 1, "igRightView": 2, "igLeftView": 3, "igFrontView": 4}
    assert not audit.is_flags(choices)
    assert not audit.accepts(choices, 0)
    assert audit.accepts(choices, 4)
