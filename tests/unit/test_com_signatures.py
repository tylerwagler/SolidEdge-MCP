"""Fail if any COM call passes the wrong number of arguments.

An audit of this repository found 143 calls whose argument count could not
match the Solid Edge type library. One of them, ``NormalCutouts.AddFiniteMulti``
missing its ``Method`` argument, was confirmed against a live Solid Edge 2026:
it raised ``0x8002000F Parameter not optional`` and the feature had never
worked. Mocked unit tests cannot catch this, because a ``MagicMock`` accepts any
call. All 143 are now fixed, so the bar here is simply zero.

The checking logic lives in ``scripts/audit_com_signatures.py`` so it can also
be run by hand while fixing a call:

    uv run python scripts/audit_com_signatures.py --filter backends/features/_holes.py

Needs ``reference/typelib_dump.json`` (gitignored, ~20 MB). Regenerate with
``uv run python scripts/scrape_typelibs.py``; without it these tests skip, so a
fresh clone and CI stay green.
"""

from __future__ import annotations

import importlib.util
import pathlib
import sys
from types import ModuleType

import pytest

REPO_ROOT = pathlib.Path(__file__).resolve().parents[2]
DUMP_PATH = REPO_ROOT / "reference" / "typelib_dump.json"
AUDIT_SCRIPT = REPO_ROOT / "scripts" / "audit_com_signatures.py"

pytestmark = pytest.mark.skipif(
    not DUMP_PATH.exists(),
    reason="reference/typelib_dump.json absent; run scripts/scrape_typelibs.py",
)


def _load_audit() -> ModuleType:
    """Import the audit script by path; scripts/ is not a package."""
    spec = importlib.util.spec_from_file_location("_audit_com_signatures", AUDIT_SCRIPT)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


@pytest.fixture(scope="module")
def audit() -> ModuleType:
    assert AUDIT_SCRIPT.exists(), f"missing {AUDIT_SCRIPT}"
    return _load_audit()


@pytest.fixture(scope="module")
def findings(audit) -> tuple[list, list]:
    precise, loose, _ = audit.collect()
    return precise, loose


def test_the_audit_is_actually_looking(audit):
    """Guard against a silent pass from a broken index or an empty scan."""
    iface_methods, any_windows, _ = audit.load_typelibs()
    assert len(iface_methods) > 1000, "type library interface index looks empty"
    assert any_windows.get("AddFiniteMulti"), "a known method is missing from the index"
    # The interface the confirmed bug lived on, with its corrected arity.
    assert iface_methods["NormalCutouts"]["AddFiniteMulti"] == (5, 5)


def test_no_call_has_an_impossible_argument_count(findings, audit):
    """Every COM call must fit the signature of the interface it targets."""
    precise, _ = findings
    assert not precise, "COM calls with an impossible argument count:\n  " + "\n  ".join(
        audit.fmt(f) for f in precise
    )


def test_no_call_mismatches_every_overload(findings, audit):
    """Calls whose receiver could not be resolved must still fit some overload."""
    _, loose = findings
    assert not loose, "COM calls matching no overload of that method name:\n  " + "\n  ".join(
        audit.fmt(f) for f in loose
    )


def test_normal_cutout_passes_its_method_argument():
    """Regression: the call that proved this whole class of bug is real."""
    source = (
        REPO_ROOT / "src" / "solidedge_mcp" / "backends" / "features" / "_cutout.py"
    ).read_text(encoding="utf-8")
    assert "NormalCutoutMethodConstants.igSMFaceCutout" in source


def test_the_receiver_is_resolved_by_type_inference(audit):
    """A call on an inferred receiver is checked against that interface alone.

    The fallback is deliberately weak: an unresolved receiver only has to fit
    *some* interface with a method of that name. ``occurrence.Replace(path)``
    passed for exactly that reason -- some other ``Replace`` takes one argument
    -- while ``Occurrence.Replace`` requires two, so replace_component could
    never have replaced anything.
    """
    import ast

    iface_methods, any_windows, data = audit.load_typelibs()
    receivers = audit.load_receivers()
    typelib = receivers.TypeLib(data)
    visitor_cls = audit.build_inferring_visitor(receivers, typelib, iface_methods, any_windows)

    source = (
        "def f(self):\n"
        "    doc = self.doc_manager.get_active_document()\n"
        "    occurrence = doc.Occurrences.Item(1)\n"
        "    occurrence.Replace('x.par')\n"
    )
    visitor = visitor_cls(REPO_ROOT / "x.py", typelib)
    visitor.visit(ast.parse(source))

    findings = visitor.checker.findings
    assert [(f[1], f[2], f[3]) for f in findings] == [("Occurrence", "Replace", 1)]
    assert findings[0][5] is True, "should be a precise finding, not a loose one"


def test_a_correct_call_on_an_inferred_receiver_passes(audit):
    import ast

    iface_methods, any_windows, data = audit.load_typelibs()
    receivers = audit.load_receivers()
    typelib = receivers.TypeLib(data)
    visitor_cls = audit.build_inferring_visitor(receivers, typelib, iface_methods, any_windows)

    source = (
        "def f(self):\n"
        "    doc = self.doc_manager.get_active_document()\n"
        "    occurrence = doc.Occurrences.Item(1)\n"
        "    occurrence.Replace('x.par', False)\n"
    )
    visitor = visitor_cls(REPO_ROOT / "x.py", typelib)
    visitor.visit(ast.parse(source))

    assert visitor.checker.findings == []
