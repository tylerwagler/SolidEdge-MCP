"""Fail when a COM member is read off an interface that does not have it.

The member-name check in test_com_members.py asks whether a name exists
anywhere in Solid Edge. A name can pass that and still be wrong where it is
used, and two real bugs got through exactly that way:

    model.RevolvedSurfaces      real, but it belongs to Constructions, so
                                every revolved surface call raised
    line.StartPoint.X           real on other interfaces; Line2d has
                                GetStartPoint(). Reading it inside a bare
                                except made sketch rotate and scale delete the
                                sketch and rebuild nothing

``scripts/audit_com_receivers.py`` infers what each receiver is by following
the declared types in the type library dump, then checks the member against
that interface. This pins the result at zero, less the entries below.
"""

from __future__ import annotations

import importlib.util
import pathlib
import sys
from types import ModuleType

import pytest

REPO_ROOT = pathlib.Path(__file__).resolve().parents[2]
AUDIT_SCRIPT = REPO_ROOT / "scripts" / "audit_com_receivers.py"
DUMP = REPO_ROOT / "reference" / "typelib_dump.json"

#: (interfaces, member) pairs that are wrong by the type library and right in
#: practice, or that the code already guards. Keep this list short and say why.
KNOWN: frozenset[tuple[str, str]] = frozenset(
    {
        # SaveAsPLMXML is declared only on the generic SolidEdgeDocument
        # interface. No concrete document answers it -- hasattr reads False on
        # a part in Solid Edge 2026 -- and export_to_plmxml checks for it and
        # returns an unsupported error rather than letting the call raise.
        (
            "AssemblyDocument,DraftDocument,PartDocument,SheetMetalDocument,WeldmentDocument",
            "SaveAsPLMXML",
        ),
    }
)

pytestmark = pytest.mark.skipif(
    not AUDIT_SCRIPT.exists() or not DUMP.exists(),
    reason="needs scripts/audit_com_receivers.py and reference/typelib_dump.json",
)


def _load() -> ModuleType:
    spec = importlib.util.spec_from_file_location("_audit_com_receivers", AUDIT_SCRIPT)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


@pytest.fixture(scope="module")
def audit() -> ModuleType:
    return _load()


@pytest.fixture(scope="module")
def result(audit) -> tuple[list, list, object]:
    return audit.audit()


def test_no_member_is_read_off_the_wrong_interface(result, audit):
    findings, _resolved, typelib = result
    unexplained = [f for f in findings if (f[1], f[2]) not in KNOWN]
    assert not unexplained, (
        "members read off an interface that does not have them:\n  "
        + "\n  ".join(audit.fmt(f, typelib) for f in unexplained)
        + "\n\nUse the interface that owns the member, or add the pair to KNOWN "
        "with a note saying why it is right anyway."
    )


def test_the_audit_resolves_a_useful_number_of_receivers(result):
    """A silent pass from broken inference would be worse than a failure."""
    _findings, resolved, _typelib = result
    assert len(resolved) > 500, (
        f"only {len(resolved)} member accesses resolved to a known interface. "
        f"The inference has probably broken, so a clean run means nothing."
    )


def test_the_audit_can_see_the_bugs_it_was_written_for(audit):
    """Both historical bugs must still be detectable."""
    import json

    typelib = audit.TypeLib(json.loads(DUMP.read_text(encoding="utf-8")))

    # Where they really live.
    assert typelib.has("Constructions", "RevolvedSurfaces")
    assert typelib.has("Line2d", "GetStartPoint")
    # Where they were wrongly read.
    assert not typelib.has("Model", "RevolvedSurfaces")
    assert not typelib.has("Line2d", "StartPoint")


def test_type_inference_follows_a_chain(audit):
    """doc.Models.Item(1).Features must resolve to the Features collection."""
    import ast
    import json

    typelib = audit.TypeLib(json.loads(DUMP.read_text(encoding="utf-8")))
    source = (
        "def f(self):\n"
        "    doc = self.doc_manager.get_active_document()\n"
        "    model = doc.Models.Item(1)\n"
        "    return model.Features\n"
    )
    visitor = audit.Visitor(pathlib.Path(REPO_ROOT / "x.py"), typelib)
    visitor.visit(ast.parse(source))

    assert not visitor.findings
    resolved = {(iface, member) for _loc, iface, member in visitor.resolved}
    assert ("Model", "Features") in resolved


def test_an_invented_member_is_caught(audit):
    """The check must fail on a member no interface has."""
    import ast
    import json

    typelib = audit.TypeLib(json.loads(DUMP.read_text(encoding="utf-8")))
    source = (
        "def f(self):\n"
        "    doc = self.doc_manager.get_active_document()\n"
        "    return doc.Models.Item(1).NotARealMember\n"
    )
    visitor = audit.Visitor(pathlib.Path(REPO_ROOT / "x.py"), typelib)
    visitor.visit(ast.parse(source))

    assert [f[2] for f in visitor.findings] == ["NotARealMember"]
