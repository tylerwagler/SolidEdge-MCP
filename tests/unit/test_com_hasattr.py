"""Fail when ``hasattr`` is used to test what a COM proxy can do.

CLAUDE.md's rule, and it is not theoretical. On a late-bound proxy the probe is
a separate ``GetIDsOfNames`` round trip that reads False both for a member that
is genuinely absent and for one whose getter merely raised, so a real error
becomes a silent "unsupported".

It cost a whole feature. Every layer call gated on ``hasattr(doc, "Layers")``,
and a DraftDocument keeps its layers on the Sheet rather than the document, so
all five answered "Active document does not support layers" for the document
type where layers matter most.

``com_get(obj, "Member")`` returns None for both cases and is one round trip.
Use ``doc.Type`` against DocumentTypeConstants to test a document kind.

This needs no type library, so unlike the COM conformance checks it never
skips.
"""

from __future__ import annotations

import ast
import importlib.util
import pathlib
import sys
from types import ModuleType

import pytest

REPO_ROOT = pathlib.Path(__file__).resolve().parents[2]
AUDIT_SCRIPT = REPO_ROOT / "scripts" / "audit_com_hasattr.py"


def _load() -> ModuleType:
    spec = importlib.util.spec_from_file_location("_audit_com_hasattr", AUDIT_SCRIPT)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


@pytest.fixture(scope="module")
def audit() -> ModuleType:
    return _load()


def test_no_com_member_is_probed_with_hasattr(audit):
    findings = audit.audit()
    assert not findings, (
        "hasattr probes of a COM member:\n  "
        + "\n  ".join(f"{rel}:{line}  {member}" for rel, line, member in findings)
        + '\n\nUse com_get(obj, "Member") and test for None, or check doc.Type.'
    )


def test_a_com_probe_is_caught(audit):
    source = 'def f(doc):\n    if hasattr(doc, "Layers"):\n        return doc.Layers\n'
    assert audit.probes(ast.parse(source)) == [(2, "Layers")]


def test_our_own_attributes_are_not_com(audit):
    """hasattr on a plain Python object stays legitimate."""
    source = (
        "def f(x, conn):\n"
        '    a = hasattr(x, "__iter__")\n'
        '    b = hasattr(conn, "on_disconnect")\n'
        '    c = hasattr(x, "active_profile")\n'
        "    return a, b, c\n"
    )
    assert audit.probes(ast.parse(source)) == []
