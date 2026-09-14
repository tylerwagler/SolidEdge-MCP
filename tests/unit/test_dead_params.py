"""Fail when a parameter is declared and never read.

The server's quietest class of bug: a call that succeeds, reports what the
caller asked for, and never passed the value to Solid Edge at all. Three were
found this way and all three had been shipping:

    add_adjustable_part(x, y, z)   accepted a position and placed the part at
                                   the origin
    create_parts_list(x, y)        accepted a table position, let Solid Edge
                                   choose one, and echoed the requested one
    create_revolve(axis_type)      offered a choice no revolve consulted

Nothing downstream can catch these. Solid Edge never hears about the value, so
it has nothing to reject, and a mocked test asserting "returns status" passes.

``scripts/audit_dead_params.py`` checks both layers and this pins it at zero.
A parameter with genuinely nowhere to go says so with ``del param``.
"""

from __future__ import annotations

import ast
import importlib.util
import pathlib
import sys
from types import ModuleType

import pytest

REPO_ROOT = pathlib.Path(__file__).resolve().parents[2]
AUDIT_SCRIPT = REPO_ROOT / "scripts" / "audit_dead_params.py"


def _load() -> ModuleType:
    spec = importlib.util.spec_from_file_location("_audit_dead_params", AUDIT_SCRIPT)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


@pytest.fixture(scope="module")
def audit() -> ModuleType:
    return _load()


@pytest.fixture(scope="module")
def result(audit) -> tuple[list, list, list]:
    return audit.audit()


def test_no_parameter_is_declared_and_never_read(result):
    findings, _resolved, _unresolved = result
    assert not findings, (
        "parameters accepted and never used:\n  "
        + "\n  ".join(str(f) for f in findings)
        + "\n\nUse the parameter, drop it from the signature, or -- if there is "
        "genuinely no COM call to pass it to -- write `del <param>` so the "
        "next reader can see it was a decision."
    )


def test_every_registered_tool_was_found(result):
    """A name the audit cannot resolve is a tool it silently skipped."""
    _findings, resolved, unresolved = result
    assert not unresolved, f"registered but not found: {unresolved}"
    assert len(resolved) > 100, (
        f"only {len(resolved)} tools resolved; the surface is ~118, so the "
        f"audit has probably stopped seeing most of it and a clean run means "
        f"nothing."
    )


def test_a_dead_parameter_is_caught(audit):
    source = "def t(a: int = 0, unused: float = 0.0):\n    return manager.go(a=a)\n"
    fn = ast.parse(source).body[0]
    assert audit.dead_params(fn, "t") == ["unused"]


def test_del_marks_a_parameter_deliberate(audit):
    """The convention the unsupported stubs already use must count as read."""
    source = "def t(a: int = 0, unused: float = 0.0):\n    del unused\n    return {'a': a}\n"
    fn = ast.parse(source).body[0]
    assert audit.dead_params(fn, "t") == []


def test_a_parameter_only_assigned_to_is_still_dead(audit):
    """Writing over a parameter is not using it."""
    source = "def t(a: int = 0, unused: float = 0.0):\n    unused = 1\n    return {'a': a}\n"
    fn = ast.parse(source).body[0]
    assert audit.dead_params(fn, "unused_is_dead") == ["unused"]


def test_a_forwarded_parameter_counts_as_read(audit):
    source = "def t(distance: float = 0.0):\n    return manager.extrude(distance=distance)\n"
    fn = ast.parse(source).body[0]
    assert audit.dead_params(fn, "t") == []
