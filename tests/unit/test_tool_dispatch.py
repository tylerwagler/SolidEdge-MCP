"""Fail when a tool-layer call lands in the wrong backend parameter.

The tool layer now calls backend managers by keyword, so this should stay at
zero. It did not always: a positional call shifts silently when a backend
signature gains or reorders a parameter, the call still type-checks, the mocked
test still passes, and Solid Edge receives the wrong values.

That is not hypothetical. ``create_extrude(distance, direction)`` was calling
``create_extrude(distance, operation="Add", direction="Normal")``, so the
direction landed in ``operation`` and *every* extrude failed with "Unknown
operation: 'Normal'". The integration tests missed it because they call the
backend directly, and the unit test missed it because it asserted the same
wrong call. It was found by driving real Solid Edge and seeing an empty part.

``scripts/audit_tool_dispatch.py`` runs the same check by hand.
"""

from __future__ import annotations

import importlib.util
import pathlib
import sys
from types import ModuleType

import pytest

REPO_ROOT = pathlib.Path(__file__).resolve().parents[2]
AUDIT_SCRIPT = REPO_ROOT / "scripts" / "audit_tool_dispatch.py"

#: Kept empty on purpose. The tool layer calls backends by keyword, so there is
#: no positional argument left to land in the wrong parameter. If this list ever
#: needs an entry again, something has gone back to positional calls.
KNOWN_RENAMES: frozenset[tuple[str, str, int]] = frozenset()
KNOWN_RENAME_PAIRS: frozenset[tuple[str, str]] = frozenset()

pytestmark = pytest.mark.skipif(
    not AUDIT_SCRIPT.exists(), reason="scripts/audit_tool_dispatch.py missing"
)


def _load() -> ModuleType:
    spec = importlib.util.spec_from_file_location("_audit_tool_dispatch", AUDIT_SCRIPT)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


@pytest.fixture(scope="module")
def audit() -> ModuleType:
    return _load()


@pytest.fixture(scope="module")
def findings(audit) -> list:
    return audit.audit()


def test_the_audit_can_see_the_backends(audit):
    """Guard against a silent pass from an import failure or empty scan."""
    managers = audit.manager_types()
    assert "feature_manager" in managers
    params = audit.backend_params(managers["feature_manager"], "create_extrude")
    assert params[:3] == ["distance", "operation", "direction"]


def test_no_argument_lands_in_the_wrong_parameter(findings, audit):
    unexplained = [
        f
        for f in findings
        if (f[2], f[4], f[3]) not in KNOWN_RENAMES and (f[4], f[5]) not in KNOWN_RENAME_PAIRS
    ]
    assert not unexplained, (
        "tool-layer arguments landing in the wrong backend parameter:\n  "
        + "\n  ".join(audit.fmt(f) for f in unexplained)
        + "\n\nPass the argument by keyword. If the names merely differ, add it "
        "to KNOWN_RENAMES or KNOWN_RENAME_PAIRS with a note."
    )


def test_extrude_passes_operation_and_direction_by_keyword():
    """Regression: the call whose misrouting produced empty parts."""
    source = (REPO_ROOT / "src" / "solidedge_mcp" / "tools" / "features" / "_extrude.py").read_text(
        encoding="utf-8"
    )
    assert "operation=operation" in source
    assert "direction=direction" in source
