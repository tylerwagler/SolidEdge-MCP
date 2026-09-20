#!/usr/bin/env python
"""Check that tool-layer calls land in the backend parameters they mean to.

The tool layer calls backend managers positionally. When a backend signature
gains or reorders a parameter, the positional arguments silently shift and the
call still type-checks, still passes its mocked test, and reaches Solid Edge
with the wrong values.

That happened: ``create_extrude(distance, direction)`` was calling
``create_extrude(distance, operation="Add", direction="Normal")``, so the
direction landed in ``operation`` and every extrude failed with "Unknown
operation: 'Normal'". Nothing caught it, because the integration tests call the
backend directly and the unit tests assert the same wrong call.

This script maps every positional argument at a tool call site onto the backend
parameter it actually fills, and reports when a bare variable name does not
match that parameter's name. A rename is a false positive; a mismatch like
``direction`` landing in ``operation`` is a bug.

Usage:
    uv run python scripts/audit_tool_dispatch.py
    uv run python scripts/audit_tool_dispatch.py --filter _extrude
"""

from __future__ import annotations

import argparse
import ast
import inspect
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
TOOLS = ROOT / "src" / "solidedge_mcp" / "tools"

sys.path.insert(0, str(ROOT / "src"))

#: Tool-layer variable names that legitimately fill a differently named
#: backend parameter. Keep this short and justified.
ALLOWED: set[tuple[str, str]] = {
    ("method", "type"),
    ("type", "method"),
    ("action", "method"),
    ("shape", "shape_type"),
}

Finding = tuple[str, str, str, int, str, str]


def manager_types() -> dict[str, object]:
    from solidedge_mcp import managers

    return {
        name: type(getattr(managers, name))
        for name in dir(managers)
        if not name.startswith("_") and hasattr(getattr(managers, name), "__class__")
    }


def backend_params(cls: object, method: str) -> list[str] | None:
    fn = getattr(cls, method, None)
    if fn is None or not callable(fn):
        return None
    try:
        sig = inspect.signature(fn)
    except (TypeError, ValueError):
        return None
    names = [
        p.name
        for p in sig.parameters.values()
        if p.kind in (p.POSITIONAL_ONLY, p.POSITIONAL_OR_KEYWORD) and p.name != "self"
    ]
    return names


def audit() -> list[Finding]:
    managers = manager_types()
    findings: list[Finding] = []

    for path in sorted(TOOLS.rglob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        for node in ast.walk(tree):
            if not isinstance(node, ast.Call):
                continue
            func = node.func
            if not isinstance(func, ast.Attribute) or not isinstance(func.value, ast.Name):
                continue
            manager_name = func.value.id
            cls = managers.get(manager_name)
            if cls is None:
                continue
            params = backend_params(cls, func.attr)
            if params is None:
                continue
            for index, arg in enumerate(node.args):
                if isinstance(arg, ast.Starred) or index >= len(params):
                    continue
                if not isinstance(arg, ast.Name):
                    continue  # a literal or expression carries no name to compare
                param = params[index]
                if arg.id == param or (arg.id, param) in ALLOWED:
                    continue
                findings.append(
                    (
                        f"{path.relative_to(ROOT).as_posix()}:{node.lineno}",
                        manager_name,
                        func.attr,
                        index,
                        arg.id,
                        param,
                    )
                )
    return findings


def fmt(f: Finding) -> str:
    loc, manager, method, index, passed, param = f
    return f"{loc}  {manager}.{method}() arg {index}: passes `{passed}` into parameter `{param}`"


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--filter", help="only show findings whose path contains this")
    args = parser.parse_args()

    findings = audit()
    if args.filter:
        findings = [f for f in findings if args.filter in f[0]]

    print(f"positional argument mismatches: {len(findings)}")
    for f in findings:
        print("  " + fmt(f))
    return 0


if __name__ == "__main__":
    sys.exit(main())
