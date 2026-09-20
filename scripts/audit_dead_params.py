"""Find parameters that no code path ever reads.

A function that advertises a parameter and never looks at it is the quietest
kind of bug on this server: the call succeeds, the result says "created", and
the value the caller cared about was dropped on the floor. Solid Edge cannot
report the mistake because it was never told about it. add_adjustable_part
took x/y/z and placed the part at the origin; create_parts_list took a table
position and let Solid Edge choose one; create_revolve offered an axis_type
that no revolve has ever consulted.

Both layers are checked: every function registered with register_tool, and
every public method on a backend manager mixin. A parameter that genuinely
has nowhere to go -- the stubs that exist only so tool dispatch keeps working
-- says so with ``del param``, which this reads as deliberate.
"""

from __future__ import annotations

import ast
import pathlib
import sys
from dataclasses import dataclass

SRC = pathlib.Path(__file__).resolve().parent.parent / "src" / "solidedge_mcp"
TOOLS = SRC / "tools"
BACKENDS = SRC / "backends"
REPO = SRC.parent.parent

#: (function or Class.method, parameter) pairs that are dead on purpose and
#: cannot say so with ``del``. Keep this short and say why.
ALLOWED: frozenset[tuple[str, str]] = frozenset()


@dataclass(frozen=True)
class Finding:
    file: str
    where: str
    param: str

    def __str__(self) -> str:
        return f"{self.file}: {self.where}({self.param})"


def registered_names(tree: ast.Module) -> set[str]:
    """Names passed to register_tool(mcp, <name>, ...)."""
    names: set[str] = set()
    for node in ast.walk(tree):
        if (
            isinstance(node, ast.Call)
            and isinstance(node.func, ast.Name)
            and node.func.id == "register_tool"
            and len(node.args) >= 2
            and isinstance(node.args[1], ast.Name)
        ):
            names.add(node.args[1].id)
    return names


def params(fn: ast.FunctionDef) -> list[str]:
    a = fn.args
    return [p.arg for p in (*a.posonlyargs, *a.args, *a.kwonlyargs) if p.arg not in ("self", "cls")]


def body_names(fn: ast.FunctionDef) -> set[str]:
    """Every identifier the body reads, plus attribute names.

    Attribute names count because a parameter is often forwarded as
    ``manager.method(distance=distance)``; the keyword itself is not a Name
    node, but the value is.

    ``del param`` counts too. A method that keeps a parameter it cannot use
    says so with ``del``, which is this codebase's existing convention and
    reads as deliberate rather than forgotten. Only a plain assignment
    (``Store``) fails to count, since writing to a name is not reading it.
    """
    used: set[str] = set()
    for node in ast.walk(fn):
        if isinstance(node, ast.Name) and not isinstance(node.ctx, ast.Store):
            used.add(node.id)
        elif isinstance(node, ast.Attribute):
            used.add(node.attr)
    return used


def dead_params(fn: ast.FunctionDef, where: str) -> list[str]:
    used = body_names(fn)
    return [p for p in params(fn) if p not in used and (where, p) not in ALLOWED]


def audit() -> tuple[list[Finding], list[str], list[str]]:
    """Return (findings, tools checked, registered names not found)."""
    # tools/features/__init__.py registers 50 functions that live in its
    # sibling modules, so names and definitions must be collected across the
    # whole package before they can be matched up.
    defined: dict[str, tuple[pathlib.Path, ast.FunctionDef]] = {}
    wanted: set[str] = set()
    for path in sorted(TOOLS.rglob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        wanted |= registered_names(tree)
        for node in tree.body:
            if isinstance(node, ast.FunctionDef):
                defined.setdefault(node.name, (path, node))

    findings: list[Finding] = []
    resolved = sorted(wanted & set(defined))
    for tool in resolved:
        path, node = defined[tool]
        rel = path.relative_to(REPO).as_posix()
        findings += [Finding(rel, tool, p) for p in dead_params(node, tool)]

    for path in sorted(BACKENDS.rglob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        rel = path.relative_to(REPO).as_posix()
        for cls in [n for n in ast.walk(tree) if isinstance(n, ast.ClassDef)]:
            for fn in [n for n in cls.body if isinstance(n, ast.FunctionDef)]:
                if fn.name.startswith("__"):
                    continue
                label = f"{cls.name}.{fn.name}"
                dead = dead_params(fn, label) or dead_params(fn, fn.name)
                findings += [Finding(rel, label, p) for p in dead]

    return findings, resolved, sorted(wanted - set(defined))


def main() -> int:
    findings, resolved, unresolved = audit()
    print(f"Checked {len(resolved)} registered tools and every backend method.")
    if unresolved:
        print(f"WARNING: {len(unresolved)} registered name(s) not found: {unresolved}")
    if not findings:
        print("No dead parameters.")
        return 1 if unresolved else 0
    print()
    print(f"{len(findings)} parameter(s) declared and never read:")
    print()
    for f in findings:
        print(f"  {f}")
    print()
    print("Use the parameter, drop it, or mark it deliberate with `del <param>`.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
