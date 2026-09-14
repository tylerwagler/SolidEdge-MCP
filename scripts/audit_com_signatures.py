#!/usr/bin/env python
"""Check COM call arity in the backends against the Solid Edge type libraries.

pywin32 returns ``[out]`` values rather than taking them, so a positional call
is well-formed when its argument count lies between the number of required
``[in]`` parameters and the total number of ``[in]`` parameters.

``[in,out]`` parameters count as optional, not required. Verified against Solid
Edge 2026: ``Body.GetRange``, whose two parameters are both ``[in,out]``
SAFEARRAYs, works both with no arguments and with two plain lists. Treating them
as required produced a false positive here, and "fixing" it broke a working
call.

Two rules for supplying such buffers, both verified live. Pass them as **plain
Python lists**; a ``VARIANT`` wrapper is rejected with "Objects for SAFEARRAYS
must be sequences". And when an out-parameter sits *between* parameters you must
supply, pass the later ones by keyword, using the names this script prints under
``--filter`` — pywin32 gives every parameter a positional slot, ``[out]`` ones
included, so filling only the ``[in]`` slots positionally misaligns the call.

Precision comes from resolving the receiver, using the same type inference as
``audit_com_receivers.py``: declared types are followed through a chain, so
``doc.Occurrences.Item(1)`` resolves to Occurrence and the call is checked
against that interface alone.

That matters, because the fallback is weak on purpose. When the receiver cannot
be resolved, a call is accepted if *any* interface defines a method of that name
with a matching arity -- and that let ``occurrence.Replace(path)`` pass for as
long as some other ``Replace`` somewhere took one argument, while
``Occurrence.Replace`` requires two. Every receiver the inference can resolve is
one fewer call relying on that coincidence.

Needs ``reference/typelib_dump.json``; regenerate it with
``uv run python scripts/scrape_typelibs.py``.

Usage:
    uv run python scripts/audit_com_signatures.py
    uv run python scripts/audit_com_signatures.py --by-file
    uv run python scripts/audit_com_signatures.py --filter backends/features/_holes.py
"""

from __future__ import annotations

import argparse
import ast
import importlib.util
import json
import pathlib
import re
import sys
from collections import defaultdict
from collections.abc import Callable
from types import ModuleType

ROOT = pathlib.Path(__file__).resolve().parent.parent
DUMP = ROOT / "reference" / "typelib_dump.json"
BACKENDS = ROOT / "src" / "solidedge_mcp" / "backends"
RECEIVERS = ROOT / "scripts" / "audit_com_receivers.py"
PASCAL = re.compile(r"^[A-Z][A-Za-z0-9]*$")


def load_receivers() -> ModuleType:
    """The receiver audit, imported for its type inference."""
    spec = importlib.util.spec_from_file_location("_audit_com_receivers", RECEIVERS)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


Finding = tuple[str, str, str, int, set[tuple[int, int]], bool]


def load_typelibs() -> tuple[dict[str, dict[str, tuple[int, int]]], dict[str, set], dict]:
    data = json.loads(DUMP.read_text(encoding="utf-8"))
    iface_methods: dict[str, dict[str, tuple[int, int]]] = defaultdict(dict)
    any_windows: dict[str, set[tuple[int, int]]] = defaultdict(set)
    for payload in data["typelibs"].values():
        for iface_name, iface in (payload.get("interfaces") or {}).items():
            for method, sig in (iface.get("methods") or {}).items():
                params = sig.get("params", [])
                flags = [set((p.get("flags") or "").split(",")) for p in params]
                # Anything the caller may supply, including [in,out] buffers.
                ins = [f for f in flags if "in" in f]
                # Required: pure [in], neither optional nor an out-buffer that
                # pywin32 is willing to allocate itself.
                required = [f for f in ins if not ({"optional", "out"} & f)]
                window = (len(required), len(ins))
                iface_methods[iface_name][method] = window
                any_windows[method].add(window)
    return dict(iface_methods), dict(any_windows), data


class CallVisitor(ast.NodeVisitor):
    """Collect COM-shaped calls whose argument count cannot be valid."""

    def __init__(
        self,
        path: pathlib.Path,
        iface_methods: dict[str, dict[str, tuple[int, int]]],
        any_windows: dict[str, set[tuple[int, int]]],
    ) -> None:
        self.path = path
        self.iface_methods = iface_methods
        self.any_windows = any_windows
        self.known_ifaces = set(iface_methods)
        self.scopes: list[dict[str, str]] = [{}]
        self.findings: list[Finding] = []
        #: Set to audit_com_receivers' inference when the dump is available.
        self.infer: Callable[[ast.expr], frozenset[str]] | None = None

    def visit_FunctionDef(self, node: ast.FunctionDef) -> None:  # noqa: N802
        self.scopes.append({})
        self.generic_visit(node)
        self.scopes.pop()

    visit_AsyncFunctionDef = visit_FunctionDef  # type: ignore[assignment]

    def _lookup(self, name: str) -> str | None:
        for scope in reversed(self.scopes):
            if name in scope:
                return scope[name]
        return None

    def visit_Assign(self, node: ast.Assign) -> None:  # noqa: N802
        if (
            len(node.targets) == 1
            and isinstance(node.targets[0], ast.Name)
            and isinstance(node.value, ast.Attribute)
            and node.value.attr in self.known_ifaces
        ):
            self.scopes[-1][node.targets[0].id] = node.value.attr
        self.generic_visit(node)

    def visit_Call(self, node: ast.Call) -> None:  # noqa: N802
        func = node.func
        if isinstance(func, ast.Attribute) and PASCAL.match(func.attr):
            starred = any(isinstance(a, ast.Starred) for a in node.args) or any(
                k.arg is None for k in node.keywords
            )
            if not starred:
                self._check(node, func)
        self.generic_visit(node)

    def _resolve_iface(self, receiver: ast.expr) -> list[str]:
        """Every interface the receiver might be, best evidence first."""
        # The full inference, when this visitor was given one.
        if self.infer is not None:
            inferred = [i for i in self.infer(receiver) if i in self.known_ifaces]
            if inferred:
                return sorted(inferred)
        if isinstance(receiver, ast.Attribute) and receiver.attr in self.known_ifaces:
            return [receiver.attr]
        if isinstance(receiver, ast.Name):
            tracked = self._lookup(receiver.id)
            if tracked:
                return [tracked]
        return []

    def _check(self, node: ast.Call, func: ast.Attribute) -> None:
        method = func.attr
        argc = len(node.args) + len(node.keywords)
        ifaces = [
            i for i in self._resolve_iface(func.value) if method in self.iface_methods.get(i, {})
        ]
        loc = f"{self.path.relative_to(ROOT).as_posix()}:{node.lineno}"

        if ifaces:
            # A receiver may be one of several interfaces -- a document is any
            # of five -- so the call is well-formed if it fits any of them.
            windows = {self.iface_methods[i][method] for i in ifaces}
            if not any(low <= argc <= high for low, high in windows):
                self.findings.append((loc, ",".join(ifaces), method, argc, windows, True))
            return

        windows = self.any_windows.get(method)
        if not windows:
            return  # unknown member name; the member-name audit covers that
        if not any(low <= argc <= high for low, high in windows):
            self.findings.append((loc, "?", method, argc, windows, False))


def build_inferring_visitor(receivers: ModuleType, typelib: object, iface_methods, any_windows):
    """A receiver-audit visitor that also checks arity as it walks.

    The inference needs live scopes -- it learns what ``occurrence`` is from
    the assignment above the call -- so the arity check has to run inside the
    same traversal rather than over a second pass.
    """

    class InferringCallVisitor(receivers.Visitor):  # type: ignore[misc, valid-type]
        def __init__(self, path: pathlib.Path, tl: object) -> None:
            super().__init__(path, tl)
            self.checker = CallVisitor(path, iface_methods, any_windows)
            self.checker.infer = self._infer

        def visit_Call(self, node: ast.Call) -> None:  # noqa: N802
            super().visit_Call(node)
            func = node.func
            if isinstance(func, ast.Attribute) and PASCAL.match(func.attr):
                starred = any(isinstance(a, ast.Starred) for a in node.args) or any(
                    k.arg is None for k in node.keywords
                )
                if not starred:
                    self.checker._check(node, func)

    return InferringCallVisitor


def collect() -> tuple[list[Finding], list[Finding], dict]:
    iface_methods, any_windows, data = load_typelibs()
    receivers = load_receivers()
    typelib = receivers.TypeLib(data)
    visitor_cls = build_inferring_visitor(receivers, typelib, iface_methods, any_windows)

    precise: list[Finding] = []
    loose: list[Finding] = []
    for path in sorted(BACKENDS.rglob("*.py")):
        visitor = visitor_cls(path, typelib)
        visitor.visit(ast.parse(path.read_text(encoding="utf-8")))
        for finding in visitor.checker.findings:
            (precise if finding[5] else loose).append(finding)
    return precise, loose, data


def fmt(finding: Finding) -> str:
    loc, iface, method, argc, windows, _ = finding
    allowed = ", ".join(f"{low}..{high}" for low, high in sorted(windows))
    return f"{loc}  {iface}.{method}({argc} args) allowed {allowed}"


def print_params(data: dict, iface: str, method: str) -> None:
    for payload in data["typelibs"].values():
        block = (payload.get("interfaces") or {}).get(iface) or {}
        sig = (block.get("methods") or {}).get(method)
        if sig:
            for prm in sig.get("params", []):
                name = str(prm.get("name"))
                type_ = str(prm.get("type"))
                print(f"        {name:30s} {type_:28s} {prm.get('flags')}")
            return


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--by-file", action="store_true", help="summarise counts per file")
    parser.add_argument("--filter", help="only show findings whose path contains this")
    args = parser.parse_args()

    if not DUMP.exists():
        print(f"missing {DUMP.relative_to(ROOT)}; run scripts/scrape_typelibs.py first")
        return 2

    precise, loose, data = collect()

    if args.filter:
        needle = args.filter.replace("\\", "/")
        shown = [f for f in precise + loose if needle in f[0]]
        print(f"=== {len(shown)} finding(s) matching {needle!r} ===")
        for finding in shown:
            print("  " + fmt(finding))
            print_params(data, finding[1], finding[2])
        return 0

    print(f"PRECISE (interface resolved): {len(precise)}")
    for finding in precise:
        print("  " + fmt(finding))
    print(f"\nLOOSE (no interface overload matched): {len(loose)}")
    for finding in loose:
        print("  " + fmt(finding))

    if args.by_file:
        counts: dict[str, int] = defaultdict(int)
        for finding in precise + loose:
            counts[finding[0].split(":")[0]] += 1
        print("\nBY FILE")
        for name, count in sorted(counts.items(), key=lambda kv: -kv[1]):
            print(f"  {count:3d}  {name}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
