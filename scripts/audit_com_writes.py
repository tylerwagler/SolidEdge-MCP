"""Find COM property writes Solid Edge cannot accept.

Assigning to a member that is a method, or to a property the type library
marks read-only, does not fail quietly in a useful way: Solid Edge answers
"Property 'Item.X' can not be set." and the surrounding try/except turns that
into a cheerful ``{"status": "set"}``. Three shipped that way --
``Body.FaceStyle.Opacity``, ``Face.Color`` and ``DrawingView.ViewOrientation``,
the last being a method with seven out-parameters.

The receiver inference is the one in ``audit_com_receivers.py``: it follows
declared types, so ``doc.Models.Item(1).Body`` resolves to Body and the write
is checked against that interface rather than against every interface with a
member of that name.
"""

from __future__ import annotations

import ast
import importlib.util
import json
import pathlib
import sys
from types import ModuleType

ROOT = pathlib.Path(__file__).resolve().parent.parent
BACKENDS = ROOT / "src" / "solidedge_mcp" / "backends"
RECEIVERS = ROOT / "scripts" / "audit_com_receivers.py"
DUMP = ROOT / "reference" / "typelib_dump.json"

#: (interface, member) pairs written on purpose despite what the dump says.
#: Keep this short and say why.
ALLOWED: frozenset[tuple[str, str]] = frozenset()


def load_receivers() -> ModuleType:
    spec = importlib.util.spec_from_file_location("_audit_com_receivers", RECEIVERS)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def build_visitor(base: type, typelib: object) -> type:
    """A Visitor that records assignments to COM members."""

    class WriteVisitor(base):  # type: ignore[misc, valid-type]
        def __init__(self, path: pathlib.Path, tl: object) -> None:
            super().__init__(path, tl)
            self.writes: list[tuple[str, str, str, str]] = []
            #: Writes whose receiver resolved, refused or not. A run that
            #: resolves almost nothing is a silent pass, not a clean bill.
            self.resolved_writes = 0

        def visit_Assign(self, node: ast.Assign) -> None:  # noqa: N802
            super().visit_Assign(node)
            for target in node.targets:
                if isinstance(target, ast.Attribute):
                    self._check_write(target)

        def visit_AugAssign(self, node: ast.AugAssign) -> None:  # noqa: N802
            self.generic_visit(node)
            if isinstance(node.target, ast.Attribute):
                self._check_write(node.target)

        def _check_write(self, target: ast.Attribute) -> None:
            member = target.attr
            if not member[:1].isupper():
                return
            receivers = [iface for iface in self._infer(target.value) if self.typelib.known(iface)]
            if not receivers:
                return
            for iface in receivers:
                sig = lookup(self.typelib, iface, member)
                if sig is None:
                    # A member that is on no interface is the name check's job.
                    continue
                self.resolved_writes += 1
                verdict = writability(sig)
                if verdict is None:
                    continue
                self.writes.append((self._loc(target.lineno), iface, member, verdict))

    return WriteVisitor


def lookup(typelib: object, iface: str, member: str) -> dict | None:
    """The declaration of ``member`` on ``iface`` or one of its bases."""
    seen: set[str] = set()
    current: str | None = iface
    while current and current not in seen:
        seen.add(current)
        sig = typelib.members.get(current, {}).get(member)  # type: ignore[attr-defined]
        if sig is not None:
            return sig
        current = typelib.bases.get(current)  # type: ignore[attr-defined]
    return None


def writability(sig: dict) -> str | None:
    """Why this member cannot be assigned to, or None when it can."""
    if "returns" in sig and "access" not in sig:
        params = sig.get("params") or []
        return f"is a method taking {len(params)} argument(s), not a settable property"
    access = sig.get("access")
    if access and "put" not in access:
        return f"is read-only (access: {access})"
    return None


def audit() -> tuple[list[tuple[str, str, str, str]], int]:
    receivers = load_receivers()
    typelib = receivers.TypeLib(json.loads(DUMP.read_text(encoding="utf-8")))
    visitor_cls = build_visitor(receivers.Visitor, typelib)

    findings: list[tuple[str, str, str, str]] = []
    checked = 0
    for path in sorted(BACKENDS.rglob("*.py")):
        visitor = visitor_cls(path, typelib)
        visitor.visit(ast.parse(path.read_text(encoding="utf-8")))
        checked += visitor.resolved_writes
        for loc, iface, member, verdict in visitor.writes:
            if (iface, member) in ALLOWED:
                continue
            findings.append((loc, iface, member, verdict))
    return findings, checked


def main() -> int:
    if not DUMP.exists():
        print("reference/typelib_dump.json is absent; nothing to check.")
        return 0
    findings, checked = audit()
    print(f"Resolved COM property writes checked: {checked}")
    if not findings:
        print("No write to a method or a read-only property.")
        return 0
    print()
    print(f"{len(findings)} write(s) Solid Edge will refuse:")
    print()
    for loc, iface, member, verdict in findings:
        print(f"  {loc}  {iface}.{member} {verdict}")
    print()
    print("Use the setter the type library declares, or drop the write.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
