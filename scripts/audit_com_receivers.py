#!/usr/bin/env python
"""Check that each COM member is read off an interface that actually has it.

The member-name audit asks "does this name exist anywhere in Solid Edge?", and
a name can pass that while being wrong where it is used. Two real bugs got
through exactly that way:

    model.RevolvedSurfaces      RevolvedSurfaces is real, but it belongs to
                                Constructions, not Model, so every revolved
                                surface call raised.
    line.StartPoint.X           StartPoint is real on other interfaces; Line2d
                                has GetStartPoint(). Reading it inside a bare
                                except made sketch rotate and scale delete the
                                sketch and rebuild nothing.

This audit answers the sharper question by inferring what each receiver *is*.
The type library dump records a declared type for every property and a return
type for every method, so a chain like

    doc.Models.Item(1).Features

resolves PartDocument -> Models -> Model -> Features, and any member read off
one of those can be checked against that interface, its bases included.

Inference is deliberately partial. A receiver whose type cannot be worked out,
or whose declared type is a bare VT_DISPATCH, is skipped rather than guessed
at, so a finding means the member really is absent from a known interface.

Needs ``reference/typelib_dump.json``; regenerate it with
``uv run python scripts/scrape_typelibs.py``.

Usage:
    uv run python scripts/audit_com_receivers.py
    uv run python scripts/audit_com_receivers.py --filter backends/features
    uv run python scripts/audit_com_receivers.py --show-resolved
"""

from __future__ import annotations

import argparse
import ast
import json
import pathlib
import re
import sys
from collections import defaultdict

ROOT = pathlib.Path(__file__).resolve().parent.parent
DUMP = ROOT / "reference" / "typelib_dump.json"
BACKENDS = ROOT / "src" / "solidedge_mcp" / "backends"
PASCAL = re.compile(r"^[A-Z][A-Za-z0-9_]*$")

#: A document could be any of these, so a member is only wrong when it is on
#: none of them.
DOCUMENT_TYPES = frozenset(
    {
        "PartDocument",
        "SheetMetalDocument",
        "AssemblyDocument",
        "DraftDocument",
        "WeldmentDocument",
    }
)

#: Expressions that seed the inference, matched on their source text.
SEEDS: dict[str, frozenset[str]] = {
    "self.doc_manager.get_active_document()": DOCUMENT_TYPES,
    "self.get_active_document()": DOCUMENT_TYPES,
    "self.doc_manager.connection.get_application()": frozenset({"Application"}),
    "self.connection.get_application()": frozenset({"Application"}),
    "self.get_application()": frozenset({"Application"}),
    "self.application": frozenset({"Application"}),
    "self.active_document": DOCUMENT_TYPES,
}

#: Our own helpers that hand back a COM object, keyed by function name so the
#: arguments do not matter. Without these the inference stops at the helper
#: and every member read off its result goes unchecked -- which is most of the
#: query layer, since it reaches geometry through body_of and all_faces.
CALL_SEEDS: dict[str, frozenset[str]] = {
    "body_of": frozenset({"Body"}),
    "all_faces": frozenset({"Faces"}),
}

#: Helpers returning ``(com_object, error_dict)``. The first name bound by a
#: tuple assignment takes the interface; the second is ours, not Solid Edge's.
TUPLE_SEEDS: dict[str, frozenset[str]] = {
    "resolve_view": frozenset({"View"}),
}

#: Members whose name is an interface but whose real type is a different one.
#: The VT_DISPATCH fallback in ``member_type`` guesses from the member name,
#: which is right almost everywhere and wrong here: ``Document.Properties``
#: hands back a PropertySets, whose Item() is what yields a Properties.
DISPATCH_OVERRIDES: dict[str, str] = {
    "Properties": "PropertySets",
}

#: Members every COM object answers, whatever its interface.
UNIVERSAL = frozenset({"Application", "Parent", "Item", "Count", "Type", "Name", "Index"})

Finding = tuple[str, str, str, list[str]]


class TypeLib:
    """Member lookup across every interface, following inheritance."""

    def __init__(self, data: dict) -> None:
        self.members: dict[str, dict[str, dict]] = {}
        self.bases: dict[str, str] = {}
        for payload in data["typelibs"].values():
            for name, iface in (payload.get("interfaces") or {}).items():
                merged = dict(iface.get("properties") or {})
                merged.update(iface.get("methods") or {})
                # A name can appear in several libraries; keep the richest.
                if len(merged) >= len(self.members.get(name, {})):
                    self.members[name] = merged
                    base = iface.get("inherits")
                    if isinstance(base, str) and base:
                        self.bases[name] = base

    def has(self, iface: str, member: str) -> bool:
        seen: set[str] = set()
        current: str | None = iface
        while current and current not in seen:
            seen.add(current)
            if member in self.members.get(current, {}):
                return True
            current = self.bases.get(current)
        return False

    def known(self, iface: str) -> bool:
        return iface in self.members

    def member_type(self, iface: str, member: str) -> str | None:
        """The interface a member yields, or None when it is not an interface.

        Plenty of collection properties are declared ``VT_DISPATCH`` rather
        than with their real type -- ``Sheet.TextBoxes`` and ``Body.Faces``
        both are -- and inference used to stop dead there, leaving everything
        downstream unchecked. Two bugs lived in exactly that blind spot:
        ``textbox.TextHeight`` and ``textbox.x``, neither of which TextBox has.

        When the declared type is useless but the member's own name is an
        interface in the dump, that name is used. Solid Edge is consistent
        about this, and a wrong guess shows up immediately as a finding on a
        member that does exist.
        """
        seen: set[str] = set()
        current: str | None = iface
        while current and current not in seen:
            seen.add(current)
            sig = self.members.get(current, {}).get(member)
            if sig is not None:
                raw = sig.get("type") or sig.get("returns") or ""
                resolved = _interface_name(raw, self)
                if resolved:
                    return resolved
                if raw.rstrip("*").strip() == "VT_DISPATCH":
                    guess = DISPATCH_OVERRIDES.get(member, member)
                    if guess in self.members:
                        return guess
                return None
            current = self.bases.get(current)
        return None

    def owners(self, member: str) -> list[str]:
        return sorted(name for name, block in self.members.items() if member in block)


def _interface_name(raw: str, typelib: TypeLib) -> str | None:
    """Turn a declared type such as ``Models*`` into an interface name."""
    if not raw:
        return None
    name = raw.rstrip("*").strip()
    if not name or name.startswith("VT_") or name.startswith("SAFEARRAY"):
        return None
    return name if name in typelib.members else None


def _source(node: ast.expr) -> str:
    try:
        return ast.unparse(node)
    except Exception:
        return ""


class Visitor(ast.NodeVisitor):
    """Infer receiver interfaces and check every member read against them."""

    def __init__(self, path: pathlib.Path, typelib: TypeLib) -> None:
        self.path = path
        self.typelib = typelib
        self.scopes: list[dict[str, frozenset[str]]] = [{}]
        self.findings: list[Finding] = []
        self.resolved: list[tuple[str, str, str]] = []

    # -- scope handling ------------------------------------------------
    def visit_FunctionDef(self, node: ast.FunctionDef) -> None:  # noqa: N802
        self.scopes.append({})
        self.generic_visit(node)
        self.scopes.pop()

    visit_AsyncFunctionDef = visit_FunctionDef  # type: ignore[assignment]

    def _lookup(self, name: str) -> frozenset[str]:
        for scope in reversed(self.scopes):
            if name in scope:
                return scope[name]
        return frozenset()

    def visit_Assign(self, node: ast.Assign) -> None:  # noqa: N802
        self.generic_visit(node)

        # view_obj, err = resolve_view(doc)
        value = node.value
        if (
            isinstance(value, ast.Call)
            and isinstance(value.func, ast.Name)
            and value.func.id in TUPLE_SEEDS
        ):
            for target in node.targets:
                if isinstance(target, ast.Tuple) and target.elts:
                    first = target.elts[0]
                    if isinstance(first, ast.Name):
                        self.scopes[-1][first.id] = TUPLE_SEEDS[value.func.id]
            return

        types = self._infer(node.value)
        if not types:
            return
        for target in node.targets:
            if isinstance(target, ast.Name):
                self.scopes[-1][target.id] = types

    # -- inference -----------------------------------------------------
    def _infer(self, node: ast.expr) -> frozenset[str]:
        seeded = SEEDS.get(_source(node))
        if seeded:
            return seeded

        if isinstance(node, ast.Name):
            return self._lookup(node.id)

        if isinstance(node, ast.Call):
            func = node.func
            if isinstance(func, ast.Attribute):
                receivers = self._infer(func.value)
                return self._member_types(receivers, func.attr)
            if isinstance(func, ast.Name):
                return CALL_SEEDS.get(func.id, frozenset())
            return frozenset()

        if isinstance(node, ast.Attribute):
            receivers = self._infer(node.value)
            return self._member_types(receivers, node.attr)

        if isinstance(node, ast.Subscript):
            # collection[i] tells us nothing about the element's interface.
            return frozenset()

        return frozenset()

    def _member_types(self, receivers: frozenset[str], member: str) -> frozenset[str]:
        out: set[str] = set()
        for iface in receivers:
            yielded = self.typelib.member_type(iface, member)
            if yielded:
                out.add(yielded)
        return frozenset(out)

    # -- checking ------------------------------------------------------
    def visit_Attribute(self, node: ast.Attribute) -> None:  # noqa: N802
        self.generic_visit(node)
        self._check(node.value, node.attr, node.lineno)

    def visit_Call(self, node: ast.Call) -> None:  # noqa: N802
        self.generic_visit(node)
        # com_get(obj, "Member") reads a member by name.
        if (
            isinstance(node.func, ast.Name)
            and node.func.id == "com_get"
            and node.args
            and len(node.args) >= 2
            and isinstance(node.args[1], ast.Constant)
            and isinstance(node.args[1].value, str)
        ):
            self._check(node.args[0], node.args[1].value, node.lineno)

    def _check(self, receiver: ast.expr, member: str, lineno: int) -> None:
        # PascalCase is how a COM member is usually spelled, but not always:
        # 857 lowercase member names exist across the Solid Edge libraries, and
        # requiring PascalCase hid two bugs of exactly the same shape --
        # ``textbox.x`` and ``balloon.x``, neither of which is a member of
        # anything, so the position was quietly missing from every result.
        #
        # A lowercase name is only checked once the receiver has resolved to a
        # known interface, which is what keeps our own Python attributes out of
        # it: ``self.active_profile`` and friends resolve to nothing.
        if member in UNIVERSAL or member.startswith("_"):
            return
        if not PASCAL.match(member) and not member.isidentifier():
            return
        receivers = self._infer(receiver)
        if not receivers:
            return
        known = [iface for iface in receivers if self.typelib.known(iface)]
        if not known:
            return
        if any(self.typelib.has(iface, member) for iface in known):
            self.resolved.append((f"{self._loc(lineno)}", ",".join(sorted(known)), member))
            return
        self.findings.append((self._loc(lineno), ",".join(sorted(known)), member, known))

    def _loc(self, lineno: int) -> str:
        return f"{self.path.relative_to(ROOT).as_posix()}:{lineno}"


def audit() -> tuple[list[Finding], list[tuple[str, str, str]], TypeLib]:
    typelib = TypeLib(json.loads(DUMP.read_text(encoding="utf-8")))
    findings: list[Finding] = []
    resolved: list[tuple[str, str, str]] = []
    for path in sorted(BACKENDS.rglob("*.py")):
        visitor = Visitor(path, typelib)
        visitor.visit(ast.parse(path.read_text(encoding="utf-8")))
        findings.extend(visitor.findings)
        resolved.extend(visitor.resolved)
    return findings, resolved, typelib


def fmt(finding: Finding, typelib: TypeLib | None = None) -> str:
    loc, ifaces, member, _known = finding
    line = f"{loc}  {ifaces}.{member} does not exist"
    if typelib is not None:
        owners = typelib.owners(member)
        if owners:
            line += f"; {member} belongs to {', '.join(owners[:4])}"
        else:
            line += f"; no Solid Edge interface has {member}"
    return line


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--filter", help="only show findings whose path contains this")
    parser.add_argument(
        "--show-resolved", action="store_true", help="list the accesses that checked out"
    )
    args = parser.parse_args()

    if not DUMP.exists():
        print(f"missing {DUMP.relative_to(ROOT)}; run scripts/scrape_typelibs.py first")
        return 2

    findings, resolved, typelib = audit()

    if args.filter:
        needle = args.filter.replace("\\", "/")
        findings = [f for f in findings if needle in f[0]]
        resolved = [r for r in resolved if needle in r[0]]

    print(f"members read off an interface that does not have them: {len(findings)}")
    for finding in findings:
        print("  " + fmt(finding, typelib))

    print(f"\naccesses checked against a resolved interface: {len(resolved)}")
    if args.show_resolved:
        for loc, iface, member in resolved:
            print(f"  {loc}  {iface}.{member}")
    else:
        by_iface: dict[str, int] = defaultdict(int)
        for _loc, iface, _member in resolved:
            by_iface[iface] += 1
        for iface, count in sorted(by_iface.items(), key=lambda kv: -kv[1])[:12]:
            print(f"  {count:4d}  {iface}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
