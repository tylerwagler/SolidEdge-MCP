"""Find a constant whose value is in no member of the enum it is passed to.

Solid Edge takes a lot of enums and does not police them. Hand it a value from
the wrong one and it usually accepts the call, does nothing useful, and reports
no error, so the feature is dead and nothing says so. Five bugs of that shape
were fixed by hand in this codebase.

**This catches a narrow slice of that, and it is worth being precise about
which.** It compares the value against the enum the type library declares for
that parameter, so it fires only when the value is in *no* member of that
enum. It therefore catches:

    a bare literal nobody checked           AddPartView(Orientation=999)
    a local bound to one                    orient = 999; AddPartView(..., orient)
    a value from the wrong constants class, when that value falls outside the
    target enum -- ExtentTypeConstants.igNone is 44, and ViewOrientationConstants
    stops well short of it

It cannot catch the wrong *member* of the right enum, which is what most of the
five actually were: AddPartView was passed 5, a real igBottomView, so every
"Front" view came out a bottom view; the helix cutout was passed 2, a real
member of FeaturePropertyConstants, just not an axis end. Nothing short of
knowing the intended semantics can see those, and live geometry verification is
what caught them.

Bitmask enums are handled: where every member is a power of two, 0 and any
combination of members are legitimate. Profile.End(0) means "no validation
criteria" and cuts a hole perfectly well -- verified live.

Both call arguments and property writes are checked, for whatever receivers
``audit_com_receivers.py`` can resolve.

Needs ``reference/typelib_dump.json``.
"""

from __future__ import annotations

import ast
import importlib.util
import json
import pathlib
import sys
from dataclasses import dataclass
from types import ModuleType

ROOT = pathlib.Path(__file__).resolve().parent.parent
DUMP = ROOT / "reference" / "typelib_dump.json"
BACKENDS = ROOT / "src" / "solidedge_mcp" / "backends"
CONSTANTS = BACKENDS / "constants.py"
RECEIVERS = ROOT / "scripts" / "audit_com_receivers.py"

#: (Interface.Method, parameter) pairs that pass a value the enum does not
#: list, on purpose. Keep this short and say why.
# (interface.method, parameter) pairs whose value is outside the declared enum on
# purpose, each driven live. Slots.Add: KeyPointExtentConstants has no null member,
# and 0 is what a finite-extent slot takes (Solid Edge 2026: 6 -> 10 faces).
ALLOWED: frozenset[tuple[str, str]] = frozenset(
    {("Slots.Add", "KeyPointFlags"), ("Slots.Add", "KeyPointFlags2")}
)


@dataclass(frozen=True)
class Finding:
    loc: str
    iface: str
    method: str
    param: str
    enum: str
    value: int
    source: str

    def __str__(self) -> str:
        return (
            f"{self.loc}  {self.iface}.{self.method}({self.param}={self.source}) "
            f"= {self.value}, which is not in {self.enum}"
        )


def load_receivers() -> ModuleType:
    spec = importlib.util.spec_from_file_location("_audit_com_receivers", RECEIVERS)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    sys.modules[spec.name] = module
    spec.loader.exec_module(module)
    return module


def load_enums(data: dict) -> dict[str, dict[str, int]]:
    enums: dict[str, dict[str, int]] = {}
    for lib in data["typelibs"].values():
        for name, values in (lib.get("enums") or {}).items():
            enums.setdefault(name, {}).update(values)
    return enums


def load_our_constants() -> dict[str, dict[str, int]]:
    """Our own constants classes, as {ClassName: {member: value}}."""
    tree = ast.parse(CONSTANTS.read_text(encoding="utf-8"))
    out: dict[str, dict[str, int]] = {}
    for node in tree.body:
        if not isinstance(node, ast.ClassDef):
            continue
        members: dict[str, int] = {}
        for stmt in node.body:
            if isinstance(stmt, ast.Assign) and len(stmt.targets) == 1:
                target = stmt.targets[0]
                value = stmt.value
                if (
                    isinstance(target, ast.Name)
                    and isinstance(value, ast.Constant)
                    and isinstance(value.value, int)
                    and not isinstance(value.value, bool)
                ):
                    members[target.id] = value.value
        if members:
            out[node.name] = members
    return out


def property_types(data: dict) -> dict[str, dict[str, str]]:
    """{interface: {property: declared type}} across every library."""
    out: dict[str, dict[str, str]] = {}
    for lib in data["typelibs"].values():
        for iface, block in (lib.get("interfaces") or {}).items():
            for name, sig in (block.get("properties") or {}).items():
                declared = sig.get("type")
                if declared:
                    out.setdefault(iface, {}).setdefault(name, declared)
    return out


def signatures(data: dict) -> dict[str, dict[str, list[dict]]]:
    """{interface: {method: params}} across every library."""
    out: dict[str, dict[str, list[dict]]] = {}
    for lib in data["typelibs"].values():
        for iface, block in (lib.get("interfaces") or {}).items():
            for method, sig in (block.get("methods") or {}).items():
                out.setdefault(iface, {}).setdefault(method, sig.get("params", []))
    return out


def is_flags(values: dict[str, int]) -> bool:
    """Does this enum look like a bitmask rather than a list of choices?

    ProfileValidationType is one: every member is a power of two, and callers
    OR them together -- our own igProfileForRevolve is igProfileClosed |
    igProfileRefAxisRequired. For such an enum, 0 means "no criteria" and any
    combination of members is legitimate, so neither is a finding. Profile.End
    is called with 0 for a hole profile and the hole is cut, verified live.
    """
    members = [v for v in values.values() if v]
    if len(members) < 2:
        return False
    return all(v > 0 and v & (v - 1) == 0 for v in members)


def accepts(values: dict[str, int], value: int) -> bool:
    """Is ``value`` something this enum can legitimately be given?"""
    if value in values.values():
        return True
    if not is_flags(values):
        return False
    if value == 0:
        return True
    covered = 0
    for member in values.values():
        covered |= member
    return value > 0 and value & ~covered == 0


def build_visitor(receivers: ModuleType, enums, sigs, props, ours):
    class EnumArgVisitor(receivers.Visitor):  # type: ignore[misc, valid-type]
        def __init__(self, path: pathlib.Path, typelib: object) -> None:
            super().__init__(path, typelib)
            self.enum_findings: list[Finding] = []
            self.checked = 0
            self._ints: list[dict[str, tuple[int, str]]] = [{}]

        # -- track locals bound to a constant -------------------------------
        def visit_FunctionDef(self, node: ast.FunctionDef) -> None:  # noqa: N802
            self._ints.append({})
            super().visit_FunctionDef(node)
            self._ints.pop()

        visit_AsyncFunctionDef = visit_FunctionDef  # type: ignore[assignment]

        def visit_Assign(self, node: ast.Assign) -> None:  # noqa: N802
            super().visit_Assign(node)
            if len(node.targets) == 1 and isinstance(node.targets[0], ast.Name):
                resolved = self._value(node.value)
                if resolved is not None:
                    self._ints[-1][node.targets[0].id] = resolved
            for target in node.targets:
                if isinstance(target, ast.Attribute) and target.attr[:1].isupper():
                    self._check_property_write(target, node.value)

        def _value(self, node: ast.expr) -> tuple[int, str] | None:
            """(value, how it was written), for anything we can pin down."""
            if (
                isinstance(node, ast.Constant)
                and isinstance(node.value, int)
                and not isinstance(node.value, bool)
            ):
                return node.value, repr(node.value)
            if isinstance(node, ast.Attribute) and isinstance(node.value, ast.Name):
                members = ours.get(node.value.id)
                if members and node.attr in members:
                    return members[node.attr], f"{node.value.id}.{node.attr}"
            if isinstance(node, ast.Name):
                for scope in reversed(self._ints):
                    if node.id in scope:
                        value, _ = scope[node.id]
                        return value, node.id
            return None

        # -- the check -------------------------------------------------------
        def _flag(self, loc_line, iface, member, param, enum_name, value, source):
            self.checked += 1
            if accepts(enums[enum_name], value):
                return
            if (f"{iface}.{member}", param) in ALLOWED:
                return
            self.enum_findings.append(
                Finding(self._loc(loc_line), iface, member, param, enum_name, value, source)
            )

        def _check_property_write(self, target: ast.Attribute, value: ast.expr) -> None:
            """``obj.Orientation = 2`` declares an enum just as an argument does."""
            resolved = self._value(value)
            if resolved is None:
                return
            for iface in sorted(i for i in self._infer(target.value) if i in props):
                declared = props[iface].get(target.attr)
                if declared in enums:
                    self._flag(
                        target.lineno,
                        iface,
                        target.attr,
                        target.attr,
                        declared,
                        resolved[0],
                        resolved[1],
                    )
                    return

        def visit_Call(self, node: ast.Call) -> None:  # noqa: N802
            super().visit_Call(node)
            func = node.func
            if not isinstance(func, ast.Attribute):
                return
            ifaces = [i for i in self._infer(func.value) if i in sigs]
            params = None
            chosen = None
            for iface in sorted(ifaces):
                if func.attr in sigs[iface]:
                    params, chosen = sigs[iface][func.attr], iface
                    break
            if params is None or chosen is None:
                return

            supplied: list[tuple[dict, ast.expr]] = list(zip(params, node.args, strict=False))
            by_name = {p["name"]: p for p in params}
            for keyword in node.keywords:
                if keyword.arg and keyword.arg in by_name:
                    supplied.append((by_name[keyword.arg], keyword.value))

            for param, arg in supplied:
                enum_name = param.get("type")
                if enum_name not in enums:
                    continue
                resolved = self._value(arg)
                if resolved is None:
                    continue
                value, source = resolved
                self.checked += 1
                if accepts(enums[enum_name], value):
                    continue
                if (f"{chosen}.{func.attr}", param["name"]) in ALLOWED:
                    continue
                self.enum_findings.append(
                    Finding(
                        self._loc(node.lineno),
                        chosen,
                        func.attr,
                        param["name"],
                        enum_name,
                        value,
                        source,
                    )
                )

    return EnumArgVisitor


def audit() -> tuple[list[Finding], int]:
    data = json.loads(DUMP.read_text(encoding="utf-8"))
    receivers = load_receivers()
    typelib = receivers.TypeLib(data)
    visitor_cls = build_visitor(
        receivers,
        load_enums(data),
        signatures(data),
        property_types(data),
        load_our_constants(),
    )

    findings: list[Finding] = []
    checked = 0
    for path in sorted(BACKENDS.rglob("*.py")):
        visitor = visitor_cls(path, typelib)
        visitor.visit(ast.parse(path.read_text(encoding="utf-8")))
        findings += visitor.enum_findings
        checked += visitor.checked
    return findings, checked


def main() -> int:
    if not DUMP.exists():
        print("reference/typelib_dump.json is absent; nothing to check.")
        return 0
    findings, checked = audit()
    print(f"Enum arguments checked against their declared enum: {checked}")
    if not findings:
        print("Every one is a member of the enum its parameter declares.")
        return 0
    print()
    print(f"{len(findings)} argument(s) from the wrong enum:")
    print()
    for finding in findings:
        print(f"  {finding}")
    print()
    print("Look the value up in the type library and use the enum the parameter declares.")
    return 1


if __name__ == "__main__":
    sys.exit(main())
