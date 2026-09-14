"""Find a COM write that is swallowed and then reported anyway.

The shape:

    with contextlib.suppress(Exception):
        obj.Member = value
    ...
    return {"status": "added", "thing": value}

If Solid Edge refuses the write, the suppress hides it and the result tells
the caller the value was applied. Nothing downstream can notice, because the
call succeeded and the dict looks right.

This heuristic has produced a bug every time it has been run:

    TextBox.TextHeight          not a member; every text box kept the default
    Leader.Text                 not a member; the leader was a bare arrow
    FeatureControlFrame.Text    not a member; the frame was placed empty
    variable.DisplayName        read-only; rename reported "not found"
    DraftPrintUtility.PaperWidth  millimetres, given meters; size never changed
    PMI.Show                    refused on a part with no PMI content

A write that genuinely may not take belongs in ALLOWED, and its result must
report what Solid Edge holds afterwards rather than what was asked for.
"""

from __future__ import annotations

import ast
import pathlib
import sys
from dataclasses import dataclass

ROOT = pathlib.Path(__file__).resolve().parent.parent
BACKENDS = ROOT / "src" / "solidedge_mcp" / "backends"

#: (Class.method, COM member) pairs that are deliberate. Each was driven
#: against live Solid Edge 2026 and does what it says.
ALLOWED: frozenset[tuple[str, str]] = frozenset(
    {
        # Balloon text reads back from the sheet after the call.
        ("AnnotationsMixin.add_balloon", "BalloonText"),
        # Feature rename reads back from the feature tree.
        ("FeatureQueryMixin.rename_feature", "Name"),
        # Custom property reads back from the Custom property set.
        ("VariablesMixin.set_custom_property", "Value"),
    }
)


@dataclass(frozen=True)
class Finding:
    file: str
    line: int
    where: str
    member: str
    var: str

    def __str__(self) -> str:
        return f"{self.file}:{self.line}  {self.where}: .{self.member} = {self.var}"


def _swallows(node: ast.AST) -> bool:
    """A ``with contextlib.suppress(...)`` or a ``try`` whose handler passes."""
    if isinstance(node, ast.With):
        for item in node.items:
            ctx = item.context_expr
            if (
                isinstance(ctx, ast.Call)
                and isinstance(ctx.func, ast.Attribute)
                and ctx.func.attr == "suppress"
            ):
                return True
    elif isinstance(node, ast.Try):
        for handler in node.handlers:
            if all(isinstance(stmt, (ast.Pass, ast.Continue)) for stmt in handler.body):
                return True
    return False


def swallowed_writes(fn: ast.FunctionDef) -> list[tuple[str, str, int]]:
    """COM property writes whose failure cannot be seen, and what fed them."""
    out: list[tuple[str, str, int]] = []
    for node in ast.walk(fn):
        if not _swallows(node):
            continue
        body = getattr(node, "body", [])
        for stmt in ast.walk(ast.Module(body=list(body), type_ignores=[])):
            if not isinstance(stmt, ast.Assign):
                continue
            for target in stmt.targets:
                if (
                    isinstance(target, ast.Attribute)
                    and target.attr[:1].isupper()
                    and isinstance(stmt.value, ast.Name)
                ):
                    out.append((target.attr, stmt.value.id, target.lineno))
    return out


def reported_names(fn: ast.FunctionDef) -> set[str]:
    """Variables handed back as values of a result dict."""
    names: set[str] = set()
    for node in ast.walk(fn):
        value = getattr(node, "value", None)
        if isinstance(node, (ast.Return, ast.Assign)) and isinstance(value, ast.Dict):
            for item in value.values:
                if isinstance(item, ast.Name):
                    names.add(item.id)
    return names


def audit() -> list[Finding]:
    findings: list[Finding] = []
    for path in sorted(BACKENDS.rglob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        rel = path.relative_to(ROOT).as_posix()
        for cls in [n for n in ast.walk(tree) if isinstance(n, ast.ClassDef)]:
            for fn in [n for n in cls.body if isinstance(n, ast.FunctionDef)]:
                writes = swallowed_writes(fn)
                if not writes:
                    continue
                reported = reported_names(fn)
                where = f"{cls.name}.{fn.name}"
                for member, var, lineno in writes:
                    if var in reported and (where, member) not in ALLOWED:
                        findings.append(Finding(rel, lineno, where, member, var))
    return findings


def main() -> int:
    findings = audit()
    if not findings:
        print("No swallowed COM write is reported as applied.")
        return 0
    print(f"{len(findings)} swallowed COM write(s) reported as applied:")
    print()
    for finding in findings:
        print(f"  {finding}")
    print()
    print(
        "Let the write raise, report what Solid Edge holds afterwards, or add "
        "the pair to ALLOWED once it has been driven live."
    )
    return 1


if __name__ == "__main__":
    sys.exit(main())
