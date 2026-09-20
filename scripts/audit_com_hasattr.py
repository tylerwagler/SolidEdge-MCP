"""Find ``hasattr`` used to test what a COM proxy can do.

CLAUDE.md's rule: never ``hasattr()`` a COM proxy to test capability. On a
late-bound proxy the probe is a separate ``GetIDsOfNames`` round trip whose
failure mode is version dependent, and it reads False both for a member that is
genuinely absent and for one whose getter merely raised -- turning a real error
into a silent "unsupported".

It cost a whole feature. Every layer call gated on ``hasattr(doc, "Layers")``,
and because a DraftDocument keeps its layers on the Sheet rather than the
document, all five answered "Active document does not support layers" for the
document type where layers matter most.

What to use instead:

    value = com_get(obj, "Member")      # None when absent or raising
    if value is None: ...

    doc.Type == DocumentTypeConstants.igPartDocument   # to test a kind

A COM member is PascalCase, which is what separates these from the legitimate
``hasattr(x, "__iter__")`` and ``hasattr(connection, "on_disconnect")``.
"""

from __future__ import annotations

import ast
import pathlib
import sys

ROOT = pathlib.Path(__file__).resolve().parent.parent
BACKENDS = ROOT / "src" / "solidedge_mcp" / "backends"

#: Probes that are deliberate and cannot be a com_get. Keep this short and say
#: why each one is here.
ALLOWED: frozenset[tuple[str, str]] = frozenset(
    {
        # The name check already pins SaveAsPLMXML as declared only on the
        # generic SolidEdgeDocument interface, which no concrete document
        # answers. The probe is the documented way this one reports itself
        # unsupported rather than raising.
        ("_file_export.py", "SaveAsPLMXML"),
    }
)


def com_member(name: str) -> bool:
    """COM members are PascalCase; our own attributes and dunders are not."""
    return bool(name) and name[0].isupper()


def probes(tree: ast.Module) -> list[tuple[int, str]]:
    """Every ``hasattr(x, "SomeComMember")`` in this module."""
    found: list[tuple[int, str]] = []
    for node in ast.walk(tree):
        if (
            isinstance(node, ast.Call)
            and isinstance(node.func, ast.Name)
            and node.func.id == "hasattr"
            and len(node.args) == 2
            and isinstance(node.args[1], ast.Constant)
            and isinstance(node.args[1].value, str)
            and com_member(node.args[1].value)
        ):
            found.append((node.lineno, node.args[1].value))
    return found


def audit() -> list[tuple[str, int, str]]:
    findings: list[tuple[str, int, str]] = []
    for path in sorted(BACKENDS.rglob("*.py")):
        tree = ast.parse(path.read_text(encoding="utf-8"))
        rel = path.relative_to(ROOT).as_posix()
        for lineno, member in probes(tree):
            if (path.name, member) in ALLOWED:
                continue
            findings.append((rel, lineno, member))
    return findings


def main() -> int:
    findings = audit()
    if not findings:
        print("No hasattr probe of a COM member.")
        return 0
    print(f"{len(findings)} hasattr probe(s) of a COM member:")
    print()
    for rel, lineno, member in findings:
        print(f"  {rel}:{lineno}  hasattr(..., {member!r})")
    print()
    print('Use com_get(obj, "Member") and test for None, or check doc.Type.')
    return 1


if __name__ == "__main__":
    sys.exit(main())
