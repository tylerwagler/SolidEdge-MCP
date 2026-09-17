"""How much of the creator surface is verified, and by which check.

Walks ``backends/`` for every ``create_*`` method and buckets it:

    verified     -- decorated with ``@verifies_geometry``,
                    ``@verifies_assembly_geometry`` or
                    ``@verifies_collection_growth(...)``, directly or through
                    the class-level ``@verify_geometry_on_creators`` /
                    ``@verify_collection_growth_on_creators(...)``
    unsupported  -- returns ``"unsupported": True`` somewhere in its body, or
                    calls a same-class helper that does
    neither      -- nothing checks that it built what it claims

A grep for one decorator name undercounts badly: the class-level forms cover
whole mixins, and there are two decorator families. This is the count
``reference/VERIFICATION_STATUS.md`` publishes; regenerate it from here.
"""

from __future__ import annotations

import ast
import pathlib
import sys
from collections import Counter

ROOT = pathlib.Path(__file__).resolve().parent.parent
BACKENDS = ROOT / "src" / "solidedge_mcp" / "backends"

METHOD_DECORATORS = {
    "verifies_geometry",
    "verifies_assembly_geometry",
    "verifies_collection_growth",
}
CLASS_DECORATORS = {
    "verify_geometry_on_creators",
    "verify_collection_growth_on_creators",
}


def _names(node: ast.ClassDef | ast.FunctionDef) -> set[str]:
    """Decorator names, whether bare (``@x``) or called (``@x(...)``)."""
    out: set[str] = set()
    for d in node.decorator_list:
        target = d.func if isinstance(d, ast.Call) else d
        if isinstance(target, ast.Name):
            out.add(target.id)
        elif isinstance(target, ast.Attribute):
            out.add(target.attr)
    return out


def audit() -> dict[str, list[str]]:
    buckets: dict[str, list[str]] = {"verified": [], "unsupported": [], "neither": []}
    for path in sorted(BACKENDS.rglob("*.py")):
        src = path.read_text(encoding="utf-8")
        tree = ast.parse(src)
        rel = path.relative_to(ROOT).as_posix()
        for cls in (n for n in ast.walk(tree) if isinstance(n, ast.ClassDef)):
            whole = bool(_names(cls) & CLASS_DECORATORS)
            # A refusal shared by several creators lives in a helper (the eight
            # flange creators return self._flanges_unsupported(...)); a creator
            # that calls one is a refusal too, whatever it is decorated with.
            refusers = {
                fn.name
                for fn in cls.body
                if isinstance(fn, ast.FunctionDef)
                and '"unsupported": True' in (ast.get_source_segment(src, fn) or "")
            }
            for fn in cls.body:
                if not isinstance(fn, ast.FunctionDef) or not fn.name.startswith("create_"):
                    continue
                body = ast.get_source_segment(src, fn) or ""
                label = f"{rel}::{fn.name}"
                # ... but only as its answer: a precondition helper that can
                # refuse (self._require_synchronous_sheet) is not a refusal.
                refuses = '"unsupported": True' in body or any(
                    f"return self.{name}(" in body for name in refusers - {fn.name}
                )
                if refuses:
                    buckets["unsupported"].append(label)
                elif whole or (_names(fn) & METHOD_DECORATORS):
                    buckets["verified"].append(label)
                else:
                    buckets["neither"].append(label)
    return buckets


def main() -> int:
    b = audit()
    total = sum(len(v) for v in b.values())
    print(f"create_* methods:  {total}")
    print(f"  verified:        {len(b['verified'])}")
    print(f"  unsupported:     {len(b['unsupported'])}")
    print(f"  neither:         {len(b['neither'])}")
    if b["neither"]:
        print("\nunverified, by file:")
        by_file = Counter(label.split("::")[0] for label in b["neither"])
        for f, n in by_file.most_common():
            print(f"  {n:3d}  {f}")
        print("\nunverified, by name:")
        for label in b["neither"]:
            print(f"  {label}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
