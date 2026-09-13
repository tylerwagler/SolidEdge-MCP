"""Fail when one manager method is defined twice in the mixin chain.

The managers are assembled from a dozen mixins each. When two of them define
the same method, Python silently picks whichever class comes first in the MRO
and the other body becomes unreachable. Nothing warns, the tests for the
shadowed copy keep passing because they call it on its own class, and the
code drifts.

That is not hypothetical. ``convert_feature_type`` was defined in both
``FeatureManagerBase`` and ``MiscFeaturesMixin``. MiscFeaturesMixin comes
first, so the base copy never ran, and the ``ConvertToType`` call inside it --
a method in no Solid Edge type library -- sat there looking live for as long
as it took someone to drive it and find the other one answering.
"""

from __future__ import annotations

import collections

import pytest

from solidedge_mcp.backends.assembly import AssemblyManager
from solidedge_mcp.backends.documents import DocumentManager
from solidedge_mcp.backends.export import ExportManager
from solidedge_mcp.backends.features import FeatureManager
from solidedge_mcp.backends.query import QueryManager
from solidedge_mcp.backends.sketching import SketchManager

MANAGERS = [
    AssemblyManager,
    DocumentManager,
    ExportManager,
    FeatureManager,
    QueryManager,
    SketchManager,
]

#: Methods a mixin is meant to override. Empty on purpose: none of these
#: managers uses inheritance to specialise behaviour, they only compose it.
INTENTIONAL_OVERRIDES: frozenset[tuple[str, str]] = frozenset()


def shadowed(cls: type) -> dict[str, list[str]]:
    """Method names defined by more than one class in the MRO, most-derived first."""
    owners: dict[str, list[str]] = collections.defaultdict(list)
    for klass in cls.__mro__:
        if klass is object:
            continue
        for name, value in vars(klass).items():
            if name.startswith("__") or not callable(value):
                continue
            owners[name].append(klass.__name__)
    return {name: who for name, who in owners.items() if len(who) > 1}


@pytest.mark.parametrize("manager", MANAGERS, ids=lambda c: c.__name__)
def test_no_method_is_shadowed(manager):
    found = {
        name: who
        for name, who in shadowed(manager).items()
        if (manager.__name__, name) not in INTENTIONAL_OVERRIDES
    }
    assert not found, (
        f"{manager.__name__} has methods defined more than once; only the first "
        f"class in the MRO runs and the rest are dead code:\n  "
        + "\n  ".join(f"{name}: {' wins over '.join(who)}" for name, who in sorted(found.items()))
        + "\n\nDelete the unreachable copy, or add it to INTENTIONAL_OVERRIDES "
        "with a note saying why the override is deliberate."
    )


def test_the_check_can_actually_see_a_duplicate():
    """Guard against a silent pass from an empty scan."""

    class First:
        def shared(self):
            return 1

    class Second:
        def shared(self):
            return 2

    class Combined(First, Second):
        pass

    assert shadowed(Combined) == {"shared": ["First", "Second"]}


def test_the_managers_actually_have_methods():
    """A manager that imported as an empty shell would pass vacuously."""
    for manager in MANAGERS:
        names = {
            name
            for klass in manager.__mro__
            for name, value in vars(klass).items()
            if not name.startswith("_") and callable(value)
        }
        assert len(names) > 5, f"{manager.__name__} exposes almost nothing: {names}"
