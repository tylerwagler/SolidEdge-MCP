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


#: One name on two managers, and why that one is tolerated. Each entry has to
#: say what the two do differently and which one the tools call.
SHARED_ACROSS_MANAGERS: dict[str, str] = {
    # FeatureManager.delete_feature(index) is the one manage_feature calls.
    # QueryManager.delete_feature(feature_name) deletes by name and nothing
    # reaches it; the signatures differ, so neither can be passed the other's
    # argument by accident.
    "delete_feature": "FeatureManager takes an index, QueryManager takes a name",
}


def test_no_method_name_is_claimed_by_two_managers():
    """Two managers answering to one name is how index spaces drift apart.

    QueryManager used to carry a list_features() of its own alongside
    FeatureManager's. They enumerated different collections: the query copy
    numbered DesignEdgebarFeatures from zero, counting the three reference
    planes as features, so every index it reported was three too high. Nothing
    called it, which is the only reason it never bit -- wiring one resource to
    the wrong manager would have been enough.
    """
    owners: dict[str, list[str]] = collections.defaultdict(list)
    for manager in MANAGERS:
        for name in dir(manager):
            if name.startswith("_"):
                continue
            if callable(getattr(manager, name, None)):
                owners[name].append(manager.__name__)

    found = {
        name: who
        for name, who in owners.items()
        if len(who) > 1 and name not in SHARED_ACROSS_MANAGERS
    }
    assert not found, (
        "these method names are defined on more than one manager:\n  "
        + "\n  ".join(f"{name}: {', '.join(sorted(who))}" for name, who in sorted(found.items()))
        + "\n\nTwo managers answering to one name drift apart: they end up "
        "enumerating different collections and reporting indices that mean "
        "different things. Delete the copy nothing calls, or add the name to "
        "SHARED_ACROSS_MANAGERS with a note on how the two differ."
    )


def test_the_shared_name_allowlist_is_not_stale():
    """An entry that no longer collides has to be deleted, not left behind."""
    for name in SHARED_ACROSS_MANAGERS:
        owners = [m.__name__ for m in MANAGERS if callable(getattr(m, name, None))]
        assert len(owners) > 1, (
            f"{name!r} is in SHARED_ACROSS_MANAGERS but is now defined on "
            f"{owners}. Remove the entry."
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
