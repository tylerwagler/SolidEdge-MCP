"""Feature indices must mean the same thing everywhere.

``list_features`` enumerates ``Models.Item(n).Features``, which holds real
features only. ``_get_feature_by_index`` used to index
``doc.DesignEdgebarFeatures``, which is the whole Pathfinder tree and begins
with three reference planes. Every index-taking operation therefore resolved a
different feature than the caller had read, and ``delete_feature(0)`` removed
a reference plane rather than the first extrusion.

Verified on Solid Edge 2026: a part with one extrusion and one cutout reports
those two through Models.Item(1).Features and five entries through
DesignEdgebarFeatures, the first three being RefPlane_1 to RefPlane_3.
"""

from __future__ import annotations

from unittest.mock import MagicMock

import pytest


class FakeCollection:
    """A 1-indexed COM collection."""

    def __init__(self, items):
        self._items = list(items)

    @property
    def Count(self):
        return len(self._items)

    def Item(self, index):
        return self._items[index - 1]


def feature(name):
    feat = MagicMock()
    feat.Name = name
    feat.Type = 0
    return feat


@pytest.fixture
def part():
    """A part whose Pathfinder tree starts with three reference planes."""
    from solidedge_mcp.backends.features import FeatureManager

    planes = [feature(f"RefPlane_{i}") for i in (1, 2, 3)]
    real = [feature("ExtrudedProtrusion_1"), feature("ExtrudedCutout_1")]

    model = MagicMock()
    model.Features = FakeCollection(real)
    doc = MagicMock()
    doc.Models = FakeCollection([model])
    doc.DesignEdgebarFeatures = FakeCollection(planes + real)

    doc_manager = MagicMock()
    doc_manager.get_active_document.return_value = doc
    return FeatureManager(doc_manager, MagicMock()), doc, real, planes


class TestIndexResolution:
    def test_index_zero_is_the_first_real_feature(self, part):
        manager, _doc, real, _planes = part

        feat, err = manager._get_feature_by_index(0)

        assert err is None
        assert feat is real[0]

    def test_index_one_is_the_second_real_feature(self, part):
        manager, _doc, real, _planes = part

        feat, _err = manager._get_feature_by_index(1)

        assert feat is real[1]

    def test_a_reference_plane_is_never_reachable_by_index(self, part):
        manager, _doc, real, planes = part

        reachable = [manager._get_feature_by_index(i)[0] for i in range(2)]

        assert all(plane not in reachable for plane in planes)

    def test_the_count_in_the_error_matches_list_features(self, part):
        manager, _doc, _real, _planes = part

        _feat, err = manager._get_feature_by_index(4)

        # Four is a valid DesignEdgebarFeatures index and an invalid one here.
        assert err is not None
        assert "2 feature(s)" in err["error"]

    def test_a_negative_index_is_refused(self, part):
        manager, _doc, _real, _planes = part

        _feat, err = manager._get_feature_by_index(-1)

        assert err is not None


class TestIndexedOperations:
    def test_delete_removes_the_feature_the_caller_named(self, part):
        manager, _doc, real, planes = part

        result = manager.delete_feature(0)

        assert result["status"] == "deleted"
        real[0].Delete.assert_called_once()
        for plane in planes:
            plane.Delete.assert_not_called()

    def test_rename_renames_the_right_feature(self, part):
        manager, _doc, real, planes = part

        manager.feature_rename(1, "Pocket")

        assert real[1].Name == "Pocket"
        assert planes[0].Name == "RefPlane_1"

    def test_reorder_uses_the_real_reorder_method(self, part):
        manager, _doc, real, _planes = part

        result = manager.feature_reorder(1, 0, after=False)

        assert result["status"] == "reordered"
        # Reorder(TargetFeature, InsertBefore); MoveAfter/MoveBefore do not exist.
        real[1].Reorder.assert_called_once_with(real[0], True)
        assert not real[1].MoveBefore.called

    def test_reorder_after_inverts_the_insert_before_flag(self, part):
        manager, _doc, real, _planes = part

        manager.feature_reorder(0, 1, after=True)

        real[0].Reorder.assert_called_once_with(real[1], False)

    def test_suppress_targets_the_right_feature(self, part):
        manager, _doc, real, planes = part

        manager.feature_suppress(0)

        assert real[0].Suppress is True
        assert planes[0].Suppress is not True


class TestListAgrees:
    def test_list_features_and_the_index_lookup_use_one_collection(self, part):
        manager, _doc, real, _planes = part

        listed = manager.list_features()

        assert listed["count"] == len(real)
        for entry in listed["features"]:
            found, err = manager._get_feature_by_index(entry["index"])
            assert err is None
            assert found.Name == entry["name"]

    def test_reference_planes_are_not_listed(self, part):
        manager, _doc, _real, _planes = part

        names = [f["name"] for f in manager.list_features()["features"]]

        assert not any(name.startswith("RefPlane") for name in names)


class TestMultipleBodies:
    def test_indices_run_across_bodies_in_order(self):
        from solidedge_mcp.backends.features import FeatureManager

        first = MagicMock()
        first.Features = FakeCollection([feature("A1"), feature("A2")])
        second = MagicMock()
        second.Features = FakeCollection([feature("B1")])

        doc = MagicMock()
        doc.Models = FakeCollection([first, second])
        doc_manager = MagicMock()
        doc_manager.get_active_document.return_value = doc
        manager = FeatureManager(doc_manager, MagicMock())

        listed = manager.list_features()["features"]

        assert [f["name"] for f in listed] == ["A1", "A2", "B1"]
        assert [f["body_index"] for f in listed] == [0, 0, 1]
        # The second body's first feature is flat index 2 but position 1 in it.
        assert manager._get_feature_by_index(2)[0].Name == "B1"
