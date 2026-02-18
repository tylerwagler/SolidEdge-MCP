"""
Unit tests for AssemblyManager features backend methods.

Tests features mixin: AssemblyExtrudedCutout, AssemblyRevolvedCutout,
AssemblyHole, AssemblyExtrudedProtrusion, AssemblyRevolvedProtrusion,
AssemblyMirror, AssemblyPattern, AssemblySweptProtrusion,
RecomputeAssemblyFeatures.
Uses unittest.mock to simulate COM objects.
"""

from unittest.mock import MagicMock

import pytest


@pytest.fixture
def asm_mgr():
    """Create AssemblyManager with mocked dependencies."""
    from solidedge_mcp.backends.assembly import AssemblyManager

    dm = MagicMock()
    doc = MagicMock()
    dm.get_active_document.return_value = doc
    return AssemblyManager(dm), doc


@pytest.fixture
def asm_mgr_with_sketch():
    """Create AssemblyManager with mocked doc and sketch manager."""
    from solidedge_mcp.backends.assembly import AssemblyManager

    dm = MagicMock()
    sm = MagicMock()
    doc = MagicMock()
    dm.get_active_document.return_value = doc
    return AssemblyManager(dm, sm), doc, sm


# ============================================================================
# ASSEMBLY FEATURES
# ============================================================================


class TestAssemblyExtrudedCutout:
    def test_success(self, asm_mgr_with_sketch):
        am, doc, sm = asm_mgr_with_sketch
        profile = MagicMock()
        sm.get_accumulated_profiles.return_value = [profile]
        occ1, occ2 = MagicMock(), MagicMock()
        doc.Occurrences.Item.side_effect = [occ1, occ2]
        cutouts = MagicMock()
        doc.AssemblyFeatures.AssemblyFeaturesExtrudedCutouts = cutouts

        result = am.create_assembly_extruded_cutout([0, 1], distance=0.02)
        assert result["status"] == "created"
        assert result["type"] == "assembly_extruded_cutout"
        cutouts.Add.assert_called_once()

    def test_no_profiles(self, asm_mgr_with_sketch):
        am, doc, sm = asm_mgr_with_sketch
        sm.get_accumulated_profiles.return_value = []

        result = am.create_assembly_extruded_cutout([0])
        assert "error" in result
        assert "No profiles" in result["error"]

    def test_com_error(self, asm_mgr_with_sketch):
        am, doc, sm = asm_mgr_with_sketch
        sm.get_accumulated_profiles.return_value = [MagicMock()]
        doc.AssemblyFeatures.AssemblyFeaturesExtrudedCutouts.Add.side_effect = Exception("COM fail")
        doc.Occurrences.Item.return_value = MagicMock()

        result = am.create_assembly_extruded_cutout([0])
        assert "error" in result


class TestAssemblyRevolvedCutout:
    def test_success(self, asm_mgr_with_sketch):
        am, doc, sm = asm_mgr_with_sketch
        sm.get_accumulated_profiles.return_value = [MagicMock()]
        doc.Occurrences.Item.return_value = MagicMock()
        cutouts = MagicMock()
        doc.AssemblyFeatures.AssemblyFeaturesRevolvedCutouts = cutouts

        result = am.create_assembly_revolved_cutout([0], angle=180.0)
        assert result["status"] == "created"
        assert result["angle"] == 180.0
        cutouts.Add.assert_called_once()

    def test_no_profiles(self, asm_mgr_with_sketch):
        am, doc, sm = asm_mgr_with_sketch
        sm.get_accumulated_profiles.return_value = []

        result = am.create_assembly_revolved_cutout([0])
        assert "error" in result


class TestAssemblyHole:
    def test_success(self, asm_mgr_with_sketch):
        am, doc, sm = asm_mgr_with_sketch
        sm.get_accumulated_profiles.return_value = [MagicMock()]
        doc.Occurrences.Item.return_value = MagicMock()
        holes = MagicMock()
        doc.AssemblyFeatures.AssemblyFeaturesHoles = holes

        result = am.create_assembly_hole([0], depth=0.005)
        assert result["status"] == "created"
        assert result["depth"] == 0.005
        holes.Add.assert_called_once()

    def test_no_profiles(self, asm_mgr_with_sketch):
        am, doc, sm = asm_mgr_with_sketch
        sm.get_accumulated_profiles.return_value = []

        result = am.create_assembly_hole([0])
        assert "error" in result


class TestAssemblyExtrudedProtrusion:
    def test_success(self, asm_mgr_with_sketch):
        am, doc, sm = asm_mgr_with_sketch
        sm.get_accumulated_profiles.return_value = [MagicMock()]
        protrusions = MagicMock()
        doc.AssemblyFeatures.ExtrudedProtrusions = protrusions

        result = am.create_assembly_extruded_protrusion(distance=0.05)
        assert result["status"] == "created"
        assert result["type"] == "assembly_extruded_protrusion"
        protrusions.Add.assert_called_once()

    def test_no_profiles(self, asm_mgr_with_sketch):
        am, doc, sm = asm_mgr_with_sketch
        sm.get_accumulated_profiles.return_value = []

        result = am.create_assembly_extruded_protrusion()
        assert "error" in result


class TestAssemblyRevolvedProtrusion:
    def test_success(self, asm_mgr_with_sketch):
        am, doc, sm = asm_mgr_with_sketch
        sm.get_accumulated_profiles.return_value = [MagicMock()]
        protrusions = MagicMock()
        doc.AssemblyFeatures.RevolvedProtrusions = protrusions

        result = am.create_assembly_revolved_protrusion(angle=90.0)
        assert result["status"] == "created"
        assert result["angle"] == 90.0
        protrusions.Add.assert_called_once()

    def test_no_profiles(self, asm_mgr_with_sketch):
        am, doc, sm = asm_mgr_with_sketch
        sm.get_accumulated_profiles.return_value = []

        result = am.create_assembly_revolved_protrusion()
        assert "error" in result


class TestAssemblyMirror:
    def test_success(self, asm_mgr_with_sketch):
        am, doc, sm = asm_mgr_with_sketch
        feat = MagicMock()
        cutouts_coll = MagicMock()
        cutouts_coll.Item.return_value = feat
        doc.AssemblyFeatures.AssemblyFeaturesExtrudedCutouts = cutouts_coll
        doc.AssemblyFeatures.AssemblyFeaturesRevolvedCutouts = MagicMock()
        doc.AssemblyFeatures.AssemblyFeaturesRevolvedCutouts.Item.side_effect = Exception
        doc.AssemblyFeatures.AssemblyFeaturesHoles = MagicMock()
        doc.AssemblyFeatures.AssemblyFeaturesHoles.Item.side_effect = Exception
        mirrors = MagicMock()
        doc.AssemblyFeatures.AssemblyFeaturesMirrors = mirrors
        plane = MagicMock()
        doc.RefPlanes.Item.return_value = plane

        result = am.create_assembly_mirror([0], plane_index=2)
        assert result["status"] == "created"
        mirrors.Add.assert_called_once()

    def test_no_features(self, asm_mgr_with_sketch):
        am, doc, sm = asm_mgr_with_sketch
        for coll_name in [
            "AssemblyFeaturesExtrudedCutouts",
            "AssemblyFeaturesRevolvedCutouts",
            "AssemblyFeaturesHoles",
        ]:
            coll = MagicMock()
            coll.Item.side_effect = Exception("not found")
            setattr(doc.AssemblyFeatures, coll_name, coll)

        result = am.create_assembly_mirror([99])
        assert "error" in result
        assert "No features" in result["error"]


class TestAssemblyPattern:
    def test_success(self, asm_mgr_with_sketch):
        am, doc, sm = asm_mgr_with_sketch
        sm.get_accumulated_profiles.return_value = [MagicMock()]
        feat = MagicMock()
        cutouts_coll = MagicMock()
        cutouts_coll.Item.return_value = feat
        doc.AssemblyFeatures.AssemblyFeaturesExtrudedCutouts = cutouts_coll
        doc.AssemblyFeatures.AssemblyFeaturesRevolvedCutouts = MagicMock()
        doc.AssemblyFeatures.AssemblyFeaturesRevolvedCutouts.Item.side_effect = Exception
        doc.AssemblyFeatures.AssemblyFeaturesHoles = MagicMock()
        doc.AssemblyFeatures.AssemblyFeaturesHoles.Item.side_effect = Exception
        patterns = MagicMock()
        doc.AssemblyFeatures.AssemblyFeaturesPatterns = patterns

        result = am.create_assembly_pattern([0], pattern_type="Circular")
        assert result["status"] == "created"
        assert result["pattern_type"] == "Circular"
        patterns.Add.assert_called_once()


class TestAssemblySweptProtrusion:
    def test_success(self, asm_mgr_with_sketch):
        am, doc, sm = asm_mgr_with_sketch
        trace = MagicMock()
        cross = MagicMock()
        sm.get_accumulated_profiles.return_value = [trace, cross]
        swept = MagicMock()
        doc.AssemblyFeatures.AssemblyFeaturesSweptProtrusions = swept

        result = am.create_assembly_swept_protrusion(1, 1)
        assert result["status"] == "created"
        swept.Add.assert_called_once()

    def test_not_enough_profiles(self, asm_mgr_with_sketch):
        am, doc, sm = asm_mgr_with_sketch
        sm.get_accumulated_profiles.return_value = [MagicMock()]

        result = am.create_assembly_swept_protrusion(1, 1)
        assert "error" in result
        assert "Need 2" in result["error"]


class TestRecomputeAssemblyFeatures:
    def test_success(self, asm_mgr_with_sketch):
        am, doc, sm = asm_mgr_with_sketch
        af = MagicMock()
        doc.AssemblyFeatures = af

        result = am.recompute_assembly_features()
        assert result["status"] == "recomputed"
        af.Recompute.assert_called_once_with(0)

    def test_com_error(self, asm_mgr_with_sketch):
        am, doc, sm = asm_mgr_with_sketch
        doc.AssemblyFeatures.Recompute.side_effect = Exception("fail")

        result = am.recompute_assembly_features()
        assert "error" in result
