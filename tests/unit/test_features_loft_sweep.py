"""
Unit tests for FeatureManager backend methods.

Uses unittest.mock to simulate COM objects so tests run without Solid Edge.
"""

from unittest.mock import MagicMock

import pytest


@pytest.fixture
def managers():
    """Create mock doc_manager and sketch_manager for FeatureManager."""
    doc_mgr = MagicMock()
    sketch_mgr = MagicMock()

    # Default: active document with models collection
    doc = MagicMock()
    doc_mgr.get_active_document.return_value = doc

    # Default: model exists (models.Count >= 1)
    models = MagicMock()
    models.Count = 1
    model = MagicMock()
    models.Item.return_value = model
    doc.Models = models

    # Default: body with faces and edges
    body = MagicMock()
    model.Body = body
    face = MagicMock()
    edge1 = MagicMock()
    edge2 = MagicMock()
    edges = MagicMock()
    edges.Count = 2
    edges.Item.side_effect = lambda i: [None, edge1, edge2][i]
    face.Edges = edges
    faces = MagicMock()
    faces.Count = 1
    faces.Item.return_value = face
    body.Faces.return_value = faces

    # Default: active profile exists
    profile = MagicMock()
    sketch_mgr.get_active_sketch.return_value = profile
    sketch_mgr.get_active_refaxis.return_value = None
    sketch_mgr.get_accumulated_profiles.return_value = []

    return doc_mgr, sketch_mgr, doc, models, model, profile


@pytest.fixture
def feature_mgr(managers):
    """Create FeatureManager with mocked dependencies."""
    from solidedge_mcp.backends.features import FeatureManager

    doc_mgr, sketch_mgr, *_ = managers
    return FeatureManager(doc_mgr, sketch_mgr)


# ============================================================================
# LOFT WITH GUIDES
# ============================================================================


class TestCreateLoftWithGuides:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, doc, models, model, _ = managers
        p1, p2, g1 = MagicMock(), MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p1, p2, g1]
        lp = MagicMock()
        model.LoftedProtrusions = lp

        result = feature_mgr.create_loft_with_guides(
            guide_profile_indices=[2], profile_indices=[0, 1]
        )
        assert result["status"] == "created"
        assert result["type"] == "loft_with_guides"
        assert result["num_profiles"] == 2
        assert result["num_guides"] == 1
        lp.AddSimple.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called_once()

    def test_no_guides(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        p1, p2 = MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p1, p2]

        result = feature_mgr.create_loft_with_guides(guide_profile_indices=None)
        assert "error" in result
        assert "guide_profile_indices is required" in result["error"]

    def test_empty_guides(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        p1, p2 = MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p1, p2]

        result = feature_mgr.create_loft_with_guides(guide_profile_indices=[])
        assert "error" in result
        assert "guide_profile_indices is required" in result["error"]

    def test_too_few_cross_sections(self, feature_mgr, managers):
        _, sketch_mgr, doc, models, model, _ = managers
        p1, g1 = MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p1, g1]

        result = feature_mgr.create_loft_with_guides(guide_profile_indices=[1], profile_indices=[0])
        assert "error" in result
        assert "at least 2 cross-section profiles" in result["error"]

    def test_auto_select_profiles(self, feature_mgr, managers):
        """When profile_indices is None, non-guide profiles are auto-selected."""
        _, sketch_mgr, doc, models, model, _ = managers
        p1, p2, p3, g1 = MagicMock(), MagicMock(), MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p1, p2, p3, g1]
        lp = MagicMock()
        model.LoftedProtrusions = lp

        result = feature_mgr.create_loft_with_guides(guide_profile_indices=[3])
        assert result["status"] == "created"
        assert result["num_profiles"] == 3
        assert result["num_guides"] == 1


# ============================================================================
# BOUNDED SURFACE (BlueSurf)
# ============================================================================


class TestCreateBoundedSurface:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, doc, _, model, _ = managers
        p1, p2 = MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p1, p2]
        blue_surfs = MagicMock()
        doc.Constructions.BlueSurfs = blue_surfs

        result = feature_mgr.create_bounded_surface()
        assert result["status"] == "created"
        assert result["type"] == "bounded_surface"
        assert result["num_profiles"] == 2
        assert result["want_end_caps"] is True
        assert result["periodic"] is False
        blue_surfs.Add.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called_once()

    def test_periodic(self, feature_mgr, managers):
        _, sketch_mgr, doc, _, model, _ = managers
        p1, p2, p3 = MagicMock(), MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p1, p2, p3]
        blue_surfs = MagicMock()
        doc.Constructions.BlueSurfs = blue_surfs

        result = feature_mgr.create_bounded_surface(want_end_caps=False, periodic=True)
        assert result["status"] == "created"
        assert result["want_end_caps"] is False
        assert result["periodic"] is True
        assert result["num_profiles"] == 3

    def test_too_few_profiles(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_accumulated_profiles.return_value = [MagicMock()]

        result = feature_mgr.create_bounded_surface()
        assert "error" in result
        assert "at least 2 profiles" in result["error"]

    def test_no_base_feature(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, _ = managers
        sketch_mgr.get_accumulated_profiles.return_value = [MagicMock(), MagicMock()]
        models.Count = 0

        result = feature_mgr.create_bounded_surface()
        assert "error" in result


# ============================================================================
# BASE HELIX (Models.AddFiniteBaseHelix* -- full parameter lists)
# ============================================================================


class TestCreateHelix:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis

        result = feature_mgr.create_helix(0.005, 0.05)
        assert result["status"] == "created"
        assert result["type"] == "helix"
        assert result["revolutions"] == 10.0
        # AddFiniteBaseHelix(HelixAxis, AxisStart, NumCrossSections,
        #   CrossSectionArray, ProfileSide, Height, Pitch, NumberOfTurns,
        #   HelixDir)
        assert models.AddFiniteBaseHelix.call_count == 1
        args = models.AddFiniteBaseHelix.call_args[0]
        assert len(args) == 9
        assert args[0] is refaxis
        assert args[1] == 29  # igStart
        assert args[2] == 1
        # Verified on SE 2026: the helix APIs take a plain list here; a
        # VARIANT(VT_ARRAY | VT_DISPATCH, ...) wrapper is rejected.
        assert args[3] == [profile]
        assert args[4:] == (2, 0.05, 0.005, 10.0, 2)

    def test_left_hand(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = MagicMock()
        result = feature_mgr.create_helix(0.005, 0.05, 4.0, direction="Left")
        assert result["status"] == "created"
        args = models.AddFiniteBaseHelix.call_args[0]
        assert args[7] == 4.0  # NumberOfTurns
        assert args[8] == 1  # HelixDir igLeft

    def test_no_refaxis(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_helix(0.005, 0.05)
        assert "axis of revolution" in result["error"]
        models.AddFiniteBaseHelix.assert_not_called()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_helix(0.005, 0.05)
        assert "error" in result
        models.AddFiniteBaseHelix.assert_not_called()


class TestCreateHelixSync:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis

        result = feature_mgr.create_helix_sync(0.005, 0.05)
        assert result["status"] == "created"
        assert models.AddFiniteBaseHelixSync.call_count == 1
        args = models.AddFiniteBaseHelixSync.call_args[0]
        assert len(args) == 9
        assert args[0] is refaxis
        assert args[1] == 29
        assert args[2] == 1
        # Verified on SE 2026: the helix APIs take a plain list here; a
        # VARIANT(VT_ARRAY | VT_DISPATCH, ...) wrapper is rejected.
        assert args[3] == [profile]
        assert args[4:] == (2, 0.05, 0.005, 10.0, 2)

    def test_no_refaxis(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_helix_sync(0.005, 0.05)
        assert "axis of revolution" in result["error"]
        models.AddFiniteBaseHelixSync.assert_not_called()


class TestCreateHelixThinWall:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis

        result = feature_mgr.create_helix_thin_wall(0.005, 0.05, 0.001)
        assert result["status"] == "created"
        # ... plus ThinWall, AddEndCaps, RemoveInsideMaterial, Thickness,
        # ThicknessSide
        assert models.AddFiniteBaseHelixWithThinWall.call_count == 1
        args = models.AddFiniteBaseHelixWithThinWall.call_args[0]
        assert len(args) == 14
        assert args[0] is refaxis
        assert args[1] == 29
        # Verified on SE 2026: the helix APIs take a plain list here; a
        # VARIANT(VT_ARRAY | VT_DISPATCH, ...) wrapper is rejected.
        assert args[3] == [profile]
        assert args[4:] == (2, 0.05, 0.005, 10.0, 2, True, False, True, 0.001, 4)

    def test_no_refaxis(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_helix_thin_wall(0.005, 0.05, 0.001)
        assert "axis of revolution" in result["error"]
        models.AddFiniteBaseHelixWithThinWall.assert_not_called()


class TestCreateHelixSyncThinWall:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, profile = managers
        refaxis = MagicMock()
        sketch_mgr.get_active_refaxis.return_value = refaxis

        result = feature_mgr.create_helix_sync_thin_wall(0.005, 0.05, 0.001)
        assert result["status"] == "created"
        assert models.AddFiniteBaseHelixSyncWithThinWall.call_count == 1
        args = models.AddFiniteBaseHelixSyncWithThinWall.call_args[0]
        assert len(args) == 14
        assert args[0] is refaxis
        assert args[1] == 29
        # Verified on SE 2026: the helix APIs take a plain list here; a
        # VARIANT(VT_ARRAY | VT_DISPATCH, ...) wrapper is rejected.
        assert args[3] == [profile]
        assert args[4:] == (2, 0.05, 0.005, 10.0, 2, True, False, True, 0.001, 4)

    def test_no_refaxis(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, _ = managers
        sketch_mgr.get_active_refaxis.return_value = None
        result = feature_mgr.create_helix_sync_thin_wall(0.005, 0.05, 0.001)
        assert "axis of revolution" in result["error"]
        models.AddFiniteBaseHelixSyncWithThinWall.assert_not_called()


# ============================================================================
# THIN-WALL LOFT / SWEEP
# ============================================================================


class TestCreateLoftThinWall:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, _ = managers
        p1, p2 = MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p1, p2]

        result = feature_mgr.create_loft_thin_wall(0.002)
        assert result["status"] == "created"
        assert result["type"] == "loft_thin_wall"
        # AddLoftedProtrusionWithThinWall: 21 required arguments, ending in
        # ThinWall, AddEndCaps, RemoveInsideMaterial, Thickness, ThicknessSide
        assert models.AddLoftedProtrusionWithThinWall.call_count == 1
        args = models.AddLoftedProtrusionWithThinWall.call_args[0]
        assert len(args) == 21
        assert args[0] == 2
        assert list(args[1]) == [p1, p2]
        assert list(args[2]) == [48, 48]
        assert args[5:] == (
            2,  # MaterialSide igRight
            44,
            0.0,
            None,  # start extent
            44,
            0.0,
            None,  # end extent
            44,
            0.0,  # start tangent
            44,
            0.0,  # end tangent
            True,
            False,
            True,
            0.002,
            4,  # thin wall
        )

    def test_too_few_profiles(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, _ = managers
        sketch_mgr.get_accumulated_profiles.return_value = [MagicMock()]
        result = feature_mgr.create_loft_thin_wall(0.002)
        assert "error" in result
        models.AddLoftedProtrusionWithThinWall.assert_not_called()


class TestCreateSweepThinWall:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, _ = managers
        path, cs = MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [path, cs]

        result = feature_mgr.create_sweep_thin_wall(0.002)
        assert result["status"] == "created"
        assert result["type"] == "sweep_thin_wall"
        # AddSweptProtrusionWithThinWall: 20 required arguments
        assert models.AddSweptProtrusionWithThinWall.call_count == 1
        args = models.AddSweptProtrusionWithThinWall.call_args[0]
        assert len(args) == 20
        assert args[0] == 1
        assert list(args[1]) == [path]
        assert list(args[2]) == [48]
        assert args[3] == 1
        assert list(args[4]) == [cs]
        assert list(args[5]) == [48]
        assert args[8:] == (
            2,  # MaterialSide igRight
            44,
            0.0,
            None,  # start extent
            44,
            0.0,
            None,  # end extent
            True,
            False,
            True,
            0.002,
            4,  # thin wall
        )

    def test_too_few_profiles(self, feature_mgr, managers):
        _, sketch_mgr, _, models, _, _ = managers
        sketch_mgr.get_accumulated_profiles.return_value = [MagicMock()]
        result = feature_mgr.create_sweep_thin_wall(0.002)
        assert "error" in result
        models.AddSweptProtrusionWithThinWall.assert_not_called()
