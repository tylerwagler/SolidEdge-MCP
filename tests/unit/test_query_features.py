"""
Unit tests for QueryManager backend methods (_features.py mixin).

Uses unittest.mock to simulate COM objects.
"""

from unittest.mock import MagicMock

import pytest


@pytest.fixture
def doc_mgr():
    """Create mock doc_manager."""
    dm = MagicMock()
    doc = MagicMock()
    dm.get_active_document.return_value = doc
    return dm, doc


@pytest.fixture
def query_mgr(doc_mgr):
    """Create QueryManager with mocked dependencies."""
    from solidedge_mcp.backends.query import QueryManager

    dm, doc = doc_mgr
    return QueryManager(dm), doc


# ============================================================================
# FEATURE SUPPRESS/UNSUPPRESS
# ============================================================================


class TestSuppressFeature:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        feat = MagicMock()
        feat.Name = "ExtrudedProtrusion_1"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.suppress_feature("ExtrudedProtrusion_1")
        assert result["status"] == "suppressed"
        feat.Suppress.assert_called_once()

    def test_not_found(self, query_mgr):
        qm, doc = query_mgr

        feat = MagicMock()
        feat.Name = "OtherFeature"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.suppress_feature("Nonexistent")
        assert "error" in result


class TestUnsuppressFeature:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        feat = MagicMock()
        feat.Name = "ExtrudedProtrusion_1"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.unsuppress_feature("ExtrudedProtrusion_1")
        assert result["status"] == "unsuppressed"
        feat.Unsuppress.assert_called_once()


# ============================================================================
# DESIGN EDGEBAR FEATURES
# ============================================================================


class TestGetDesignEdgebarFeatures:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        feat1 = MagicMock()
        feat1.Name = "Sketch 1"
        feat1.Type = 1
        feat2 = MagicMock()
        feat2.Name = "ExtrudedProtrusion 1"
        feat2.Type = 3

        features = MagicMock()
        features.Count = 2
        features.Item.side_effect = lambda i: [None, feat1, feat2][i]
        doc.DesignEdgebarFeatures = features

        result = qm.get_design_edgebar_features()
        assert result["count"] == 2
        assert result["features"][0]["name"] == "Sketch 1"
        assert result["features"][1]["name"] == "ExtrudedProtrusion 1"

    def test_not_available(self, query_mgr):
        qm, doc = query_mgr
        del doc.DesignEdgebarFeatures

        result = qm.get_design_edgebar_features()
        assert "error" in result


# ============================================================================
# RENAME FEATURE
# ============================================================================


class TestRenameFeature:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "OldName"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.rename_feature("OldName", "NewName")
        assert result["status"] == "renamed"
        assert feat.Name == "NewName"

    def test_not_found(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Other"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.rename_feature("Nonexistent", "NewName")
        assert "error" in result


# ============================================================================
# DELETE FEATURE
# ============================================================================


class TestDeleteFeature:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        feat = MagicMock()
        feat.Name = "Extrude1"
        debf = MagicMock()
        debf.Count = 1
        debf.Item.return_value = feat
        doc.DesignEdgebarFeatures = debf

        result = qm.delete_feature("Extrude1")
        assert result["status"] == "deleted"
        feat.Delete.assert_called_once()

    def test_not_found(self, query_mgr):
        qm, doc = query_mgr

        feat = MagicMock()
        feat.Name = "Extrude1"
        debf = MagicMock()
        debf.Count = 1
        debf.Item.return_value = feat
        doc.DesignEdgebarFeatures = debf

        result = qm.delete_feature("NonExistent")
        assert "error" in result


# ============================================================================
# GET FEATURE STATUS
# ============================================================================


class TestGetFeatureStatus:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude1"
        feat.Status = 1
        feat.IsSuppressed = False
        feat.Type = 25

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_feature_status("Extrude1")
        assert result["feature_name"] == "Extrude1"
        assert result["index"] == 0
        assert result["status"] == 1
        assert result["is_suppressed"] is False
        assert result["type"] == 25

    def test_not_found(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude1"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_feature_status("NonExistent")
        assert "error" in result
        assert "not found" in result["error"]

    def test_no_design_features(self, query_mgr):
        qm, doc = query_mgr
        del doc.DesignEdgebarFeatures

        result = qm.get_feature_status("Extrude1")
        assert "error" in result


# ============================================================================
# GET FEATURE PROFILES
# ============================================================================


class TestGetFeatureProfiles:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        profile = MagicMock()
        profile.Name = "Profile1"
        profile.Status = 0

        feat = MagicMock()
        feat.Name = "Extrude1"
        feat.GetProfiles.return_value = [profile]

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_feature_profiles("Extrude1")
        assert result["feature_name"] == "Extrude1"
        assert result["count"] == 1
        assert result["profiles"][0]["name"] == "Profile1"

    def test_not_found(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude1"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_feature_profiles("NonExistent")
        assert "error" in result

    def test_no_profiles(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude1"
        feat.GetProfiles.return_value = None

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_feature_profiles("Extrude1")
        assert result["count"] == 0


# ============================================================================
# GET FEATURE PARENTS
# ============================================================================


class TestGetFeatureParents:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        parent = MagicMock()
        parent.Name = "Extrusion1"

        feature = MagicMock()
        feature.Name = "Round1"
        parents = MagicMock()
        parents.Count = 1
        parents.Item.return_value = parent
        feature.Parents = parents

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feature
        doc.DesignEdgebarFeatures = features

        result = qm.get_feature_parents("Round1")
        assert isinstance(result, dict)

    def test_feature_not_found(self, query_mgr):
        qm, doc = query_mgr

        feature = MagicMock()
        feature.Name = "Extrusion1"
        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feature
        doc.DesignEdgebarFeatures = features

        result = qm.get_feature_parents("NonExistent")
        assert "error" in result

    def test_no_features(self, query_mgr):
        qm, doc = query_mgr
        del doc.DesignEdgebarFeatures

        result = qm.get_feature_parents("Round1")
        assert "error" in result


# ============================================================================
# TIER 3: GET FEATURE DIMENSIONS
# ============================================================================


class TestGetFeatureDimensions:
    def test_success(self, query_mgr):
        """Test getting dimensions from a feature."""
        qm, doc = query_mgr

        dim1 = MagicMock()
        dim1.Name = "Depth"
        dim1.Value = 0.05
        dim1.Formula = "0.05"

        dim2 = MagicMock()
        dim2.Name = "Width"
        dim2.Value = 0.1
        dim2.Formula = "0.1"

        feat = MagicMock()
        feat.Name = "ExtrudedProtrusion_1"
        feat.GetDimensions.return_value = (2, [dim1, dim2])

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_feature_dimensions("ExtrudedProtrusion_1")
        assert result["feature_name"] == "ExtrudedProtrusion_1"
        assert result["count"] == 2
        assert result["dimensions"][0]["name"] == "Depth"
        assert result["dimensions"][1]["value"] == 0.1

    def test_feature_not_found(self, query_mgr):
        """Test when feature doesn't exist."""
        qm, doc = query_mgr

        feat = MagicMock()
        feat.Name = "Other_1"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_feature_dimensions("NonExistent")
        assert "error" in result
        assert "not found" in result["error"]

    def test_no_dimensions(self, query_mgr):
        """Test feature that doesn't support GetDimensions."""
        qm, doc = query_mgr

        feat = MagicMock()
        feat.Name = "Round_1"
        feat.GetDimensions.side_effect = Exception("Not supported")

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_feature_dimensions("Round_1")
        assert result["feature_name"] == "Round_1"
        assert "not supported" in result.get("note", "").lower() or result.get("dimensions") == []


# ============================================================================
# FEATURE EDITING: GET DIRECTION 1 EXTENT
# ============================================================================


class TestGetDirection1Extent:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"
        feat.GetDirection1Extent.return_value = (13, 0.05, None)

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_direction1_extent("Extrude 1")
        assert result["feature_name"] == "Extrude 1"
        assert result["extent_type"] == 13
        assert result["distance"] == 0.05

    def test_feature_not_found(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Other"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_direction1_extent("NonExistent")
        assert "error" in result
        assert "not found" in result["error"]

    def test_exception(self, query_mgr):
        qm, doc = query_mgr
        doc.DesignEdgebarFeatures = None

        result = qm.get_direction1_extent("test")
        assert "error" in result


# ============================================================================
# FEATURE EDITING: SET DIRECTION 1 EXTENT
# ============================================================================


class TestSetDirection1Extent:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.set_direction1_extent("Extrude 1", 13, 0.05)
        assert result["status"] == "updated"
        assert result["extent_type"] == 13
        assert result["distance"] == 0.05
        feat.ApplyDirection1Extent.assert_called_once_with(13, 0.05, None)

    def test_feature_not_found(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Other"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.set_direction1_extent("NonExistent", 13, 0.05)
        assert "error" in result
        assert "not found" in result["error"]

    def test_exception(self, query_mgr):
        qm, doc = query_mgr
        doc.DesignEdgebarFeatures = None

        result = qm.set_direction1_extent("test", 13, 0.05)
        assert "error" in result


# ============================================================================
# FEATURE EDITING: GET DIRECTION 2 EXTENT
# ============================================================================


class TestGetDirection2Extent:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"
        feat.GetDirection2Extent.return_value = (16, 0.0, None)

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_direction2_extent("Extrude 1")
        assert result["feature_name"] == "Extrude 1"
        assert result["extent_type"] == 16

    def test_feature_not_found(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Other"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_direction2_extent("NonExistent")
        assert "error" in result

    def test_exception(self, query_mgr):
        qm, doc = query_mgr
        doc.DesignEdgebarFeatures = None

        result = qm.get_direction2_extent("test")
        assert "error" in result


# ============================================================================
# FEATURE EDITING: SET DIRECTION 2 EXTENT
# ============================================================================


class TestSetDirection2Extent:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.set_direction2_extent("Extrude 1", 16, 0.0)
        assert result["status"] == "updated"
        assert result["extent_type"] == 16
        feat.ApplyDirection2Extent.assert_called_once_with(16, 0.0, None)

    def test_feature_not_found(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Other"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.set_direction2_extent("NonExistent", 16, 0.0)
        assert "error" in result

    def test_exception(self, query_mgr):
        qm, doc = query_mgr
        doc.DesignEdgebarFeatures = None

        result = qm.set_direction2_extent("test", 16, 0.0)
        assert "error" in result


# ============================================================================
# FEATURE EDITING: GET THIN WALL OPTIONS
# ============================================================================


class TestGetThinWallOptions:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"
        feat.GetThinWallOptions.return_value = (1, 0.002, 0.003)

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_thin_wall_options("Extrude 1")
        assert result["feature_name"] == "Extrude 1"
        assert result["wall_type"] == 1
        assert result["thickness1"] == 0.002
        assert result["thickness2"] == 0.003

    def test_feature_not_found(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Other"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_thin_wall_options("NonExistent")
        assert "error" in result

    def test_exception(self, query_mgr):
        qm, doc = query_mgr
        doc.DesignEdgebarFeatures = None

        result = qm.get_thin_wall_options("test")
        assert "error" in result


# ============================================================================
# FEATURE EDITING: SET THIN WALL OPTIONS
# ============================================================================


class TestSetThinWallOptions:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.set_thin_wall_options("Extrude 1", 1, 0.002, 0.003)
        assert result["status"] == "updated"
        assert result["wall_type"] == 1
        assert result["thickness1"] == 0.002
        assert result["thickness2"] == 0.003
        feat.SetThinWallOptions.assert_called_once_with(1, 0.002, 0.003)

    def test_feature_not_found(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Other"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.set_thin_wall_options("NonExistent", 1, 0.002)
        assert "error" in result

    def test_exception(self, query_mgr):
        qm, doc = query_mgr
        doc.DesignEdgebarFeatures = None

        result = qm.set_thin_wall_options("test", 1, 0.002)
        assert "error" in result


# ============================================================================
# FEATURE EDITING: GET FROM FACE OFFSET
# ============================================================================


class TestGetFromFaceOffset:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"
        feat.GetFromFaceOffsetData.return_value = (0.01, None)

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_from_face_offset("Extrude 1")
        assert result["feature_name"] == "Extrude 1"
        assert result["offset"] == 0.01

    def test_feature_not_found(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Other"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_from_face_offset("NonExistent")
        assert "error" in result

    def test_exception(self, query_mgr):
        qm, doc = query_mgr
        doc.DesignEdgebarFeatures = None

        result = qm.get_from_face_offset("test")
        assert "error" in result


# ============================================================================
# FEATURE EDITING: SET FROM FACE OFFSET
# ============================================================================


class TestSetFromFaceOffset:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.set_from_face_offset("Extrude 1", 0.01)
        assert result["status"] == "updated"
        assert result["offset"] == 0.01
        feat.SetFromFaceOffsetData.assert_called_once_with(0.01)

    def test_feature_not_found(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Other"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.set_from_face_offset("NonExistent", 0.01)
        assert "error" in result

    def test_exception(self, query_mgr):
        qm, doc = query_mgr
        doc.DesignEdgebarFeatures = None

        result = qm.set_from_face_offset("test", 0.01)
        assert "error" in result


# ============================================================================
# TO-FACE OFFSET: GET
# ============================================================================


class TestGetToFaceOffset:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"
        to_face = MagicMock()
        feat.GetToFaceOffsetData.return_value = (to_face, 2, 0.005)

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_to_face_offset("Extrude 1")
        assert result["feature"] == "Extrude 1"
        assert result["to_face_offset_side"] == 2
        assert result["to_face_offset_distance"] == 0.005

    def test_feature_not_found(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Other"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_to_face_offset("NonExistent")
        assert "error" in result
        assert "not found" in result["error"]

    def test_com_error(self, query_mgr):
        qm, doc = query_mgr
        doc.DesignEdgebarFeatures = None

        result = qm.get_to_face_offset("test")
        assert "error" in result


# ============================================================================
# TO-FACE OFFSET: SET
# ============================================================================


class TestSetToFaceOffset:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"
        to_face = MagicMock()
        feat.GetToFaceOffsetData.return_value = (to_face, 1, 0.001)

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.set_to_face_offset("Extrude 1", 2, 0.01)
        assert result["status"] == "set"
        assert result["feature"] == "Extrude 1"
        assert result["offset_side"] == 2
        assert result["distance"] == 0.01
        feat.SetToFaceOffsetData.assert_called_once_with(to_face, 2, 0.01)

    def test_feature_not_found(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Other"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.set_to_face_offset("NonExistent", 2, 0.01)
        assert "error" in result
        assert "not found" in result["error"]

    def test_com_error(self, query_mgr):
        qm, doc = query_mgr
        doc.DesignEdgebarFeatures = None

        result = qm.set_to_face_offset("test", 2, 0.01)
        assert "error" in result


# ============================================================================
# DIRECTION 1 TREATMENT: GET
# ============================================================================


class TestGetDirection1Treatment:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"
        feat.GetDirection1Treatment.return_value = (1, 2, 0.087, 0, 0, 0, 0.0, 0.0)

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_direction1_treatment("Extrude 1")
        assert result["feature"] == "Extrude 1"
        assert result["treatment_type"] == 1
        assert result["draft_side"] == 2
        assert result["draft_angle"] == 0.087
        assert result["crown_type"] == 0
        assert result["crown_side"] == 0
        assert result["crown_curvature_side"] == 0
        assert result["crown_radius_or_offset"] == 0.0
        assert result["crown_takeoff_angle"] == 0.0

    def test_feature_not_found(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Other"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_direction1_treatment("NonExistent")
        assert "error" in result
        assert "not found" in result["error"]

    def test_non_tuple_result(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"
        feat.GetDirection1Treatment.return_value = "unexpected"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_direction1_treatment("Extrude 1")
        assert result["feature"] == "Extrude 1"
        assert "raw_result" in result


# ============================================================================
# DIRECTION 1 TREATMENT: APPLY
# ============================================================================


class TestApplyDirection1Treatment:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.apply_direction1_treatment(
            "Extrude 1",
            treatment_type=1,
            draft_side=2,
            draft_angle=0.087,
        )
        assert result["status"] == "applied"
        assert result["feature"] == "Extrude 1"
        assert result["treatment_type"] == 1
        feat.ApplyDirection1Treatment.assert_called_once_with(
            1, 2, 0.087, 0, 0, 0, 0.0, 0.0
        )

    def test_all_params(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.apply_direction1_treatment(
            "Extrude 1",
            treatment_type=2,
            draft_side=1,
            draft_angle=0.1,
            crown_type=3,
            crown_side=1,
            crown_curvature_side=2,
            crown_radius_or_offset=0.005,
            crown_takeoff_angle=0.02,
        )
        assert result["status"] == "applied"
        feat.ApplyDirection1Treatment.assert_called_once_with(
            2, 1, 0.1, 3, 1, 2, 0.005, 0.02
        )

    def test_feature_not_found(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Other"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.apply_direction1_treatment("NonExistent")
        assert "error" in result
        assert "not found" in result["error"]

    def test_com_error(self, query_mgr):
        qm, doc = query_mgr
        doc.DesignEdgebarFeatures = None

        result = qm.apply_direction1_treatment("test")
        assert "error" in result


# ============================================================================
# FEATURE EDITING: GET BODY ARRAY
# ============================================================================


class TestGetBodyArray:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        body1 = MagicMock()
        body1.Name = "Body_1"
        body1.Volume = 0.001

        feat = MagicMock()
        feat.Name = "Extrude 1"
        feat.GetBodyArray.return_value = [body1]

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_body_array("Extrude 1")
        assert result["feature_name"] == "Extrude 1"
        assert result["count"] == 1
        assert result["bodies"][0]["name"] == "Body_1"

    def test_feature_not_found(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Other"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_body_array("NonExistent")
        assert "error" in result

    def test_exception(self, query_mgr):
        qm, doc = query_mgr
        doc.DesignEdgebarFeatures = None

        result = qm.get_body_array("test")
        assert "error" in result


# ============================================================================
# FEATURE EDITING: SET BODY ARRAY
# ============================================================================


class TestSetBodyArray:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        model = MagicMock()
        models = MagicMock()
        models.Count = 2
        models.Item.return_value = model
        doc.Models = models

        result = qm.set_body_array("Extrude 1", [0, 1])
        assert result["status"] == "updated"
        assert result["body_count"] == 2
        feat.SetBodyArray.assert_called_once()

    def test_feature_not_found(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Other"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.set_body_array("NonExistent", [0])
        assert "error" in result

    def test_invalid_body_index(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        models = MagicMock()
        models.Count = 1
        doc.Models = models

        result = qm.set_body_array("Extrude 1", [5])
        assert "error" in result
        assert "Invalid body index" in result["error"]
