"""
Unit tests for QueryManager backend methods (_features.py mixin).

Uses unittest.mock to simulate COM objects.
"""

from unittest.mock import MagicMock

import pytest

from solidedge_mcp.backends.constants import FeatureStatusConstants


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
        # Suppress is a read/write VT_BOOL property, not a method.
        assert feat.Suppress is True

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
        assert feat.Suppress is False
        feat.Unsuppress.assert_not_called()


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

    def test_the_tree_position_is_not_offered_as_a_tool_index(self, query_mgr):
        """Only a real feature gets an index a tool will accept.

        This listing is the whole Pathfinder, sketches and planes included.
        Numbering those 0, 1, 2 under the name "index" invited a caller to
        pass one to manage_feature, which speaks a different index entirely.
        """
        qm, doc = query_mgr
        plane = MagicMock(Name="Front")
        extrude = MagicMock(Name="ExtrudedProtrusion 1", Type=3)

        edgebar = MagicMock()
        edgebar.Count = 2
        edgebar.Item.side_effect = lambda i: [None, plane, extrude][i]
        doc.DesignEdgebarFeatures = edgebar

        model_features = MagicMock()
        model_features.Count = 1
        model_features.Item.return_value = extrude
        model = MagicMock()
        model.Features = model_features
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        rows = qm.get_design_edgebar_features()["features"]

        assert [r["tree_position"] for r in rows] == [0, 1]
        assert rows[0]["index"] is None, "a reference plane is not a tool index"
        assert rows[1]["index"] == 0, "the only real feature is index 0"


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
    """Feature.Status is a FeatureStatusConstants value, not a small integer.

    igFeatureOK is 1216476310, which tells a caller nothing on its own.
    """

    def _feature(self, doc, status, planes_before=0):
        """A document holding one feature, optionally behind some ref planes.

        A real part keeps its features in Models.Item(n).Features and repeats
        them in DesignEdgebarFeatures behind the reference planes, so the two
        collections number the same feature differently.
        """
        feat = MagicMock()
        feat.Name = "Extrude1"
        feat.Status = status
        feat.Suppress = False
        feat.Type = 25

        tree = [MagicMock(Name=f"Plane{i}") for i in range(planes_before)] + [feat]
        edgebar = MagicMock()
        edgebar.Count = len(tree)
        edgebar.Item.side_effect = lambda i: tree[i - 1]
        doc.DesignEdgebarFeatures = edgebar

        model_features = MagicMock()
        model_features.Count = 1
        model_features.Item.return_value = feat
        model = MagicMock()
        model.Features = model_features
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models
        return feat

    def test_success(self, query_mgr):
        qm, doc = query_mgr
        self._feature(doc, FeatureStatusConstants.igFeatureOK)

        result = qm.get_feature_status("Extrude1")

        assert result["feature_name"] == "Extrude1"
        assert result["index"] == 0
        assert result["tree_position"] == 0
        assert result["status"] == "ok"
        assert result["status_code"] == FeatureStatusConstants.igFeatureOK
        assert result["is_suppressed"] is False
        assert result["type"] == 25

    def test_the_index_is_the_one_tools_take_not_the_tree_position(self, query_mgr):
        """The tree carries the reference planes; the tool index does not.

        Reporting the tree position as "index" sent a caller to a different
        feature: live, the base extrusion read as index 3 and renaming index 3
        hit the third cutout and reported success.
        """
        qm, doc = query_mgr
        self._feature(doc, FeatureStatusConstants.igFeatureOK, planes_before=3)

        result = qm.get_feature_status("Extrude1")

        assert result["index"] == 0, "the index a tool would accept"
        assert result["tree_position"] == 3, "where it sits in the Pathfinder"

    def test_every_named_status_is_decoded(self, query_mgr):
        qm, doc = query_mgr
        expected = {
            FeatureStatusConstants.igFeatureOK: "ok",
            FeatureStatusConstants.igFeatureFailed: "failed",
            FeatureStatusConstants.igFeatureWarned: "warned",
            FeatureStatusConstants.igFeatureSuppressed: "suppressed",
            FeatureStatusConstants.igFeatureRolledBack: "rolled_back",
        }
        for code, name in expected.items():
            self._feature(doc, code)
            assert qm.get_feature_status("Extrude1")["status"] == name, name

    def test_a_status_pair_is_unwrapped(self, query_mgr):
        """pywin32 returns Status as (code, None); reading the pair decoded nothing."""
        qm, doc = query_mgr
        self._feature(doc, (FeatureStatusConstants.igFeatureFailed, None))

        result = qm.get_feature_status("Extrude1")

        assert result["status"] == "failed"
        assert result["status_code"] == FeatureStatusConstants.igFeatureFailed

    def test_an_unreadable_status_does_not_crash(self, query_mgr):
        qm, doc = query_mgr
        self._feature(doc, object())

        result = qm.get_feature_status("Extrude1")

        assert result["status"] == "unknown"

    def test_an_unknown_code_is_reported_as_unknown(self, query_mgr):
        qm, doc = query_mgr
        self._feature(doc, 999)

        result = qm.get_feature_status("Extrude1")

        assert result["status"] == "unknown"
        assert result["status_code"] == 999

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

        # GetDimensions(NumDimensions [out], Dimensions SAFEARRAY(DISPATCH)
        # [in,out]) - the buffer must be supplied, by keyword because the
        # count precedes it.
        feat.GetDimensions.assert_called_once()
        args, kwargs = feat.GetDimensions.call_args
        assert args == ()
        assert set(kwargs) == {"Dimensions"}

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
    """Part.tlb: GetDirection1Extent returns (ExtentType, ExtentSide, value).

    Three out-params and no face. The code used to read slot 1 as the
    distance and slot 2 as a face, so the side constant was reported as the
    distance and the real value was stringified into face_ref -- live, an
    extrude read distance=2 and face_ref='0.03'. This test used to feed that
    wrong shape.
    """

    def test_success(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"
        feat.GetDirection1Extent.return_value = (13, 2, 0.05)

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_direction1_extent("Extrude 1")
        assert result["feature_name"] == "Extrude 1"
        assert result["extent_type"] == 13
        assert result["extent_side"] == 2
        assert result["distance"] == 0.05
        assert "face_ref" not in result

    def test_a_short_tuple_is_reported_not_guessed(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"
        feat.GetDirection1Extent.return_value = (13, 2)
        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_direction1_extent("Extrude 1")
        assert "distance" not in result
        assert "expected (type, side, value)" in result["error_detail"]


def _revolved_type_code():
    """The FeatureTypeConstants value describe_feature_type names as revolved."""
    from solidedge_mcp.backends.comutil import FEATURE_TYPE_NAMES

    for code, name in FEATURE_TYPE_NAMES.items():
        if "revolv" in str(name).lower():
            return code
    raise AssertionError("no revolved feature type in FEATURE_TYPE_NAMES")


class TestARevolvedExtentIsAnAngle:
    """On a revolved feature slot 2 of GetDirection1Extent is an angle in radians.

    Verified live: a 90-degree revolve read 1.5707963 there. Degrees is the
    unit at this boundary, so it is reported as angle_degrees, and the setter
    takes degrees and converts before COM.
    """

    def _revolve(self, doc):
        import math

        feat = MagicMock()
        feat.Name = "Revolve 1"
        feat.Type = _revolved_type_code()
        feat.GetDirection1Extent.return_value = (13, 2, math.pi / 2)
        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features
        return feat

    def test_the_getter_reports_degrees(self, query_mgr):
        import math

        qm, doc = query_mgr
        self._revolve(doc)

        result = qm.get_direction1_extent("Revolve 1")

        assert result["angle_degrees"] == pytest.approx(90.0)
        assert result["angle_radians"] == pytest.approx(math.pi / 2)
        assert "distance" not in result

    def test_the_setter_takes_degrees_and_sends_radians(self, query_mgr):
        import math

        qm, doc = query_mgr
        feat = self._revolve(doc)

        qm.set_direction1_extent("Revolve 1", 13, 45.0)

        feat.ApplyDirection1Extent.assert_called_once_with(
            13, 2, pytest.approx(math.radians(45.0)), None, 0
        )

    def test_an_extrude_still_sends_meters_untouched(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"
        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        qm.set_direction1_extent("Extrude 1", 13, 0.05)

        feat.ApplyDirection1Extent.assert_called_once_with(13, 2, 0.05, None, 0)

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
        assert result["extent_side"] == 2  # DirectionConstants.igRight
        # ApplyDirection1Extent(ExtentType, ExtentSide, FiniteDepth,
        #                       KeyPointOrTangentFace, KeyPointFlags)
        feat.ApplyDirection1Extent.assert_called_once_with(13, 2, 0.05, None, 0)

    def test_extent_side_is_passed_through(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        qm.set_direction1_extent("Extrude 1", 13, 0.05, extent_side=3)
        feat.ApplyDirection1Extent.assert_called_once_with(13, 3, 0.05, None, 0)

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
        feat.GetDirection2Extent.return_value = (16, 2, 0.0)

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.get_direction2_extent("Extrude 1")
        assert result["feature_name"] == "Extrude 1"
        assert result["extent_type"] == 16
        assert result["extent_side"] == 2
        assert result["distance"] == 0.0

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
        # ApplyDirection2Extent(ExtentType, ExtentSide, FiniteDepth,
        #                       KeyPointOrTangentFace, KeyPointFlags)
        feat.ApplyDirection2Extent.assert_called_once_with(16, 2, 0.0, None, 0)

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

        result = qm.set_thin_wall_options("Extrude 1", 0.002, 1)
        assert result["status"] == "updated"
        assert result["thickness"] == 0.002
        assert result["thickness_side"] == 1
        assert result["thin_wall"] is True
        # SetThinWallOptions(ThinWall, AddEndCaps, RemoveInsideMaterial,
        #                    Thickness, ThicknessSide)
        feat.SetThinWallOptions.assert_called_once_with(True, False, False, 0.002, 1)

    def test_flags_are_passed_through(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        qm.set_thin_wall_options(
            "Extrude 1",
            0.004,
            2,
            thin_wall=False,
            add_end_caps=True,
            remove_inside_material=True,
        )
        feat.SetThinWallOptions.assert_called_once_with(False, True, True, 0.004, 2)

    def test_feature_not_found(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Other"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.set_thin_wall_options("NonExistent", 0.002, 1)
        assert "error" in result

    def test_exception(self, query_mgr):
        qm, doc = query_mgr
        doc.DesignEdgebarFeatures = None

        result = qm.set_thin_wall_options("test", 0.002, 1)
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

        from_face = MagicMock()
        feat.GetFromFaceOffsetData.return_value = (from_face, 2, 0.005)

        result = qm.set_from_face_offset("Extrude 1", 0.01)
        assert result["status"] == "updated"
        assert result["offset"] == 0.01
        assert result["offset_side"] == 2
        # SetFromFaceOffsetData(FromFaceOrPlane, FromFaceOffsetSide,
        #                       FromFaceOffsetDistance) - the face already on
        # the feature is reused, since this server cannot select one.
        feat.SetFromFaceOffsetData.assert_called_once_with(from_face, 2, 0.01)

    def test_explicit_offset_side_overrides(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        from_face = MagicMock()
        feat.GetFromFaceOffsetData.return_value = (from_face, 2, 0.005)

        qm.set_from_face_offset("Extrude 1", 0.01, offset_side=1)
        feat.SetFromFaceOffsetData.assert_called_once_with(from_face, 1, 0.01)

    def test_no_from_face_is_unsupported(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        feat.GetFromFaceOffsetData.return_value = (None, 44, 0.0)

        result = qm.set_from_face_offset("Extrude 1", 0.01)
        assert result["unsupported"] is True
        assert "cannot select" in result["error"]
        feat.SetFromFaceOffsetData.assert_not_called()

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
        feat.ApplyDirection1Treatment.assert_called_once_with(1, 2, 0.087, 0, 0, 0, 0.0, 0.0)

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
        feat.ApplyDirection1Treatment.assert_called_once_with(2, 1, 0.1, 3, 1, 2, 0.005, 0.02)

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
        assert result["multi_body_cut"] is True

        # SetBodyArray(MultiBodyCut, NumberOfBodies, BodyArray)
        feat.SetBodyArray.assert_called_once()
        args, kwargs = feat.SetBodyArray.call_args
        assert kwargs == {}
        assert len(args) == 3
        assert args[0] is True
        assert args[1] == 2
        # Plain sequence, not a VARIANT: Solid Edge 2026 accepts a list wherever
        # the wrapper was used and rejects the wrapper on several APIs.
        assert isinstance(args[2], list)
        assert list(args[2]) == [model.Body, model.Body]

    def test_multi_body_cut_flag_is_passed(self, query_mgr):
        qm, doc = query_mgr
        feat = MagicMock()
        feat.Name = "Extrude 1"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        models = MagicMock()
        models.Count = 1
        models.Item.return_value = MagicMock()
        doc.Models = models

        qm.set_body_array("Extrude 1", [0], multi_body_cut=False)
        assert feat.SetBodyArray.call_args[0][0] is False

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
