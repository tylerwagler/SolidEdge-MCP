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
# REF PLANE NORMAL TO CURVE
# ============================================================================


class TestRefPlaneNormalToCurve:
    def test_success(self):
        from solidedge_mcp.backends.features import FeatureManager

        dm = MagicMock()
        sm = MagicMock()
        fm = FeatureManager(dm, sm)

        doc = MagicMock()
        dm.get_active_document.return_value = doc
        profile = MagicMock()
        sm.get_active_sketch.return_value = profile

        pivot_plane = MagicMock()
        doc.RefPlanes.Item.return_value = pivot_plane
        doc.RefPlanes.Count = 4

        result = fm.create_ref_plane_normal_to_curve("End", 2)
        assert result["status"] == "created"
        assert result["type"] == "ref_plane_normal_to_curve"
        assert result["new_plane_index"] == 4
        doc.RefPlanes.AddNormalToCurve.assert_called_once()

    def test_no_profile(self):
        from solidedge_mcp.backends.features import FeatureManager

        dm = MagicMock()
        sm = MagicMock()
        fm = FeatureManager(dm, sm)

        dm.get_active_document.return_value = MagicMock()
        sm.get_active_sketch.return_value = None

        result = fm.create_ref_plane_normal_to_curve()
        assert "error" in result


# ============================================================================
# REFERENCE PLANE BY ANGLE
# ============================================================================


class TestRefPlaneByAngle:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 4  # 3 default + 1 new
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_by_angle(1, 45.0)
        assert result["status"] == "created"
        assert result["type"] == "reference_plane"
        assert result["method"] == "angular_by_angle"
        assert result["angle_degrees"] == 45.0
        ref_planes.AddAngularByAngle.assert_called_once()

    def test_invalid_plane_index(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_by_angle(5, 30.0)
        assert "error" in result
        assert "Invalid plane index" in result["error"]

    def test_reverse_side(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 4
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_by_angle(2, 60.0, normal_side="Reverse")
        assert result["status"] == "created"
        assert result["normal_side"] == "Reverse"


# ============================================================================
# REFERENCE PLANE BY 3 POINTS
# ============================================================================


class TestRefPlaneBy3Points:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 4
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_by_3_points(0, 0, 0, 0.1, 0, 0, 0, 0.1, 0)
        assert result["status"] == "created"
        assert result["type"] == "reference_plane"
        assert result["method"] == "by_3_points"
        assert result["point1"] == [0, 0, 0]
        assert result["point2"] == [0.1, 0, 0]
        assert result["point3"] == [0, 0.1, 0]
        ref_planes.AddBy3Points.assert_called_once_with(0, 0, 0, 0.1, 0, 0, 0, 0.1, 0)


# ============================================================================
# REFERENCE PLANE MID-PLANE
# ============================================================================


class TestRefPlaneMidPlane:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 4
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_midplane(1, 3)
        assert result["status"] == "created"
        assert result["type"] == "reference_plane"
        assert result["method"] == "mid_plane"
        assert result["plane1_index"] == 1
        assert result["plane2_index"] == 3
        ref_planes.AddMidPlane.assert_called_once()

    def test_invalid_plane1(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_midplane(5, 1)
        assert "error" in result
        assert "Invalid plane1" in result["error"]

    def test_invalid_plane2(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_midplane(1, 5)
        assert "error" in result
        assert "Invalid plane2" in result["error"]


# ============================================================================
# REF PLANE NORMAL AT DISTANCE
# ============================================================================


class TestCreateRefPlaneNormalAtDistance:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, _, profile = managers
        ref_planes = MagicMock()
        ref_planes.Count = 4
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_normal_at_distance(0.05)
        assert result["status"] == "created"
        assert result["type"] == "ref_plane_normal_at_distance"
        assert result["distance"] == 0.05
        ref_planes.AddNormalToCurveAtDistance.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_ref_plane_normal_at_distance(0.05)
        assert "error" in result


# ============================================================================
# REF PLANE NORMAL AT ARC RATIO
# ============================================================================


class TestCreateRefPlaneNormalAtArcRatio:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, _, profile = managers
        ref_planes = MagicMock()
        ref_planes.Count = 4
        pivot = MagicMock()
        ref_planes.Item.return_value = pivot
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_normal_at_arc_ratio(0.5)
        assert result["status"] == "created"
        assert result["type"] == "ref_plane_normal_at_arc_ratio"
        assert result["ratio"] == 0.5
        ref_planes.AddNormalToCurveAtArcLengthRatio.assert_called_once()

    def test_invalid_ratio(self, feature_mgr, managers):
        _, _, doc, _, _, profile = managers
        ref_planes = MagicMock()
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_normal_at_arc_ratio(1.5)
        assert "error" in result
        assert "Ratio" in result["error"]

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_ref_plane_normal_at_arc_ratio(0.5)
        assert "error" in result


# ============================================================================
# REF PLANE NORMAL AT DISTANCE ALONG
# ============================================================================


class TestCreateRefPlaneNormalAtDistanceAlong:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, _, profile = managers
        ref_planes = MagicMock()
        ref_planes.Count = 4
        pivot = MagicMock()
        ref_planes.Item.return_value = pivot
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_normal_at_distance_along(0.03)
        assert result["status"] == "created"
        assert result["type"] == "ref_plane_normal_at_distance_along"
        assert result["distance_along"] == 0.03
        ref_planes.AddNormalToCurveAtDistanceAlongCurve.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_ref_plane_normal_at_distance_along(0.03)
        assert "error" in result


# ============================================================================
# REF PLANE PARALLEL BY TANGENT
# ============================================================================


class TestCreateRefPlaneParallelByTangent:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        parent = MagicMock()
        ref_planes.Item.return_value = parent
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_parallel_by_tangent(1, 0)
        assert result["status"] == "created"
        assert result["type"] == "ref_plane_parallel_by_tangent"
        assert result["parent_plane_index"] == 1
        assert result["face_index"] == 0
        ref_planes.AddParallelByTangent.assert_called_once()

    def test_invalid_plane_index(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_parallel_by_tangent(10, 0)
        assert "error" in result
        assert "parent_plane_index" in result["error"]

    def test_invalid_face_index(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_parallel_by_tangent(1, 99)
        assert "error" in result
        assert "face_index" in result["error"]

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes
        models.Count = 0

        result = feature_mgr.create_ref_plane_parallel_by_tangent(1, 0)
        assert "error" in result


# ============================================================================
# REF PLANE NORMAL AT KEYPOINT
# ============================================================================


class TestCreateRefPlaneNormalAtKeypoint:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, doc, _, _, profile = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes
        ref_planes.Item.return_value = MagicMock()

        result = feature_mgr.create_ref_plane_normal_at_keypoint("End", 2)
        assert result["status"] == "created"
        assert result["type"] == "ref_plane_normal_at_keypoint"
        assert result["keypoint_type"] == "End"
        ref_planes.AddNormalToCurveAtKeyPoint.assert_called_once()

    def test_no_active_sketch(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None

        result = feature_mgr.create_ref_plane_normal_at_keypoint("End", 2)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_invalid_pivot_plane_index(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_normal_at_keypoint("End", 10)
        assert "error" in result
        assert "Invalid pivot_plane_index" in result["error"]


# ============================================================================
# REF PLANE TANGENT CYLINDER AT ANGLE
# ============================================================================


class TestCreateRefPlaneTangentCylinderAngle:
    def test_success(self, feature_mgr, managers):
        _, _, doc, models, model, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes
        ref_planes.Item.return_value = MagicMock()

        result = feature_mgr.create_ref_plane_tangent_cylinder_angle(0, 45.0, 1)
        assert result["status"] == "created"
        assert result["type"] == "ref_plane_tangent_cylinder_angle"
        assert result["angle_degrees"] == 45.0
        ref_planes.AddTangentToCylinderOrConeAtAngle.assert_called_once()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes
        models.Count = 0

        result = feature_mgr.create_ref_plane_tangent_cylinder_angle(0, 45.0, 1)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_invalid_face_index(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes
        faces = MagicMock()
        faces.Count = 1
        model.Body.Faces.return_value = faces

        result = feature_mgr.create_ref_plane_tangent_cylinder_angle(5, 45.0, 1)
        assert "error" in result
        assert "Invalid face_index" in result["error"]


# ============================================================================
# REF PLANE TANGENT CYLINDER AT KEYPOINT
# ============================================================================


class TestCreateRefPlaneTangentCylinderKeypoint:
    def test_success(self, feature_mgr, managers):
        _, _, doc, models, model, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes
        ref_planes.Item.return_value = MagicMock()

        result = feature_mgr.create_ref_plane_tangent_cylinder_keypoint(0, "End", 1)
        assert result["status"] == "created"
        assert result["type"] == "ref_plane_tangent_cylinder_keypoint"
        assert result["keypoint_type"] == "End"
        ref_planes.AddTangentToCylinderOrConeAtKeyPoint.assert_called_once()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes
        models.Count = 0

        result = feature_mgr.create_ref_plane_tangent_cylinder_keypoint(0, "End", 1)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_invalid_face_index(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes
        faces = MagicMock()
        faces.Count = 1
        model.Body.Faces.return_value = faces

        result = feature_mgr.create_ref_plane_tangent_cylinder_keypoint(5, "End", 1)
        assert "error" in result
        assert "Invalid face_index" in result["error"]


# ============================================================================
# REF PLANE TANGENT SURFACE AT KEYPOINT
# ============================================================================


class TestCreateRefPlaneTangentSurfaceKeypoint:
    def test_success(self, feature_mgr, managers):
        _, _, doc, models, model, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes
        ref_planes.Item.return_value = MagicMock()

        result = feature_mgr.create_ref_plane_tangent_surface_keypoint(0, "Start", 1)
        assert result["status"] == "created"
        assert result["type"] == "ref_plane_tangent_surface_keypoint"
        assert result["keypoint_type"] == "Start"
        ref_planes.AddTangentToCurvedSurfaceAtKeyPoint.assert_called_once()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes
        models.Count = 0

        result = feature_mgr.create_ref_plane_tangent_surface_keypoint(0, "Start", 1)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_invalid_parent_plane_index(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_tangent_surface_keypoint(0, "End", 10)
        assert "error" in result
        assert "Invalid parent_plane_index" in result["error"]


# ============================================================================
# REF PLANE NORMAL AT DISTANCE V2
# ============================================================================


class TestCreateRefPlaneNormalAtDistanceV2:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        body = model.Body
        edges = MagicMock()
        edges.Count = 5
        edge = MagicMock()
        edges.Item.return_value = edge
        body.Edges.return_value = edges
        ref_planes = MagicMock()
        ref_planes.Count = 4
        ref_planes.Item.return_value = MagicMock()
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_normal_at_distance_v2(0, 1, 0.05)
        assert result["status"] == "created"
        assert result["type"] == "ref_plane_normal_at_distance_v2"
        assert result["distance"] == 0.05
        ref_planes.AddNormalToCurveAtDistance.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_ref_plane_normal_at_distance_v2(0, 1, 0.05)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_invalid_edge(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        body = model.Body
        edges = MagicMock()
        edges.Count = 2
        body.Edges.return_value = edges
        ref_planes = MagicMock()
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_normal_at_distance_v2(5, 1, 0.05)
        assert "error" in result
        assert "Invalid curve_edge_index" in result["error"]


class TestCreateRefPlaneNormalAtArcRatioV2:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        body = model.Body
        edges = MagicMock()
        edges.Count = 3
        edges.Item.return_value = MagicMock()
        body.Edges.return_value = edges
        ref_planes = MagicMock()
        ref_planes.Count = 4
        ref_planes.Item.return_value = MagicMock()
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_normal_at_arc_ratio_v2(0, 1, 0.5)
        assert result["status"] == "created"
        assert result["type"] == "ref_plane_normal_at_arc_ratio_v2"
        assert result["ratio"] == 0.5
        ref_planes.AddNormalToCurveAtArcLengthRatio.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_ref_plane_normal_at_arc_ratio_v2(0, 1, 0.5)
        assert "error" in result

    def test_invalid_ratio(self, feature_mgr, managers):
        result = feature_mgr.create_ref_plane_normal_at_arc_ratio_v2(0, 1, 1.5)
        assert "error" in result
        assert "Ratio must be between" in result["error"]


class TestCreateRefPlaneNormalAtDistanceAlongV2:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        body = model.Body
        edges = MagicMock()
        edges.Count = 3
        edges.Item.return_value = MagicMock()
        body.Edges.return_value = edges
        ref_planes = MagicMock()
        ref_planes.Count = 4
        ref_planes.Item.return_value = MagicMock()
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_normal_at_distance_along_v2(0, 1, 0.02)
        assert result["status"] == "created"
        assert result["type"] == "ref_plane_normal_at_distance_along_v2"
        assert result["distance"] == 0.02
        ref_planes.AddNormalToCurveAtDistanceAlongCurve.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_ref_plane_normal_at_distance_along_v2(0, 1, 0.02)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_invalid_edge(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        body = model.Body
        edges = MagicMock()
        edges.Count = 1
        body.Edges.return_value = edges
        ref_planes = MagicMock()
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_normal_at_distance_along_v2(5, 1, 0.02)
        assert "error" in result
        assert "Invalid curve_edge_index" in result["error"]


class TestCreateRefPlaneTangentParallel:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        body = model.Body
        faces = MagicMock()
        faces.Count = 3
        faces.Item.return_value = MagicMock()
        body.Faces.return_value = faces
        ref_planes = MagicMock()
        ref_planes.Count = 4
        ref_planes.Item.return_value = MagicMock()
        doc.RefPlanes = ref_planes

        # Backend signature: (face_index, parent_plane_index, normal_side)
        result = feature_mgr.create_ref_plane_tangent_parallel(0, 1)
        assert result["status"] == "created"
        assert result["type"] == "ref_plane_tangent_parallel"
        ref_planes.AddParallelByTangent.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        models.Count = 0
        ref_planes = MagicMock()
        ref_planes.Count = 4
        doc.RefPlanes = ref_planes
        result = feature_mgr.create_ref_plane_tangent_parallel(0, 1)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_invalid_face(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        body = model.Body
        faces = MagicMock()
        faces.Count = 1
        body.Faces.return_value = faces
        ref_planes = MagicMock()
        ref_planes.Count = 4
        ref_planes.Item.return_value = MagicMock()
        doc.RefPlanes = ref_planes

        result = feature_mgr.create_ref_plane_tangent_parallel(5, 1)
        assert "error" in result
        assert "Invalid face_index" in result["error"]
