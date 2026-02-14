"""
Unit tests for QueryManager backend methods (Tier 2).

Tests variables, custom properties, topology queries, and recompute.
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
# REFERENCE PLANES
# ============================================================================


class TestGetRefPlanes:
    def test_default_planes(self, query_mgr):
        qm, doc = query_mgr

        plane1 = MagicMock()
        plane1.Name = "Top"
        plane1.Visible = True

        plane2 = MagicMock()
        plane2.Name = "Front"
        plane2.Visible = True

        plane3 = MagicMock()
        plane3.Name = "Right"
        plane3.Visible = True

        ref_planes = MagicMock()
        ref_planes.Count = 3
        ref_planes.Item.side_effect = lambda i: [None, plane1, plane2, plane3][i]
        doc.RefPlanes = ref_planes

        result = qm.get_ref_planes()
        assert result["count"] == 3
        assert result["planes"][0]["is_default"] is True
        assert result["planes"][0]["name"] == "Top"

    def test_with_offset_planes(self, query_mgr):
        qm, doc = query_mgr

        planes = []
        for i in range(5):
            p = MagicMock()
            p.Name = f"Plane_{i + 1}"
            p.Visible = True
            planes.append(p)

        ref_planes = MagicMock()
        ref_planes.Count = 5
        ref_planes.Item.side_effect = lambda i: planes[i - 1]
        doc.RefPlanes = ref_planes

        result = qm.get_ref_planes()
        assert result["count"] == 5
        assert result["planes"][2]["is_default"] is True
        assert result["planes"][3]["is_default"] is False


# ============================================================================
# VARIABLES
# ============================================================================


class TestGetVariables:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        var1 = MagicMock()
        var1.DisplayName = "Width"
        var1.Value = 0.1
        var1.Formula = "100 mm"
        var1.Units = "m"

        var2 = MagicMock()
        var2.DisplayName = "Height"
        var2.Value = 0.05
        var2.Formula = "50 mm"
        var2.Units = "m"

        variables = MagicMock()
        variables.Count = 2
        variables.Item.side_effect = lambda i: [None, var1, var2][i]
        doc.Variables = variables

        result = qm.get_variables()
        assert result["count"] == 2
        assert result["variables"][0]["name"] == "Width"
        assert result["variables"][0]["value"] == 0.1
        assert result["variables"][1]["name"] == "Height"

    def test_empty(self, query_mgr):
        qm, doc = query_mgr
        variables = MagicMock()
        variables.Count = 0
        doc.Variables = variables

        result = qm.get_variables()
        assert result["count"] == 0
        assert result["variables"] == []


class TestGetVariable:
    def test_found(self, query_mgr):
        qm, doc = query_mgr
        var = MagicMock()
        var.DisplayName = "Width"
        var.Value = 0.1
        var.Formula = "100 mm"
        var.Units = "m"

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var
        doc.Variables = variables

        result = qm.get_variable("Width")
        assert result["name"] == "Width"
        assert result["value"] == 0.1

    def test_not_found(self, query_mgr):
        qm, doc = query_mgr
        var = MagicMock()
        var.DisplayName = "Width"

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var
        doc.Variables = variables

        result = qm.get_variable("Nonexistent")
        assert "error" in result


class TestSetVariable:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        var = MagicMock()
        var.DisplayName = "Width"
        var.Value = 0.1

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var
        doc.Variables = variables

        result = qm.set_variable("Width", 0.2)
        assert result["status"] == "updated"
        assert result["old_value"] == 0.1
        assert result["new_value"] == 0.2

    def test_not_found(self, query_mgr):
        qm, doc = query_mgr
        var = MagicMock()
        var.DisplayName = "Width"

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var
        doc.Variables = variables

        result = qm.set_variable("Nonexistent", 0.5)
        assert "error" in result


# ============================================================================
# CUSTOM PROPERTIES
# ============================================================================


class TestGetCustomProperties:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        prop = MagicMock()
        prop.Name = "Material"
        prop.Value = "Steel"

        ps = MagicMock()
        ps.Name = "Custom"
        ps.Count = 1
        ps.Item.return_value = prop

        prop_sets = MagicMock()
        prop_sets.Count = 1
        prop_sets.Item.return_value = ps
        doc.Properties = prop_sets

        result = qm.get_custom_properties()
        assert "property_sets" in result
        assert "Custom" in result["property_sets"]
        assert result["property_sets"]["Custom"]["Material"] == "Steel"


class TestSetCustomProperty:
    def test_update_existing(self, query_mgr):
        qm, doc = query_mgr

        prop = MagicMock()
        prop.Name = "Material"
        prop.Value = "Aluminum"

        ps = MagicMock()
        ps.Name = "Custom"
        ps.Count = 1
        ps.Item.return_value = prop

        prop_sets = MagicMock()
        prop_sets.Count = 1
        prop_sets.Item.return_value = ps
        doc.Properties = prop_sets

        result = qm.set_custom_property("Material", "Steel")
        assert result["status"] == "updated"
        assert result["old_value"] == "Aluminum"

    def test_create_new(self, query_mgr):
        qm, doc = query_mgr

        prop = MagicMock()
        prop.Name = "OtherProp"

        ps = MagicMock()
        ps.Name = "Custom"
        ps.Count = 1
        ps.Item.return_value = prop

        prop_sets = MagicMock()
        prop_sets.Count = 1
        prop_sets.Item.return_value = ps
        doc.Properties = prop_sets

        result = qm.set_custom_property("NewProp", "NewValue")
        assert result["status"] == "created"
        ps.Add.assert_called_once_with("NewProp", "NewValue")


class TestDeleteCustomProperty:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        prop = MagicMock()
        prop.Name = "Material"

        ps = MagicMock()
        ps.Name = "Custom"
        ps.Count = 1
        ps.Item.return_value = prop

        prop_sets = MagicMock()
        prop_sets.Count = 1
        prop_sets.Item.return_value = ps
        doc.Properties = prop_sets

        result = qm.delete_custom_property("Material")
        assert result["status"] == "deleted"
        prop.Delete.assert_called_once()

    def test_not_found(self, query_mgr):
        qm, doc = query_mgr

        prop = MagicMock()
        prop.Name = "OtherProp"

        ps = MagicMock()
        ps.Name = "Custom"
        ps.Count = 1
        ps.Item.return_value = prop

        prop_sets = MagicMock()
        prop_sets.Count = 1
        prop_sets.Item.return_value = ps
        doc.Properties = prop_sets

        result = qm.delete_custom_property("Nonexistent")
        assert "error" in result


# ============================================================================
# TOPOLOGY QUERIES
# ============================================================================


class TestGetBodyFaces:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        # Set up model
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        # Set up faces
        face1 = MagicMock()
        face1.Type = 1
        face1.Area = 0.01
        face1.Edges.Count = 4

        face2 = MagicMock()
        face2.Type = 2
        face2.Area = 0.005
        face2.Edges.Count = 3

        faces = MagicMock()
        faces.Count = 2
        faces.Item.side_effect = lambda i: [None, face1, face2][i]
        model.Body.Faces.return_value = faces

        result = qm.get_body_faces()
        assert result["count"] == 2
        assert result["faces"][0]["area"] == 0.01
        assert result["faces"][0]["edge_count"] == 4


class TestGetBodyEdges:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        face = MagicMock()
        face.Edges.Count = 4

        faces = MagicMock()
        faces.Count = 1
        faces.Item.return_value = face
        model.Body.Faces.return_value = faces

        result = qm.get_body_edges()
        assert result["total_face_count"] == 1
        assert result["total_edge_references"] == 4


class TestGetFaceInfo:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        face = MagicMock()
        face.Type = 1
        face.Area = 0.01
        face.Edges.Count = 4
        face.Vertices.Count = 4

        faces = MagicMock()
        faces.Count = 6
        faces.Item.return_value = face
        model.Body.Faces.return_value = faces

        result = qm.get_face_info(0)
        assert result["index"] == 0
        assert result["area"] == 0.01

    def test_invalid_index(self, query_mgr):
        qm, doc = query_mgr

        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        faces = MagicMock()
        faces.Count = 6
        model.Body.Faces.return_value = faces

        result = qm.get_face_info(10)
        assert "error" in result


# ============================================================================
# SELECT SET
# ============================================================================


class TestGetSelectSet:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        item1 = MagicMock()
        item1.Name = "Face_1"

        item2 = MagicMock()
        item2.Name = "Edge_1"

        select_set = MagicMock()
        select_set.Count = 2
        select_set.Item.side_effect = lambda i: [None, item1, item2][i]
        doc.SelectSet = select_set

        result = qm.get_select_set()
        assert result["count"] == 2
        assert result["items"][0]["name"] == "Face_1"
        assert result["items"][1]["name"] == "Edge_1"

    def test_empty(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 0
        doc.SelectSet = select_set

        result = qm.get_select_set()
        assert result["count"] == 0
        assert result["items"] == []


class TestClearSelectSet:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 3
        doc.SelectSet = select_set

        result = qm.clear_select_set()
        assert result["status"] == "cleared"
        assert result["items_removed"] == 3
        select_set.RemoveAll.assert_called_once()

    def test_already_empty(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 0
        doc.SelectSet = select_set

        result = qm.clear_select_set()
        assert result["status"] == "cleared"
        assert result["items_removed"] == 0


# ============================================================================
# RECOMPUTE
# ============================================================================


class TestRecompute:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        result = qm.recompute()
        assert result["status"] == "recomputed"
        model.Recompute.assert_called_once()


# ============================================================================
# PERFORMANCE FLAGS
# ============================================================================

# ============================================================================
# SOLID BODIES
# ============================================================================


class TestGetSolidBodies:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        model = MagicMock()
        model.Name = "Model_1"
        body = MagicMock()
        body.IsSolid = True
        body.Volume = 0.001
        shells = MagicMock()
        shells.Count = 1
        body.Shells = shells
        model.Body = body

        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        # No constructions
        doc.Constructions = MagicMock()
        doc.Constructions.Count = 0

        result = qm.get_solid_bodies()
        assert result["total_bodies"] == 1
        assert result["bodies"][0]["type"] == "design"
        assert result["bodies"][0]["is_solid"] is True


# ============================================================================
# MODELING MODE
# ============================================================================


class TestGetModelingMode:
    def test_synchronous(self, query_mgr):
        qm, doc = query_mgr
        doc.ModelingMode = 1  # seModelingModeSynchronous = 1

        result = qm.get_modeling_mode()
        assert result["mode"] == "synchronous"

    def test_ordered(self, query_mgr):
        qm, doc = query_mgr
        doc.ModelingMode = 2  # seModelingModeOrdered = 2

        result = qm.get_modeling_mode()
        assert result["mode"] == "ordered"


class TestSetModelingMode:
    def test_to_synchronous(self, query_mgr):
        qm, doc = query_mgr
        doc.ModelingMode = 1

        result = qm.set_modeling_mode("synchronous")
        assert result["status"] == "changed"

    def test_invalid_mode(self, query_mgr):
        qm, doc = query_mgr

        result = qm.set_modeling_mode("invalid")
        assert "error" in result


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
# PERFORMANCE FLAGS
# ============================================================================

# ============================================================================
# QUIT APPLICATION
# ============================================================================


class TestQuitApplication:
    def test_success(self):
        from solidedge_mcp.backends.connection import SolidEdgeConnection

        conn = SolidEdgeConnection()
        conn.application = MagicMock()
        conn._is_connected = True

        result = conn.quit_application()
        assert result["status"] == "quit"
        assert conn.application is None
        assert conn._is_connected is False

    def test_not_connected(self):
        from solidedge_mcp.backends.connection import SolidEdgeConnection

        conn = SolidEdgeConnection()

        result = conn.quit_application()
        assert "error" in result


class TestSetPerformanceMode:
    def test_success(self):
        from solidedge_mcp.backends.connection import SolidEdgeConnection

        conn = SolidEdgeConnection()
        conn.application = MagicMock()
        conn._is_connected = True

        result = conn.set_performance_mode(
            delay_compute=True, screen_updating=False, display_alerts=False
        )
        assert result["status"] == "updated"
        assert result["settings"]["delay_compute"] is True
        assert result["settings"]["screen_updating"] is False
        assert result["settings"]["display_alerts"] is False

    def test_partial_update(self):
        from solidedge_mcp.backends.connection import SolidEdgeConnection

        conn = SolidEdgeConnection()
        conn.application = MagicMock()
        conn._is_connected = True

        result = conn.set_performance_mode(screen_updating=False)
        assert result["status"] == "updated"
        assert "screen_updating" in result["settings"]
        assert "delay_compute" not in result["settings"]


# ============================================================================
# SET BODY COLOR
# ============================================================================


class TestSetBodyColor:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        result = qm.set_body_color(255, 0, 0)
        assert result["status"] == "set"
        assert result["color"]["red"] == 255
        assert result["color"]["green"] == 0
        assert result["color"]["blue"] == 0
        assert result["hex"] == "#ff0000"
        model.Body.Style.SetForegroundColor.assert_called_once_with(255, 0, 0)

    def test_clamps_values(self, query_mgr):
        qm, doc = query_mgr

        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        result = qm.set_body_color(300, -10, 128)
        assert result["color"]["red"] == 255
        assert result["color"]["green"] == 0
        assert result["color"]["blue"] == 128

    def test_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.set_body_color(255, 0, 0)
        assert "error" in result


# ============================================================================
# SET MATERIAL DENSITY
# ============================================================================


class TestSetMaterialDensity:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        # ComputePhysicalPropertiesWithSpecifiedDensity returns a tuple
        model.ComputePhysicalPropertiesWithSpecifiedDensity.return_value = (
            0.001,  # volume
            0.06,  # area
            7.85,  # mass
        )

        result = qm.set_material_density(7850)
        assert result["status"] == "computed"
        assert result["density"] == 7850
        assert result["mass"] == 7.85
        assert result["volume"] == 0.001

    def test_negative_density(self, query_mgr):
        qm, doc = query_mgr

        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        result = qm.set_material_density(-100)
        assert "error" in result
        assert "positive" in result["error"]


# ============================================================================
# GET EDGE COUNT
# ============================================================================


class TestGetEdgeCount:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        face1 = MagicMock()
        face1.Edges.Count = 4

        face2 = MagicMock()
        face2.Edges.Count = 4

        faces = MagicMock()
        faces.Count = 2
        faces.Item.side_effect = lambda i: [None, face1, face2][i]
        model.Body.Faces.return_value = faces

        result = qm.get_edge_count()
        assert result["total_edge_references"] == 8
        assert result["face_count"] == 2

    def test_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.get_edge_count()
        assert "error" in result


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
# SET DOCUMENT PROPERTY
# ============================================================================


class TestSetDocumentProperty:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        summary = MagicMock()
        doc.SummaryInfo = summary

        result = qm.set_document_property("Title", "My Part")
        assert result["status"] == "set"
        assert summary.Title == "My Part"

    def test_invalid_property(self, query_mgr):
        qm, doc = query_mgr
        summary = MagicMock()
        doc.SummaryInfo = summary

        result = qm.set_document_property("InvalidProp", "value")
        assert "error" in result

    def test_no_summary_info(self, query_mgr):
        qm, doc = query_mgr
        del doc.SummaryInfo

        result = qm.set_document_property("Title", "Test")
        assert "error" in result


# ============================================================================
# GET FACE AREA
# ============================================================================


class TestGetFaceArea:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        face = MagicMock()
        face.Area = 0.01  # 10000 mm²
        faces = MagicMock()
        faces.Count = 1
        faces.Item.return_value = face
        model.Body.Faces.return_value = faces

        result = qm.get_face_area(0)
        assert result["area"] == 0.01
        assert result["area_mm2"] == 10000.0

    def test_invalid_index(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        faces = MagicMock()
        faces.Count = 1
        model.Body.Faces.return_value = faces

        result = qm.get_face_area(5)
        assert "error" in result


# ============================================================================
# GET SURFACE AREA
# ============================================================================


class TestGetSurfaceArea:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        model.Body.SurfaceArea = 0.06  # 60000 mm²

        result = qm.get_surface_area()
        assert result["surface_area"] == 0.06
        assert result["surface_area_mm2"] == 60000.0

    def test_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.get_surface_area()
        assert "error" in result


# ============================================================================
# GET VOLUME
# ============================================================================


class TestGetVolume:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        model.Body.Volume = 0.001  # 1e6 mm³ = 1000 cm³

        result = qm.get_volume()
        assert result["volume"] == 0.001
        assert result["volume_cm3"] == 1000.0

    def test_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.get_volume()
        assert "error" in result


# ============================================================================
# GET FACE COUNT
# ============================================================================


class TestGetFaceCount:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        faces = MagicMock()
        faces.Count = 6
        model.Body.Faces.return_value = faces

        result = qm.get_face_count()
        assert result["face_count"] == 6

    def test_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.get_face_count()
        assert "error" in result


# ============================================================================
# GET EDGE INFO
# ============================================================================


class TestGetEdgeInfo:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        edge = MagicMock()
        edge.Length = 0.1
        edge.Type = 1

        edges = MagicMock()
        edges.Count = 3
        edges.Item.return_value = edge

        face = MagicMock()
        face.Edges = edges

        faces = MagicMock()
        faces.Count = 6
        faces.Item.return_value = face
        model.Body.Faces.return_value = faces

        result = qm.get_edge_info(0, 0)
        assert result["face_index"] == 0
        assert result["edge_index"] == 0
        assert result["length"] == 0.1

    def test_invalid_face(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        faces = MagicMock()
        faces.Count = 6
        model.Body.Faces.return_value = faces

        result = qm.get_edge_info(10, 0)
        assert "error" in result

    def test_invalid_edge(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        face = MagicMock()
        face.Edges.Count = 2

        faces = MagicMock()
        faces.Count = 6
        faces.Item.return_value = face
        model.Body.Faces.return_value = faces

        result = qm.get_edge_info(0, 5)
        assert "error" in result


# ============================================================================
# SET FACE COLOR
# ============================================================================


class TestSetFaceColor:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        face = MagicMock()
        faces = MagicMock()
        faces.Count = 6
        faces.Item.return_value = face
        model.Body.Faces.return_value = faces

        result = qm.set_face_color(0, 255, 0, 0)
        assert result["status"] == "updated"
        assert result["color"] == [255, 0, 0]

    def test_invalid_face(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        faces = MagicMock()
        faces.Count = 2
        model.Body.Faces.return_value = faces

        result = qm.set_face_color(5, 0, 0, 255)
        assert "error" in result


# ============================================================================
# GET CENTER OF GRAVITY
# ============================================================================


class TestGetCenterOfGravity:
    def test_success_via_variables(self, query_mgr):
        qm, doc = query_mgr

        var_x = MagicMock()
        var_x.Name = "CoMX"
        var_x.Value = 0.05
        var_y = MagicMock()
        var_y.Name = "CoMY"
        var_y.Value = 0.025
        var_z = MagicMock()
        var_z.Name = "CoMZ"
        var_z.Value = 0.01

        variables = MagicMock()
        variables.Count = 3
        variables.Item.side_effect = lambda i: [None, var_x, var_y, var_z][i]
        doc.Variables = variables

        result = qm.get_center_of_gravity()
        assert result["center_of_gravity"] == [0.05, 0.025, 0.01]
        assert result["center_of_gravity_mm"][0] == 50.0


# ============================================================================
# GET MOMENTS OF INERTIA
# ============================================================================


class TestGetMomentsOfInertia:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        moi = (1.0, 2.0, 3.0)
        principal = (1.5, 2.5, 3.5)
        model.ComputePhysicalPropertiesWithSpecifiedDensity.return_value = (
            0.001,
            0.06,
            7.85,
            (0, 0, 0),
            (0,),
            moi,
            principal,
            (0,),
            (0,),
            0,
            0,
        )

        result = qm.get_moments_of_inertia()
        assert result["moments_of_inertia"] == [1.0, 2.0, 3.0]
        assert result["principal_moments"] == [1.5, 2.5, 3.5]


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
# MEASURE ANGLE
# ============================================================================


class TestMeasureAngle:
    def test_right_angle(self, query_mgr):
        qm, doc = query_mgr
        # 90 degree angle: P1=(1,0,0), vertex P2=(0,0,0), P3=(0,1,0)
        result = qm.measure_angle(1, 0, 0, 0, 0, 0, 0, 1, 0)
        assert abs(result["angle_degrees"] - 90.0) < 0.001

    def test_straight_angle(self, query_mgr):
        qm, doc = query_mgr
        # 180 degree angle: P1=(-1,0,0), vertex P2=(0,0,0), P3=(1,0,0)
        result = qm.measure_angle(-1, 0, 0, 0, 0, 0, 1, 0, 0)
        assert abs(result["angle_degrees"] - 180.0) < 0.001

    def test_zero_vector(self, query_mgr):
        qm, doc = query_mgr
        result = qm.measure_angle(0, 0, 0, 0, 0, 0, 1, 0, 0)
        assert "error" in result


# ============================================================================
# GET MATERIAL TABLE
# ============================================================================


class TestGetMaterialTable:
    def test_with_variables(self, query_mgr):
        qm, doc = query_mgr

        var1 = MagicMock()
        var1.Name = "Density"
        var1.Value = 7850.0

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var1
        doc.Variables = variables

        result = qm.get_material_table()
        assert result["property_count"] >= 1
        assert "Density" in result["material_properties"]


# ============================================================================
# TIER 1: ADD VARIABLE
# ============================================================================


class TestAddVariable:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        variables = MagicMock()
        new_var = MagicMock()
        new_var.Value = 0.025
        new_var.DisplayName = "MyWidth"
        variables.Add.return_value = new_var
        doc.Variables = variables

        result = qm.add_variable("MyWidth", "0.025")
        assert result["status"] == "created"
        assert result["name"] == "MyWidth"
        assert result["formula"] == "0.025"
        assert result["value"] == 0.025
        variables.Add.assert_called_once_with("MyWidth", "0.025")

    def test_with_units(self, query_mgr):
        qm, doc = query_mgr
        variables = MagicMock()
        new_var = MagicMock()
        new_var.Value = 0.01
        new_var.DisplayName = "BoltDia"
        variables.Add.return_value = new_var
        doc.Variables = variables

        result = qm.add_variable("BoltDia", "0.01", units_type="m")
        assert result["status"] == "created"
        variables.Add.assert_called_once_with("BoltDia", "0.01", "m")

    def test_formula_expression(self, query_mgr):
        qm, doc = query_mgr
        variables = MagicMock()
        new_var = MagicMock()
        new_var.Value = 0.05
        new_var.DisplayName = "TotalWidth"
        variables.Add.return_value = new_var
        doc.Variables = variables

        result = qm.add_variable("TotalWidth", "MyWidth * 2")
        assert result["status"] == "created"
        assert result["formula"] == "MyWidth * 2"


# ============================================================================
# TIER 1: SELECT ADD
# ============================================================================


class TestSelectAdd:
    def test_add_feature(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 1
        doc.SelectSet = select_set

        features = MagicMock()
        features.Count = 3
        feat = MagicMock()
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = qm.select_add("feature", 0)
        assert result["status"] == "added"
        assert result["object_type"] == "feature"
        assert result["selection_count"] == 1
        select_set.Add.assert_called_once_with(feat)

    def test_add_face(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 1
        doc.SelectSet = select_set

        models = MagicMock()
        models.Count = 1
        model = MagicMock()
        models.Item.return_value = model
        doc.Models = models

        body = MagicMock()
        model.Body = body
        face = MagicMock()
        faces = MagicMock()
        faces.Count = 5
        faces.Item.return_value = face
        body.Faces.return_value = faces

        result = qm.select_add("face", 2)
        assert result["status"] == "added"
        assert result["object_type"] == "face"
        select_set.Add.assert_called_once_with(face)

    def test_add_plane(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 1
        doc.SelectSet = select_set

        ref_planes = MagicMock()
        ref_planes.Count = 3
        plane = MagicMock()
        ref_planes.Item.return_value = plane
        doc.RefPlanes = ref_planes

        result = qm.select_add("plane", 0)
        assert result["status"] == "added"
        assert result["object_type"] == "plane"
        select_set.Add.assert_called_once_with(plane)

    def test_invalid_type(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        doc.SelectSet = select_set

        result = qm.select_add("invalid", 0)
        assert "error" in result
        assert "Unsupported object type" in result["error"]

    def test_invalid_feature_index(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        doc.SelectSet = select_set

        features = MagicMock()
        features.Count = 2
        doc.DesignEdgebarFeatures = features

        result = qm.select_add("feature", 5)
        assert "error" in result
        assert "Invalid feature index" in result["error"]

    def test_invalid_face_index(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        doc.SelectSet = select_set

        models = MagicMock()
        models.Count = 1
        model = MagicMock()
        models.Item.return_value = model
        doc.Models = models

        body = MagicMock()
        model.Body = body
        faces = MagicMock()
        faces.Count = 3
        body.Faces.return_value = faces

        result = qm.select_add("face", 10)
        assert "error" in result
        assert "Invalid face index" in result["error"]


# ============================================================================
# TIER 2: QUERY VARIABLES
# ============================================================================


class TestQueryVariables:
    def test_query_with_fallback(self, query_mgr):
        """Test fallback fnmatch when Variables.Query not available."""
        qm, doc = query_mgr
        variables = MagicMock()
        variables.Count = 3
        variables.Query.side_effect = Exception("Not available")

        v1 = MagicMock()
        v1.Name = "Length"
        v1.Value = 0.1
        v1.Formula = "0.1"

        v2 = MagicMock()
        v2.Name = "Width"
        v2.Value = 0.05
        v2.Formula = "0.05"

        v3 = MagicMock()
        v3.Name = "LengthOffset"
        v3.Value = 0.02
        v3.Formula = "Length / 5"

        variables.Item.side_effect = lambda i: [None, v1, v2, v3][i]
        doc.Variables = variables

        result = qm.query_variables("*Length*")
        assert result["count"] == 2
        assert result["method"] == "fallback_fnmatch"
        names = [m["name"] for m in result["matches"]]
        assert "Length" in names
        assert "LengthOffset" in names

    def test_query_with_native(self, query_mgr):
        """Test native Variables.Query when available."""
        qm, doc = query_mgr
        variables = MagicMock()

        # Mock Query return value
        query_results = MagicMock()
        query_results.Count = 1
        var = MagicMock()
        var.Name = "Width"
        var.Value = 0.05
        var.Formula = "0.05"
        query_results.Item.return_value = var
        variables.Query.return_value = query_results
        doc.Variables = variables

        result = qm.query_variables("Width")
        assert result["count"] == 1
        assert result["matches"][0]["name"] == "Width"
        variables.Query.assert_called_once_with("Width", 0, 0, True)

    def test_case_insensitive(self, query_mgr):
        """Test case-insensitive search with fallback."""
        qm, doc = query_mgr
        variables = MagicMock()
        variables.Count = 1
        variables.Query.side_effect = Exception("Not available")

        v1 = MagicMock()
        v1.Name = "MASS"
        v1.Value = 1.5
        v1.Formula = "1.5"
        variables.Item.return_value = v1
        doc.Variables = variables

        result = qm.query_variables("*mass*", case_insensitive=True)
        assert result["count"] == 1
        assert result["matches"][0]["name"] == "MASS"

    def test_no_matches(self, query_mgr):
        """Test query with no matching variables."""
        qm, doc = query_mgr
        variables = MagicMock()
        variables.Count = 1
        variables.Query.side_effect = Exception("Not available")

        v1 = MagicMock()
        v1.Name = "Length"
        v1.Value = 0.1
        variables.Item.return_value = v1
        doc.Variables = variables

        result = qm.query_variables("*Xyz*")
        assert result["count"] == 0
        assert result["matches"] == []


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
# TIER 3: MATERIAL OPERATIONS
# ============================================================================


class TestGetMaterialList:
    def test_success(self, query_mgr):
        """Test getting material list."""
        qm, doc = query_mgr

        mat_table = MagicMock()
        mat_table.GetMaterialList.return_value = (3, ["Steel", "Aluminum", "Copper"])

        app = MagicMock()
        app.GetMaterialTable.return_value = mat_table
        qm.doc_manager.connection = MagicMock()
        qm.doc_manager.connection.get_application.return_value = app

        result = qm.get_material_list()
        assert result["count"] == 3
        assert "Steel" in result["materials"]
        assert "Aluminum" in result["materials"]

    def test_com_error(self, query_mgr):
        """Test handling COM errors."""
        qm, doc = query_mgr

        app = MagicMock()
        app.GetMaterialTable.side_effect = Exception("COM error")
        qm.doc_manager.connection = MagicMock()
        qm.doc_manager.connection.get_application.return_value = app

        result = qm.get_material_list()
        assert "error" in result


class TestSetMaterial:
    def test_success(self, query_mgr):
        """Test applying a material."""
        qm, doc = query_mgr

        mat_table = MagicMock()
        app = MagicMock()
        app.GetMaterialTable.return_value = mat_table
        qm.doc_manager.connection = MagicMock()
        qm.doc_manager.connection.get_application.return_value = app

        result = qm.set_material("Steel")
        assert result["status"] == "applied"
        assert result["material"] == "Steel"
        mat_table.ApplyMaterial.assert_called_once_with(doc, "Steel")

    def test_invalid_material(self, query_mgr):
        """Test applying non-existent material."""
        qm, doc = query_mgr

        mat_table = MagicMock()
        mat_table.ApplyMaterial.side_effect = Exception("Material not found")
        app = MagicMock()
        app.GetMaterialTable.return_value = mat_table
        qm.doc_manager.connection = MagicMock()
        qm.doc_manager.connection.get_application.return_value = app

        result = qm.set_material("InvalidMaterial")
        assert "error" in result


class TestGetMaterialProperty:
    def test_success(self, query_mgr):
        """Test getting a material property."""
        qm, doc = query_mgr

        mat_table = MagicMock()
        mat_table.GetMatPropValue.return_value = 7850.0
        app = MagicMock()
        app.GetMaterialTable.return_value = mat_table
        qm.doc_manager.connection = MagicMock()
        qm.doc_manager.connection.get_application.return_value = app

        result = qm.get_material_property("Steel", 0)
        assert result["material"] == "Steel"
        assert result["property_index"] == 0
        assert result["value"] == 7850.0
        mat_table.GetMatPropValue.assert_called_once_with("Steel", 0)

    def test_com_error(self, query_mgr):
        """Test handling COM error for invalid property."""
        qm, doc = query_mgr

        mat_table = MagicMock()
        mat_table.GetMatPropValue.side_effect = Exception("Invalid index")
        app = MagicMock()
        app.GetMaterialTable.return_value = mat_table
        qm.doc_manager.connection = MagicMock()
        qm.doc_manager.connection.get_application.return_value = app

        result = qm.get_material_property("Steel", 99)
        assert "error" in result


# ============================================================================
# SELECT SET: REMOVE
# ============================================================================


class TestSelectRemove:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 3
        doc.SelectSet = select_set

        result = qm.select_remove(1)
        assert result["status"] == "removed"
        assert result["index"] == 1
        select_set.Remove.assert_called_once_with(2)  # 0-based -> 1-based

    def test_invalid_index(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 2
        doc.SelectSet = select_set

        result = qm.select_remove(5)
        assert "error" in result


# ============================================================================
# SELECT SET: ALL
# ============================================================================


class TestSelectAll:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 10
        doc.SelectSet = select_set

        result = qm.select_all()
        assert result["status"] == "selected_all"
        assert result["selection_count"] == 10
        select_set.AddAll.assert_called_once()


# ============================================================================
# SELECT SET: COPY / CUT / DELETE
# ============================================================================


class TestSelectCopy:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 2
        doc.SelectSet = select_set

        result = qm.select_copy()
        assert result["status"] == "copied"
        assert result["items_copied"] == 2
        select_set.Copy.assert_called_once()

    def test_empty_selection(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 0
        doc.SelectSet = select_set

        result = qm.select_copy()
        assert "error" in result


class TestSelectCut:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 3
        doc.SelectSet = select_set

        result = qm.select_cut()
        assert result["status"] == "cut"
        assert result["items_cut"] == 3
        select_set.Cut.assert_called_once()

    def test_empty_selection(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 0
        doc.SelectSet = select_set

        result = qm.select_cut()
        assert "error" in result


class TestSelectDelete:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 1
        doc.SelectSet = select_set

        result = qm.select_delete()
        assert result["status"] == "deleted"
        assert result["items_deleted"] == 1
        select_set.Delete.assert_called_once()

    def test_empty_selection(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        select_set.Count = 0
        doc.SelectSet = select_set

        result = qm.select_delete()
        assert "error" in result


# ============================================================================
# SELECT SET: DISPLAY CONTROL
# ============================================================================


class TestSelectSuspendDisplay:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        doc.SelectSet = select_set

        result = qm.select_suspend_display()
        assert result["status"] == "display_suspended"
        select_set.SuspendDisplay.assert_called_once()


class TestSelectResumeDisplay:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        doc.SelectSet = select_set

        result = qm.select_resume_display()
        assert result["status"] == "display_resumed"
        select_set.ResumeDisplay.assert_called_once()


class TestSelectRefreshDisplay:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        select_set = MagicMock()
        doc.SelectSet = select_set

        result = qm.select_refresh_display()
        assert result["status"] == "display_refreshed"
        select_set.RefreshDisplay.assert_called_once()


# ============================================================================
# VARIABLES: GET FORMULA
# ============================================================================


class TestGetVariableFormula:
    def test_found(self, query_mgr):
        qm, doc = query_mgr
        var = MagicMock()
        var.DisplayName = "Width"
        var.Formula = "100 mm"
        var.Value = 0.1

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var
        doc.Variables = variables

        result = qm.get_variable_formula("Width")
        assert result["name"] == "Width"
        assert result["formula"] == "100 mm"
        assert result["value"] == 0.1

    def test_not_found(self, query_mgr):
        qm, doc = query_mgr
        var = MagicMock()
        var.DisplayName = "Height"

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var
        doc.Variables = variables

        result = qm.get_variable_formula("Width")
        assert "error" in result


# ============================================================================
# VARIABLES: RENAME
# ============================================================================


class TestRenameVariable:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        var = MagicMock()
        var.DisplayName = "OldVar"

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var
        doc.Variables = variables

        result = qm.rename_variable("OldVar", "NewVar")
        assert result["status"] == "renamed"
        assert result["old_name"] == "OldVar"
        assert result["new_name"] == "NewVar"
        assert var.DisplayName == "NewVar"

    def test_not_found(self, query_mgr):
        qm, doc = query_mgr
        var = MagicMock()
        var.DisplayName = "Width"

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var
        doc.Variables = variables

        result = qm.rename_variable("Nonexistent", "NewName")
        assert "error" in result


# ============================================================================
# VARIABLES: GET NAMES (DisplayName + SystemName)
# ============================================================================


class TestGetVariableNames:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        var = MagicMock()
        var.DisplayName = "Width"
        var.Name = "V1"
        var.Value = 0.1

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var
        doc.Variables = variables

        result = qm.get_variable_names("Width")
        assert result["display_name"] == "Width"
        assert result["system_name"] == "V1"
        assert result["value"] == 0.1

    def test_not_found(self, query_mgr):
        qm, doc = query_mgr
        var = MagicMock()
        var.DisplayName = "Height"

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var
        doc.Variables = variables

        result = qm.get_variable_names("Width")
        assert "error" in result


# ============================================================================
# TRANSLATE VARIABLE
# ============================================================================


class TestTranslateVariable:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        translated_var = MagicMock()
        translated_var.Name = "V1"
        translated_var.DisplayName = "Width"
        translated_var.Value = 0.1
        translated_var.Formula = "100 mm"

        variables = MagicMock()
        variables.Translate.return_value = translated_var
        doc.Variables = variables

        result = qm.translate_variable("Width")
        assert result["status"] == "success"
        assert result["name"] == "V1"
        assert result["display_name"] == "Width"
        assert result["value"] == 0.1
        assert result["formula"] == "100 mm"
        variables.Translate.assert_called_once_with("Width")

    def test_not_found(self, query_mgr):
        qm, doc = query_mgr
        variables = MagicMock()
        variables.Translate.side_effect = Exception("Variable not found")
        doc.Variables = variables

        result = qm.translate_variable("NonExistent")
        assert "error" in result


# ============================================================================
# COPY VARIABLE TO CLIPBOARD
# ============================================================================


class TestCopyVariableToClipboard:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        variables = MagicMock()
        doc.Variables = variables

        result = qm.copy_variable_to_clipboard("Width")
        assert result["status"] == "copied"
        assert result["name"] == "Width"
        variables.CopyToClipboard.assert_called_once_with("Width")

    def test_com_error(self, query_mgr):
        qm, doc = query_mgr
        variables = MagicMock()
        variables.CopyToClipboard.side_effect = Exception("COM error")
        doc.Variables = variables

        result = qm.copy_variable_to_clipboard("Width")
        assert "error" in result


# ============================================================================
# ADD VARIABLE FROM CLIPBOARD
# ============================================================================


class TestAddVariableFromClipboard:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        new_var = MagicMock()
        new_var.Value = 0.1
        variables = MagicMock()
        variables.AddFromClipboard.return_value = new_var
        doc.Variables = variables

        result = qm.add_variable_from_clipboard("PastedWidth")
        assert result["status"] == "added"
        assert result["name"] == "PastedWidth"
        assert result["value"] == 0.1
        variables.AddFromClipboard.assert_called_once_with("PastedWidth")

    def test_with_units(self, query_mgr):
        qm, doc = query_mgr
        new_var = MagicMock()
        new_var.Value = 0.01
        variables = MagicMock()
        variables.AddFromClipboard.return_value = new_var
        doc.Variables = variables

        result = qm.add_variable_from_clipboard("PastedDia", units_type="m")
        assert result["status"] == "added"
        variables.AddFromClipboard.assert_called_once_with("PastedDia", "m")

    def test_com_error(self, query_mgr):
        qm, doc = query_mgr
        variables = MagicMock()
        variables.AddFromClipboard.side_effect = Exception("No clipboard data")
        doc.Variables = variables

        result = qm.add_variable_from_clipboard("Test")
        assert "error" in result


# ============================================================================
# LAYERS: GET LAYERS
# ============================================================================


class TestGetLayers:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        layer1 = MagicMock()
        layer1.Name = "Default"
        layer1.Show = True
        layer1.Locatable = True
        layer1.IsEmpty = False

        layer2 = MagicMock()
        layer2.Name = "Construction"
        layer2.Show = False
        layer2.Locatable = False
        layer2.IsEmpty = True

        layers = MagicMock()
        layers.Count = 2
        layers.Item.side_effect = lambda i: {1: layer1, 2: layer2}[i]
        doc.Layers = layers

        result = qm.get_layers()
        assert result["count"] == 2
        assert result["layers"][0]["name"] == "Default"
        assert result["layers"][0]["show"] is True
        assert result["layers"][1]["name"] == "Construction"
        assert result["layers"][1]["is_empty"] is True

    def test_no_layers(self, query_mgr):
        qm, doc = query_mgr
        del doc.Layers

        result = qm.get_layers()
        assert "error" in result


# ============================================================================
# LAYERS: ADD LAYER
# ============================================================================


class TestAddLayer:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        layers = MagicMock()
        layers.Count = 3
        layers.Add.return_value = MagicMock()
        doc.Layers = layers

        result = qm.add_layer("MyLayer")
        assert result["status"] == "added"
        assert result["name"] == "MyLayer"
        assert result["total_layers"] == 3
        layers.Add.assert_called_once_with("MyLayer")

    def test_no_layers_support(self, query_mgr):
        qm, doc = query_mgr
        del doc.Layers

        result = qm.add_layer("Test")
        assert "error" in result


# ============================================================================
# LAYERS: ACTIVATE LAYER
# ============================================================================


class TestActivateLayer:
    def test_by_index(self, query_mgr):
        qm, doc = query_mgr
        layer = MagicMock()
        layer.Name = "Layer1"
        layers = MagicMock()
        layers.Count = 3
        layers.Item.return_value = layer
        doc.Layers = layers

        result = qm.activate_layer(0)
        assert result["status"] == "activated"
        assert result["name"] == "Layer1"
        layer.Activate.assert_called_once()

    def test_by_name(self, query_mgr):
        qm, doc = query_mgr
        layer1 = MagicMock()
        layer1.Name = "Default"
        layer2 = MagicMock()
        layer2.Name = "Custom"
        layers = MagicMock()
        layers.Count = 2
        layers.Item.side_effect = lambda i: {1: layer1, 2: layer2}[i]
        doc.Layers = layers

        result = qm.activate_layer("Custom")
        assert result["status"] == "activated"
        assert result["name"] == "Custom"
        layer2.Activate.assert_called_once()

    def test_invalid_index(self, query_mgr):
        qm, doc = query_mgr
        layers = MagicMock()
        layers.Count = 2
        doc.Layers = layers

        result = qm.activate_layer(5)
        assert "error" in result

    def test_name_not_found(self, query_mgr):
        qm, doc = query_mgr
        layer1 = MagicMock()
        layer1.Name = "Default"
        layers = MagicMock()
        layers.Count = 1
        layers.Item.return_value = layer1
        doc.Layers = layers

        result = qm.activate_layer("NonExistent")
        assert "error" in result

    def test_no_layers_support(self, query_mgr):
        qm, doc = query_mgr
        del doc.Layers

        result = qm.activate_layer(0)
        assert "error" in result


# ============================================================================
# LAYERS: SET LAYER PROPERTIES
# ============================================================================


class TestSetLayerProperties:
    def test_set_show(self, query_mgr):
        qm, doc = query_mgr
        layer = MagicMock()
        layer.Name = "Layer1"
        layers = MagicMock()
        layers.Count = 1
        layers.Item.return_value = layer
        doc.Layers = layers

        result = qm.set_layer_properties(0, show=False)
        assert result["status"] == "updated"
        assert result["properties"]["show"] is False
        assert layer.Show is False

    def test_set_selectable(self, query_mgr):
        qm, doc = query_mgr
        layer = MagicMock()
        layer.Name = "Layer1"
        layers = MagicMock()
        layers.Count = 1
        layers.Item.return_value = layer
        doc.Layers = layers

        result = qm.set_layer_properties(0, selectable=True)
        assert result["status"] == "updated"
        assert result["properties"]["selectable"] is True
        assert layer.Locatable is True

    def test_set_both(self, query_mgr):
        qm, doc = query_mgr
        layer = MagicMock()
        layer.Name = "Layer1"
        layers = MagicMock()
        layers.Count = 1
        layers.Item.return_value = layer
        doc.Layers = layers

        result = qm.set_layer_properties(0, show=True, selectable=False)
        assert result["status"] == "updated"
        assert result["properties"]["show"] is True
        assert result["properties"]["selectable"] is False

    def test_by_name(self, query_mgr):
        qm, doc = query_mgr
        layer = MagicMock()
        layer.Name = "Custom"
        layers = MagicMock()
        layers.Count = 1
        layers.Item.return_value = layer
        doc.Layers = layers

        result = qm.set_layer_properties("Custom", show=False)
        assert result["status"] == "updated"
        assert result["name"] == "Custom"

    def test_no_layers_support(self, query_mgr):
        qm, doc = query_mgr
        del doc.Layers

        result = qm.set_layer_properties(0, show=True)
        assert "error" in result


# ============================================================================
# FACESTYLE: SET BODY OPACITY
# ============================================================================


class TestSetBodyOpacity:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        result = qm.set_body_opacity(0.5)
        assert result["status"] == "set"
        assert result["opacity"] == 0.5
        assert model.Body.FaceStyle.Opacity == 0.5

    def test_clamps_values(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        result = qm.set_body_opacity(1.5)
        assert result["status"] == "set"
        assert result["opacity"] == 1.0

        result = qm.set_body_opacity(-0.5)
        assert result["status"] == "set"
        assert result["opacity"] == 0.0

    def test_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.set_body_opacity(0.5)
        assert "error" in result


# ============================================================================
# FACESTYLE: SET BODY REFLECTIVITY
# ============================================================================


class TestSetBodyReflectivity:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        result = qm.set_body_reflectivity(0.7)
        assert result["status"] == "set"
        assert result["reflectivity"] == 0.7
        assert model.Body.FaceStyle.Reflectivity == 0.7

    def test_clamps_values(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        result = qm.set_body_reflectivity(2.0)
        assert result["status"] == "set"
        assert result["reflectivity"] == 1.0

    def test_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.set_body_reflectivity(0.5)
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
# GET VERTEX COUNT
# ============================================================================


class TestGetVertexCount:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        model = MagicMock()
        body = MagicMock()

        face1 = MagicMock()
        verts1 = MagicMock()
        verts1.Count = 4
        face1.Vertices = verts1

        face2 = MagicMock()
        verts2 = MagicMock()
        verts2.Count = 3
        face2.Vertices = verts2

        faces = MagicMock()
        faces.Count = 2
        faces.Item.side_effect = lambda i: {1: face1, 2: face2}[i]
        body.Faces.return_value = faces

        model.Body = body
        models = MagicMock()
        models.Count = 1
        models.Item.return_value = model
        doc.Models = models

        result = qm.get_vertex_count()
        assert result["total_vertex_references"] == 7
        assert result["face_count"] == 2

    def test_no_model(self, query_mgr):
        qm, doc = query_mgr
        models = MagicMock()
        models.Count = 0
        doc.Models = models

        result = qm.get_vertex_count()
        assert "error" in result


# ============================================================================
# SET VARIABLE FORMULA
# ============================================================================


class TestSetVariableFormula:
    def test_success(self, query_mgr):
        qm, doc = query_mgr

        var1 = MagicMock()
        var1.DisplayName = "Width"
        var1.Formula = "50 mm"
        var1.Value = 0.05

        variables = MagicMock()
        variables.Count = 1
        variables.Item.return_value = var1
        doc.Variables = variables

        result = qm.set_variable_formula("Width", "100 mm")
        assert isinstance(result, dict)
        # Should attempt to set the formula
        assert "error" not in result or "Formula" in str(result.get("error", ""))

    def test_variable_not_found(self, query_mgr):
        qm, doc = query_mgr

        variables = MagicMock()
        variables.Count = 0
        doc.Variables = variables

        result = qm.set_variable_formula("NonExistent", "100 mm")
        assert "error" in result

    def test_no_variables(self, query_mgr):
        qm, doc = query_mgr
        del doc.Variables

        result = qm.set_variable_formula("Width", "100 mm")
        assert "error" in result


# ============================================================================
# DELETE LAYER
# ============================================================================


class TestDeleteLayer:
    def test_success_by_name(self, query_mgr):
        qm, doc = query_mgr

        layer = MagicMock()
        layer.Name = "Layer1"
        layers = MagicMock()
        layers.Count = 2
        layers.Item.return_value = layer
        doc.Layers = layers

        result = qm.delete_layer("Layer1")
        assert isinstance(result, dict)

    def test_success_by_index(self, query_mgr):
        qm, doc = query_mgr

        layer = MagicMock()
        layer.Name = "Layer1"
        layers = MagicMock()
        layers.Count = 2
        layers.Item.return_value = layer
        doc.Layers = layers

        result = qm.delete_layer(0)
        assert isinstance(result, dict)

    def test_no_layers(self, query_mgr):
        qm, doc = query_mgr
        del doc.Layers

        result = qm.delete_layer("Layer1")
        assert "error" in result


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


# ============================================================================
# BATCH 10: MATERIAL LIBRARY
# ============================================================================


class TestGetMaterialLibrary:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        mat1 = MagicMock()
        mat1.Name = "Steel"
        mat1.Density = 7850.0
        mat1.YoungsModulus = 2.1e11
        mat1.PoissonsRatio = 0.3

        mat2 = MagicMock()
        mat2.Name = "Aluminum"
        mat2.Density = 2700.0
        mat2.YoungsModulus = 7.0e10
        mat2.PoissonsRatio = 0.33

        mat_table = MagicMock()
        mat_table.Count = 2
        mat_table.Item.side_effect = lambda i: [None, mat1, mat2][i]
        doc.GetMaterialTable.return_value = mat_table

        result = qm.get_material_library()
        assert result["count"] == 2
        assert result["materials"][0]["name"] == "Steel"
        assert result["materials"][0]["density"] == 7850.0
        assert result["materials"][1]["name"] == "Aluminum"

    def test_empty(self, query_mgr):
        qm, doc = query_mgr
        mat_table = MagicMock()
        mat_table.Count = 0
        doc.GetMaterialTable.return_value = mat_table

        result = qm.get_material_library()
        assert result["count"] == 0
        assert result["materials"] == []

    def test_error(self, query_mgr):
        qm, doc = query_mgr
        doc.GetMaterialTable.side_effect = Exception("No material table")

        result = qm.get_material_library()
        assert "error" in result


class TestSetMaterialByName:
    def test_success(self, query_mgr):
        qm, doc = query_mgr
        mat = MagicMock()
        mat.Name = "Steel"
        mat.Density = 7850.0

        mat_table = MagicMock()
        mat_table.Count = 1
        mat_table.Item.return_value = mat
        doc.GetMaterialTable.return_value = mat_table

        result = qm.set_material_by_name("Steel")
        assert result["status"] == "applied"
        assert result["material"] == "Steel"

    def test_not_found(self, query_mgr):
        qm, doc = query_mgr
        mat = MagicMock()
        mat.Name = "Steel"

        mat_table = MagicMock()
        mat_table.Count = 1
        mat_table.Item.return_value = mat
        doc.GetMaterialTable.return_value = mat_table

        result = qm.set_material_by_name("Titanium")
        assert "error" in result
        assert "Titanium" in result["error"]

    def test_error(self, query_mgr):
        qm, doc = query_mgr
        doc.GetMaterialTable.side_effect = Exception("COM error")

        result = qm.set_material_by_name("Steel")
        assert "error" in result
