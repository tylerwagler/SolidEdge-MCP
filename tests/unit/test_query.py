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
            p.Name = f"Plane_{i+1}"
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
    def test_ordered(self, query_mgr):
        qm, doc = query_mgr
        doc.ModelingMode = 1

        result = qm.get_modeling_mode()
        assert result["mode"] == "ordered"

    def test_synchronous(self, query_mgr):
        qm, doc = query_mgr
        doc.ModelingMode = 2

        result = qm.get_modeling_mode()
        assert result["mode"] == "synchronous"


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
            delay_compute=True,
            screen_updating=False,
            display_alerts=False
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
            0.001,   # volume
            0.06,    # area
            7.85,    # mass
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
