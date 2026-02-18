"""
Unit tests for QueryManager backend methods (_physical_props.py mixin).

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
# USER PHYSICAL PROPERTIES
# ============================================================================


class TestGetUserPhysicalProperties:
    def test_success_full_tuple(self, query_mgr):
        qm, doc = query_mgr
        cog = (0.01, 0.02, 0.03)
        cov = (0.04, 0.05, 0.06)
        doc.GetUserPhysicalProperties.return_value = (
            1.5e-5,  # volume
            0.001,  # area
            0.12,  # mass
            cog,
            cov,
        )

        result = qm.get_user_physical_properties()
        assert result["status"] == "success"
        assert result["volume"] == 1.5e-5
        assert result["surface_area"] == 0.001
        assert result["mass"] == 0.12
        assert result["center_of_gravity"] == [0.01, 0.02, 0.03]
        assert result["center_of_volume"] == [0.04, 0.05, 0.06]

    def test_success_short_tuple(self, query_mgr):
        qm, doc = query_mgr
        doc.GetUserPhysicalProperties.return_value = (1.0e-5, 0.002, 0.05)

        result = qm.get_user_physical_properties()
        assert result["status"] == "success"
        assert result["volume"] == 1.0e-5
        assert result["surface_area"] == 0.002
        assert result["mass"] == 0.05

    def test_non_tuple_result(self, query_mgr):
        qm, doc = query_mgr
        doc.GetUserPhysicalProperties.return_value = "unexpected"

        result = qm.get_user_physical_properties()
        assert result["status"] == "success"
        assert "raw_result" in result

    def test_com_error(self, query_mgr):
        qm, doc = query_mgr
        doc.GetUserPhysicalProperties.side_effect = Exception("Not a part doc")

        result = qm.get_user_physical_properties()
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
