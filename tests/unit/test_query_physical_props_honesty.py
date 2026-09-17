"""A physical-property answer is the computation's result, or it is an error.

Three ways the old code handed back a plausible number that was not one:

    a short COM tuple was padded with zeros and reported as "computed", so a
    caller was told the part had no volume and no mass;
    a face that could not be read was skipped and the partial sum reported
    as the surface area;
    the moments of inertia assumed steel whatever material the part carried,
    with no unit and no density in the result to say so.

Fakes are hand-written: they return exactly the eleven-tuple Solid Edge
returns, or deliberately not.
"""

from __future__ import annotations

from unittest.mock import MagicMock

from solidedge_mcp.backends.query import QueryManager

ELEVEN = (
    1.152e-4,  # volume
    0.01536,  # area
    0.1152,  # mass
    (0.04, 0.024, 0.015),  # cog
    (0.04, 0.024, 0.015),  # cov
    (1.0, 2.0, 3.0, 4.0, 5.0, 6.0),  # global moi
    (0.1, 0.2, 0.3),  # principal
    (0, 0, 1, 0, 1, 0, -1, 0, 0),  # principal axes
    (0.01, 0.02, 0.03),  # radii of gyration
    0.0,  # relative accuracy achieved
    0,  # status
)


def _manager(compute_result, faces=()):
    model = MagicMock()
    model.ComputePhysicalPropertiesWithSpecifiedDensity.return_value = compute_result
    body = MagicMock()
    face_list = MagicMock()
    face_list.Count = len(faces)
    face_list.Item.side_effect = lambda i: faces[i - 1]
    body.Faces.return_value = face_list
    model.Body = body
    models = MagicMock()
    models.Count = 1
    models.Item.return_value = model
    doc = MagicMock()
    doc.Models = models
    doc_manager = MagicMock()
    doc_manager.get_active_document.return_value = doc
    return QueryManager(doc_manager)


class TestAShortTupleIsNotAComputation:
    def test_mass_properties_refuses_a_short_tuple(self):
        result = _manager((1.152e-4, 0.01536)).get_mass_properties()

        assert "error" in result
        assert "2 value(s)" in result["error"]
        assert "volume" not in result, "no zero-padded volume must be reported"

    def test_moments_refuse_a_short_tuple(self):
        result = _manager((1.152e-4,)).get_moments_of_inertia()

        assert "error" in result
        assert "moments_of_inertia" not in result

    def test_the_full_tuple_is_named_field_by_field(self):
        result = _manager(ELEVEN).get_mass_properties(density=1000.0)

        assert result["status"] == "computed"
        assert result["volume"] == 1.152e-4
        assert result["mass"] == 0.1152
        assert result["moments_of_inertia"] == {
            "Ixx": 1.0,
            "Iyy": 2.0,
            "Izz": 3.0,
            "Ixy": 4.0,
            "Ixz": 5.0,
            "Iyz": 6.0,
        }
        assert result["radii_of_gyration"] == [0.01, 0.02, 0.03]
        assert result["relative_accuracy_achieved"] == 0.0
        assert result["compute_status"] == 0


class TestMomentsOfInertiaSayWhatTheyAssume:
    def test_density_is_a_parameter_and_is_reported(self):
        manager = _manager(ELEVEN)

        result = manager.get_moments_of_inertia(density=2700.0)

        model = manager.doc_manager.get_active_document().Models.Item(1)
        kwargs = model.ComputePhysicalPropertiesWithSpecifiedDensity.call_args.kwargs
        assert kwargs["Density"] == 2700.0
        assert result["density"] == 2700.0
        assert result["units"]["moments_of_inertia"] == "kg·m²"
        assert "origin" in result["frame"]

    def test_the_default_is_still_steel_but_now_visible(self):
        result = _manager(ELEVEN).get_moments_of_inertia()

        assert result["density"] == 7850.0

    def test_the_labelled_shape_matches_mass_properties(self):
        manager = _manager(ELEVEN)

        moi = manager.get_moments_of_inertia()["moments_of_inertia"]
        via_mass = manager.get_mass_properties()["moments_of_inertia"]

        assert (
            moi
            == via_mass
            == {"Ixx": 1.0, "Iyy": 2.0, "Izz": 3.0, "Ixy": 4.0, "Ixz": 5.0, "Iyz": 6.0}
        )


class _Face:
    def __init__(self, area):
        self._area = area

    @property
    def Area(self):  # noqa: N802 - COM spelling
        if isinstance(self._area, Exception):
            raise self._area
        return self._area


class TestAPartialSumIsNotTheSurfaceArea:
    def test_every_face_readable_sums_normally(self):
        result = _manager(ELEVEN, faces=[_Face(1.0), _Face(2.0), _Face(3.0)]).get_surface_area()

        assert result["surface_area"] == 6.0
        assert result["face_count"] == 3

    def test_one_unreadable_face_is_an_error_not_a_smaller_number(self):
        faces = [_Face(1.0), _Face(RuntimeError("COM error")), _Face(3.0)]

        result = _manager(ELEVEN, faces=faces).get_surface_area()

        assert "error" in result
        assert "1 of 3 faces" in result["error"]
        assert result["faces_unreadable"] == 1
        assert "surface_area" not in result
