"""Assembly feature side constants, and refusing to report a no-op as success.

``AssemblyFeaturePropertyConstants`` is not a Solid Edge enum. Every side
argument on ``AssemblyFeatures*.Add`` is typed ``FeaturePropertyConstants``,
and the values here used to be an invented 0/1/2 numbering whose "Left" was 0
-- ``igNullConstant``, not a side at all.

That was not cosmetic. Verified on Solid Edge 2026 against a one-part
assembly, each attempt in a fresh document:

    extent=0, profile=0   Add succeeds, feature recorded, faces 6 -> 6
    extent=1, profile=0   Add succeeds, feature recorded, faces 6 -> 6
    extent=2, profile=1   Add succeeds, faces 6 -> 7  -- an actual cutout

So the assembly cutout never cut, and said it had.
"""

from __future__ import annotations

from unittest.mock import MagicMock

from solidedge_mcp.backends.assembly._base import (
    occurrence_face_count,
    verifies_assembly_geometry,
)
from solidedge_mcp.backends.constants import (
    AssemblyFeaturePropertyConstants as Sides,
)
from solidedge_mcp.backends.constants import (
    DirectionConstants,
)


class TestTheSidesAreFeaturePropertyConstants:
    def test_profile_sides_match_the_real_enum(self):
        assert Sides.igAssemblyFeatureProfileLeft == DirectionConstants.igLeft == 1
        assert Sides.igAssemblyFeatureProfileRight == DirectionConstants.igRight == 2
        assert Sides.igAssemblyFeatureProfileSymmetric == DirectionConstants.igSymmetric == 3

    def test_extent_sides_match_the_real_enum(self):
        assert Sides.igAssemblyFeatureOneSide == DirectionConstants.igRight
        assert Sides.igAssemblyFeatureBothSides == DirectionConstants.igSymmetric

    def test_no_side_is_the_null_constant(self):
        """0 is igNullConstant. Passing it is what made the cutout do nothing."""
        sides = [
            value
            for name, value in vars(Sides).items()
            if name.startswith("igAssemblyFeature") and isinstance(value, int)
        ]
        assert sides, "the constants class went empty"
        assert 0 not in sides


class _Occurrence:
    def __init__(self, faces: int) -> None:
        self._faces = faces

    @property
    def Body(self):  # noqa: N802 - COM spelling
        holder = MagicMock()
        holder.Faces.return_value.Count = self._faces
        return holder


class _Occurrences:
    def __init__(self, *occurrences: _Occurrence) -> None:
        self._items = occurrences
        self.Count = len(occurrences)

    def Item(self, index: int) -> _Occurrence:  # noqa: N802 - COM spelling
        return self._items[index - 1]


class _Assembly:
    """An assembly whose occurrence faces can be made to change, or not."""

    def __init__(self, faces: int) -> None:
        self._faces = faces
        self.updated = 0

    @property
    def Occurrences(self):  # noqa: N802 - COM spelling
        return _Occurrences(_Occurrence(self._faces))

    def UpdateAll(self) -> None:  # noqa: N802 - COM spelling
        self.updated += 1


class _Manager:
    def __init__(self, doc) -> None:
        self.doc_manager = MagicMock()
        self.doc_manager.get_active_document.return_value = doc

    @verifies_assembly_geometry
    def create_thing(self, cuts: bool = False) -> dict:
        if cuts:
            self.doc_manager.get_active_document.return_value._faces += 1
        return {"status": "created", "type": "assembly_extruded_cutout"}


class TestNoOpIsNotSuccess:
    def test_an_unchanged_assembly_is_reported_as_an_error(self):
        manager = _Manager(_Assembly(6))

        result = manager.create_thing()

        assert "error" in result
        assert "changed no geometry" in result["error"]
        assert result["unsupported"] is True
        assert result["faces_before"] == 6
        assert result["faces_after"] == 6

    def test_a_real_change_passes_straight_through(self):
        manager = _Manager(_Assembly(6))

        result = manager.create_thing(cuts=True)

        assert result == {"status": "created", "type": "assembly_extruded_cutout"}

    def test_the_assembly_is_updated_before_measuring(self):
        """Without UpdateAll the occurrence body can still read as it was."""
        doc = _Assembly(6)
        manager = _Manager(doc)

        manager.create_thing()

        assert doc.updated == 1

    def test_an_error_result_is_left_alone(self):
        class Failing(_Manager):
            @verifies_assembly_geometry
            def create_thing(self, cuts: bool = False) -> dict:
                return {"error": "COM said no"}

        assert Failing(_Assembly(6)).create_thing() == {"error": "COM said no"}

    def test_a_mocked_document_is_never_second_guessed(self):
        """Unit tests elsewhere drive these with MagicMock; stay inert there."""
        manager = _Manager(MagicMock())

        assert manager.create_thing() == {
            "status": "created",
            "type": "assembly_extruded_cutout",
        }


class TestOccurrenceFaceCount:
    def test_sums_every_occurrence(self):
        doc = MagicMock()
        doc.Occurrences = _Occurrences(_Occurrence(6), _Occurrence(4))

        assert occurrence_face_count(doc) == 10

    def test_an_unreadable_occurrence_yields_none(self):
        """A partial sum must never pass for a measurement."""
        doc = MagicMock()
        broken = MagicMock()
        type(broken).Body = property(lambda self: (_ for _ in ()).throw(Exception("COM error")))
        broken.OccurrenceDocument.Models.Count = 1
        broken.OccurrenceDocument.Models.Item.return_value.Body.Faces.return_value.Count = (
            MagicMock()
        )
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.side_effect = [_Occurrence(6), broken]
        doc.Occurrences = occurrences

        assert occurrence_face_count(doc) is None

    def test_a_document_with_no_occurrences_reads_zero(self):
        doc = MagicMock()
        doc.Occurrences = _Occurrences()

        assert occurrence_face_count(doc) == 0

    def test_a_document_that_is_not_an_assembly_yields_none(self):
        """A part has no Occurrences at all, which reads as AttributeError."""
        doc = MagicMock()
        del doc.Occurrences

        assert occurrence_face_count(doc) is None

    def test_an_unreadable_count_is_not_taken_as_a_number(self):
        """A MagicMock Count must not be mistaken for a real measurement."""
        doc = MagicMock()  # Occurrences.Count is a MagicMock, not an int

        assert occurrence_face_count(doc) is None
