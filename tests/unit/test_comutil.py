"""Tests for backends/comutil.py helpers."""


class TestProfileOriginElement:
    """The element + keypoint form of a profile origin (SweptSurfaces.Add)."""

    @staticmethod
    def _profile(**counts):
        from unittest.mock import MagicMock

        profile = MagicMock()
        for collection in ("Lines2d", "Arcs2d", "Circles2d", "Ellipses2d"):
            getattr(profile, collection).Count = counts.get(collection, 0)
        return profile

    def test_circle_is_placed_by_its_centre(self):
        from solidedge_mcp.backends.comutil import profile_origin_element
        from solidedge_mcp.backends.constants import KeyPointTypeConstants

        profile = self._profile(Circles2d=1)
        element, keypoint = profile_origin_element(profile)
        assert element is profile.Circles2d.Item.return_value
        assert keypoint == KeyPointTypeConstants.igKeyPointCenter == 4
        profile.Circles2d.Item.assert_called_once_with(1)

    def test_line_is_placed_by_its_start(self):
        from solidedge_mcp.backends.comutil import profile_origin_element
        from solidedge_mcp.backends.constants import KeyPointTypeConstants

        profile = self._profile(Lines2d=4, Circles2d=1)
        element, keypoint = profile_origin_element(profile)
        assert element is profile.Lines2d.Item.return_value
        assert keypoint == KeyPointTypeConstants.igKeyPointStart

    def test_no_geometry_is_none(self):
        from solidedge_mcp.backends.comutil import profile_origin_element

        assert profile_origin_element(self._profile()) is None
