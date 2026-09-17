"""
Unit tests for FeatureManager backend methods.

Uses unittest.mock to simulate COM objects so tests run without Solid Edge.
"""

import math
from unittest.mock import MagicMock

import pytest

from solidedge_mcp.backends.constants import FaceRotateConstants


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
# DELETE FACES
# ============================================================================


class TestDeleteFaces:
    """DeleteFaces.Add takes a single FaceSet object that cannot be built here."""

    def test_unsupported(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.delete_faces([0])
        assert result["unsupported"] is True
        assert "FaceSet" in result["error"]
        assert result["face_indices"] == [0]
        model.DeleteFaces.Add.assert_not_called()

    def test_unsupported_multiple_faces(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.delete_faces([0, 2])
        assert result["unsupported"] is True
        assert result["face_indices"] == [0, 2]
        model.DeleteFaces.Add.assert_not_called()


# ============================================================================
# DELETE FACES NO HEAL
# ============================================================================


class TestDeleteFacesNoHeal:
    """DeleteFaces.AddNoHeal also takes a single FaceSet object."""

    def test_unsupported(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.delete_faces_no_heal([0])
        assert result["unsupported"] is True
        assert "FaceSet" in result["error"]
        assert result["face_indices"] == [0]
        model.DeleteFaces.AddNoHeal.assert_not_called()


# ============================================================================
# GET FEATURE INFO
# ============================================================================


def _with_features(model, names):
    """Give a mock body a Features collection, as Solid Edge exposes it."""
    from unittest.mock import MagicMock

    feats = []
    for n in names:
        f = MagicMock()
        f.Name = n
        f.Type = 462094706
        f.Suppress = False
        f.Visible = True
        feats.append(f)
    collection = MagicMock()
    collection.Count = len(feats)
    collection.Item.side_effect = lambda i: feats[i - 1]
    model.Features = collection
    return feats


class TestListFeatures:
    def test_lists_features_not_bodies(self, feature_mgr, managers):
        """Regression: this walked doc.Models, so it always returned one entry.

        doc.Models is the list of bodies. A part has one, so the resource
        reported a single "Design Model" however many features existed.
        Verified against Solid Edge 2026, where a box with a hole has two.
        """
        _, _, _, models, model, _ = managers
        _with_features(model, ["ExtrudedProtrusion_1", "ExtrudedCutout_1"])

        result = feature_mgr.list_features()
        assert result["count"] == 2
        assert [f["name"] for f in result["features"]] == [
            "ExtrudedProtrusion_1",
            "ExtrudedCutout_1",
        ]
        assert [f["index"] for f in result["features"]] == [0, 1]

    def test_no_features_collection_is_not_an_error(self, feature_mgr, managers):
        _, _, doc, _models, model, _ = managers
        del model.Features
        doc.DesignEdgebarFeatures.Count = 0

        result = feature_mgr.list_features()

        assert result == {"features": [], "count": 0}

    def test_falls_back_to_the_edgebar_when_the_ordered_tree_is_empty(self, feature_mgr, managers):
        """A part switched from ordered to synchronous empties Models.Features.

        Verified on Solid Edge 2026: the Pathfinder still shows the features
        and DesignEdgebarFeatures still lists them, so reporting none was
        wrong. Reference planes share that collection and must stay out.
        """
        _, _, doc, _models, model, _ = managers
        model.Features.Count = 0

        planes = MagicMock()
        planes.Count = 3
        plane_names = ["RefPlane_1", "RefPlane_2", "RefPlane_3"]
        planes.Item.side_effect = lambda i: MagicMock(Name=plane_names[i - 1])
        doc.RefPlanes = planes

        entries = [
            MagicMock(Name="RefPlane_1"),
            MagicMock(Name="RefPlane_2"),
            MagicMock(Name="RefPlane_3"),
            MagicMock(Name="ExtrudedProtrusion_1"),
        ]
        edgebar = MagicMock()
        edgebar.Count = len(entries)
        edgebar.Item.side_effect = lambda i: entries[i - 1]
        doc.DesignEdgebarFeatures = edgebar

        result = feature_mgr.list_features()

        assert [f["name"] for f in result["features"]] == ["ExtrudedProtrusion_1"]
        assert result["features"][0]["source"] == "edgebar"
        # It must resolve to the edgebar entry, not to a reference plane.
        found, err = feature_mgr._get_feature_by_index(0)
        assert err is None
        assert found is entries[3]


class TestGetFeatureInfo:
    def test_success(self, feature_mgr, managers):
        _, _, _, models, model, _ = managers
        _with_features(model, ["ExtrudedProtrusion_1", "ExtrudedCutout_1"])

        result = feature_mgr.get_feature_info(1)
        assert result["index"] == 1
        assert result["name"] == "ExtrudedCutout_1"
        assert result["suppressed"] is False
        assert result["visible"] is True

    def test_invalid_index(self, feature_mgr, managers):
        _, _, _, models, model, _ = managers
        _with_features(model, ["ExtrudedProtrusion_1"])
        result = feature_mgr.get_feature_info(99)
        assert "error" in result
        assert "1 feature" in result["error"]

    def test_negative_index(self, feature_mgr, managers):
        _, _, _, models, model, _ = managers
        _with_features(model, ["ExtrudedProtrusion_1"])
        result = feature_mgr.get_feature_info(-1)
        assert "error" in result


# ============================================================================
# DIMPLE
# ============================================================================


class TestCreateDimple:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        dimples = MagicMock()
        model.Dimples = dimples
        result = feature_mgr.create_dimple(0.01)
        assert result["status"] == "created"
        assert result["type"] == "dimple"
        assert result["depth"] == 0.01
        dimples.Add.assert_called_once_with(profile, 0.01, 1, 2)

    def test_reverse_direction(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        dimples = MagicMock()
        model.Dimples = dimples
        result = feature_mgr.create_dimple(0.005, "Reverse")
        assert result["status"] == "created"
        dimples.Add.assert_called_once_with(profile, 0.005, 2, 1)

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_dimple(0.01)
        assert "error" in result

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_dimple(0.01)
        assert "error" in result


# ============================================================================
# ETCH
# ============================================================================


class TestCreateEtch:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        etches = MagicMock()
        model.Etches = etches
        result = feature_mgr.create_etch()
        assert result["status"] == "created"
        assert result["type"] == "etch"
        etches.Add.assert_called_once_with(profile)

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_etch()
        assert "error" in result

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_etch()
        assert "error" in result


# ============================================================================
# RIB
# ============================================================================


class TestCreateRib:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        ribs = MagicMock()
        model.Ribs = ribs
        result = feature_mgr.create_rib(0.005)
        assert result["status"] == "created"
        assert result["type"] == "rib"
        # Ribs.Add(RibProfile, ProfileExtensionType, ThicknessType,
        # MaterialSide, ThicknessSide, Thickness)
        ribs.Add.assert_called_once_with(profile, 9, 12, 2, 3, 0.005)

    def test_symmetric(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        ribs = MagicMock()
        model.Ribs = ribs
        result = feature_mgr.create_rib(0.01, "Symmetric")
        assert result["status"] == "created"
        assert result["direction"] == "Symmetric"

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_rib(0.005)
        assert "error" in result


# ============================================================================
# LIP
# ============================================================================


class TestCreateLip:
    """Lips.Add needs body edges plus a side face and a cap face."""

    def test_unsupported(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        lips = MagicMock()
        model.Lips = lips
        result = feature_mgr.create_lip(0.003)
        assert result["unsupported"] is True
        assert result["depth"] == 0.003
        assert "cap face" in result["error"]
        lips.Add.assert_not_called()


# ============================================================================
# DRAWN CUTOUT
# ============================================================================


class TestCreateDrawnCutout:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        drawn_cutouts = MagicMock()
        model.DrawnCutouts = drawn_cutouts
        result = feature_mgr.create_drawn_cutout(0.002)
        assert result["status"] == "created"
        assert result["type"] == "drawn_cutout"
        # DrawnCutouts.Add(Profile, Depth, ProfileSide, DepthSide, MaterialSide)
        drawn_cutouts.Add.assert_called_once_with(profile, 0.002, 6, 2, 3)

    def test_reverse(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        drawn_cutouts = MagicMock()
        model.DrawnCutouts = drawn_cutouts
        result = feature_mgr.create_drawn_cutout(0.002, "Reverse")
        assert result["status"] == "created"
        drawn_cutouts.Add.assert_called_once_with(profile, 0.002, 6, 1, 3)

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_drawn_cutout(0.002)
        assert "error" in result


# ============================================================================
# BEAD
# ============================================================================


class TestCreateBead:
    """Beads.Add needs a full 13-argument bead cross-section."""

    def test_unsupported(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        beads = MagicMock()
        model.Beads = beads
        result = feature_mgr.create_bead(0.003)
        assert result["unsupported"] is True
        assert result["depth"] == 0.003
        assert "cross-section" in result["error"]
        beads.Add.assert_not_called()


# ============================================================================
# LOUVER
# ============================================================================


class TestCreateLouver:
    """Louvers.Add records a louver Solid Edge 2026 never solves; no call is made."""

    def test_refuses_with_the_evidence(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        louvers = MagicMock()
        model.Louvers = louvers

        result = feature_mgr.create_louver(0.005, "Reverse", height=0.008)

        assert result["unsupported"] is True
        assert "never solves" in result["error"]
        assert result["depth"] == 0.005
        assert result["direction"] == "Reverse"
        assert result["height"] == 0.008
        louvers.Add.assert_not_called()


# ============================================================================
# GUSSET
# ============================================================================


class TestCreateGusset:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        gussets = MagicMock()
        model.Gussets = gussets
        result = feature_mgr.create_gusset(0.002)
        assert result["status"] == "created"
        assert result["type"] == "gusset"
        # Gussets has no Add; AddByProfile is the real one.
        gussets.Add.assert_not_called()
        gussets.AddByProfile.assert_called_once()
        args = gussets.AddByProfile.call_args.args
        assert args[0] is profile
        assert args[1] == 2  # igRight
        assert args[3] == 0.002  # gusset width

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_gusset(0.002)
        assert "error" in result


# ============================================================================
# THREAD
# ============================================================================


# ============================================================================
# SLOT
# ============================================================================


class TestCreateSlot:
    """Slots.Add needs 22 arguments including KeyPointOrTangentFace objects."""

    def test_unsupported(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        slots = MagicMock()
        model.Slots = slots
        result = feature_mgr.create_slot(0.01)
        assert result["unsupported"] is True
        assert result["depth"] == 0.01
        assert "KeyPointOrTangentFace" in result["error"]
        slots.Add.assert_not_called()

    def test_unsupported_reverse(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        slots = MagicMock()
        model.Slots = slots
        result = feature_mgr.create_slot(0.01, "Reverse")
        assert result["unsupported"] is True
        assert result["direction"] == "Reverse"
        slots.Add.assert_not_called()


# ============================================================================
# SPLIT
# ============================================================================


class TestCreateSplit:
    """Splits.Add(1, [Body], 1, [RefPlane], multiple-bodies options) splits the box."""

    def test_splits_the_body_with_a_reference_plane(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        doc.RefPlanes.Count = 4
        plane = doc.RefPlanes.Item.return_value
        model.Splits.Count = 1

        result = feature_mgr.create_split(plane_index=4)

        assert result["status"] == "created"
        assert result["plane_index"] == 4
        model.Splits.Add.assert_called_once_with(1, [model.Body], 1, [plane], 0, 0)
        doc.RefPlanes.Item.assert_called_once_with(4)

    def test_a_plane_index_out_of_range_is_refused(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        doc.RefPlanes.Count = 3

        result = feature_mgr.create_split(plane_index=4)

        assert "Invalid plane index" in result["error"]
        model.Splits.Add.assert_not_called()

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, _, models, model, _ = managers
        models.Count = 0
        assert "error" in feature_mgr.create_split()
        model.Splits.Add.assert_not_called()


# ============================================================================
# DO IDLE
# ============================================================================


class TestDoIdle:
    def test_success(self):
        from solidedge_mcp.backends.connection import SolidEdgeConnection

        cm = SolidEdgeConnection()
        cm._is_connected = True
        cm.application = MagicMock()

        result = cm.do_idle()
        assert result["status"] == "success"
        cm.application.DoIdle.assert_called_once()

    def test_not_connected(self):
        from solidedge_mcp.backends.connection import SolidEdgeConnection

        cm = SolidEdgeConnection()
        cm._is_connected = False
        cm.application = None

        result = cm.do_idle()
        assert "error" in result


# ============================================================================
# SKETCH FILLET
# ============================================================================


def _two_meeting_lines(profile):
    """Two lines sharing the corner at (0, 0), as Solid Edge reports them."""
    from unittest.mock import MagicMock

    line1, line2 = MagicMock(), MagicMock()
    line1.GetStartPoint.return_value = (0.0, 0.0)
    line1.GetEndPoint.return_value = (0.08, 0.0)
    line2.GetStartPoint.return_value = (0.0, 0.0)
    line2.GetEndPoint.return_value = (0.0, 0.05)
    lines = MagicMock()
    lines.Count = 2
    lines.Item.side_effect = lambda i: [None, line1, line2][i]
    profile.Lines2d = lines
    return line1, line2


class TestSketchFillet:
    def test_success(self):
        """Arcs2d.AddAsFillet, not AddByFillet, which does not exist.

        The old loop swallowed the failure per pair and returned "created"
        with a count of zero, so a fillet never appeared and nothing said so.
        """
        from solidedge_mcp.backends.sketching import SketchManager

        dm = MagicMock()
        sm = SketchManager(dm)
        profile = MagicMock()
        line1, line2 = _two_meeting_lines(profile)
        sm.active_profile = profile

        result = sm.sketch_fillet(0.005)
        assert result["status"] == "created"
        assert result["fillet_count"] == 1
        # The vertex itself is accepted for a fillet.
        profile.Arcs2d.AddAsFillet.assert_called_once_with(line1, line2, 0.005, 0.0, 0.0)
        profile.Arcs2d.AddByFillet.assert_not_called()

    def test_reports_when_nothing_could_be_filleted(self):
        from solidedge_mcp.backends.sketching import SketchManager

        dm = MagicMock()
        sm = SketchManager(dm)
        profile = MagicMock()
        _two_meeting_lines(profile)
        profile.Arcs2d.AddAsFillet.side_effect = Exception("radius too large")
        sm.active_profile = profile

        result = sm.sketch_fillet(0.005)
        assert "error" in result
        assert result["failures"]

    def test_rejects_a_non_positive_radius(self):
        from solidedge_mcp.backends.sketching import SketchManager

        dm = MagicMock()
        sm = SketchManager(dm)
        profile = MagicMock()
        _two_meeting_lines(profile)
        sm.active_profile = profile

        assert "must be positive" in sm.sketch_fillet(0.0)["error"]


class TestSketchChamfer:
    def test_success(self):
        """Lines2d.AddAsChamfer wants a point nudged inside the corner.

        The vertex itself gives E_INVALIDARG, unlike the fillet. Verified
        against Solid Edge 2026.
        """
        from solidedge_mcp.backends.sketching import SketchManager

        dm = MagicMock()
        sm = SketchManager(dm)
        profile = MagicMock()
        line1, line2 = _two_meeting_lines(profile)
        sm.active_profile = profile

        result = sm.sketch_chamfer(0.005)
        assert result["status"] == "created"
        assert result["chamfer_count"] == 1
        args = profile.Lines2d.AddAsChamfer.call_args[0]
        assert args[0] is line1 and args[1] is line2
        assert (args[2], args[3]) != (0.0, 0.0)  # nudged off the vertex
        assert args[4:] == (0.005, 0.005)
        profile.Lines2d.AddByChamfer.assert_not_called()


class TestSketchMirror:
    """Line2d has GetStartPoint, not StartPoint.X.

    Reading the property that does not exist raised inside a bare except, so
    every element was skipped and the result still said "created" with a count
    of zero.
    """

    def _profile(self, lines=(), circles=(), arcs=()):
        profile = MagicMock()
        for name, items in (("Lines2d", lines), ("Circles2d", circles), ("Arcs2d", arcs)):
            collection = MagicMock()
            collection.Count = len(items)
            collection.Item.side_effect = lambda i, items=items: items[i - 1]
            setattr(profile, name, collection)
        return profile

    def _line(self, x1, y1, x2, y2):
        line = MagicMock()
        line.GetStartPoint.return_value = (x1, y1)
        line.GetEndPoint.return_value = (x2, y2)
        del line.StartPoint
        del line.EndPoint
        return line

    def _manager(self, profile):
        from solidedge_mcp.backends.sketching import SketchManager

        manager = SketchManager(MagicMock())
        manager.active_profile = profile
        return manager

    def test_mirror_x_flips_y(self):
        profile = self._profile(lines=[self._line(0.01, 0.02, 0.03, 0.04)])
        manager = self._manager(profile)

        result = manager.sketch_mirror("X")

        assert result["status"] == "created"
        assert result["mirrored_elements"] == 1
        profile.Lines2d.AddBy2Points.assert_called_once_with(0.01, -0.02, 0.03, -0.04)

    def test_mirror_y_flips_x(self):
        profile = self._profile(lines=[self._line(0.01, 0.02, 0.03, 0.04)])
        manager = self._manager(profile)

        manager.sketch_mirror("Y")

        profile.Lines2d.AddBy2Points.assert_called_once_with(-0.01, 0.02, -0.03, 0.04)

    def test_mirrors_a_circle(self):
        circle = MagicMock()
        circle.GetCenterPoint.return_value = (0.05, 0.02)
        circle.Radius = 0.01
        profile = self._profile(circles=[circle])
        manager = self._manager(profile)

        result = manager.sketch_mirror("X")

        assert result["mirrored_elements"] == 1
        profile.Circles2d.AddByCenterRadius.assert_called_once_with(0.05, -0.02, 0.01)

    def test_mirrors_an_arc_with_its_ends_swapped(self):
        """Mirroring reverses the sweep, so start and end change places."""
        arc = MagicMock()
        arc.GetCenterPoint.return_value = (0.0, 0.0)
        arc.GetStartPoint.return_value = (0.01, 0.0)
        arc.GetEndPoint.return_value = (0.0, 0.01)
        profile = self._profile(arcs=[arc])
        manager = self._manager(profile)

        result = manager.sketch_mirror("X")

        assert result["mirrored_elements"] == 1
        profile.Arcs2d.AddByCenterStartEnd.assert_called_once_with(0.0, 0.0, 0.0, -0.01, 0.01, 0.0)

    def test_an_empty_sketch_is_an_error_not_a_zero_count(self):
        manager = self._manager(self._profile())

        result = manager.sketch_mirror("X")

        assert "error" in result
        assert "Nothing was mirrored" in result["error"]

    def test_invalid_axis(self):
        manager = self._manager(self._profile(lines=[self._line(0, 0, 1, 1)]))

        assert "error" in manager.sketch_mirror("Z")

    def test_no_sketch(self):
        from solidedge_mcp.backends.sketching import SketchManager

        assert "error" in SketchManager(MagicMock()).sketch_mirror()


# ============================================================================
# FACE ROTATE BY EDGE
# ============================================================================


class TestFaceRotateByEdge:
    def _make_fm_with_body(self):
        from solidedge_mcp.backends.features import FeatureManager

        dm = MagicMock()
        sm = MagicMock()
        fm = FeatureManager(dm, sm)

        doc = MagicMock()
        dm.get_active_document.return_value = doc
        model = MagicMock()
        doc.Models.Count = 1
        doc.Models.Item.return_value = model
        body = MagicMock()
        model.Body = body

        edge = MagicMock()
        face = MagicMock()
        face_edges = MagicMock()
        face_edges.Count = 4
        face_edges.Item.return_value = edge
        face.Edges = face_edges

        faces = MagicMock()
        faces.Count = 6
        faces.Item.return_value = face
        body.Faces.return_value = faces

        return fm, model

    def test_success(self):
        fm, model = self._make_fm_with_body()
        result = fm.create_face_rotate_by_edge(1, 0, 5.0)
        assert result["status"] == "created"
        assert result["type"] == "face_rotate"
        assert result["method"] == "by_edge"
        # FaceRotateConstants: ByGeometry=2, RecreateBlends=6, AxisEnd=4. The
        # literals 1, 1 and 2 this replaced raised E_INVALIDARG on Solid Edge 2026.
        args = model.FaceRotates.Add.call_args.args
        assert args[1:3] == (
            FaceRotateConstants.igFaceRotateByGeometry,
            FaceRotateConstants.igFaceRotateRecreateBlends,
        )
        assert args[3] is None and args[4] is None
        assert args[6] == FaceRotateConstants.igFaceRotateAxisEnd
        assert args[7] == pytest.approx(math.radians(5.0))

    def test_invalid_face(self):
        fm, model = self._make_fm_with_body()
        result = fm.create_face_rotate_by_edge(10, 0, 5.0)
        assert "error" in result

    def test_no_features(self):
        from solidedge_mcp.backends.features import FeatureManager

        dm = MagicMock()
        sm = MagicMock()
        fm = FeatureManager(dm, sm)
        doc = MagicMock()
        dm.get_active_document.return_value = doc
        doc.Models.Count = 0
        result = fm.create_face_rotate_by_edge(0, 0, 5.0)
        assert "error" in result


# ============================================================================
# FACE ROTATE BY POINTS
# ============================================================================


class TestFaceRotateByPoints:
    def test_success(self):
        from solidedge_mcp.backends.features import FeatureManager

        dm = MagicMock()
        sm = MagicMock()
        fm = FeatureManager(dm, sm)

        doc = MagicMock()
        dm.get_active_document.return_value = doc
        model = MagicMock()
        doc.Models.Count = 1
        doc.Models.Item.return_value = model
        body = MagicMock()
        model.Body = body

        face = MagicMock()
        v1, v2 = MagicMock(), MagicMock()
        vertices = MagicMock()
        vertices.Count = 4
        vertices.Item.side_effect = lambda i: [None, v1, v2, MagicMock(), MagicMock()][i]
        face.Vertices = vertices

        faces = MagicMock()
        faces.Count = 6
        faces.Item.return_value = face
        body.Faces.return_value = faces

        result = fm.create_face_rotate_by_points(0, 0, 1, 5.0)
        assert result["status"] == "created"
        assert result["method"] == "by_points"
        # ByPoints=1, RecreateBlends=6, the two vertices, no axis object, None=0.
        model.FaceRotates.Add.assert_called_once_with(
            face,
            FaceRotateConstants.igFaceRotateByPoints,
            FaceRotateConstants.igFaceRotateRecreateBlends,
            v1,
            v2,
            None,
            FaceRotateConstants.igFaceRotateNone,
            pytest.approx(math.radians(5.0)),
        )

    def test_no_features(self):
        from solidedge_mcp.backends.features import FeatureManager

        dm = MagicMock()
        sm = MagicMock()
        fm = FeatureManager(dm, sm)
        doc = MagicMock()
        dm.get_active_document.return_value = doc
        doc.Models.Count = 0
        result = fm.create_face_rotate_by_points(0, 0, 1, 5.0)
        assert "error" in result


# ============================================================================
# DRAFT ANGLE
# ============================================================================


class TestDraftAngle:
    def test_success(self):
        from solidedge_mcp.backends.features import FeatureManager

        dm = MagicMock()
        sm = MagicMock()
        fm = FeatureManager(dm, sm)

        doc = MagicMock()
        dm.get_active_document.return_value = doc
        model = MagicMock()
        doc.Models.Count = 1
        doc.Models.Item.return_value = model
        body = MagicMock()
        model.Body = body

        face = MagicMock()
        faces = MagicMock()
        faces.Count = 6
        faces.Item.return_value = face
        body.Faces.return_value = faces

        ref_plane = MagicMock()
        doc.RefPlanes.Item.return_value = ref_plane

        result = fm.create_draft_angle(0, 3.0, plane_index=1)
        assert result["status"] == "created"
        assert result["type"] == "draft_angle"
        assert result["angle_degrees"] == 3.0
        model.Drafts.Add.assert_called_once()

    def test_no_features(self):
        from solidedge_mcp.backends.features import FeatureManager

        dm = MagicMock()
        sm = MagicMock()
        fm = FeatureManager(dm, sm)
        doc = MagicMock()
        dm.get_active_document.return_value = doc
        doc.Models.Count = 0
        result = fm.create_draft_angle(0, 3.0)
        assert "error" in result


# ============================================================================
# CONVERT FEATURE TYPE
# ============================================================================


class TestConvertFeatureType:
    def test_convert_to_cutout(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        feat = MagicMock()
        feat.Name = "Protrusion_1"
        cutout = MagicMock()
        cutout.Name = "Cutout_1"
        feat.ConvertToCutout.return_value = cutout

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = feature_mgr.convert_feature_type("Protrusion_1", "cutout")
        assert result["status"] == "converted"
        assert result["target_type"] == "cutout"
        assert result["new_name"] == "Cutout_1"
        feat.ConvertToCutout.assert_called_once()

    def test_convert_to_protrusion(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        feat = MagicMock()
        feat.Name = "Cutout_1"
        protrusion = MagicMock()
        protrusion.Name = "Protrusion_1"
        feat.ConvertToProtrusion.return_value = protrusion

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = feature_mgr.convert_feature_type("Cutout_1", "protrusion")
        assert result["status"] == "converted"
        assert result["target_type"] == "protrusion"
        assert result["new_name"] == "Protrusion_1"

    def test_feature_not_found(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        feat = MagicMock()
        feat.Name = "Other_1"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = feature_mgr.convert_feature_type("NonExistent", "cutout")
        assert "error" in result
        assert "not found" in result["error"]

    def test_invalid_target_type(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        feat = MagicMock()
        feat.Name = "Protrusion_1"

        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = feature_mgr.convert_feature_type("Protrusion_1", "invalid")
        assert "error" in result
        assert "Invalid target_type" in result["error"]


# ============================================================================
# EMBOSS
# ============================================================================


class TestCreateEmboss:
    def test_basic_emboss(self, feature_mgr, managers):
        _, _, doc, models, model, _ = managers
        emboss_features = MagicMock()
        model.EmbossFeatures = emboss_features

        result = feature_mgr.create_emboss([0])
        assert result["status"] == "created"
        assert result["type"] == "emboss"
        assert result["face_count"] == 1
        emboss_features.Add.assert_called_once()

    def test_multiple_faces(self, feature_mgr, managers):
        _, _, doc, models, model, _ = managers
        body = model.Body
        MagicMock()
        faces = body.Faces.return_value
        faces.Count = 3
        faces.Item.side_effect = lambda i: MagicMock()

        emboss_features = MagicMock()
        model.EmbossFeatures = emboss_features

        result = feature_mgr.create_emboss(
            [0, 1, 2],
            clearance=0.002,
            thickness=0.001,
            thicken=True,
        )
        assert result["status"] == "created"
        assert result["face_count"] == 3
        assert result["clearance"] == 0.002
        assert result["thickness"] == 0.001
        assert result["thicken"] is True

    def test_no_base_feature(self, feature_mgr, managers):
        _, _, doc, models, model, _ = managers
        models.Count = 0

        result = feature_mgr.create_emboss([0])
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_invalid_face_index(self, feature_mgr, managers):
        _, _, doc, models, model, _ = managers

        result = feature_mgr.create_emboss([99])
        assert "error" in result
        assert "Invalid face index" in result["error"]

    def test_empty_face_indices(self, feature_mgr, managers):
        _, _, doc, models, model, _ = managers

        result = feature_mgr.create_emboss([])
        assert "error" in result
        assert "at least one" in result["error"]


# ============================================================================
# BATCH 9: DIMPLE EX
# ============================================================================


class TestCreateDimpleEx:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        result = feature_mgr.create_dimple_ex(0.003)
        assert result["status"] == "created"
        assert result["type"] == "dimple_ex"
        assert result["depth"] == 0.003
        model.Dimples.AddEx.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_dimple_ex(0.003)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_dimple_ex(0.003)
        assert "error" in result
        assert "No base feature" in result["error"]


# ============================================================================
# BATCH 9: THREAD EX
# ============================================================================


# ============================================================================
# BATCH 9: SLOT EX
# ============================================================================


class TestCreateSlotEx:
    """Slots.AddEx needs 18 arguments including a KeyPointOrTangentFace."""

    def test_unsupported(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_slot_ex(0.005, 0.01)
        assert result["unsupported"] is True
        assert result["width"] == 0.005
        assert result["depth"] == 0.01
        assert "KeyPointOrTangentFace" in result["error"]
        model.Slots.AddEx.assert_not_called()


# ============================================================================
# BATCH 9: SLOT SYNC
# ============================================================================


class TestCreateSlotSync:
    """Slots.AddSync needs the same 18 arguments as Slots.AddEx."""

    def test_unsupported(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_slot_sync(0.005, 0.01)
        assert result["unsupported"] is True
        assert result["width"] == 0.005
        assert result["depth"] == 0.01
        assert "KeyPointOrTangentFace" in result["error"]
        model.Slots.AddSync.assert_not_called()


# ============================================================================
# BATCH 9: DRAWN CUTOUT EX
# ============================================================================


class TestCreateDrawnCutoutEx:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, _, _, model, _ = managers
        result = feature_mgr.create_drawn_cutout_ex(0.005)
        assert result["status"] == "created"
        assert result["type"] == "drawn_cutout_ex"
        assert result["depth"] == 0.005
        assert result["profile_count"] == 1
        # AddEx(NumberOfProfiles, ProfileArray, Depth, ProfileSide, DepthSide,
        # MaterialSide)
        args = model.DrawnCutouts.AddEx.call_args.args
        assert len(args) == 6
        assert args[0] == 1
        assert (args[2], args[3], args[4], args[5]) == (0.005, 6, 2, 3)
        sketch_mgr.clear_accumulated_profiles.assert_called()

    def test_reverse(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_drawn_cutout_ex(0.005, "Reverse")
        assert result["status"] == "created"
        args = model.DrawnCutouts.AddEx.call_args.args
        assert args[4] == 1  # DepthSide -> seDrawnCutoutDepthLeft

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_drawn_cutout_ex(0.005)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_drawn_cutout_ex(0.005)
        assert "error" in result
        assert "No base feature" in result["error"]


# ============================================================================
# BATCH 9: LOUVER SYNC
# ============================================================================


class TestCreateLouverSync:
    """Louvers.AddSync is face-based with origin/orientation arrays."""

    def test_unsupported(self, feature_mgr, managers):
        _, _, _, _, model, _ = managers
        result = feature_mgr.create_louver_sync(0.003)
        assert result["unsupported"] is True
        assert result["depth"] == 0.003
        assert "orientation" in result["error"]
        model.Louvers.AddSync.assert_not_called()


# ============================================================================
# BATCH 9: THICKEN SYNC
# ============================================================================


class TestCreateThickenSync:
    """The same AddThickenFeature call, in either mode."""

    def test_delegates_to_the_basic_thicken(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        doc.Constructions.Count = 1
        faces = doc.Constructions.Item.return_value.Body.Faces.return_value
        faces.Count = 2
        face_objects = [MagicMock(), MagicMock()]
        faces.Item.side_effect = lambda i: face_objects[i - 1]

        result = feature_mgr.create_thicken_sync(0.003, direction="Reverse")

        assert result["status"] == "created"
        assert result["type"] == "thicken_sync"
        assert result["direction"] == "Reverse"
        models.AddThickenFeature.assert_called_once_with(1, 0.003, 2, face_objects)


# ============================================================================
# BATCH 9: MIRROR SYNC EX
# ============================================================================


def _setup_feature_lookup(doc, feature_name="Extrude1"):
    """Helper to set up DesignEdgebarFeatures mock for feature lookup."""
    feat = MagicMock()
    feat.Name = feature_name
    features = MagicMock()
    features.Count = 1
    features.Item.return_value = feat
    doc.DesignEdgebarFeatures = features

    ref_planes = MagicMock()
    ref_planes.Count = 3
    ref_planes.Item.return_value = MagicMock()
    doc.RefPlanes = ref_planes

    return feat


class TestCreateMirrorSyncEx:
    """AddSyncEx needs an [in,out] SAFEARRAY that late binding cannot pass."""

    def test_unsupported(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        _setup_feature_lookup(doc, "Extrude1")
        mc = MagicMock()
        model.MirrorCopies = mc

        result = feature_mgr.create_mirror_sync_ex("Extrude1", 1)
        assert result["unsupported"] is True
        assert result["feature"] == "Extrude1"
        assert result["mirror_plane"] == 1
        assert "AddSyncEx" in result["error"]
        mc.AddSyncEx.assert_not_called()


# ============================================================================
# BATCH 9: PATTERN RECTANGULAR EX
# ============================================================================


class TestCreatePatternRectangularEx:
    def test_success(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        _setup_feature_lookup(doc, "Hole1")
        pattern = MagicMock()
        pattern.Name = "Pattern1"
        model.Patterns.AddByRectangularEx.return_value = pattern

        result = feature_mgr.create_pattern_rectangular_ex("Hole1", 3, 2, 0.01, 0.02)
        assert result["status"] == "created"
        assert result["type"] == "pattern_rectangular_ex"
        assert result["x_count"] == 3
        assert result["y_count"] == 2
        assert result["plane_index"] == 1
        # AddByRectangularEx(NumberOfFeatures, FeatureArray, ReferencePlane,
        # XDirectionCount, YDirectionCount, XDirectionSpacing,
        # YDirectionSpacing, RectangleAngle, PatternMethod, ReferenceIndex,
        # PatternType)
        args = model.Patterns.AddByRectangularEx.call_args.args
        assert len(args) == 11
        assert args[0] == 1
        assert args[2] is doc.RefPlanes.Item.return_value
        assert args[3:] == (3, 2, 0.01, 0.02, 0.0, 2, 0, 0)

    def test_invalid_plane_index(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        _setup_feature_lookup(doc, "Hole1")

        result = feature_mgr.create_pattern_rectangular_ex(
            "Hole1", 3, 2, 0.01, 0.02, plane_index=99
        )
        assert "error" in result
        assert "Invalid plane_index" in result["error"]
        model.Patterns.AddByRectangularEx.assert_not_called()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_pattern_rectangular_ex("Hole1", 3, 2, 0.01, 0.02)
        assert "error" in result
        assert "No base feature" in result["error"]

    def test_feature_not_found(self, feature_mgr, managers):
        _, _, doc, _, _, _ = managers
        feat = MagicMock()
        feat.Name = "OtherFeature"
        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = feature_mgr.create_pattern_rectangular_ex("Missing", 3, 2, 0.01, 0.02)
        assert "error" in result
        assert "not found" in result["error"]


# ============================================================================
# BATCH 9: PATTERN CIRCULAR EX
# ============================================================================


class TestCreatePatternCircularEx:
    """AddByCircularEx needs a reference plane and an AxisPoint array."""

    def test_unsupported(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        _setup_feature_lookup(doc, "Hole1")

        result = feature_mgr.create_pattern_circular_ex("Hole1", 6, 360.0, 0)
        assert result["unsupported"] is True
        assert result["count"] == 6
        assert result["angle"] == 360.0
        assert "AxisPoint" in result["error"]
        model.Patterns.AddByCircularEx.assert_not_called()


# ============================================================================
# BATCH 9: PATTERN DUPLICATE
# ============================================================================


class TestCreatePatternDuplicate:
    """AddDuplicate needs a FromReference and instance references."""

    def test_unsupported(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        _setup_feature_lookup(doc, "Extrude1")

        result = feature_mgr.create_pattern_duplicate("Extrude1")
        assert result["unsupported"] is True
        assert result["feature"] == "Extrude1"
        assert "FromReference" in result["error"]
        model.Patterns.AddDuplicate.assert_not_called()


# ============================================================================
# BATCH 9: PATTERN BY FILL
# ============================================================================


class TestCreatePatternByFill:
    """AddByFill needs region profiles, not a body face."""

    def test_unsupported(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        _setup_feature_lookup(doc, "Hole1")

        result = feature_mgr.create_pattern_by_fill("Hole1", 0, 0.01, 0.01)
        assert result["unsupported"] is True
        assert result["fill_region_face_index"] == 0
        assert "region profiles" in result["error"]
        model.Patterns.AddByFill.assert_not_called()


# ============================================================================
# BATCH 9: PATTERN BY TABLE
# ============================================================================


class TestCreatePatternByTable:
    """AddPatternByTable is Excel- and KeyPoint-driven."""

    def test_unsupported(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        _setup_feature_lookup(doc, "Hole1")

        result = feature_mgr.create_pattern_by_table("Hole1", [0.01, 0.02], [0.01, 0.02])
        assert result["unsupported"] is True
        assert result["x_offsets"] == [0.01, 0.02]
        assert "Excel" in result["error"]
        model.Patterns.AddPatternByTable.assert_not_called()


# ============================================================================
# PATTERN BY TABLE SYNC
# ============================================================================


class TestCreatePatternByTableSync:
    """AddPatternByTableSync is Excel- and KeyPoint-driven."""

    def test_unsupported(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        _setup_feature_lookup(doc, "Hole1")

        result = feature_mgr.create_pattern_by_table_sync("Hole1", [0.01, 0.02], [0.01, 0.02])
        assert result["unsupported"] is True
        assert result["y_offsets"] == [0.01, 0.02]
        assert "Excel" in result["error"]
        model.Patterns.AddPatternByTableSync.assert_not_called()


# ============================================================================
# PATTERN BY FILL EX
# ============================================================================


class TestCreatePatternByFillEx:
    """AddByFillEx needs region profiles, not a body face."""

    def test_unsupported(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        _setup_feature_lookup(doc, "Hole1")

        result = feature_mgr.create_pattern_by_fill_ex("Hole1", 0, 0.01, 0.01)
        assert result["unsupported"] is True
        assert result["stagger_offset"] == 0.0
        assert "region profiles" in result["error"]
        model.Patterns.AddByFillEx.assert_not_called()


# ============================================================================
# PATTERN BY CURVE EX
# ============================================================================


class TestCreatePatternByCurveEx:
    """AddByCurveEx needs 23 arguments including an anchor KeyPoint."""

    def test_unsupported(self, feature_mgr, managers):
        _, _, doc, _, model, _ = managers
        _setup_feature_lookup(doc, "Hole1")

        result = feature_mgr.create_pattern_by_curve_ex("Hole1", 0, 5, 0.01)
        assert result["unsupported"] is True
        assert result["count"] == 5
        assert result["spacing"] == 0.01
        assert "KeyPoint" in result["error"]
        model.Patterns.AddByCurveEx.assert_not_called()


# ============================================================================
# SAVE AS MIRROR PART
# ============================================================================


class TestSaveAsMirrorPart:
    def test_success(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        plane = MagicMock()
        ref_planes = MagicMock()
        ref_planes.Count = 3
        ref_planes.Item.return_value = plane
        doc.RefPlanes = ref_planes

        result = feature_mgr.save_as_mirror_part("C:/temp/mirror.par", 3, True)
        assert result["status"] == "saved"
        assert result["type"] == "mirror_part"
        assert result["path"] == "C:/temp/mirror.par"
        assert result["linked"] is True
        models.SaveAsMirrorPart.assert_called_once_with("C:/temp/mirror.par", plane, True)

    def test_no_model(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        models.Count = 0

        result = feature_mgr.save_as_mirror_part("C:/temp/mirror.par")
        assert "error" in result
        assert "No model" in result["error"]

    def test_invalid_plane_index(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        ref_planes = MagicMock()
        ref_planes.Count = 3
        doc.RefPlanes = ref_planes

        result = feature_mgr.save_as_mirror_part("C:/temp/mirror.par", 5)
        assert "error" in result
        assert "Invalid mirror_plane_index" in result["error"]

    def test_unlinked(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        plane = MagicMock()
        ref_planes = MagicMock()
        ref_planes.Count = 3
        ref_planes.Item.return_value = plane
        doc.RefPlanes = ref_planes

        result = feature_mgr.save_as_mirror_part("C:/temp/mirror.par", 2, False)
        assert result["status"] == "saved"
        assert result["linked"] is False
        models.SaveAsMirrorPart.assert_called_once_with("C:/temp/mirror.par", plane, False)


# ============================================================================
# USER DEFINED PATTERN
# ============================================================================


class TestCreateUserDefinedPattern:
    def test_success(self, feature_mgr, managers):
        _, sketch_mgr, doc, _, model, _ = managers
        seed = MagicMock()
        seed.Name = "Extrude 1"
        features = MagicMock()
        features.Count = 1
        features.Item.return_value = seed
        doc.DesignEdgebarFeatures = features

        p1, p2 = MagicMock(), MagicMock()
        sketch_mgr.get_accumulated_profiles.return_value = [p1, p2]

        udp = MagicMock()
        model.UserDefinedPatterns = udp

        result = feature_mgr.create_user_defined_pattern("Extrude 1")
        assert result["status"] == "created"
        assert result["type"] == "user_defined_pattern"
        assert result["seed_feature"] == "Extrude 1"
        assert result["num_occurrences"] == 2
        udp.AddByProfiles.assert_called_once()
        sketch_mgr.clear_accumulated_profiles.assert_called_once()

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0

        result = feature_mgr.create_user_defined_pattern("Extrude 1")
        assert "error" in result
        assert "No model" in result["error"]

    def test_feature_not_found(self, feature_mgr, managers):
        _, sketch_mgr, doc, _, model, _ = managers
        feat = MagicMock()
        feat.Name = "Other Feature"
        features = MagicMock()
        features.Count = 1
        features.Item.return_value = feat
        doc.DesignEdgebarFeatures = features

        result = feature_mgr.create_user_defined_pattern("NonExistent")
        assert "error" in result
        assert "not found" in result["error"]

    def test_no_profiles(self, feature_mgr, managers):
        _, sketch_mgr, doc, _, model, _ = managers
        seed = MagicMock()
        seed.Name = "Extrude 1"
        features = MagicMock()
        features.Count = 1
        features.Item.return_value = seed
        doc.DesignEdgebarFeatures = features

        sketch_mgr.get_accumulated_profiles.return_value = []

        result = feature_mgr.create_user_defined_pattern("Extrude 1")
        assert "error" in result
        assert "at least 1 profile" in result["error"]


# ============================================================================
# SLOT MULTI BODY
# ============================================================================


class TestCreateSlotMultiBody:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        model.Slots = MagicMock()

        result = feature_mgr.create_slot_multi_body(0.005, 0.01)
        assert result["status"] == "created"
        assert result["type"] == "slot_multi_body"
        assert result["width"] == 0.005
        assert result["depth"] == 0.01
        model.Slots.AddMultiBody.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_slot_multi_body(0.005, 0.01)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_no_model(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        models.Count = 0
        result = feature_mgr.create_slot_multi_body(0.005, 0.01)
        assert "error" in result
        assert "No base feature" in result["error"]


# ============================================================================
# SLOT SYNC MULTI BODY
# ============================================================================


class TestCreateSlotSyncMultiBody:
    def test_success(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        model.Slots = MagicMock()

        result = feature_mgr.create_slot_sync_multi_body(0.005, 0.01)
        assert result["status"] == "created"
        assert result["type"] == "slot_sync_multi_body"
        assert result["width"] == 0.005
        assert result["depth"] == 0.01
        model.Slots.AddSyncMultiBody.assert_called_once()

    def test_no_profile(self, feature_mgr, managers):
        _, sketch_mgr, _, _, _, _ = managers
        sketch_mgr.get_active_sketch.return_value = None
        result = feature_mgr.create_slot_sync_multi_body(0.005, 0.01)
        assert "error" in result
        assert "No active sketch" in result["error"]

    def test_reverse_direction(self, feature_mgr, managers):
        _, _, _, _, model, profile = managers
        model.Slots = MagicMock()

        result = feature_mgr.create_slot_sync_multi_body(0.005, 0.01, "Reverse")
        assert result["status"] == "created"
        assert result["direction"] == "Reverse"


# ============================================================================
# BODY CREATION
# ============================================================================


class TestAddBody:
    def test_success(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        result = feature_mgr.add_body()
        assert result["status"] == "created"
        assert result["body_type"] == "Solid"
        assert result["body_name"] == "Body"
        # Models.AddBody(igBodyType, BodyName); igPartType = 1
        models.AddBody.assert_called_once_with(1, "Body")

    def test_sheet_metal_with_name(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        result = feature_mgr.add_body("SheetMetal", "Panel")
        assert result["status"] == "created"
        # igSheetMetalType = 2
        models.AddBody.assert_called_once_with(2, "Panel")

    def test_unknown_body_type(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        result = feature_mgr.add_body("Surface")
        assert "error" in result
        assert "Unknown body_type" in result["error"]
        models.AddBody.assert_not_called()


class TestAddBodyByMesh:
    def test_unsupported(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        result = feature_mgr.add_body_by_mesh()
        assert result["unsupported"] is True
        assert "facet vertex" in result["error"]
        models.AddBodyByMeshFacets.assert_not_called()


class TestAddBodyFeature:
    def test_success(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        result = feature_mgr.add_body_feature("C:/parts/base.par")
        assert result["status"] == "created"
        assert result["import_file_name"] == "C:/parts/base.par"
        # Models.AddBodyFeature(ImportFileName)
        models.AddBodyFeature.assert_called_once_with("C:/parts/base.par")

    def test_file_name_required(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        result = feature_mgr.add_body_feature()
        assert "error" in result
        assert "import_file_name" in result["error"]
        models.AddBodyFeature.assert_not_called()


class TestAddByConstruction:
    def test_success(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        constructions = MagicMock()
        constructions.Count = 2
        construction = MagicMock()
        constructions.Item.return_value = construction
        doc.Constructions = constructions

        result = feature_mgr.add_by_construction(1)
        assert result["status"] == "created"
        assert result["construction_index"] == 1
        constructions.Item.assert_called_once_with(2)
        # Models.AddByConstruction(ConstructionSolid)
        models.AddByConstruction.assert_called_once_with(construction)

    def test_no_constructions(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        constructions = MagicMock()
        constructions.Count = 0
        doc.Constructions = constructions

        result = feature_mgr.add_by_construction()
        assert "error" in result
        assert "No construction bodies" in result["error"]
        models.AddByConstruction.assert_not_called()

    def test_invalid_index(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        constructions = MagicMock()
        constructions.Count = 1
        doc.Constructions = constructions

        result = feature_mgr.add_by_construction(7)
        assert "error" in result
        assert "Invalid construction_index" in result["error"]
        models.AddByConstruction.assert_not_called()


# ============================================================================
# THICKEN / SIMPLIFY (unsupported)
# ============================================================================


class TestThickenSurface:
    """AddThickenFeature over a construction body's faces; verified live for every side."""

    @staticmethod
    def _surface(doc, n_faces=6):
        doc.Constructions.Count = 1
        faces = doc.Constructions.Item.return_value.Body.Faces.return_value
        faces.Count = n_faces
        face_objects = [MagicMock(name=f"face{i}") for i in range(n_faces)]
        faces.Item.side_effect = lambda i: face_objects[i - 1]
        return face_objects

    @pytest.mark.parametrize("direction, side", [("Both", 3), ("Normal", 2), ("Reverse", 1)])
    def test_thickens_every_face_of_the_first_construction_body(
        self, feature_mgr, managers, direction, side
    ):
        _, _, doc, models, _, _ = managers
        face_objects = self._surface(doc)

        result = feature_mgr.thicken_surface(0.002, direction=direction)

        assert result["status"] == "created"
        assert result["faces_thickened"] == 6
        models.AddThickenFeature.assert_called_once_with(side, 0.002, 6, face_objects)
        doc.Constructions.Item.assert_called_once_with(1)

    def test_a_missing_surface_is_refused(self, feature_mgr, managers):
        _, _, doc, models, _, _ = managers
        doc.Constructions.Count = 0

        result = feature_mgr.thicken_surface(0.002)

        assert "Invalid surface index" in result["error"]
        models.AddThickenFeature.assert_not_called()

    def test_a_bad_direction_is_refused(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        assert "Invalid direction" in feature_mgr.thicken_surface(0.002, direction="Up")["error"]
        models.AddThickenFeature.assert_not_called()


class TestSimplify:
    def test_auto_simplify_unsupported(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        result = feature_mgr.auto_simplify()
        assert result["unsupported"] is True
        assert "occurrences" in result["error"]
        models.AddAutoSimplify.assert_not_called()

    def test_simplify_enclosure_unsupported(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        result = feature_mgr.simplify_enclosure()
        assert result["unsupported"] is True
        models.AddSimplifyEnclosure.assert_not_called()

    def test_simplify_duplicate_unsupported(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        result = feature_mgr.simplify_duplicate()
        assert result["unsupported"] is True
        models.AddSimplifyDuplicate.assert_not_called()

    def test_local_simplify_enclosure_unsupported(self, feature_mgr, managers):
        _, _, _, models, _, _ = managers
        result = feature_mgr.local_simplify_enclosure()
        assert result["unsupported"] is True
        assert "topology proxy" in result["error"]
        models.AddLocalSimplifyEnclosure.assert_not_called()


class TestDraftSide:
    """Drafts.Add takes igInside (4) or igOutside (5), never igLeft or igRight.

    Verified on Solid Edge 2026: 4 and 5 both work on every face of a box,
    while 1, 2 and 3 fail with a bare E_FAIL whatever plane is named. The
    face array goes as a plain list; no VARIANT wrapper is needed.
    """

    def _part(self, managers):
        _, _, doc, models, model, _ = managers
        faces = MagicMock()
        faces.Count = 6
        face = MagicMock()
        faces.Item.return_value = face
        model.Body.Faces.return_value = faces
        doc.RefPlanes.Count = 3
        return model, face

    def test_inside_is_the_default(self, feature_mgr, managers):
        model, face = self._part(managers)

        result = feature_mgr.create_draft_angle(face_index=0, angle=3.0)

        assert result["status"] == "created"
        assert result["side"] == "inside"
        args = model.Drafts.Add.call_args.args
        assert args[2] == [face]  # a plain list, not a VARIANT
        assert args[4] == 4  # igInside

    def test_outside_is_five(self, feature_mgr, managers):
        model, _face = self._part(managers)

        feature_mgr.create_draft_angle(face_index=0, angle=3.0, side="outside")

        assert model.Drafts.Add.call_args.args[4] == 5  # igOutside

    def test_the_angle_reaches_com_in_radians(self, feature_mgr, managers):
        model, _face = self._part(managers)

        feature_mgr.create_draft_angle(face_index=0, angle=90.0)

        assert model.Drafts.Add.call_args.args[3] == [pytest.approx(math.pi / 2)]

    def test_an_unknown_side_is_refused(self, feature_mgr, managers):
        model, _face = self._part(managers)

        result = feature_mgr.create_draft_angle(face_index=0, angle=3.0, side="sideways")

        assert "error" in result
        model.Drafts.Add.assert_not_called()
