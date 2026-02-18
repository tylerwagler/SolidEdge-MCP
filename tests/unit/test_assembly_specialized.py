"""
Unit tests for AssemblyManager specialized backend methods.

Tests specialized mixin: AddVirtualComponent, AddVirtualComponentPredefined,
AddVirtualComponentBIDM, GetTube, AddTube,
AddStructuralFrame, AddStructuralFrameByOrientation,
AddSplice, AddWire, AddCable, AddBundle.
Uses unittest.mock to simulate COM objects.
"""

from unittest.mock import MagicMock

import pytest


@pytest.fixture
def asm_mgr():
    """Create AssemblyManager with mocked dependencies."""
    from solidedge_mcp.backends.assembly import AssemblyManager

    dm = MagicMock()
    doc = MagicMock()
    dm.get_active_document.return_value = doc
    return AssemblyManager(dm), doc


@pytest.fixture
def asm_mgr_with_sketch():
    """Create AssemblyManager with mocked doc and sketch manager."""
    from solidedge_mcp.backends.assembly import AssemblyManager

    dm = MagicMock()
    sm = MagicMock()
    doc = MagicMock()
    dm.get_active_document.return_value = doc
    return AssemblyManager(dm, sm), doc, sm


# ============================================================================
# VIRTUAL COMPONENTS
# ============================================================================


class TestAddVirtualComponent:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        vc_occ = MagicMock()
        vc_occ.Name = "VirtualPart_1"
        vc_occs = MagicMock()
        vc_occs.Add.return_value = vc_occ
        doc.VirtualComponentOccurrences = vc_occs

        result = am.add_virtual_component("MyVirtualPart", "Part")
        assert result["status"] == "created"
        assert result["name"] == "MyVirtualPart"
        assert result["component_type"] == "Part"
        vc_occs.Add.assert_called_once_with("MyVirtualPart", 3)

    def test_assembly_type(self, asm_mgr):
        am, doc = asm_mgr
        vc_occ = MagicMock()
        vc_occs = MagicMock()
        vc_occs.Add.return_value = vc_occ
        doc.VirtualComponentOccurrences = vc_occs

        result = am.add_virtual_component("VAsm", "Assembly")
        assert result["status"] == "created"
        vc_occs.Add.assert_called_once_with("VAsm", 2)

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.add_virtual_component("test")
        assert "error" in result

    def test_com_error(self, asm_mgr):
        am, doc = asm_mgr
        doc.VirtualComponentOccurrences.Add.side_effect = Exception("fail")

        result = am.add_virtual_component("test")
        assert "error" in result


class TestAddVirtualComponentPredefined:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        vc_occ = MagicMock()
        vc_occ.Name = "PreDef_1"
        vc_occs = MagicMock()
        vc_occs.AddAsPreDefined.return_value = vc_occ
        doc.VirtualComponentOccurrences = vc_occs

        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.add_virtual_component_predefined("C:\\vc\\part.par")
        assert result["status"] == "created"
        assert result["filename"] == "C:\\vc\\part.par"
        vc_occs.AddAsPreDefined.assert_called_once_with("C:\\vc\\part.par")

    def test_file_not_found(self, asm_mgr):
        am, doc = asm_mgr
        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=False):
            result = am.add_virtual_component_predefined("C:\\missing.par")
        assert "error" in result
        assert "File not found" in result["error"]

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences
        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.add_virtual_component_predefined("C:\\vc\\part.par")
        assert "error" in result


class TestAddVirtualComponentBIDM:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        vc_occ = MagicMock()
        vc_occ.Name = "BIDM_1"
        vc_occs = MagicMock()
        vc_occs.AddBIDM.return_value = vc_occ
        doc.VirtualComponentOccurrences = vc_occs

        result = am.add_virtual_component_bidm("DOC001", "REV_A", "Part")
        assert result["status"] == "created"
        assert result["doc_number"] == "DOC001"
        assert result["revision_id"] == "REV_A"
        vc_occs.AddBIDM.assert_called_once_with("DOC001", "REV_A", 3)

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.add_virtual_component_bidm("DOC001", "REV_A")
        assert "error" in result

    def test_com_error(self, asm_mgr):
        am, doc = asm_mgr
        doc.VirtualComponentOccurrences.AddBIDM.side_effect = Exception("fail")

        result = am.add_virtual_component_bidm("DOC001", "REV_A")
        assert "error" in result


# ============================================================================
# GET TUBE
# ============================================================================


class TestGetTube:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        tube = MagicMock()
        tube.OuterDiameter = 0.025
        tube.WallThickness = 0.002
        tube.BendRadius = 0.05
        tube.Material = "Steel"
        tube.IsSolid = False
        tube.Length = 0.5

        occ = MagicMock()
        occ.GetTube.return_value = tube
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.get_tube(0)
        assert result["is_tube"] is True
        assert result["outer_diameter"] == 0.025
        assert result["wall_thickness"] == 0.002
        assert result["bend_radius"] == 0.05
        assert result["material"] == "Steel"
        assert result["length"] == 0.5

    def test_not_tube(self, asm_mgr):
        am, doc = asm_mgr
        occ = MagicMock()
        occ.GetTube.return_value = None
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = occ
        doc.Occurrences = occurrences

        result = am.get_tube(0)
        assert "error" in result
        assert "not a tube" in result["error"]

    def test_invalid_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences

        result = am.get_tube(5)
        assert "error" in result

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.get_tube(0)
        assert "error" in result


# ============================================================================
# ADD TUBE
# ============================================================================


class TestAddTube:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        seg1, seg2 = MagicMock(), MagicMock()
        occ = MagicMock()
        occ.Name = "Tube_1"
        occurrences = MagicMock()
        occurrences.Count = 3
        occurrences.Item.side_effect = lambda i: {1: seg1, 2: seg2, 3: MagicMock()}[i]
        occurrences.AddTube.return_value = occ
        doc.Occurrences = occurrences

        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.add_tube([0, 1], "C:\\tubes\\tube.par")
        assert result["status"] == "created"
        assert result["type"] == "tube"
        assert result["num_segments"] == 2
        occurrences.AddTube.assert_called_once()

    def test_file_not_found(self, asm_mgr):
        am, doc = asm_mgr
        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=False):
            result = am.add_tube([0], "C:\\missing.par")
        assert "error" in result
        assert "File not found" in result["error"]

    def test_invalid_segment_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 2
        doc.Occurrences = occurrences

        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.add_tube([0, 5], "C:\\tubes\\tube.par")
        assert "error" in result
        assert "Invalid segment index" in result["error"]

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.add_tube([0], "C:\\tubes\\tube.par")
        assert "error" in result


# ============================================================================
# STRUCTURAL FRAMES
# ============================================================================


class TestAddStructuralFrame:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        path1, path2 = MagicMock(), MagicMock()
        frame = MagicMock()
        frame.Name = "Frame_1"
        frames = MagicMock()
        frames.Add.return_value = frame
        doc.StructuralFrames = frames

        occurrences = MagicMock()
        occurrences.Count = 3
        occurrences.Item.side_effect = lambda i: {1: path1, 2: path2, 3: MagicMock()}[i]
        doc.Occurrences = occurrences

        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.add_structural_frame("C:\\frames\\beam.par", [0, 1])
        assert result["status"] == "created"
        assert result["type"] == "structural_frame"
        assert result["num_paths"] == 2
        frames.Add.assert_called_once()

    def test_file_not_found(self, asm_mgr):
        am, doc = asm_mgr
        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=False):
            result = am.add_structural_frame("C:\\missing.par", [0])
        assert "error" in result
        assert "File not found" in result["error"]

    def test_invalid_path_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences

        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.add_structural_frame("C:\\frames\\beam.par", [0, 5])
        assert "error" in result
        assert "Invalid path index" in result["error"]

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.add_structural_frame("C:\\frames\\beam.par", [0])
        assert "error" in result


class TestAddStructuralFrameByOrientation:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        path1 = MagicMock()
        frame = MagicMock()
        frame.Name = "OrientedFrame_1"
        frames = MagicMock()
        frames.AddByOrientation.return_value = frame
        doc.StructuralFrames = frames

        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = path1
        doc.Occurrences = occurrences

        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.add_structural_frame_by_orientation(
                "C:\\frames\\beam.par", "CoordSys1", [0]
            )
        assert result["status"] == "created"
        assert result["type"] == "structural_frame_oriented"
        assert result["coord_system"] == "CoordSys1"
        frames.AddByOrientation.assert_called_once()

    def test_file_not_found(self, asm_mgr):
        am, doc = asm_mgr
        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=False):
            result = am.add_structural_frame_by_orientation(
                "C:\\missing.par", "CS1", [0]
            )
        assert "error" in result

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        import unittest.mock

        with unittest.mock.patch("os.path.exists", return_value=True):
            result = am.add_structural_frame_by_orientation(
                "C:\\frames\\beam.par", "CS1", [0]
            )
        assert "error" in result


# ============================================================================
# SPLICES
# ============================================================================


class TestAddSplice:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        cond1, cond2 = MagicMock(), MagicMock()
        splice = MagicMock()
        splice.Name = "Splice_1"
        splices = MagicMock()
        splices.Add.return_value = splice
        doc.Splices = splices

        occurrences = MagicMock()
        occurrences.Count = 3
        occurrences.Item.side_effect = lambda i: {1: cond1, 2: cond2, 3: MagicMock()}[i]
        doc.Occurrences = occurrences

        result = am.add_splice(0.1, 0.2, 0.3, [0, 1], "Test splice")
        assert result["status"] == "created"
        assert result["type"] == "splice"
        assert result["position"] == [0.1, 0.2, 0.3]
        assert result["num_conductors"] == 2
        assert result["description"] == "Test splice"
        splices.Add.assert_called_once()

    def test_invalid_conductor_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences

        result = am.add_splice(0, 0, 0, [0, 5], "")
        assert "error" in result
        assert "Invalid conductor index" in result["error"]

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.add_splice(0, 0, 0, [0], "")
        assert "error" in result


# ============================================================================
# WIRES
# ============================================================================


class TestAddWire:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        p1, p2 = MagicMock(), MagicMock()
        wire = MagicMock()
        wire.Name = "Wire_1"
        wires = MagicMock()
        wires.Add.return_value = wire
        doc.Wires = wires

        occurrences = MagicMock()
        occurrences.Count = 3
        occurrences.Item.side_effect = lambda i: {1: p1, 2: p2, 3: MagicMock()}[i]
        doc.Occurrences = occurrences

        result = am.add_wire([0, 1], [True, False], "Test wire")
        assert result["status"] == "created"
        assert result["type"] == "wire"
        assert result["num_paths"] == 2
        assert result["description"] == "Test wire"
        wires.Add.assert_called_once()

    def test_mismatched_lengths(self, asm_mgr):
        am, doc = asm_mgr
        doc.Occurrences = MagicMock()

        result = am.add_wire([0, 1], [True], "")
        assert "error" in result
        assert "same length" in result["error"]

    def test_invalid_path_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 1
        doc.Occurrences = occurrences

        result = am.add_wire([0, 5], [True, True], "")
        assert "error" in result
        assert "Invalid path index" in result["error"]

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.add_wire([0], [True], "")
        assert "error" in result


# ============================================================================
# CABLES
# ============================================================================


class TestAddCable:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        p1, w1 = MagicMock(), MagicMock()
        cable = MagicMock()
        cable.Name = "Cable_1"
        cables = MagicMock()
        cables.Add.return_value = cable
        doc.Cables = cables

        occurrences = MagicMock()
        occurrences.Count = 3
        occurrences.Item.side_effect = lambda i: {1: p1, 2: w1, 3: MagicMock()}[i]
        doc.Occurrences = occurrences

        result = am.add_cable([0], [True], [1], description="Test cable")
        assert result["status"] == "created"
        assert result["type"] == "cable"
        assert result["num_paths"] == 1
        assert result["num_wires"] == 1
        cables.Add.assert_called_once()

    def test_mismatched_path_lengths(self, asm_mgr):
        am, doc = asm_mgr
        doc.Occurrences = MagicMock()

        result = am.add_cable([0, 1], [True], [0])
        assert "error" in result
        assert "same length" in result["error"]

    def test_invalid_wire_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = MagicMock()
        doc.Occurrences = occurrences

        result = am.add_cable([0], [True], [5])
        assert "error" in result
        assert "Invalid wire index" in result["error"]

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.add_cable([0], [True], [0])
        assert "error" in result


# ============================================================================
# BUNDLES
# ============================================================================


class TestAddBundle:
    def test_success(self, asm_mgr):
        am, doc = asm_mgr
        p1, c1 = MagicMock(), MagicMock()
        bundle = MagicMock()
        bundle.Name = "Bundle_1"
        bundles = MagicMock()
        bundles.Add.return_value = bundle
        doc.Bundles = bundles

        occurrences = MagicMock()
        occurrences.Count = 3
        occurrences.Item.side_effect = lambda i: {1: p1, 2: c1, 3: MagicMock()}[i]
        doc.Occurrences = occurrences

        result = am.add_bundle([0], [True], [1], description="Test bundle")
        assert result["status"] == "created"
        assert result["type"] == "bundle"
        assert result["num_paths"] == 1
        assert result["num_conductors"] == 1
        bundles.Add.assert_called_once()

    def test_mismatched_path_lengths(self, asm_mgr):
        am, doc = asm_mgr
        doc.Occurrences = MagicMock()

        result = am.add_bundle([0, 1], [True], [0])
        assert "error" in result
        assert "same length" in result["error"]

    def test_invalid_conductor_index(self, asm_mgr):
        am, doc = asm_mgr
        occurrences = MagicMock()
        occurrences.Count = 2
        occurrences.Item.return_value = MagicMock()
        doc.Occurrences = occurrences

        result = am.add_bundle([0], [True], [5])
        assert "error" in result
        assert "Invalid conductor index" in result["error"]

    def test_not_assembly(self, asm_mgr):
        am, doc = asm_mgr
        del doc.Occurrences

        result = am.add_bundle([0], [True], [0])
        assert "error" in result
