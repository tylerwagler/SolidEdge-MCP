#!/usr/bin/env python
"""
Comprehensive test of ALL Solid Edge MCP tools.
Tests every major category with real Solid Edge instance.

Usage:
    python comprehensive_test.py
"""

import sys
import os
import tempfile
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

def test_category(name):
    """Print category header"""
    print("\n" + "=" * 70)
    print(f"  {name}")
    print("=" * 70)

def test_step(name, func):
    """Run a test step and report result"""
    try:
        result = func()
        if isinstance(result, dict) and "error" in result:
            print(f"[FAIL] {name}: {result['error']}")
            return False
        else:
            print(f"[OK] {name}")
            if isinstance(result, dict) and "status" in result:
                print(f"     -> {result['status']}")
            return True
    except Exception as e:
        print(f"[FAIL] {name}: {e}")
        return False

def main():
    """Run comprehensive tests"""
    print("=" * 70)
    print("SOLID EDGE MCP SERVER - COMPREHENSIVE TEST SUITE")
    print("=" * 70)
    print("\nThis will test all major tool categories with real Solid Edge.")
    print("Press Ctrl+C to cancel...\n")

    passed = []
    failed = []

    # Import all managers
    from solidedge_mcp.backends.connection import SolidEdgeConnection
    from solidedge_mcp.backends.documents import DocumentManager
    from solidedge_mcp.backends.sketching import SketchManager
    from solidedge_mcp.backends.features import FeatureManager
    from solidedge_mcp.backends.assembly import AssemblyManager
    from solidedge_mcp.backends.query import QueryManager
    from solidedge_mcp.backends.export import ExportManager, ViewModel

    # Initialize managers
    connection = SolidEdgeConnection()
    doc_manager = DocumentManager(connection)
    sketch_manager = SketchManager(doc_manager)
    feature_manager = FeatureManager(doc_manager, sketch_manager)
    assembly_manager = AssemblyManager(doc_manager)
    query_manager = QueryManager(doc_manager)
    export_manager = ExportManager(doc_manager)
    view_manager = ViewModel(doc_manager)

    temp_dir = tempfile.gettempdir()

    # ========================================================================
    # CATEGORY 1: CONNECTION
    # ========================================================================
    test_category("CONNECTION & APPLICATION")

    if test_step("Connect to Solid Edge", lambda: connection.connect(start_if_needed=True)):
        passed.append("connect")
    else:
        failed.append("connect")
        print("\n[CRITICAL] Cannot connect to Solid Edge. Aborting tests.")
        return 1

    if test_step("Get application info", lambda: connection.get_application_info()):
        passed.append("app_info")
    else:
        failed.append("app_info")

    # ========================================================================
    # CATEGORY 2: DOCUMENT MANAGEMENT
    # ========================================================================
    test_category("DOCUMENT MANAGEMENT")

    if test_step("Create part document", lambda: doc_manager.create_part()):
        passed.append("create_part")
    else:
        failed.append("create_part")
        return 1

    # ========================================================================
    # CATEGORY 3: SKETCHING - BASIC SHAPES
    # ========================================================================
    test_category("SKETCHING - Basic Shapes")

    if test_step("Create sketch on Top plane", lambda: sketch_manager.create_sketch("Top")):
        passed.append("create_sketch")
    else:
        failed.append("create_sketch")

    if test_step("Draw line", lambda: sketch_manager.draw_line(0, 0, 0.05, 0)):
        passed.append("draw_line")
    else:
        failed.append("draw_line")

    if test_step("Draw circle", lambda: sketch_manager.draw_circle(0.1, 0.1, 0.02)):
        passed.append("draw_circle")
    else:
        failed.append("draw_circle")

    if test_step("Draw rectangle", lambda: sketch_manager.draw_rectangle(0.15, 0, 0.25, 0.08)):
        passed.append("draw_rectangle")
    else:
        failed.append("draw_rectangle")

    if test_step("Draw arc", lambda: sketch_manager.draw_arc(0.3, 0.1, 0.03, 0, 180)):
        passed.append("draw_arc")
    else:
        failed.append("draw_arc")

    if test_step("Draw polygon (hexagon)", lambda: sketch_manager.draw_polygon(0.4, 0.05, 0.025, 6)):
        passed.append("draw_polygon")
    else:
        failed.append("draw_polygon")

    if test_step("Draw ellipse", lambda: sketch_manager.draw_ellipse(0.5, 0.1, 0.04, 0.02, 45)):
        passed.append("draw_ellipse")
    else:
        failed.append("draw_ellipse")

    if test_step("Draw spline", lambda: sketch_manager.draw_spline([[0.6, 0], [0.65, 0.05], [0.7, 0.02]])):
        passed.append("draw_spline")
    else:
        failed.append("draw_spline")

    if test_step("Close sketch", lambda: sketch_manager.close_sketch()):
        passed.append("close_sketch")
    else:
        failed.append("close_sketch")

    # ========================================================================
    # CATEGORY 4: 3D FEATURES - PRIMITIVES
    # ========================================================================
    test_category("3D FEATURES - Primitives")

    if test_step("Create box by center", lambda: feature_manager.create_box_by_center(0, 0, 0, 0.1, 0.1, 0.1)):
        passed.append("box_center")
    else:
        failed.append("box_center")

    if test_step("Create cylinder", lambda: feature_manager.create_cylinder(0.2, 0, 0, 0.03, 0.15)):
        passed.append("cylinder")
    else:
        failed.append("cylinder")

    if test_step("Create sphere", lambda: feature_manager.create_sphere(0.4, 0, 0, 0.04)):
        passed.append("sphere")
    else:
        failed.append("sphere")

    # ========================================================================
    # CATEGORY 5: 3D FEATURES - EXTRUDE
    # ========================================================================
    test_category("3D FEATURES - Extrude")

    # Create a simple sketch for extrusion
    sketch_manager.create_sketch("Front")
    sketch_manager.draw_rectangle(0, 0, 0.05, 0.05)
    sketch_manager.close_sketch()

    if test_step("Create extrude (finite)", lambda: feature_manager.create_extrude(0.03, "Add", "Normal")):
        passed.append("extrude")
    else:
        failed.append("extrude")

    # ========================================================================
    # CATEGORY 6: 3D FEATURES - REVOLVE
    # ========================================================================
    test_category("3D FEATURES - Revolve")

    # Create sketch for revolve
    sketch_manager.create_sketch("Front")
    sketch_manager.draw_line(0, 0, 0.02, 0)
    sketch_manager.draw_line(0.02, 0, 0.02, 0.04)
    sketch_manager.draw_line(0.02, 0.04, 0, 0.04)
    sketch_manager.draw_line(0, 0.04, 0, 0)
    sketch_manager.close_sketch()

    if test_step("Create revolve (360 degrees)", lambda: feature_manager.create_revolve(360, "Add")):
        passed.append("revolve")
    else:
        failed.append("revolve")

    # ========================================================================
    # CATEGORY 7: QUERY & ANALYSIS
    # ========================================================================
    test_category("QUERY & ANALYSIS")

    if test_step("Get feature count", lambda: query_manager.get_feature_count()):
        passed.append("feature_count")
    else:
        failed.append("feature_count")

    if test_step("List features", lambda: feature_manager.list_features()):
        passed.append("list_features")
    else:
        failed.append("list_features")

    if test_step("Get mass properties", lambda: query_manager.get_mass_properties()):
        passed.append("mass_properties")
    else:
        failed.append("mass_properties")

    if test_step("Get bounding box", lambda: query_manager.get_bounding_box()):
        passed.append("bounding_box")
    else:
        failed.append("bounding_box")

    if test_step("Measure distance", lambda: query_manager.measure_distance(0, 0, 0, 0.1, 0.1, 0.1)):
        passed.append("measure_distance")
    else:
        failed.append("measure_distance")

    if test_step("Get document properties", lambda: query_manager.get_document_properties()):
        passed.append("doc_properties")
    else:
        failed.append("doc_properties")

    # ========================================================================
    # CATEGORY 8: VIEW OPERATIONS
    # ========================================================================
    test_category("VIEW OPERATIONS")

    if test_step("Set view to Iso", lambda: view_manager.set_view("Iso")):
        passed.append("set_view")
    else:
        failed.append("set_view")

    if test_step("Zoom fit", lambda: view_manager.zoom_fit()):
        passed.append("zoom_fit")
    else:
        failed.append("zoom_fit")

    if test_step("Set display mode", lambda: view_manager.set_display_mode("Shaded")):
        passed.append("display_mode")
    else:
        failed.append("display_mode")

    # ========================================================================
    # CATEGORY 9: EXPORT OPERATIONS
    # ========================================================================
    test_category("EXPORT OPERATIONS")

    # Save the document first
    test_file = os.path.join(temp_dir, "solidedge_comprehensive_test.par")
    if test_step("Save document", lambda: doc_manager.save_document(test_file)):
        passed.append("save_document")
    else:
        failed.append("save_document")

    if test_step("Export to STEP", lambda: export_manager.export_step(os.path.join(temp_dir, "test.step"))):
        passed.append("export_step")
    else:
        failed.append("export_step")

    if test_step("Export to STL", lambda: export_manager.export_stl(os.path.join(temp_dir, "test.stl"))):
        passed.append("export_stl")
    else:
        failed.append("export_stl")

    if test_step("Capture screenshot", lambda: export_manager.capture_screenshot(os.path.join(temp_dir, "test.png"))):
        passed.append("screenshot")
    else:
        failed.append("screenshot")

    # ========================================================================
    # CATEGORY 10: ASSEMBLY (if time permits)
    # ========================================================================
    test_category("ASSEMBLY OPERATIONS")

    # Close part and create assembly
    doc_manager.close_document()

    if test_step("Create assembly document", lambda: doc_manager.create_assembly()):
        passed.append("create_assembly")
    else:
        failed.append("create_assembly")

    # Try to place the part we just created
    if os.path.exists(test_file):
        if test_step("Place component", lambda: assembly_manager.add_component(test_file, 0, 0, 0)):
            passed.append("place_component")
        else:
            failed.append("place_component")

        if test_step("List components", lambda: assembly_manager.list_components()):
            passed.append("list_components")
        else:
            failed.append("list_components")

    # Close assembly (don't save test document)
    doc_manager.close_document(save=False)

    # ========================================================================
    # SUMMARY
    # ========================================================================
    print("\n" + "=" * 70)
    print("TEST SUMMARY")
    print("=" * 70)
    print(f"\nTotal Tests: {len(passed) + len(failed)}")
    print(f"Passed: {len(passed)} ({100*len(passed)//(len(passed)+len(failed)) if passed or failed else 0}%)")
    print(f"Failed: {len(failed)}")

    if passed:
        print(f"\n[OK] Passed ({len(passed)}):")
        for i, test in enumerate(passed, 1):
            print(f"  {i}. {test}")

    if failed:
        print(f"\n[FAIL] Failed ({len(failed)}):")
        for i, test in enumerate(failed, 1):
            print(f"  {i}. {test}")

    print("\n" + "=" * 70)

    if not failed:
        print("\nSUCCESS! All tests passed!")
        print("The Solid Edge MCP server is fully operational!")
        return 0
    else:
        print(f"\n{len(failed)} test(s) failed. Review output above for details.")
        return 1

if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n\nTests cancelled by user.")
        sys.exit(1)
    except Exception as e:
        print(f"\n\nUnexpected error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
