#!/usr/bin/env python
"""
Integration test for Solid Edge MCP Server.
Requires Solid Edge to be installed and the MCP server dependencies.

This script tests basic workflows with a real Solid Edge instance.

Usage:
    python test_with_solidedge.py
"""

import os
import sys
import tempfile
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))


def test_connection():
    """Test connecting to Solid Edge"""
    print("Testing connection to Solid Edge...")
    try:
        from solidedge_mcp.backends.connection import SolidEdgeConnection

        connection = SolidEdgeConnection()
        result = connection.connect(start_if_needed=True)

        if "error" in result:
            print(f"[FAIL] Connection failed: {result['error']}")
            return False

        print(f"[OK] Connected: {result}")
        return True
    except Exception as e:
        print(f"[FAIL] Connection test failed: {e}")
        import traceback

        traceback.print_exc()
        return False


def test_create_simple_part():
    """Test creating a simple extruded part"""
    print("\nTesting simple part creation...")
    try:
        from solidedge_mcp.backends.connection import SolidEdgeConnection
        from solidedge_mcp.backends.documents import DocumentManager
        from solidedge_mcp.backends.features import FeatureManager
        from solidedge_mcp.backends.sketching import SketchManager

        # Connect
        connection = SolidEdgeConnection()
        connection.connect(start_if_needed=True)

        # Create managers
        doc_manager = DocumentManager(connection)
        sketch_manager = SketchManager(doc_manager)
        feature_manager = FeatureManager(doc_manager, sketch_manager)

        # Create part
        result = doc_manager.create_part()
        if "error" in result:
            print(f"[FAIL] Create part failed: {result['error']}")
            return False
        print(f"[OK] Part created: {result}")

        # Create sketch
        result = sketch_manager.create_sketch("Top")
        if "error" in result:
            print(f"[FAIL] Create sketch failed: {result['error']}")
            return False
        print(f"[OK] Sketch created: {result}")

        # Draw rectangle (100mm x 100mm, convert to meters)
        result = sketch_manager.draw_rectangle(0, 0, 0.1, 0.1)
        if "error" in result:
            print(f"[FAIL] Draw rectangle failed: {result['error']}")
            return False
        print(f"[OK] Rectangle drawn: {result}")

        # Close sketch
        result = sketch_manager.close_sketch()
        if "error" in result:
            print(f"[FAIL] Close sketch failed: {result['error']}")
            return False
        print(f"[OK] Sketch closed: {result}")

        # Extrude (50mm height)
        result = feature_manager.create_extrude(distance=0.05, operation="Add")
        if "error" in result:
            print(f"[FAIL] Extrude failed: {result['error']}")
            return False
        print(f"[OK] Extrude created: {result}")

        # Save to temp file
        temp_dir = tempfile.gettempdir()
        test_file = os.path.join(temp_dir, "solidedge_mcp_test.par")
        result = doc_manager.save_document(test_file)
        if "error" in result:
            print(f"[FAIL] Save failed: {result['error']}")
            return False
        print(f"[OK] Part saved: {test_file}")

        # Close document (non-fatal if it fails)
        result = doc_manager.close_document()
        if "error" in result:
            print(f"[WARN] Close document warning: {result['error']}")
        else:
            print("[OK] Document closed")

        # Main functionality worked!
        return True

    except Exception as e:
        print(f"[FAIL] Part creation test failed: {e}")
        import traceback

        traceback.print_exc()
        return False


def test_query_operations():
    """Test querying a part"""
    print("\nTesting query operations...")
    try:
        from solidedge_mcp.backends.connection import SolidEdgeConnection
        from solidedge_mcp.backends.documents import DocumentManager
        from solidedge_mcp.backends.features import FeatureManager
        from solidedge_mcp.backends.query import QueryManager
        from solidedge_mcp.backends.sketching import SketchManager

        # Connect
        connection = SolidEdgeConnection()
        connection.connect(start_if_needed=True)

        # Create managers
        doc_manager = DocumentManager(connection)
        query_manager = QueryManager(doc_manager)
        sketch_manager = SketchManager(doc_manager)
        feature_manager = FeatureManager(doc_manager, sketch_manager)

        # Create a simple part first
        doc_manager.create_part()

        # Test get_feature_count
        result = query_manager.get_feature_count()
        if "error" in result:
            print(f"[FAIL] Get feature count failed: {result['error']}")
            return False
        print(f"[OK] Feature count: {result.get('count', 0)} features")

        # Test list_features
        result = feature_manager.list_features()
        if "error" in result:
            print(f"[WARN] List features failed (may be empty): {result['error']}")
        else:
            print(f"[OK] Features listed: {result.get('feature_count', 0)} features")

        # Test get_bounding_box
        result = query_manager.get_bounding_box()
        if "error" in result:
            print(f"[WARN] Bounding box query failed (may be empty): {result['error']}")
        else:
            print(f"[OK] Bounding box: {result}")

        # Close document
        doc_manager.close_document()

        return True

    except Exception as e:
        print(f"[FAIL] Query test failed: {e}")
        import traceback

        traceback.print_exc()
        return False


def main():
    """Run all integration tests"""
    print("=" * 70)
    print("Solid Edge MCP Server - Integration Tests")
    print("=" * 70)
    print("\nNOTE: These tests require:")
    print("  - Solid Edge installed and licensed")
    print("  - Python dependencies installed (uv sync --all-extras)")
    print("  - Windows OS")
    print()

    input("Press Enter to start tests (or Ctrl+C to cancel)...")

    results = []
    results.append(("Connection Test", test_connection()))
    results.append(("Simple Part Creation", test_create_simple_part()))
    results.append(("Query Operations", test_query_operations()))

    print("\n" + "=" * 70)
    print("Test Results:")
    print("=" * 70)

    all_passed = True
    for test_name, result in results:
        status = "[OK] PASS" if result else "[FAIL] FAIL"
        print(f"{test_name}: {status}")
        if not result:
            all_passed = False

    print("=" * 70)

    if all_passed:
        print("\nSUCCESS! All integration tests passed!")
        print("\nYour Solid Edge MCP server is working correctly!")
        return 0
    else:
        print("\n[WARN] Some tests failed.")
        print("This may be due to:")
        print("  - Solid Edge not running or not licensed")
        print("  - COM automation permissions")
        print("  - Missing dependencies")
        return 1


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n\nTests cancelled by user.")
        sys.exit(1)
