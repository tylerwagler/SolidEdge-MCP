#!/usr/bin/env python
"""
Test script to validate the MCP server loads correctly.
Run this to verify all tools are registered properly.

Usage:
    python test_server.py
"""

import sys
from pathlib import Path

# Add src to path
sys.path.insert(0, str(Path(__file__).parent / "src"))

def test_imports():
    """Test that all modules can be imported"""
    print("Testing imports...")
    try:
        from solidedge_mcp.backends import connection
        from solidedge_mcp.backends import documents
        from solidedge_mcp.backends import sketching
        from solidedge_mcp.backends import features
        from solidedge_mcp.backends import assembly
        from solidedge_mcp.backends import query
        from solidedge_mcp.backends import export as export_module
        from solidedge_mcp.backends import constants
        from solidedge_mcp.backends import diagnostics
        print("[OK] All backend modules imported successfully")
        return True
    except Exception as e:
        print(f"[FAIL] Import failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_server_structure():
    """Test that server.py has the expected structure"""
    print("\nTesting server structure...")
    try:
        # Read server.py and count tools
        server_path = Path(__file__).parent / "src" / "solidedge_mcp" / "server.py"
        with open(server_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # Count @mcp.tool() decorators
        tool_count = content.count('@mcp.tool()')
        print(f"[OK] Found {tool_count} MCP tools")

        if tool_count != 89:
            print(f"  [WARN] Expected 89 tools, found {tool_count}")
            return False

        # Check for manager instances
        managers = [
            'connection = ',
            'doc_manager = ',
            'sketch_manager = ',
            'feature_manager = ',
            'assembly_manager = ',
            'query_manager = ',
            'export_manager = ',
            'view_manager = '
        ]

        for manager in managers:
            if manager not in content:
                print(f"[FAIL] Missing manager: {manager}")
                return False

        print(f"[OK] All {len(managers)} managers found")
        return True
    except Exception as e:
        print(f"[FAIL] Structure test failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def test_tool_names():
    """Extract and display all tool names"""
    print("\nExtracting tool names...")
    try:
        server_path = Path(__file__).parent / "src" / "solidedge_mcp" / "server.py"
        with open(server_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()

        tools = []
        for i, line in enumerate(lines):
            if '@mcp.tool()' in line and i + 1 < len(lines):
                next_line = lines[i + 1].strip()
                if next_line.startswith('def '):
                    func_name = next_line.split('(')[0].replace('def ', '')
                    tools.append(func_name)

        print(f"[OK] Found {len(tools)} tool definitions:")

        # Group by category
        categories = {
            'Connection': [],
            'Documents': [],
            'Sketching': [],
            'Primitives': [],
            'Features': [],
            'Sheet Metal': [],
            'Body': [],
            'Simplification': [],
            'View': [],
            'Query': [],
            'Export': [],
            'Assembly': [],
            'Diagnostics': []
        }

        for tool in tools:
            if 'connect' in tool or 'application' in tool:
                categories['Connection'].append(tool)
            elif any(x in tool for x in ['document', 'save', 'open', 'close', 'list_documents']):
                categories['Documents'].append(tool)
            elif any(x in tool for x in ['sketch', 'draw', 'constraint']):
                categories['Sketching'].append(tool)
            elif any(x in tool for x in ['box', 'cylinder', 'sphere']):
                categories['Primitives'].append(tool)
            elif any(x in tool for x in ['extrude', 'revolve', 'loft', 'sweep', 'helix']):
                categories['Features'].append(tool)
            elif any(x in tool for x in ['flange', 'tab', 'web']):
                categories['Sheet Metal'].append(tool)
            elif 'body' in tool or 'construction' in tool or 'thicken' in tool:
                categories['Body'].append(tool)
            elif 'simplify' in tool:
                categories['Simplification'].append(tool)
            elif any(x in tool for x in ['view', 'zoom', 'display']):
                categories['View'].append(tool)
            elif any(x in tool for x in ['mass', 'bounding', 'measure', 'feature', 'properties']):
                categories['Query'].append(tool)
            elif 'export' in tool or 'screenshot' in tool or 'drawing' in tool:
                categories['Export'].append(tool)
            elif any(x in tool for x in ['component', 'mate', 'assembly', 'place', 'pattern', 'suppress']):
                categories['Assembly'].append(tool)
            elif 'diagnose' in tool:
                categories['Diagnostics'].append(tool)

        print()
        for category, tool_list in categories.items():
            if tool_list:
                print(f"  {category} ({len(tool_list)}): {', '.join(tool_list[:3])}{'...' if len(tool_list) > 3 else ''}")

        return True
    except Exception as e:
        print(f"[FAIL] Tool name extraction failed: {e}")
        import traceback
        traceback.print_exc()
        return False

def main():
    """Run all tests"""
    print("=" * 70)
    print("Solid Edge MCP Server - Validation Tests")
    print("=" * 70)

    results = []
    results.append(("Import Test", test_imports()))
    results.append(("Structure Test", test_server_structure()))
    results.append(("Tool Names Test", test_tool_names()))

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
        print("\nSUCCESS! All validation tests passed!")
        print("\nNext steps:")
        print("1. Ensure Solid Edge is installed and licensed")
        print("2. Test with MCP client (Claude Code, etc.)")
        print("3. Try the example workflows in CLAUDE.md")
        return 0
    else:
        print("\n[WARN] Some tests failed. Please review the output above.")
        return 1

if __name__ == "__main__":
    sys.exit(main())
