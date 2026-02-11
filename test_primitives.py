#!/usr/bin/env python
"""Test primitive creation signatures"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from solidedge_mcp.backends.connection import SolidEdgeConnection
from solidedge_mcp.backends.documents import DocumentManager
import win32com.client

connection = SolidEdgeConnection()
connection.connect(start_if_needed=True)
doc_manager = DocumentManager(connection)
doc_manager.create_part()
doc = doc_manager.get_active_document()
models = doc.Models

def create_variant_array(values):
    """Create a COM VARIANT array"""
    import pythoncom
    return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, values)

print("Testing AddBoxByCenter...")
try:
    # Maybe it needs (BaseXOffset, BaseYOffset, BaseZOffset, Length, Width, Height)?
    box = models.AddBoxByCenter(0.0, 0.0, 0.0, 0.1, 0.1, 0.1)
    print("SUCCESS with 6 float params!")
except Exception as e:
    print(f"Failed with 6 params: {e}")
    try:
        # Or maybe the method name is wrong? Try other box methods
        box = models.AddBoxByTwoPoints(0.0, 0.0, 0.0, 0.1, 0.1, 0.1)
        print("SUCCESS with AddBoxByTwoPoints!")
    except Exception as e2:
        print(f"Failed with AddBoxByTwoPoints: {e2}")

print("\nTesting AddCylinderByCenterAndRadius...")
try:
    # Try with tuples
    center = (0.2, 0.0, 0.0)
    axis = (0.0, 0.0, 1.0)
    cyl = models.AddCylinderByCenterAndRadius(center, axis, 0.03, 0.15)
    print("SUCCESS with tuples!")
except Exception as e:
    print(f"Failed: {e}")

print("\nTesting AddSphereByCenterAndRadius...")
try:
    # Try with tuple
    center = (0.4, 0.0, 0.0)
    sphere = models.AddSphereByCenterAndRadius(center, 0.04)
    print("SUCCESS with tuple!")
except Exception as e:
    print(f"Failed: {e}")
