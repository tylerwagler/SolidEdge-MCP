#!/usr/bin/env python
"""Test ellipse signature"""

import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).parent / "src"))

from solidedge_mcp.backends.connection import SolidEdgeConnection
from solidedge_mcp.backends.documents import DocumentManager
import math

connection = SolidEdgeConnection()
connection.connect(start_if_needed=True)
doc_manager = DocumentManager(connection)
doc_manager.create_part()
doc = doc_manager.get_active_document()

# Create sketch
ref_planes = doc.RefPlanes
ref_plane = ref_planes.Item(1)
profile_sets = doc.ProfileSets
profile_set = profile_sets.Add()
profiles = profile_set.Profiles
profile = profiles.Add(ref_plane)

ellipses = profile.Ellipses2d

print("Testing AddByCenter with different parameter counts...")

try:
    # Try 5 params (cx, cy, major, minor, angle)
    e = ellipses.AddByCenter(0.5, 0.1, 0.04, 0.02, math.radians(45))
    print("SUCCESS with 5 params (cx, cy, major, minor, angle)!")
except Exception as ex:
    print(f"Failed with 5 params: {ex}")
    try:
        # Try 4 params (cx, cy, major, minor)
        e = ellipses.AddByCenter(0.5, 0.1, 0.04, 0.02)
        print("SUCCESS with 4 params (cx, cy, major, minor)!")
    except Exception as ex2:
        print(f"Failed with 4 params: {ex2}")
        try:
            # Try with different order or params?
            e = ellipses.AddByCenter(0.5, 0.1, 0.04, 0.02, 0.0, 1.0)
            print("SUCCESS with 6 params!")
        except Exception as ex3:
            print(f"Failed with 6 params: {ex3}")
