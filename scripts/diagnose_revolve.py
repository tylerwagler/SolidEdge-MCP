#!/usr/bin/env python
"""Diagnose revolve COM API in detail"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))

import inspect
import math

from solidedge_mcp.backends.connection import SolidEdgeConnection
from solidedge_mcp.backends.documents import DocumentManager

connection = SolidEdgeConnection()
connection.connect(start_if_needed=True)
doc_manager = DocumentManager(connection)
doc_manager.create_part()
doc = doc_manager.get_active_document()

# Create sketch for revolve
ref_planes = doc.RefPlanes
ref_plane = ref_planes.Item(2)  # Front plane
profile_sets = doc.ProfileSets
profile_set = profile_sets.Add()
profiles = profile_set.Profiles
profile = profiles.Add(ref_plane)

# Draw a simple L-shape
lines = profile.Lines2d
lines.AddBy2Points(0, 0, 0.02, 0)
lines.AddBy2Points(0.02, 0, 0.02, 0.04)
lines.AddBy2Points(0.02, 0.04, 0, 0.04)
lines.AddBy2Points(0, 0.04, 0, 0)

profile.End(0)

models = doc.Models
angle_rad = math.radians(360)

print("=" * 70)
print("DETAILED COM API DIAGNOSTICS FOR AddRevolvedProtrusion")
print("=" * 70)

# Get the signature
try:
    sig = inspect.signature(models.AddRevolvedProtrusion)
    print(f"\nSignature: {sig}")
except Exception as e:
    print(f"\nCannot get signature: {e}")

# Try to get help
try:
    print("\nHelp text:")
    print(help(models.AddRevolvedProtrusion))
except Exception as e:
    print(f"Cannot get help: {e}")

# Look at what a successful extrude does for comparison
print("\n" + "=" * 70)
print("COMPARISON: How AddExtrudedProtrusion works")
print("=" * 70)

try:
    sig = inspect.signature(models.AddExtrudedProtrusion)
    print(f"\nAddExtrudedProtrusion signature: {sig}")
except Exception as e:
    print(f"Cannot get signature: {e}")

# Test different array formats
print("\n" + "=" * 70)
print("TESTING DIFFERENT ARRAY FORMATS")
print("=" * 70)

# Format 1: Python list
print("\nTest 1: Python list [profile]")
try:
    model = models.AddRevolvedProtrusion(1, [profile], angle_rad)
    print("SUCCESS!")
except Exception as e:
    print(f"Failed: {e}")

# Format 2: Python tuple
print("\nTest 2: Python tuple (profile,)")
try:
    model = models.AddRevolvedProtrusion(1, (profile,), angle_rad)
    print("SUCCESS!")
except Exception as e:
    print(f"Failed: {e}")

# Format 3: No count, just array and angle
print("\nTest 3: Just array and angle")
try:
    model = models.AddRevolvedProtrusion([profile], angle_rad)
    print("SUCCESS!")
except Exception as e:
    print(f"Failed: {e}")

# Format 4: Try with 0 for options
print("\nTest 4: Count, array, angle, 0")
try:
    model = models.AddRevolvedProtrusion(1, [profile], angle_rad, 0)
    print("SUCCESS!")
except Exception as e:
    print(f"Failed: {e}")

doc_manager.close_document(save=False)
