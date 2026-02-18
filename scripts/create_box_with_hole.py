#!/usr/bin/env python
"""
Create a 200mm x 300mm x 50mm box with a 10mm diameter through hole in the center
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "src"))

from solidedge_mcp.backends.connection import SolidEdgeConnection
from solidedge_mcp.backends.documents import DocumentManager
from solidedge_mcp.backends.features import FeatureManager
from solidedge_mcp.backends.sketching import SketchManager

print("Creating 200mm x 300mm x 50mm box with 10mm center hole...")

# Initialize
connection = SolidEdgeConnection()
doc_manager = DocumentManager(connection)
sketch_manager = SketchManager(doc_manager)
feature_manager = FeatureManager(doc_manager, sketch_manager)

# Connect
print("\n[1/5] Connecting...")
connection.connect(start_if_needed=True)

# Create part
print("[2/5] Creating part...")
doc_manager.create_part()

# Create base box sketch (200mm x 300mm rectangle)
print("[3/5] Creating base sketch (200mm x 300mm)...")
sketch_manager.create_sketch("Top")
# Draw rectangle centered at origin
# 200mm = 0.2m, 300mm = 0.3m
# -100mm to +100mm in X, -150mm to +150mm in Y
sketch_manager.draw_rectangle(-0.1, -0.15, 0.1, 0.15)
sketch_manager.close_sketch()

# Extrude to create box (50mm = 0.05m)
print("[4/5] Extruding box (50mm height)...")
result = feature_manager.create_extrude(0.05, "Add", "Normal")
if "error" in result:
    print(f"ERROR: Extrude failed - {result['error']}")
    doc_manager.close_document(save=False)
    sys.exit(1)
else:
    print(f"   Box created: {result}")

# Create hole sketch (10mm diameter circle at center)
print("[5/5] Creating hole sketch (10mm diameter at center)...")
sketch_manager.create_sketch("Top")
# 10mm diameter = 5mm radius = 0.005m radius
sketch_manager.draw_circle(0, 0, 0.005)
sketch_manager.close_sketch()

# Cut through hole
print("[6/6] Cutting through hole...")
# Use large depth to ensure it cuts through
result = feature_manager.create_extrude(0.1, "Cut", "Normal")
if "error" in result:
    print(f"ERROR: Cut failed - {result['error']}")
else:
    print(f"   Hole created: {result}")

print("\n" + "=" * 60)
print("COMPLETE!")
print("=" * 60)
print("Created:")
print("  - Box: 200mm x 300mm x 50mm")
print("  - Center hole: 10mm diameter, through")
print("=" * 60)
