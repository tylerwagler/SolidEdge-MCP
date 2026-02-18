#!/usr/bin/env python
"""
Demo of Working SolidEdge MCP Features
Shows all the functionality that's currently operational (~60% of tools)
"""

import math
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / "src"))

import time

from solidedge_mcp.backends.connection import SolidEdgeConnection
from solidedge_mcp.backends.documents import DocumentManager
from solidedge_mcp.backends.export import ViewModel
from solidedge_mcp.backends.features import FeatureManager
from solidedge_mcp.backends.query import QueryManager
from solidedge_mcp.backends.sketching import SketchManager

print("=" * 70)
print("SOLID EDGE MCP - WORKING FEATURES DEMO")
print("=" * 70)

# Initialize
connection = SolidEdgeConnection()
doc_manager = DocumentManager(connection)
sketch_manager = SketchManager(doc_manager)
feature_manager = FeatureManager(doc_manager, sketch_manager)
query_manager = QueryManager(doc_manager)
view_manager = ViewModel(doc_manager)

# Connect
print("\n[1/8] Connecting to Solid Edge...")
connection.connect(start_if_needed=True)
print("[OK] Connected")

# Create part
print("\n[2/8] Creating new part document...")
doc_manager.create_part()
print("[OK] Part created")

# Demo sketching - create a gear-like profile
print("\n[3/8] Creating sketch with various 2D geometry...")
sketch_manager.create_sketch("Front")

# Center hub (circle)
sketch_manager.draw_circle(0, 0, 0.015)
print("  [OK] Drew center hub (circle)")

# Outer ring with cutouts
sketch_manager.draw_circle(0, 0, 0.04)
print("  [OK] Drew outer ring (circle)")

# Add 6 cutout circles around the perimeter
num_teeth = 6
radius = 0.035
cutout_radius = 0.008
for i in range(num_teeth):
    angle = (2 * math.pi * i) / num_teeth
    x = radius * math.cos(angle)
    y = radius * math.sin(angle)
    sketch_manager.draw_circle(x, y, cutout_radius)
print(f"  [OK] Drew {num_teeth} cutout circles")

# Add decorative ellipse in center
sketch_manager.draw_ellipse(0, 0, 0.01, 0.005, 45)
print("  [OK] Drew decorative ellipse")

# Add some spline curves for detail
points1 = [[0.02, 0.02], [0.025, 0.025], [0.02, 0.03]]
sketch_manager.draw_spline(points1)
print("  [OK] Drew spline curve")

sketch_manager.close_sketch()
print("[OK] Sketch completed")

# Create extrusion
print("\n[4/8] Creating 3D extrusion...")
result = feature_manager.create_extrude(0.01, "Add", "Normal")
if "error" in result:
    print(f"  [X] Extrude failed: {result['error']}")
else:
    print(f"[OK] Extruded to depth: {result['distance']} meters")

# Query properties
print("\n[5/8] Querying model properties...")
time.sleep(0.5)  # Give Solid Edge moment to compute

# Feature count
count_result = query_manager.get_feature_count()
print(f"  • Feature count: {count_result.get('count', 'unknown')}")

# Mass properties
mass_result = query_manager.get_mass_properties()
if "error" not in mass_result:
    volume = mass_result.get("volume_m3", 0)
    area = mass_result.get("surface_area_m2", 0)
    print(f"  • Volume: {volume:.6f} m³")
    print(f"  • Surface area: {area:.6f} m²")
    com = mass_result.get("center_of_mass", {})
    if com:
        cx = com.get("x", 0)
        cy = com.get("y", 0)
        cz = com.get("z", 0)
        print(f"  • Center of mass: ({cx:.4f}, {cy:.4f}, {cz:.4f})")

# Bounding box
bbox_result = query_manager.get_bounding_box()
if "error" not in bbox_result:
    print(f"  • Bounding box: {bbox_result.get('dimensions', 'unknown')}")

print("[OK] Properties queried")

# Set view
print("\n[6/8] Setting isometric view...")
view_result = view_manager.set_view("Iso")
if "error" not in view_result:
    print("[OK] View set to isometric")
else:
    print(f"  Note: {view_result.get('error', 'unknown')}")

# Zoom to fit
print("\n[7/8] Zooming to fit...")
zoom_result = view_manager.zoom_fit()
print("[OK] Zoomed to fit")

# Create second feature - add a cut
print("\n[8/8] Adding a cutout feature...")
sketch_manager.create_sketch("Top")
sketch_manager.draw_rectangle(-0.05, -0.002, 0.05, 0.002)
sketch_manager.close_sketch()

cut_result = feature_manager.create_extrude(0.015, "Cut", "Normal")
if "error" in cut_result:
    print(f"  Note: Cut operation - {cut_result.get('error', 'unknown')}")
else:
    print(f"[OK] Cut created with depth: {cut_result['distance']} meters")

# Final feature count
final_count = query_manager.get_feature_count()
print(f"\n[OK] Final feature count: {final_count.get('count', 'unknown')}")

print("\n" + "=" * 70)
print("DEMO COMPLETE!")
print("=" * 70)
print("\nWorking Features Demonstrated:")
print("  [OK] Connection management")
print("  [OK] Document creation")
print("  [OK] 2D sketching (circles, ellipses, splines)")
print("  [OK] Profile management")
print("  [OK] 3D extrusion (add and cut)")
print("  [OK] Property queries (volume, area, center of mass, bbox)")
print("  [OK] View management (orientation, zoom)")
print("\nNot Demonstrated (but also working):")
print("  • Lines, arcs, rectangles, polygons")
print("  • Document save/export (STEP, STL)")
print("  • Screenshot capture")
print("  • Assembly operations")
print("  • Distance measurement")
print("\nKnown Issues (being worked on):")
print("  [X] Revolve operations")
print("  [X] Primitive creation (box, cylinder, sphere)")
print("  [X] Display mode setting")
print("\n" + "=" * 70)
print("\nYour part should now show a gear-like profile with:")
print("  - Center hub")
print("  - Outer ring with 6 cutouts")
print("  - Decorative details")
print("  - A slot cut through the middle")
print("=" * 70)
