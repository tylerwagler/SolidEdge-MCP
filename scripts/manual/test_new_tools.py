"""
Test new tools: extruded cutout, through-all cutout, offset reference plane,
create_sketch_on_plane_index.
"""

import sys

sys.path.insert(0, "src")

from solidedge_mcp.backends.connection import SolidEdgeConnection
from solidedge_mcp.backends.documents import DocumentManager
from solidedge_mcp.backends.features import FeatureManager
from solidedge_mcp.backends.sketching import SketchManager

# Initialize
conn = SolidEdgeConnection()
result = conn.connect(True)
print(f"Connect: {result.get('status', result.get('error'))}")

doc_mgr = DocumentManager(conn)
sketch_mgr = SketchManager(doc_mgr)
feat_mgr = FeatureManager(doc_mgr, sketch_mgr)

# Create a new part
result = doc_mgr.create_part()
print(f"Create part: {result.get('status', result.get('error'))}")

# ============================================================
# TEST 1: Create base box, then extruded cutout (hole in box)
# ============================================================
print("\n" + "=" * 60)
print("TEST 1: EXTRUDED CUTOUT")
print("=" * 60)

# Create a 100mm x 100mm x 50mm box
result = feat_mgr.create_box_by_two_points(-0.05, -0.05, 0, 0.05, 0.05, 0.05)
print(f"Create box: {result.get('status', result.get('error'))}")

# Create a sketch on the top plane with a circle for the cutout
result = sketch_mgr.create_sketch("Top")
print(f"Create sketch: {result.get('status', result.get('error'))}")

# Draw a 20mm radius circle centered at origin
result = sketch_mgr.draw_circle(0, 0, 0.02)
print(f"Draw circle: {result.get('status', result.get('error'))}")

# Close sketch
result = sketch_mgr.close_sketch()
print(f"Close sketch: {result.get('status', result.get('error'))}")

# Create extruded cutout (30mm deep)
result = feat_mgr.create_extruded_cutout(0.03, "Normal")
print(f"Extruded cutout: {result}")

# ============================================================
# TEST 2: Through-all cutout
# ============================================================
print("\n" + "=" * 60)
print("TEST 2: THROUGH-ALL CUTOUT")
print("=" * 60)

# Create a sketch for a smaller through-all hole
result = sketch_mgr.create_sketch("Top")
print(f"Create sketch: {result.get('status', result.get('error'))}")

# Draw a 5mm radius circle offset from center
result = sketch_mgr.draw_circle(0.03, 0.03, 0.005)
print(f"Draw circle: {result.get('status', result.get('error'))}")

result = sketch_mgr.close_sketch()
print(f"Close sketch: {result.get('status', result.get('error'))}")

result = feat_mgr.create_extruded_cutout_through_all("Normal")
print(f"Through-all cutout: {result}")

# ============================================================
# TEST 3: Offset reference plane + sketch on it
# ============================================================
print("\n" + "=" * 60)
print("TEST 3: OFFSET REFERENCE PLANE")
print("=" * 60)

# Create an offset plane 25mm above the top plane
result = feat_mgr.create_ref_plane_by_offset(1, 0.025, "Normal")
print(f"Offset plane: {result}")

if "error" not in result:
    new_plane_idx = result["new_plane_index"]
    print(f"  New plane index: {new_plane_idx}")

    # Create a sketch on the new offset plane
    result = sketch_mgr.create_sketch_on_plane_index(new_plane_idx)
    print(f"  Sketch on offset plane: {result.get('status', result.get('error'))}")

    # Draw something on it
    result = sketch_mgr.draw_rectangle(-0.02, -0.02, 0.02, 0.02)
    print(f"  Draw rectangle: {result.get('status', result.get('error'))}")

    result = sketch_mgr.close_sketch()
    print(f"  Close sketch: {result.get('status', result.get('error'))}")

    # Extrude from the offset plane
    result = feat_mgr.create_extrude(0.02, "Add", "Normal")
    print(f"  Extrude from offset plane: {result}")

print("\n" + "=" * 60)
print("ALL TESTS COMPLETE")
print("=" * 60)
