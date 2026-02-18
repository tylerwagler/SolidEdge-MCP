# Composite Tool Consolidation Plan (Revised)

## Goal
Reduce tool count from **514 → ~130** by merging related tools into **composite tools** that use a `method` or `type` discriminator parameter. No functionality is lost — the backend calls remain the same.

> [!NOTE]
> The [Resource Migration](file:///c:/Users/tyler/Dev/repos/SolidEdge_MCP/reference/RESOURCE_MIGRATION_PLAN.md) already converted 52 read-only tools to MCP Resources.
> This plan operates on the remaining **514 action tools**. Additionally, ~30 remaining
> read-only `get_*` tools in query.py and export.py could be migrated to resources in
> a future Phase 2 (see Appendix A), bringing the final total to ~100 tools + ~82 resources.

---

## Composite Tool Pattern

Each composite tool dispatches to the appropriate backend method based on a `method` or `type` string parameter:

```python
@mcp.tool()
def create_extrude(
    method: str = "finite",
    distance: float = 0.01,
    direction: str = "Normal",
    wall_thickness: float | None = None,
    from_plane_index: int | None = None,
    to_plane_index: int | None = None,
    keypoint_or_body: int | None = None,
) -> dict:
    """Create an extruded protrusion.

    method: 'finite' | 'infinite' | 'through_next' | 'from_to' | 'thin_wall'
            | 'symmetric' | 'through_next_v2' | 'from_to_v2' | 'by_keypoint'
            | 'from_to_single' | 'through_next_single'
    """
    match method:
        case "finite":
            return feature_manager.create_extrude(distance, direction)
        case "infinite":
            return feature_manager.create_extrude_infinite(direction)
        # ... etc
```

---

## Part 1: Feature Tools (features.py) — 210 → 55

### Group 1: `create_extrude` — 11 → 1

| Current tool | Method value |
|---|---|
| `create_extrude` | `finite` (default) |
| `create_extrude_infinite` | `infinite` |
| `create_extrude_through_next` | `through_next` |
| `create_extrude_from_to` | `from_to` |
| `create_extrude_thin_wall` | `thin_wall` |
| `create_extrude_symmetric` | `symmetric` |
| `create_extrude_through_next_v2` | `through_next_v2` |
| `create_extrude_from_to_v2` | `from_to_v2` |
| `create_extrude_by_keypoint` | `by_keypoint` |
| `create_extrude_from_to_single` | `from_to_single` |
| `create_extrude_through_next_single` | `through_next_single` |

**Saves: 10**

### Group 2: `create_revolve` — 8 → 1

| Current tool | Method value |
|---|---|
| `create_revolve` | `full` (default) |
| `create_revolve_finite` | `finite` |
| `create_revolve_sync` | `sync` |
| `create_revolve_finite_sync` | `finite_sync` |
| `create_revolve_thin_wall` | `thin_wall` |
| `create_revolve_by_keypoint` | `by_keypoint` |
| `create_revolve_full` | `full_360` |
| `create_revolve_by_keypoint_sync` | `by_keypoint_sync` |

**Saves: 7**

### Group 3: `create_extruded_cutout` — 10 → 1

| Current tool | Method value |
|---|---|
| `create_extruded_cutout` | `finite` (default) |
| `create_extruded_cutout_through_all` | `through_all` |
| `create_extruded_cutout_through_next` | `through_next` |
| `create_extruded_cutout_from_to` | `from_to` |
| `create_extruded_cutout_from_to_v2` | `from_to_v2` |
| `create_extruded_cutout_by_keypoint` | `by_keypoint` |
| `create_extruded_cutout_through_next_single` | `through_next_single` |
| `create_extruded_cutout_multi_body` | `multi_body` |
| `create_extruded_cutout_from_to_multi_body` | `from_to_multi_body` |
| `create_extruded_cutout_through_all_multi_body` | `through_all_multi_body` |

**Saves: 9**

### Group 4: `create_revolved_cutout` — 6 → 1

| Current tool | Method value |
|---|---|
| `create_revolved_cutout` | `finite` (default) |
| `create_revolved_cutout_sync` | `sync` |
| `create_revolved_cutout_by_keypoint` | `by_keypoint` |
| `create_revolved_cutout_multi_body` | `multi_body` |
| `create_revolved_cutout_full` | `full` |
| `create_revolved_cutout_full_sync` | `full_sync` |

**Saves: 5**

### Group 5: `create_normal_cutout` — 5 → 1

| Current tool | Method value |
|---|---|
| `create_normal_cutout` | `finite` (default) |
| `create_normal_cutout_through_all` | `through_all` |
| `create_normal_cutout_from_to` | `from_to` |
| `create_normal_cutout_through_next` | `through_next` |
| `create_normal_cutout_by_keypoint` | `by_keypoint` |

**Saves: 4**

### Group 6: `create_lofted_cutout` — 2 → 1

| Current tool | Method value |
|---|---|
| `create_lofted_cutout` | `basic` (default) |
| `create_lofted_cutout_full` | `full` |

**Saves: 1**

### Group 7: `create_swept_cutout` — 2 → 1

| Current tool | Method value |
|---|---|
| `create_swept_cutout` | `basic` (default) |
| `create_swept_cutout_multi_body` | `multi_body` |

**Saves: 1**

### Group 8: `create_helix` — 8 → 1

| Current tool | Method value |
|---|---|
| `create_helix` | `finite` (default) |
| `create_helix_sync` | `sync` |
| `create_helix_thin_wall` | `thin_wall` |
| `create_helix_sync_thin_wall` | `sync_thin_wall` |
| `create_helix_from_to` | `from_to` |
| `create_helix_from_to_thin_wall` | `from_to_thin_wall` |
| `create_helix_from_to_sync` | `from_to_sync` |
| `create_helix_from_to_sync_thin_wall` | `from_to_sync_thin_wall` |

**Saves: 7**

### Group 9: `create_helix_cutout` — 4 → 1

| Current tool | Method value |
|---|---|
| `create_helix_cutout` | `finite` (default) |
| `create_helix_cutout_sync` | `sync` |
| `create_helix_cutout_from_to` | `from_to` |
| `create_helix_cutout_from_to_sync` | `from_to_sync` |

**Saves: 3**

### Group 10: `create_loft` — 2 → 1

| Current tool | Method value |
|---|---|
| `create_loft` | `solid` (default) |
| `create_loft_thin_wall` | `thin_wall` |

**Saves: 1**

### Group 11: `create_sweep` — 2 → 1

| Current tool | Method value |
|---|---|
| `create_sweep` | `solid` (default) |
| `create_sweep_thin_wall` | `thin_wall` |

**Saves: 1**

### Group 12: `create_extruded_surface` — 5 → 1

| Current tool | Method value |
|---|---|
| `create_extruded_surface` | `finite` (default) |
| `create_extruded_surface_from_to` | `from_to` |
| `create_extruded_surface_by_keypoint` | `by_keypoint` |
| `create_extruded_surface_by_curves` | `by_curves` |
| `create_extruded_surface_full` | `full` |

**Saves: 4**

### Group 13: `create_revolved_surface` — 5 → 1

| Current tool | Method value |
|---|---|
| `create_revolved_surface` | `finite` (default) |
| `create_revolved_surface_sync` | `sync` |
| `create_revolved_surface_by_keypoint` | `by_keypoint` |
| `create_revolved_surface_full` | `full` |
| `create_revolved_surface_full_sync` | `full_sync` |

**Saves: 4**

### Group 14: `create_lofted_surface` — 2 → 1

| Current tool | Method value |
|---|---|
| `create_lofted_surface` | `basic` (default) |
| `create_lofted_surface_v2` | `v2` |

**Saves: 1**

### Group 15: `create_swept_surface` — 2 → 1

| Current tool | Method value |
|---|---|
| `create_swept_surface` | `basic` (default) |
| `create_swept_surface_ex` | `ex` |

**Saves: 1**

### Group 16: `thicken` — 2 → 1

| Current tool | Method value |
|---|---|
| `thicken_surface` | `basic` (default) |
| `create_thicken_sync` | `sync` |

**Saves: 1**

### Group 17: `create_primitive` — 5 → 1

| Current tool | Shape value |
|---|---|
| `create_box_by_center` | `box_center` |
| `create_box_by_two_points` | `box_two_points` (default) |
| `create_box_by_three_points` | `box_three_points` |
| `create_cylinder` | `cylinder` |
| `create_sphere` | `sphere` |

**Saves: 4**

### Group 18: `create_primitive_cutout` — 3 → 1

| Current tool | Shape value |
|---|---|
| `create_box_cutout` | `box` (default) |
| `create_cylinder_cutout` | `cylinder` |
| `create_sphere_cutout` | `sphere` |

**Saves: 2**

### Group 19: `create_hole` — 12 → 1

| Current tool | Method value |
|---|---|
| `create_hole` | `finite` (default) |
| `create_hole_through_all` | `through_all` |
| `create_hole_from_to` | `from_to` |
| `create_hole_through_next` | `through_next` |
| `create_hole_sync` | `sync` |
| `create_hole_finite_ex` | `finite_ex` |
| `create_hole_from_to_ex` | `from_to_ex` |
| `create_hole_through_next_ex` | `through_next_ex` |
| `create_hole_through_all_ex` | `through_all_ex` |
| `create_hole_sync_ex` | `sync_ex` |
| `create_hole_multi_body` | `multi_body` |
| `create_hole_sync_multi_body` | `sync_multi_body` |

**Saves: 11**

### Group 20: `create_round` — 5 → 1

| Current tool | Method value |
|---|---|
| `create_round` | `all_edges` (default) |
| `create_round_on_face` | `on_face` |
| `create_variable_round` | `variable` |
| `create_round_blend` | `blend` |
| `create_round_surface_blend` | `surface_blend` |

**Saves: 4**

### Group 21: `create_chamfer` — 5 → 1

| Current tool | Method value |
|---|---|
| `create_chamfer` | `equal` (default) |
| `create_chamfer_on_face` | `on_face` |
| `create_chamfer_unequal` | `unequal` |
| `create_chamfer_unequal_on_face` | `unequal_on_face` |
| `create_chamfer_angle` | `angle` |

**Saves: 4**

### Group 22: `create_blend` — 3 → 1

| Current tool | Method value |
|---|---|
| `create_blend` | `basic` (default) |
| `create_blend_variable` | `variable` |
| `create_blend_surface` | `surface` |

**Saves: 2**

### Group 23: `delete_topology` — 4 → 1

| Current tool | Type value |
|---|---|
| `create_delete_hole` | `hole` (default) |
| `create_delete_hole_by_face` | `hole_by_face` |
| `create_delete_blend` | `blend` |
| `delete_faces` | `faces` |

**Saves: 3**

### Group 24: `create_ref_plane` — 17 → 1

| Current tool | Method value |
|---|---|
| `create_ref_plane_by_offset` | `offset` (default) |
| `create_ref_plane_by_angle` | `angle` |
| `create_ref_plane_by_3_points` | `three_points` |
| `create_ref_plane_midplane` | `midplane` |
| `create_ref_plane_normal_to_curve` | `normal_to_curve` |
| `create_ref_plane_normal_at_distance` | `normal_at_distance` |
| `create_ref_plane_normal_at_arc_ratio` | `normal_at_arc_ratio` |
| `create_ref_plane_normal_at_distance_along` | `normal_at_distance_along` |
| `create_ref_plane_parallel_by_tangent` | `parallel_by_tangent` |
| `create_ref_plane_normal_at_keypoint` | `normal_at_keypoint` |
| `create_ref_plane_tangent_cylinder_angle` | `tangent_cylinder_angle` |
| `create_ref_plane_tangent_cylinder_keypoint` | `tangent_cylinder_keypoint` |
| `create_ref_plane_tangent_surface_keypoint` | `tangent_surface_keypoint` |
| `create_ref_plane_normal_at_distance_v2` | `normal_at_distance_v2` |
| `create_ref_plane_normal_at_arc_ratio_v2` | `normal_at_arc_ratio_v2` |
| `create_ref_plane_normal_at_distance_along_v2` | `normal_at_distance_along_v2` |
| `create_ref_plane_tangent_parallel` | `tangent_parallel` |

**Saves: 16**

### Group 25: `create_flange` — 8 → 1

| Current tool | Method value |
|---|---|
| `create_flange` | `basic` (default) |
| `create_flange_by_match_face` | `by_match_face` |
| `create_flange_sync` | `sync` |
| `create_flange_by_face` | `by_face` |
| `create_flange_with_bend_calc` | `with_bend_calc` |
| `create_flange_sync_with_bend_calc` | `sync_with_bend_calc` |
| `create_flange_match_face_with_bend` | `match_face_with_bend` |
| `create_flange_by_face_with_bend` | `by_face_with_bend` |

**Saves: 7**

### Group 26: `create_contour_flange` — 5 → 1

| Current tool | Method value |
|---|---|
| `create_contour_flange_ex` | `ex` (default) |
| `create_contour_flange_sync` | `sync` |
| `create_contour_flange_sync_with_bend` | `sync_with_bend` |
| `create_contour_flange_v3` | `v3` |
| `create_contour_flange_sync_ex` | `sync_ex` |

**Saves: 4**

### Group 27: `create_sheet_metal_base` — 4 → 1

| Current tool | Type value |
|---|---|
| `create_base_flange` | `flange` (default) |
| `create_base_tab` | `tab` |
| `create_base_contour_flange_advanced` | `contour_advanced` |
| `create_base_tab_multi_profile` | `tab_multi_profile` |

**Saves: 3**

### Group 28: `create_lofted_flange` — 3 → 1

| Current tool | Method value |
|---|---|
| `create_lofted_flange` | `basic` (default) |
| `create_lofted_flange_advanced` | `advanced` |
| `create_lofted_flange_ex` | `ex` |

**Saves: 2**

### Group 29: `create_bend` — 2 → 1

| Current tool | Method value |
|---|---|
| `create_bend` | `basic` (default) |
| `create_bend_with_calc` | `with_calc` |

**Saves: 1**

### Group 30: `create_slot` — 5 → 1

| Current tool | Method value |
|---|---|
| `create_slot` | `basic` (default) |
| `create_slot_ex` | `ex` |
| `create_slot_sync` | `sync` |
| `create_slot_multi_body` | `multi_body` |
| `create_slot_sync_multi_body` | `sync_multi_body` |

**Saves: 4**

### Group 31: `create_thread` — 2 → 1

| Current tool | Method value |
|---|---|
| `create_thread` | `basic` (default) |
| `create_thread_ex` | `ex` |

**Saves: 1**

### Group 32: `create_drawn_cutout` — 2 → 1

| Current tool | Method value |
|---|---|
| `create_drawn_cutout` | `basic` (default) |
| `create_drawn_cutout_ex` | `ex` |

**Saves: 1**

### Group 33: `create_dimple` — 2 → 1

| Current tool | Method value |
|---|---|
| `create_dimple` | `basic` (default) |
| `create_dimple_ex` | `ex` |

**Saves: 1**

### Group 34: `create_louver` — 2 → 1

| Current tool | Method value |
|---|---|
| `create_louver` | `basic` (default) |
| `create_louver_sync` | `sync` |

**Saves: 1**

### Group 35: `create_pattern` — 10 → 1

| Current tool | Method value |
|---|---|
| `create_pattern_rectangular` | `rectangular` |
| `create_pattern_circular` | `circular` |
| `create_pattern_rectangular_ex` | `rectangular_ex` (default) |
| `create_pattern_circular_ex` | `circular_ex` |
| `create_pattern_duplicate` | `duplicate` |
| `create_pattern_by_fill` | `by_fill` |
| `create_pattern_by_table` | `by_table` |
| `create_pattern_by_table_sync` | `by_table_sync` |
| `create_pattern_by_fill_ex` | `by_fill_ex` |
| `create_pattern_by_curve_ex` | `by_curve_ex` |

**Saves: 9**

### Group 36: `create_mirror` — 2 → 1

| Current tool | Method value |
|---|---|
| `create_mirror` | `basic` (default) |
| `create_mirror_sync_ex` | `sync_ex` |

**Saves: 1**

### Group 37: `create_thin_wall` — 2 → 1

| Current tool | Method value |
|---|---|
| `create_thin_wall` | `basic` (default) |
| `create_thin_wall_with_open_faces` | `with_open_faces` |

**Saves: 1**

### Group 38: `face_operation` — 2 → 1

| Current tool | Type value |
|---|---|
| `create_face_rotate_by_points` | `rotate_by_points` |
| `create_face_rotate_by_edge` | `rotate_by_edge` |

**Saves: 1**

### Group 39: `add_body` — 5 → 1

| Current tool | Method value |
|---|---|
| `add_body` | `basic` (default) |
| `add_body_by_mesh` | `by_mesh` |
| `add_body_feature` | `feature` |
| `add_by_construction` | `construction` |
| `add_body_by_tag` | `by_tag` |

**Saves: 4**

### Group 40: `simplify` — 4 → 1

| Current tool | Method value |
|---|---|
| `auto_simplify` | `auto` (default) |
| `simplify_enclosure` | `enclosure` |
| `simplify_duplicate` | `duplicate` |
| `local_simplify_enclosure` | `local_enclosure` |

**Saves: 3**

### Group 41: `manage_feature` — 6 → 1

| Current tool | Action value |
|---|---|
| `delete_feature` | `delete` |
| `feature_suppress` | `suppress` |
| `feature_unsuppress` | `unsuppress` |
| `feature_reorder` | `reorder` |
| `feature_rename` | `rename` |
| `convert_feature_type` | `convert` |

**Saves: 5**

### Standalone features (14 tools, no merge)

These tools have unique parameter signatures and no close relatives:

| Tool | Reason standalone |
|---|---|
| `create_web_network` | Unique sheet metal feature |
| `create_hem` | Unique edge treatment |
| `create_jog` | Unique bend operation |
| `create_close_corner` | Unique corner treatment |
| `create_multi_edge_flange` | Unique multi-edge feature |
| `convert_part_to_sheet_metal` | Unique conversion operation |
| `create_bead` | Unique sheet metal forming |
| `create_gusset` | Unique reinforcement |
| `create_emboss` | Raised surface imprint |
| `create_etch` | Recessed surface imprint |
| `create_rib` | Unique reinforcement |
| `create_lip` | Unique edge feature |
| `create_split` | Unique split operation |
| `create_draft_angle` | Unique mold draft |

**features.py total: 41 composites + 14 standalone = 55 tools (from 210). Saves 155.**

---

## Part 2: Sketching Tools (sketching.py) — 37 → 8

### Group 42: `manage_sketch` — 3 → 1

| Current tool | Action value |
|---|---|
| `create_sketch` | `create` (default) |
| `close_sketch` | `close` |
| `create_sketch_on_plane` | `create_on_plane` |

**Saves: 2**

### Group 43: `draw` — 12 → 1

| Current tool | Shape value |
|---|---|
| `draw_line` | `line` (default) |
| `draw_circle` | `circle` |
| `draw_rectangle` | `rectangle` |
| `draw_arc` | `arc` |
| `draw_polygon` | `polygon` |
| `draw_ellipse` | `ellipse` |
| `draw_spline` | `spline` |
| `draw_arc_by_3_points` | `arc_3pt` |
| `draw_circle_by_2_points` | `circle_2pt` |
| `draw_circle_by_3_points` | `circle_3pt` |
| `draw_point` | `point` |
| `draw_construction_line` | `construction_line` |

**Saves: 11**

### Group 44: `sketch_modify` — 9 → 1

| Current tool | Action value |
|---|---|
| `sketch_fillet` | `fillet` |
| `sketch_chamfer` | `chamfer` |
| `sketch_offset` | `offset` |
| `sketch_rotate` | `rotate` |
| `sketch_scale` | `scale` |
| `sketch_mirror` | `mirror` |
| `sketch_paste` | `paste` |
| `mirror_spline` | `mirror_spline` |
| `offset_sketch_2d` | `offset_2d` |

**Saves: 8**

### Group 45: `sketch_constraint` — 2 → 1

| Current tool | Type value |
|---|---|
| `add_constraint` | `geometric` (default) |
| `add_keypoint_constraint` | `keypoint` |

**Saves: 1**

### Group 46: `sketch_project` — 7 → 1

| Current tool | Source value |
|---|---|
| `project_edge` | `edge` |
| `include_edge` | `include_edge` |
| `project_ref_plane` | `ref_plane` |
| `project_silhouette_edges` | `silhouette` |
| `include_region_faces` | `region_faces` |
| `chain_locate` | `chain` |
| `convert_to_curve` | `to_curve` |

**Saves: 6**

### Standalone sketching (4 tools, no merge)

| Tool | Reason standalone |
|---|---|
| `set_axis_of_revolution` | Unique revolve axis setup |
| `set_profile_visibility` | Unique visibility toggle |
| `clean_sketch_geometry` | Unique utility |
| `get_ordered_geometry` | Unique query |

**sketching.py total: 5 composites + 4 standalone = 9 tools (from 37). Saves 28.**

---

## Part 3: Export & Drawing Tools (export.py) — 90 → 20

### Group 47: `export_file` — 8 → 1

| Current tool | Format value |
|---|---|
| `export_step` | `step` (default) |
| `export_stl` | `stl` |
| `export_iges` | `iges` |
| `export_pdf` | `pdf` |
| `export_dxf` | `dxf` |
| `export_parasolid` | `parasolid` |
| `export_jt` | `jt` |
| `export_flat_dxf` | `flat_dxf` |

**Saves: 7**

### Group 48: `add_drawing_view` — 8 → 1

| Current tool | Type value |
|---|---|
| `add_assembly_drawing_view` | `assembly` |
| `add_assembly_drawing_view_ex` | `assembly_ex` |
| `add_drawing_view_with_config` | `with_config` |
| `add_projected_view` | `projected` |
| `add_detail_view` | `detail` |
| `add_auxiliary_view` | `auxiliary` |
| `add_draft_view` | `draft` |
| `add_by_draft_view` | `by_draft_view` |

**Saves: 7**

### Group 49: `manage_drawing_view` — 12 → 1

| Current tool | Action value |
|---|---|
| `get_drawing_view_model_link` | `get_model_link` |
| `show_tangent_edges` | `show_tangent_edges` |
| `set_drawing_view_scale` | `set_scale` |
| `delete_drawing_view` | `delete` |
| `update_drawing_view` | `update` |
| `move_drawing_view` | `move` |
| `show_hidden_edges` | `show_hidden_edges` |
| `set_drawing_view_display_mode` | `set_display_mode` |
| `set_drawing_view_orientation` | `set_orientation` |
| `activate_drawing_view` | `activate` |
| `deactivate_drawing_view` | `deactivate` |
| `get_drawing_view_dimensions` | `get_dimensions` |

**Saves: 11**

### Group 50: `add_annotation` — 14 → 1

| Current tool | Type value |
|---|---|
| `add_text_box` | `text_box` |
| `add_leader` | `leader` |
| `add_dimension` | `dimension` |
| `add_balloon` | `balloon` |
| `add_note` | `note` |
| `add_angular_dimension` | `angular_dimension` |
| `add_radial_dimension` | `radial_dimension` |
| `add_diameter_dimension` | `diameter_dimension` |
| `add_ordinate_dimension` | `ordinate_dimension` |
| `add_center_mark` | `center_mark` |
| `add_centerline` | `centerline` |
| `add_surface_finish_symbol` | `surface_finish` |
| `add_weld_symbol` | `weld_symbol` |
| `add_geometric_tolerance` | `geometric_tolerance` |

**Saves: 13**

### Group 51: `add_2d_dimension` — 4 → 1

| Current tool | Type value |
|---|---|
| `add_distance_dimension` | `distance` |
| `add_length_dimension` | `length` |
| `add_radius_dimension_2d` | `radius` |
| `add_angle_dimension_2d` | `angle` |

**Saves: 3**

### Group 52: `camera_control` — 10 → 1

| Current tool | Action value |
|---|---|
| `set_view_orientation` | `set_orientation` |
| `zoom_fit` | `zoom_fit` |
| `zoom_to_selection` | `zoom_to_selection` |
| `rotate_view` | `rotate` |
| `pan_view` | `pan` |
| `zoom_view` | `zoom` |
| `refresh_view` | `refresh` |
| `set_camera` | `set_camera` |
| `begin_camera_dynamics` | `begin_dynamics` |
| `end_camera_dynamics` | `end_dynamics` |

**Saves: 9**

### Group 53: `display_control` — 4 → 1

| Current tool | Action value |
|---|---|
| `set_display_mode` | `set_mode` |
| `set_view_background` | `set_background` |
| `transform_model_to_screen` | `model_to_screen` |
| `transform_screen_to_model` | `screen_to_model` |

**Saves: 3**

### Group 54: `manage_sheet` — 3 → 1

| Current tool | Action value |
|---|---|
| `activate_sheet` | `activate` |
| `rename_sheet` | `rename` |
| `delete_sheet` | `delete` |

**Saves: 2**

### Group 55: `print_control` — 4 → 1

| Current tool | Action value |
|---|---|
| `print_drawing` | `print` |
| `set_printer` | `set_printer` |
| `get_printer` | `get_printer` |
| `set_paper_size` | `set_paper_size` |

**Saves: 3**

### Group 56: `query_sheet` — 9 → 1

| Current tool | Type value |
|---|---|
| `get_sheet_dimensions` | `dimensions` |
| `get_sheet_balloons` | `balloons` |
| `get_sheet_text_boxes` | `text_boxes` |
| `get_sheet_drawing_objects` | `drawing_objects` |
| `get_sheet_sections` | `sections` |
| `get_lines2d` | `lines2d` |
| `get_circles2d` | `circles2d` |
| `get_arcs2d` | `arcs2d` |
| `get_section_cuts` | `section_cuts` |

**Saves: 8**

### Group 57: `manage_annotation_data` — 4 → 1

| Current tool | Action value |
|---|---|
| `add_symbol` | `add_symbol` |
| `get_symbols` | `get_symbols` |
| `get_pmi_info` | `get_pmi` |
| `set_pmi_visibility` | `set_pmi_visibility` |

**Saves: 3**

### Group 58: `add_smart_frame` — 2 → 1

| Current tool | Method value |
|---|---|
| `add_smart_frame` | `two_point` (default) |
| `add_smart_frame_by_origin` | `by_origin` |

**Saves: 1**

### Standalone export/drawing (8 tools, no merge)

| Tool | Reason standalone |
|---|---|
| `create_drawing` | Unique drawing creation |
| `add_draft_sheet` | Unique sheet creation |
| `capture_screenshot` | Unique image capture |
| `create_parts_list` | Unique BOM creation |
| `set_face_texture` | Unique appearance |
| `create_bend_table` | Unique table creation |
| `align_drawing_views` | Unique alignment |
| `add_section_cut` | Unique section creation |

**export.py total: 12 composites + 8 standalone = 20 tools (from 90). Saves 70.**

---

## Part 4: Query & Edit Tools (query.py) — 69 → 17

### Group 59: `measure` — 2 → 1

| Current tool | Type value |
|---|---|
| `measure_distance` | `distance` (default) |
| `measure_angle` | `angle` |

**Saves: 1**

### Group 60: `manage_variable` — 8 → 1

| Current tool | Action value |
|---|---|
| `set_variable` | `set` (default) |
| `add_variable` | `add` |
| `query_variables` | `query` |
| `rename_variable` | `rename` |
| `translate_variable` | `translate` |
| `copy_variable_to_clipboard` | `copy_clipboard` |
| `add_variable_from_clipboard` | `add_from_clipboard` |
| `set_variable_formula` | `set_formula` |

**Saves: 7**

### Group 61: `manage_property` — 3 → 1

| Current tool | Action value |
|---|---|
| `set_document_property` | `set_document` |
| `set_custom_property` | `set_custom` |
| `delete_custom_property` | `delete_custom` |

**Saves: 2**

### Group 62: `manage_material` — 4 → 1

| Current tool | Action value |
|---|---|
| `set_material_density` | `set_density` |
| `set_material` | `set` (default) |
| `set_material_by_name` | `set_by_name` |
| `get_material_library` | `get_library` |

**Saves: 3**

### Group 63: `set_appearance` — 4 → 1

| Current tool | Target value |
|---|---|
| `set_body_color` | `body_color` |
| `set_face_color` | `face_color` |
| `set_body_opacity` | `opacity` |
| `set_body_reflectivity` | `reflectivity` |

**Saves: 3**

### Group 64: `manage_layer` — 4 → 1

| Current tool | Action value |
|---|---|
| `add_layer` | `add` |
| `activate_layer` | `activate` |
| `set_layer_properties` | `set_properties` |
| `delete_layer` | `delete` |

**Saves: 3**

### Group 65: `select_set` — 10 → 1

| Current tool | Action value |
|---|---|
| `clear_select_set` | `clear` |
| `select_add` | `add` |
| `select_remove` | `remove` |
| `select_all` | `all` |
| `select_copy` | `copy` |
| `select_cut` | `cut` |
| `select_delete` | `delete` |
| `select_suspend_display` | `suspend_display` |
| `select_resume_display` | `resume_display` |
| `select_refresh_display` | `refresh_display` |

**Saves: 9**

### Group 66: `edit_feature_extent` — 10 → 1

| Current tool | Property value |
|---|---|
| `get_direction1_extent` | `get_direction1` |
| `set_direction1_extent` | `set_direction1` |
| `get_direction2_extent` | `get_direction2` |
| `set_direction2_extent` | `set_direction2` |
| `get_thin_wall_options` | `get_thin_wall` |
| `set_thin_wall_options` | `set_thin_wall` |
| `get_from_face_offset` | `get_from_face` |
| `set_from_face_offset` | `set_from_face` |
| `get_body_array` | `get_body_array` |
| `set_body_array` | `set_body_array` |

**Saves: 9**

### Group 67: `manage_feature_tree` — 3 → 1

| Current tool | Action value |
|---|---|
| `rename_feature` | `rename` |
| `suppress_feature` | `suppress` |
| `unsuppress_feature` | `unsuppress` |

**Saves: 2**

### Group 68: `query_edge` — 5 → 1

| Current tool | Property value |
|---|---|
| `get_edge_endpoints` | `endpoints` |
| `get_edge_length` | `length` |
| `get_edge_tangent` | `tangent` |
| `get_edge_geometry` | `geometry` |
| `get_edge_curvature` | `curvature` |

**Saves: 4**

### Group 69: `query_face` — 4 → 1

| Current tool | Property value |
|---|---|
| `get_face_normal` | `normal` |
| `get_face_geometry` | `geometry` |
| `get_face_loops` | `loops` |
| `get_face_curvature` | `curvature` |

**Saves: 3**

### Group 70: `query_body` — 6 → 1

| Current tool | Property value |
|---|---|
| `get_body_extreme_point` | `extreme_point` |
| `get_faces_by_ray` | `faces_by_ray` |
| `get_body_shells` | `shells` |
| `get_body_vertices` | `vertices` |
| `get_shell_info` | `shell_info` |
| `is_point_inside_body` | `point_inside` |

**Saves: 5**

### Group 71: `query_bspline` — 2 → 1

| Current tool | Type value |
|---|---|
| `get_bspline_curve_info` | `curve` |
| `get_bspline_surface_info` | `surface` |

**Saves: 1**

### Standalone query (5 tools, no merge)

| Tool | Reason standalone |
|---|---|
| `get_vertex_point` | Single vertex query |
| `set_modeling_mode` | Unique mode toggle |
| `recompute` | Unique rebuild |
| `get_body_facet_data` | Unique tessellation query |

**query.py total: 13 composites + 4 standalone = 17 tools (from 69). Saves 52.**

---

## Part 5: Assembly Tools (assembly.py) — 77 → 14

### Group 72: `add_assembly_component` — 7 → 1

| Current tool | Method value |
|---|---|
| `assembly_add_component` | `basic` (default) |
| `assembly_add_component_with_transform` | `with_transform` |
| `assembly_add_family_member` | `family` |
| `assembly_add_family_with_transform` | `family_with_transform` |
| `assembly_add_family_with_matrix` | `family_with_matrix` |
| `assembly_add_by_template` | `by_template` |
| `assembly_add_adjustable_part` | `adjustable` |

**Saves: 6**

### Group 73: `manage_component` — 6 → 1

| Current tool | Action value |
|---|---|
| `assembly_delete_component` | `delete` |
| `assembly_replace_component` | `replace` |
| `assembly_suppress_component` | `suppress` |
| `assembly_reorder_occurrence` | `reorder` |
| `assembly_make_writable` | `make_writable` |
| `assembly_swap_family_member` | `swap_family` |

**Saves: 5**

### Group 74: `query_component` — 18 → 1

| Current tool | Property value |
|---|---|
| `assembly_list_components` | `list` (default) |
| `assembly_get_component_info` | `info` |
| `assembly_get_bounding_box` | `bounding_box` |
| `assembly_get_bom` | `bom` |
| `assembly_get_structured_bom` | `structured_bom` |
| `assembly_get_tree` | `tree` |
| `assembly_get_component_transform` | `transform` |
| `assembly_get_occurrence_count` | `count` |
| `assembly_is_subassembly` | `is_subassembly` |
| `assembly_get_component_display_name` | `display_name` |
| `assembly_get_occurrence_document` | `document` |
| `assembly_get_sub_occurrences` | `sub_occurrences` |
| `assembly_get_occurrence_bodies` | `bodies` |
| `assembly_get_occurrence_style` | `style` |
| `assembly_is_tube` | `is_tube` |
| `assembly_get_adjustable_part` | `adjustable_part` |
| `assembly_get_face_style` | `face_style` |
| `assembly_get_occurrence` | `occurrence` |

**Saves: 17**

### Group 75: `set_component_appearance` — 2 → 1

| Current tool | Property value |
|---|---|
| `assembly_set_component_visibility` | `visibility` |
| `assembly_set_component_color` | `color` |

**Saves: 1**

### Group 76: `transform_component` — 7 → 1

| Current tool | Method value |
|---|---|
| `assembly_update_component_position` | `update_position` |
| `assembly_set_component_transform` | `set_transform` |
| `assembly_set_component_origin` | `set_origin` |
| `assembly_occurrence_move` | `move` |
| `assembly_occurrence_rotate` | `rotate` |
| `assembly_put_transform_euler` | `put_euler` |
| `assembly_put_origin` | `put_origin` |

**Saves: 6**

### Group 77: `add_assembly_constraint` — 5 → 1

| Current tool | Type value |
|---|---|
| `assembly_create_mate` | `mate` |
| `assembly_add_align` | `align` |
| `assembly_add_planar_align` | `planar_align` |
| `assembly_add_axial_align` | `axial_align` |
| `assembly_add_angle` | `angle` |

**Saves: 4**

### Group 78: `add_assembly_relation` — 6 → 1

| Current tool | Type value |
|---|---|
| `assembly_add_planar_relation` | `planar` |
| `assembly_add_axial_relation` | `axial` |
| `assembly_add_angular_relation` | `angular` |
| `assembly_add_point_relation` | `point` |
| `assembly_add_tangent_relation` | `tangent` |
| `assembly_add_gear_relation` | `gear` |

**Saves: 5**

### Group 79: `manage_relation` — 13 → 1

| Current tool | Action value |
|---|---|
| `assembly_get_relations` | `list` |
| `assembly_get_relation_info` | `info` |
| `assembly_delete_relation` | `delete` |
| `assembly_get_relation_offset` | `get_offset` |
| `assembly_set_relation_offset` | `set_offset` |
| `assembly_get_relation_angle` | `get_angle` |
| `assembly_set_relation_angle` | `set_angle` |
| `assembly_get_normals_aligned` | `get_normals` |
| `assembly_set_normals_aligned` | `set_normals` |
| `assembly_suppress_relation` | `suppress` |
| `assembly_unsuppress_relation` | `unsuppress` |
| `assembly_get_relation_geometry` | `get_geometry` |
| `assembly_get_gear_ratio` | `get_gear_ratio` |

**Saves: 12**

### Group 80: `assembly_feature` — 9 → 1

| Current tool | Type value |
|---|---|
| `assembly_create_extruded_cutout` | `extruded_cutout` |
| `assembly_create_revolved_cutout` | `revolved_cutout` |
| `assembly_create_hole` | `hole` |
| `assembly_create_extruded_protrusion` | `extruded_protrusion` |
| `assembly_create_revolved_protrusion` | `revolved_protrusion` |
| `assembly_create_mirror` | `mirror` |
| `assembly_create_pattern` | `pattern` |
| `assembly_create_swept_protrusion` | `swept_protrusion` |
| `assembly_recompute_features` | `recompute` |

**Saves: 8**

### Standalone assembly (4 tools, no merge)

| Tool | Reason standalone |
|---|---|
| `assembly_pattern_component` | Unique pattern placement |
| `assembly_mirror_component` | Unique mirror placement |
| `assembly_ground_component` | Unique ground constraint |
| `assembly_check_interference` | Unique analysis |

**assembly.py total: 9 composites + 4 standalone = 13 tools (from 77). Saves 64.**

---

## Part 6: Document & Connection Tools — 29 → 14

### Group 81: `create_document` — 5 → 1

| Current tool | Type value |
|---|---|
| `create_part_document` | `part` (default) |
| `create_assembly_document` | `assembly` |
| `create_sheet_metal_document` | `sheet_metal` |
| `create_draft_document` | `draft` |
| `create_weldment_document` | `weldment` |

**Saves: 4**

### Group 82: `open_document` — 2 → 1

| Current tool | Method value |
|---|---|
| `open_document` | `foreground` (default) |
| `open_in_background` | `background` |

**Saves: 1**

### Group 83: `close_document` — 2 → 1

| Current tool | Scope value |
|---|---|
| `close_document` | `active` (default) |
| `close_all_documents` | `all` |

**Saves: 1**

### Group 84: `app_command` — 3 → 1

| Current tool | Action value |
|---|---|
| `start_command` | `start` |
| `abort_command` | `abort` |
| `do_idle` | `idle` |

**Saves: 2**

### Group 85: `app_config` — 8 → 1

| Current tool | Property value |
|---|---|
| `set_performance_mode` | `set_performance` |
| `get_active_environment` | `get_environment` |
| `get_status_bar` | `get_status_bar` |
| `set_status_bar` | `set_status_bar` |
| `get_visible` | `get_visible` |
| `set_visible` | `set_visible` |
| `get_global_parameter` | `get_global` |
| `set_global_parameter` | `set_global` |

**Saves: 7**

### Standalone (9 tools)

| Tool | Reason standalone |
|---|---|
| `connect_to_solidedge` | Entry point |
| `quit_application` | Exit point |
| `disconnect_from_solidedge` | Session cleanup |
| `activate_application` | Window activation |
| `save_document` | Document persistence |
| `activate_document` | Document switching |
| `undo` | Undo action |
| `redo` | Redo action |
| `import_file` | File import |

**documents.py + connection.py total: 5 composites + 9 standalone = 14 tools (from 29). Saves 15.**

---

## Part 7: Diagnostics (diagnostics.py) — 2 → 2

No merges. Both tools have distinct purposes.

| Tool |
|---|
| `diagnose_api` |
| `diagnose_feature_tool` |

---

## Summary

| Module | Before | After | Saved |
|---|---|---|---|
| features.py | 210 | 55 | 155 |
| sketching.py | 37 | 9 | 28 |
| export.py | 90 | 20 | 70 |
| query.py | 69 | 17 | 52 |
| assembly.py | 77 | 13 | 64 |
| documents.py + connection.py | 29 | 14 | 15 |
| diagnostics.py | 2 | 2 | 0 |
| **Total** | **514** | **130** | **384** |

**This plan: 514 → 130 tools (75% reduction).**

Combined with existing 52 MCP Resources: **130 tools + 52 resources = 182 endpoints**.

With optional Phase 2 resource migration (~30 more read-only tools): **~100 tools + ~82 resources = 182 endpoints**.

---

## Implementation Order

### Phase 1: High-Impact, Low-Risk (saves ~120 tools)
1. **`export_file`** (Group 47) — simplest, isolated, good proof of concept
2. **`draw`** (Group 43) — high impact, self-contained
3. **`create_document`** (Group 81) — trivial merge
4. **`select_set`** (Group 65) — self-contained
5. **`create_ref_plane`** (Group 24) — biggest single group (17→1)
6. **`create_hole`** (Group 19) — 12→1
7. **`add_annotation`** (Group 50) — 14→1
8. **`query_component`** (Group 74) — 18→1
9. **`manage_relation`** (Group 79) — 13→1
10. **`manage_drawing_view`** (Group 49) — 12→1

### Phase 2: Feature Groups (saves ~150 tools)
11. Extrude/Revolve groups (Groups 1-2)
12. Cutout groups (Groups 3-9)
13. Surface groups (Groups 12-16)
14. Sheet metal groups (Groups 25-34)
15. Remaining feature groups (Groups 17-23, 35-41)

### Phase 3: Remaining Merges (saves ~110 tools)
16. Sketching groups (Groups 42-46)
17. Assembly groups (Groups 72-80)
18. Query groups (Groups 59-71)
19. View/camera groups (Groups 52-53)
20. Connection/document groups (Groups 82-85)

---

## Files Modified

| File | Groups |
|---|---|
| `tools/features.py` | 1-41 |
| `tools/sketching.py` | 42-46 |
| `tools/export.py` | 47-58 |
| `tools/query.py` | 59-71 |
| `tools/assembly.py` | 72-80 |
| `tools/documents.py` | 81-83 |
| `tools/connection.py` | 84-85 |

---

## Verification

1. **Golden Master**: Dump current tool names to `baseline_tools.txt` before starting
2. **Functional Check**: For each composite tool, verify all `method` values dispatch correctly
3. **Tool Count**: `grep -c "mcp.tool()" src/solidedge_mcp/tools/*.py` should equal ~130
4. **Unit Tests**: `uv run pytest tests/unit` must remain green — update test calls to use new composite tool signatures
5. **Integration Smoke**: Connect to Solid Edge and run basic workflow (create doc → sketch → extrude → export)

---

## Appendix A: Phase 2 Resource Migration Candidates

These ~30 read-only `get_*` tools in query.py and export.py could be migrated to MCP Resources:

**From query.py:**
- `get_direction1_extent`, `get_direction2_extent`, `get_thin_wall_options`, `get_from_face_offset`, `get_body_array`
- `get_material_library`
- `get_edge_endpoints`, `get_edge_length`, `get_edge_tangent`, `get_edge_geometry`, `get_edge_curvature`
- `get_face_normal`, `get_face_geometry`, `get_face_loops`, `get_face_curvature`
- `get_vertex_point`, `get_body_extreme_point`, `get_body_shells`, `get_body_vertices`
- `get_shell_info`, `get_bspline_curve_info`, `get_bspline_surface_info`
- `get_body_facet_data`

**From export.py:**
- `get_drawing_view_model_link`, `get_drawing_view_dimensions`
- `get_sheet_dimensions`, `get_sheet_balloons`, `get_sheet_text_boxes`, `get_sheet_drawing_objects`, `get_sheet_sections`
- `get_lines2d`, `get_circles2d`, `get_arcs2d`, `get_section_cuts`
- `get_symbols`, `get_pmi_info`, `get_printer`
