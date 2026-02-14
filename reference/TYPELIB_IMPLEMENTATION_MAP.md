# Solid Edge Type Library Implementation Map

Generated: 2026-02-13 | Source: 40 type libraries, 2,240 interfaces, 21,237 methods
Current: 556 MCP tools implemented

This document maps every actionable COM API surface from the Solid Edge type libraries
against our current MCP tool coverage. It identifies gaps and prioritizes what to implement next.

## Coverage Summary

|     Category          | Sections | Complete | Partial | Not Started | Methods (impl/total) |
|-----------------------|----------|----------|---------|-------------|----------------------|
| **Part Features**     |    52    |    36    |    13   |      3      |      192 / 201       |
| **Assembly**          |    11    |     4    |     4   |      3      |       45 /  61       |
| **Draft/Drawing**     |     5    |     3    |     2   |      0      |       51 /  53       |
| **Framework/App**     |     7    |     5    |     2   |      0      |       53 /  55       |
| **Geometry/Topology** |     2    |     0    |     2   |      0      |       15 /  19       |
| **Total**             | **77**   |  **48**  |  **23** |    **6**    | **356 / 389 (92%)**  |

**556 MCP tools** registered (many tools cover multiple methods or provide capabilities
beyond what the type library tracks, e.g. primitives, view controls, export formats).

## Tool Count by Category (556 total)

| Category                  | Count | Tools |
|:--------------------------|:-----:|:---|
| **Connection/Application**| 19    | Connect, disconnect, app info, quit, is_connected, process_info, install_info, start_command, set_performance_mode, do_idle, activate, abort_command, active_environment, status_bar (get/set), visible (get/set), global_parameter (get/set) |
| **Document Management**   | 13    | Create (part, assembly, sheet metal, draft), open, save, close, list, activate, undo, redo |
| **Sketching**             | 30    | Lines, circles, arcs (multiple), rects, polygons, ellipses, splines, points, constraints (9 types), fillet, chamfer, mirror, construction, hide profile, project_edge, include_edge, get_sketch_matrix, clean_sketch_geometry, project_silhouette_edges, include_region_faces, chain_locate, convert_to_curve |
| **Basic Primitives**      | 10    | Box (3 variants), cylinder, sphere, box cutout (3 variants), cylinder cutout, sphere cutout |
| **Extrusions**            | 8     | Finite, infinite, symmetric, thin-wall, extruded surface, through-next v2, from-to v2, by-keypoint |
| **Revolves**              | 7     | Basic, finite, sync, thin-wall, by-keypoint, full |
| **Cutouts**               | 20    | Extruded finite/through-all/through-next/from-to v2/by-keypoint, revolved/revolved-sync/revolved-by-keypoint, normal/normal-through-all/normal-from-to/normal-through-next/normal-by-keypoint, lofted/lofted-full, swept/swept-multi-body, helix/helix-sync/helix-from-to |
| **Rounds/Chamfers/Holes** | 24    | Round (all/face), variable round, round blend, round surface blend, blend variable, blend surface, chamfer (equal/unequal/angle/face), chamfer unequal on face, hole (finite/from-to/through-next/through-all/sync), hole-ex (finite/from-to/through-next/through-all/sync), hole multi-body, hole sync multi-body |
| **Reference Planes**      | 9     | Offset, normal-to-curve, angle, 3-points, mid-plane, normal-at-keypoint, tangent-cylinder-angle, tangent-cylinder-keypoint, tangent-surface-keypoint |
| **Loft**                  | 3     | Basic, thin-wall, lofted-cutout-full |
| **Sweep**                 | 3     | Basic, thin-wall, swept-cutout-multi-body |
| **Helix/Spiral**          | 8     | Basic, sync, thin-wall variants, from-to, from-to-thin-wall, helix-cutout-sync, helix-cutout-from-to |
| **Construction Surfaces** | 11    | Extruded surface (basic, from-to, by-keypoint, by-curves, full), revolved surface (basic, sync, by-keypoint), lofted surface (basic, v2), swept surface (basic, ex) |
| **Sheet Metal**           | 25    | Base flange/tab, lofted flange, web network, advanced variants, emboss, flange, flange-by-match-face, flange-sync, flange-by-face, flange-with-bend-calc, flange-sync-with-bend-calc, contour-flange-ex, contour-flange-sync, contour-flange-sync-with-bend, hem, jog, close-corner, multi-edge-flange, bend-with-calc, convert-part-to-sheet-metal, dimple-ex |
| **Body Operations**       | 11    | Add body, thicken, mesh, tag, construction, delete faces (2), delete holes (2), delete blends |
| **Simplification**        | 4     | Auto-simplify, enclosure, duplicate |
| **Mirror**                | 2     | Mirror copy, mirror sync ex |
| **Patterns**              | 5     | Rectangular ex, circular ex, duplicate, by fill, by table |
| **View/Display**          | 15    | Orientation, zoom, display mode, background color, get/set camera, rotate, pan, zoom factor, refresh, transform model-to-screen, transform screen-to-model, begin/end camera dynamics |
| **Variables**             | 12    | Get all, get by name, set value, set formula, add variable, query/search, get formula, rename, get names (display+system), translate, copy to clipboard, add from clipboard |
| **Custom Properties**     | 3     | Get all, set/create, delete |
| **Body Topology**         | 3     | Body faces, body edges, face info |
| **Performance**           | 1     | Recompute (set_performance_mode moved to Connection/Application) |
| **Query/Analysis**        | 67    | Mass properties, bounding box, features, measurements, facet data, solid bodies, modeling mode, face/edge info, colors, angles, volume, delete feature, material table, feature dimensions, material list/set/property, feature status, feature profiles, vertex count, layers (get/add/activate/set properties/delete), body opacity, body reflectivity, variable formula, feature parents, direction1/2 extent (get/set), thin wall options (get/set), from-face offset (get/set), body array (get/set), material library, set material by name, **edge endpoints/length/tangent/geometry/curvature, face normal/geometry/loops/curvature, vertex point, body extreme point/shells/vertices, faces by ray, shell info, point inside body, bspline curve/surface info** |
| **Feature Management**    | 6     | Suppress, unsuppress, face rotate (2), draft angle, convert feature type |
| **Part Features**         | 16    | Dimple, dimple-ex, etch, rib, lip, drawn cutout, drawn-cutout-ex, bead, louver, louver-sync, gusset, thread, thread-ex, slot, slot-ex, slot-sync, split, thicken-sync |
| **Export**                | 10    | STEP, STL, IGES, PDF, DXF, flat DXF, Parasolid, JT, drawing, screenshot |
| **Assembly**              | 61    | Place, place with transform, list, constraints, patterns, suppress, BOM, structured BOM, interference, bbox, relations (get/add planar/axial/angular/point/tangent/gear), relation editing (offset get/set, angle get/set, normals get/set, suppress/unsuppress, geometry, gear ratio), delete relation, doc tree, replace, delete, visibility, color, transform, count, move, rotate, is_subassembly, display_name, occurrence_document, sub_occurrences, add family member, add family with transform, add by template, add adjustable part, reorder occurrence, put transform euler, put origin, make writable, swap family member, occurrence bodies, occurrence style, is tube, get adjustable part, get face style |
| **Draft/Drawing**         | 67    | Sheets (add, activate, delete, rename, dimensions, balloons, text-boxes, drawing-objects, sections), views (add, projected, count, get/set scale, delete, update, move, info, orientation, hidden edges, display mode, model-link, tangent-edges, detail, auxiliary, draft, align, assembly-view-ex, with-config, activate, deactivate), annotations (dimension, angular, radial, diameter, ordinate, balloon, note, leader, text box, center mark, centerline, surface finish, weld symbol, geometric tolerance), parts list, printing (print, set printer, get printer, paper size), face texture, **2D geometry (lines2d, circles2d, arcs2d), dimensions (distance, length, radius, angle), smart frames (2-point, by-origin), symbols (add, list), PMI (info, visibility)** |
| **Diagnostics**           | 2     | API and feature inspection |
| **Select Set**            | 11    | Get selection, clear selection, add, remove, select all, copy, cut, delete, suspend/resume/refresh display |

## Part 1: Part Feature Collections (Part.tlb)

### Legend
- [x] = Implemented as MCP tool
- [ ] = NOT implemented - available in COM API
- ~~ = Known broken / infeasible via COM

### 1.1 Protrusion Features

#### ExtrudedProtrusions Collection (13 Add methods)
- [x] `AddFinite` / `AddFiniteMulti` - via `create_extrude`
- [x] `AddThroughAll` - via `create_extrude_infinite`
- [x] `AddWithThinWall` - via `create_extrude_thin_wall`
- [x] `AddFromTo` - via `create_extrude_from_to_single`
- [x] `AddThroughNext` - via `create_extrude_through_next_single`
- [x] `AddFiniteByKeyPoint` - via `create_extrude_by_keypoint`
- [x] `AddFromToMulti` - via `create_extrude_from_to_v2`
- [x] `AddThroughNextMulti` - via `create_extrude_through_next_v2`

#### RevolvedProtrusions Collection (12 Add methods)
- [x] `AddFinite` / `AddFiniteMulti` - via `create_revolve`, `create_revolve_finite`
- [x] `AddFiniteSync` - via `create_revolve_sync`, `create_revolve_finite_sync`
- [x] `AddWithThinWall` - via `create_revolve_thin_wall`
- [x] `AddFiniteByKeyPoint` - via `create_revolve_by_keypoint`
- [x] `AddFiniteByKeyPointSync` - via `create_revolve_by_keypoint_sync`
- [x] `Add` (full params with treatment) - via `create_revolve_full`

#### LoftedProtrusions Collection (4 Add methods)
- [x] `AddSimple` / Models.`AddLoftedProtrusion` - via `create_loft`
- [x] `AddWithThinWall` - via `create_loft_thin_wall`
- [ ] `Add` (full params) - Loft with guide curves

#### SweptProtrusions Collection (3 Add methods)
- [x] `Add` / Models.`AddSweptProtrusion` - via `create_sweep`
- [x] `AddWithThinWall` - via `create_sweep_thin_wall`

#### HelixProtrusions Collection (9 Add methods)
- [x] `AddFinite` - via `create_helix`
- [x] `AddFiniteSync` - via `create_helix_sync`
- [x] `AddFiniteWithThinWall` - via `create_helix_thin_wall`
- [x] `AddFiniteSyncWithThinWall` - via `create_helix_sync_thin_wall`
- [x] `AddFromTo` - via `create_helix_from_to`
- [x] `AddFromToWithThinWall` - via `create_helix_from_to_thin_wall`
- [x] `AddFromToSync` - via `create_helix_from_to_sync`
- [x] `AddFromToSyncWithThinWall` - via `create_helix_from_to_sync_thin_wall`

### 1.2 Cutout Features

#### ExtrudedCutouts Collection (17 Add methods)
- [x] `AddFiniteMulti` - via `create_extruded_cutout`
- [x] `AddThroughAllMulti` - via `create_extruded_cutout_through_all`
- [x] `AddFromToMulti` - via `create_extruded_cutout_from_to_v2`
- [x] `AddThroughNextMulti` - via `create_extruded_cutout_through_next`
- [x] `AddThroughNext` - via `create_extruded_cutout_through_next_single`
- [x] `AddFiniteByKeyPointMulti` - via `create_extruded_cutout_by_keypoint`
- [x] `AddFiniteMultiBody` - via `create_extruded_cutout_multi_body`
- [x] `AddFromToMultiBody` - via `create_extruded_cutout_from_to_multi_body`
- [x] `AddThroughAllMultiBody` - via `create_extruded_cutout_through_all_multi_body`

#### RevolvedCutouts Collection (17 Add methods)
- [x] `AddFiniteMulti` - via `create_revolved_cutout`
- [x] `AddFiniteMultiSync` - via `create_revolved_cutout_sync`
- [x] `AddFiniteByKeyPointMulti` - via `create_revolved_cutout_by_keypoint`
- [x] `AddFiniteMultiBody` - via `create_revolved_cutout_multi_body`
- [x] `Add` / `AddSync` (full params) - via `create_revolved_cutout_full`, `create_revolved_cutout_full_sync`

#### NormalCutouts Collection (6 Add methods)
- [x] `AddFiniteMulti` - via `create_normal_cutout`
- [x] `AddFromToMulti` - via `create_normal_cutout_from_to`
- [x] `AddThroughNextMulti` - via `create_normal_cutout_through_next`
- [x] `AddThroughAllMulti` - via `create_normal_cutout_through_all`
- [x] `AddFiniteByKeyPointMulti` - via `create_normal_cutout_by_keypoint`

#### LoftedCutouts Collection (3 Add methods)
- [x] `AddSimple` - via `create_lofted_cutout`
- [x] `Add` (full params) - via `create_lofted_cutout_full`

#### SweptCutouts Collection (3 Add methods)
- [x] `Add` - via `create_swept_cutout`
- [x] `AddMultiBody` - via `create_swept_cutout_multi_body`

#### HelixCutouts Collection (5 Add methods)
- [x] `AddFinite` - via `create_helix_cutout`
- [x] `AddFiniteSync` - via `create_helix_cutout_sync`
- [x] `AddFromTo` - via `create_helix_cutout_from_to`
- [x] `AddFromToSync` - via `create_helix_cutout_from_to_sync`

### 1.3 Rounds, Chamfers, Holes

#### Rounds Collection (5 Add methods)
- [x] `Add` (constant radius) - via `create_round`, `create_round_on_face`
- [x] `AddVariable` - via `create_variable_round`
- [x] `AddBlend` - via `create_round_blend`
- [x] `AddSurfaceBlend` - via `create_round_surface_blend`

#### Chamfers Collection (5 Add methods)
- [x] `AddEqualSetback` - via `create_chamfer`, `create_chamfer_on_face`
- [x] `AddUnequalSetback` - via `create_chamfer_unequal`, `create_chamfer_unequal_on_face`
- [x] `AddSetbackAngle` - via `create_chamfer_angle`

#### Holes Collection (14 Add methods)
- [x] `AddFinite` (via circular cutout workaround) - via `create_hole`
- [x] `AddFromTo` - via `create_hole_from_to`
- [x] `AddThroughNext` - via `create_hole_through_next`
- [x] `AddThroughAll` - via `create_hole_through_all` (circular cutout through-all)
- [x] `AddSync` - via `create_hole_sync`
- [x] `AddMultiBody` - via `create_hole_multi_body`
- [x] `AddFiniteEx` - via `create_hole_finite_ex`
- [x] `AddFromToEx` - via `create_hole_from_to_ex`
- [x] `AddThroughNextEx` - via `create_hole_through_next_ex`
- [x] `AddThroughAllEx` - via `create_hole_through_all_ex`
- [x] `AddSyncEx` - via `create_hole_sync_ex`
- [x] `AddSyncMultiBody` - via `create_hole_sync_multi_body`

### 1.4 Patterns

#### Patterns Collection (18 Add methods)
- ~~`Add` / `AddSync` - Feature pattern (SAFEARRAY marshaling broken)~~
- ~~`AddByRectangular` / `AddByCircular` - Rectangular/circular pattern (broken)~~
- [ ] `AddByCurve` / `AddByCurveSync` - Pattern along curve
- [x] `AddByFill` - via `create_pattern_by_fill`
- [x] `AddPatternByTable` - via `create_pattern_by_table`
- [x] `AddPatternByTableSync` - via `create_pattern_by_table_sync`
- [x] `AddDuplicate` - via `create_pattern_duplicate`
- [ ] `RecognizeAndCreatePatterns` - Auto-recognize patterns
- [x] `AddByRectangularEx` - via `create_pattern_rectangular_ex`
- [x] `AddByCircularEx` - via `create_pattern_circular_ex`
- [x] `AddByFillEx` / `AddByCurveEx` - via `create_pattern_by_fill_ex`, `create_pattern_by_curve_ex`

**Note**: The non-Ex feature pattern methods (`AddByRectangular`, `AddByCircular`) are known broken
due to SAFEARRAY(VT_DISPATCH) marshaling issues in late binding. The `Ex` variants
(`AddByRectangularEx`, `AddByCircularEx`) are now implemented and work correctly.

### 1.5 Reference Planes

#### RefPlanes Collection (14 Add methods)
- [x] `AddParallelByDistance` - via `create_ref_plane_by_offset`
- [x] `AddNormalToCurve` - via `create_ref_plane_normal_to_curve`
- [x] `AddAngularByAngle` - via `create_ref_plane_by_angle`
- [x] `AddBy3Points` - via `create_ref_plane_by_3_points`
- [x] `AddNormalToCurveAtDistance` - via `create_ref_plane_normal_at_distance_v2`
- [x] `AddNormalToCurveAtArcLengthRatio` - via `create_ref_plane_normal_at_arc_ratio_v2`
- [x] `AddNormalToCurveAtDistanceAlongCurve` - via `create_ref_plane_normal_at_distance_along_v2`
- [x] `AddNormalToCurveAtKeyPoint` - via `create_ref_plane_normal_at_keypoint`
- [x] `AddParallelByTangent` - via `create_ref_plane_tangent_parallel`
- [x] `AddTangentToCylinderOrConeAtAngle` - via `create_ref_plane_tangent_cylinder_angle`
- [x] `AddTangentToCylinderOrConeAtKeyPoint` - via `create_ref_plane_tangent_cylinder_keypoint`
- [x] `AddTangentToCurvedSurfaceAtKeyPoint` - via `create_ref_plane_tangent_surface_keypoint`
- [x] `AddMidPlane` - via `create_ref_plane_midplane`

### 1.6 Surface Features

#### ExtrudedSurfaces Collection (6 Add methods)
- [x] `AddFinite` - via `create_extruded_surface`
- [x] `AddFromTo` - via `create_extruded_surface_from_to`
- [x] `AddFiniteByKeyPoint` - via `create_extruded_surface_by_keypoint`
- [x] `Add` (full params with treatment) - via `create_extruded_surface_full`
- [x] `AddByCurves` - via `create_extruded_surface_by_curves`

#### RevolvedSurfaces Collection (7 Add methods)
- [x] `AddFinite` - via `create_revolved_surface`
- [x] `AddFiniteSync` - via `create_revolved_surface_sync`
- [x] `AddFiniteByKeyPoint` - via `create_revolved_surface_by_keypoint`
- [x] `Add` / `AddSync` (full params) - via `create_revolved_surface_full`, `create_revolved_surface_full_sync`

#### LoftedSurfaces Collection (3 Add methods)
- [x] `Add` - via `create_lofted_surface`
- [x] `Add2` - via `create_lofted_surface_v2`

#### SweptSurfaces Collection (3 Add methods)
- [x] `Add` - via `create_swept_surface`
- [x] `AddEx` - via `create_swept_surface_ex`

#### BlueSurf Interface (21 methods)
- [ ] **BlueSurf (bounded surface)** - Advanced surface creation, not exposed

### 1.7 Other Part Features

#### Threads Collection (3 methods)
- [x] `Add` - via `create_thread`
- [x] `AddEx` - via `create_thread_ex`

#### Ribs Collection (2 methods)
- [x] `Add` - via `create_rib`

#### Slots Collection (6 methods)
- [x] `Add` - via `create_slot`
- [x] `AddEx` - via `create_slot_ex`
- [x] `AddSync` - via `create_slot_sync`
- [x] `AddMultiBody` / `AddSyncMultiBody` - via `create_slot_multi_body`, `create_slot_sync_multi_body`

#### Blends Collection (4 methods)
- [x] `Add` - via `create_blend`
- [x] `AddVariable` - via `create_blend_variable`
- [x] `AddSurfaceBlend` - via `create_blend_surface`

#### Beads Collection (2 methods)
- [x] `Add` - via `create_bead`

#### DeleteFaces Collection (4 methods)
- [x] `Add` - via `delete_faces`
- [x] `AddNoHeal` - via `delete_faces_no_heal`

#### DeleteHoles Collection (3 methods)
- [x] `Add` - via `create_delete_hole`
- [x] `AddByFace` - via `delete_hole_by_face`

#### DeleteBlends Collection (2 methods)
- [x] `Add` - via `create_delete_blend`

#### Dimples Collection (3 methods)
- [x] `Add` - via `create_dimple`
- [x] `AddEx` - via `create_dimple_ex`

#### DrawnCutouts Collection (3 methods)
- [x] `Add` - via `create_drawn_cutout`
- [x] `AddEx` - via `create_drawn_cutout_ex`

#### EmbossFeatures Collection (2 methods)
- [x] `Add` - via `create_emboss`

#### Lips Collection (2 methods)
- [x] `Add` - via `create_lip`

#### Louvers Collection (3 methods)
- [x] `Add` - via `create_louver`
- [x] `AddSync` - via `create_louver_sync`

#### Gussets Collection (3 methods)
- [x] `AddByProfile` / `AddByBend` - via `create_gusset`

#### Thickens Collection (3 methods)
- [x] `Add` - via `thicken_surface`
- [x] `AddSync` - via `create_thicken_sync`

#### MirrorCopies Collection (4 methods)
- ~~`Add` / `AddSync` - Mirror copy (partially broken - feature tree entry but no geometry)~~
- [x] `AddSyncEx` - via `create_mirror_sync_ex` (implemented, though may have same geometry limitation)

#### Etches Collection
- [x] `Add` - via `create_etch`

#### Splits Collection
- [x] `Add` - via `create_split`

#### BoxFeatures Collection (7 methods)
- [x] `AddByCenter` / `AddByTwoPoints` / `AddByThreePoints` - via box tools
- [x] `AddCutoutByCenter` - via `create_box_cutout_by_center`
- [x] `AddCutoutByTwoPoints` - via `create_box_cutout`
- [x] `AddCutoutByThreePoints` - via `create_box_cutout_by_three_points`

#### CylinderFeatures Collection (3 methods)
- [x] `AddByCenterAndRadius` - via `create_cylinder`
- [x] `AddCutoutByCenterAndRadius` - via `create_cylinder_cutout`

#### SphereFeatures Collection (3 methods)
- [x] `AddByCenterAndRadius` - via `create_sphere`
- [x] `AddCutoutByCenterAndRadius` - via `create_sphere_cutout`

### 1.8 Sheet Metal Features

#### Flanges Collection (9 Add methods)
- [x] `Add` - via `create_flange`
- [x] `AddByMatchFace` - via `create_flange_by_match_face`
- [x] `AddSync` - via `create_flange_sync`
- [x] `AddFlangeByFace` - via `create_flange_by_face`
- [x] `AddByBendDeductionOrBendAllowance` - via `create_flange_with_bend_calc`
- [x] `AddByMatchFaceAndBendDeductionOrBendAllowance` - via `create_flange_match_face_with_bend`
- [x] `AddFlangeByFaceAndBendDeductionOrBendAllowance` - via `create_flange_by_face_with_bend`
- [x] `AddSyncByBendDeductionOrBendAllowance` - via `create_flange_sync_with_bend_calc`

#### ContourFlanges Collection (8 Add methods)
- [x] `Add` (via Models.AddBaseContourFlange) - via `create_base_flange`
- [x] `AddByBendDeductionOrBendAllowance` - via `create_base_contour_flange_advanced`
- [x] `AddEx` - via `create_contour_flange_ex`
- [x] `AddSync` - via `create_contour_flange_sync`
- [x] `Add3` - via `create_contour_flange_v3`
- [x] `AddSyncEx` - via `create_contour_flange_sync_ex`
- [x] `AddSyncByBendDeductionOrBendAllowance` - via `create_contour_flange_sync_with_bend`

#### Bends Collection (3 methods)
- [x] `Add` - via `create_bend`
- [x] `AddByBendDeductionOrBendAllowance` - via `create_bend_with_calc`

#### MultiEdgeFlanges (not a collection, but a feature type)
- [x] **Multi-edge flange** - via `create_multi_edge_flange`

#### Jogs Collection
- [x] **Jog feature** - via `create_jog`

#### Hems Collection
- [x] **Hem feature** - via `create_hem`

#### CloseCorners Collection
- [x] **Close corner** - via `create_close_corner`

#### ConvertPartToSM
- [x] **Convert solid part to sheet metal** - via `convert_part_to_sheet_metal`

### 1.9 Profile/Sketch Interface (26 methods)

- [x] `End` (close sketch) - via `close_sketch`
- [x] `SetAxisOfRevolution` - via `set_axis_of_revolution`
- [x] `ToggleConstruction` - via `draw_construction_line` (internal)
- [x] `IsConstructionElement` - used internally
- [x] `ProjectEdge` - via `project_edge`
- [x] `IncludeEdge` - via `include_edge`
- [x] `ProjectSilhouetteEdges` - via `project_silhouette_edges`
- [x] `ProjectRefPlane` - via `project_ref_plane`
- [x] `IncludeRegionFaces` - via `include_region_faces`
- [x] `Offset2d` - via `offset_sketch_2d`
- [x] `ChainLocate` - via `chain_locate`
- [x] `ConvertToCurve` - via `convert_to_curve`
- [x] `CleanGeometry2d` - via `clean_sketch_geometry`
- [x] `Paste` - via `sketch_paste`
- [x] `OrderedGeometry` - via `get_ordered_geometry`
- [x] `GetMatrix` - via `get_sketch_matrix`

### 1.10 Feature Editing (Common Methods on Feature Objects)

These methods exist on nearly every individual feature object (ExtrudedProtrusion,
RevolvedCutout, Round, Chamfer, Hole, etc.) and allow editing after creation:

- [x] `GetDimensions` - via `get_feature_dimensions`
- [x] `GetProfiles` - via `get_feature_profiles`; `SetProfiles` not implemented
- [x] `GetDirection1Extent` / `ApplyDirection1Extent` - via `get_direction1_extent`, `set_direction1_extent`
- [x] `GetDirection2Extent` / `ApplyDirection2Extent` - via `get_direction2_extent`, `set_direction2_extent`
- [x] `GetFromFaceOffsetData` / `SetFromFaceOffsetData` - via `get_from_face_offset`, `set_from_face_offset`
- [x] `GetThinWallOptions` / `SetThinWallOptions` - via `get_thin_wall_options`, `set_thin_wall_options`
- [x] `GetBodyArray` / `SetBodyArray` - via `get_body_array`, `set_body_array`
- [x] `ConvertToCutout` / `ConvertToProtrusion` - via `convert_feature_type`
- [x] `GetStatusEx` - via `get_feature_status`
- [x] `GetTopologyParents` - via `get_feature_parents`

**Impact**: These enable parametric editing of existing features, not just creation.

---

## Part 2: Assembly (assembly.tlb)

### 2.1 Occurrences Collection (13 Add methods)

- [x] `AddByFilename` - via `place_component`
- [x] `AddWithMatrix` - via `place_component` (with position)
- [x] `AddWithTransform` - via `assembly_add_component_with_transform`
- [x] `AddFamilyByFilename` - via `assembly_add_family_member`
- [x] `AddFamilyWithTransform` - via `assembly_add_family_with_transform`
- [x] `AddFamilyWithMatrix` - via `assembly_add_family_with_matrix`
- [x] `AddByTemplate` - via `assembly_add_by_template`
- [x] `AddAsAdjustablePart` - via `assembly_add_adjustable_part`
- [ ] `AddTube` - **Create tube/pipe between segments**
- [x] `ReorderOccurrence` - via `assembly_reorder_occurrence`
- [x] `GetOccurrence` - via `assembly_get_occurrence`

### 2.2 Occurrence Interface (66 methods, 73 properties)

- [x] `Delete` - via `delete_component`
- [x] `GetTransform` / `GetMatrix` - via `get_component_transform`
- [x] `PutMatrix` - via `update_component_position`
- [x] `Replace` - via `replace_component`
- [x] `Select` (implicit) - via selection tools
- [x] `PutTransform` - via `assembly_put_transform_euler`
- [x] `PutOrigin` - via `assembly_put_origin`
- [x] `Move` - via `occurrence_move`
- [x] `Rotate` - via `occurrence_rotate`
- [ ] `Mirror` - **Mirror occurrence across plane**
- [x] `MakeWritable` - via `assembly_make_writable`
- [x] `IsTube` - via `assembly_is_tube`
- [ ] `GetTube` - Query tube info
- [x] `SwapFamilyMember` - via `assembly_swap_family_member`
- [x] `GetAdjustablePart` - via `assembly_get_adjustable_part`
- [x] `GetFaceStyle2` - via `assembly_get_face_style`

Key Properties NOT exposed:
- [x] `Visible` (get/put) - via `set_component_visibility`
- [x] `Subassembly` (get) - via `is_subassembly`
- [x] `SubOccurrences` (get) - via `get_sub_occurrences`
- [x] `Bodies` (get) - via `assembly_get_occurrence_bodies`
- [x] `OccurrenceDocument` (get) - via `get_occurrence_document`
- [x] `DisplayName` (get) - via `get_component_display_name`
- [x] `Style` (get) - via `assembly_get_occurrence_style`

### 2.3 Assembly Relations (6 relation types)

**Relations3d Collection**:
- [x] Iterate and read relations - via `get_assembly_relations`
- [x] `AddPlanar` - via `assembly_add_planar_relation`
- [x] `AddAxial` - via `assembly_add_axial_relation`
- [x] `AddAngular` - via `assembly_add_angular_relation`
- [x] `AddPoint` - via `assembly_add_point_relation`
- [x] `AddTangent` - via `assembly_add_tangent_relation`
- [x] `AddGear` - via `assembly_add_gear_relation`

**Individual Relation Editing** (PlanarRelation3d, AxialRelation3d, etc.):
- [x] `Offset` (get/put) - via `assembly_get_relation_offset`, `assembly_set_relation_offset`
- [x] `Angle` (get/put) - via `assembly_get_relation_angle`, `assembly_set_relation_angle`
- [x] `NormalsAligned` (get/put) - via `assembly_get_normals_aligned`, `assembly_set_normals_aligned`
- [x] `Suppress` (put True/False) - via `assembly_suppress_relation`, `assembly_unsuppress_relation`
- [x] `Delete` - via `assembly_delete_relation`
- [x] `GetGeometry1` / `GetGeometry2` - via `assembly_get_relation_geometry`
- [x] `RatioValue1` / `RatioValue2` - via `assembly_get_gear_ratio`

### 2.4 Assembly Features

#### AssemblyFeaturesExtrudedCutout (23 methods)
- [ ] **Assembly-level extruded cutout** - Cut across multiple components

#### AssemblyFeaturesHole (23 methods)
- [ ] **Assembly-level hole** - Drill through multiple components

#### AssemblyFeaturesRevolvedCutout (21 methods)
- [ ] **Assembly-level revolved cutout**

#### AssemblyFeaturesExtrudedProtrusion (21 methods)
- [ ] **Assembly-level extruded protrusion**

#### AssemblyFeaturesRevolvedProtrusion (19 methods)
- [ ] **Assembly-level revolved protrusion**

#### AssemblyFeaturesMirror (13 methods)
- [ ] **Assembly-level mirror** - Mirror components across plane

#### AssemblyFeaturesPattern (13 methods)
- [ ] **Assembly-level pattern** - Pattern components

### 2.5 Specialized Assembly

- [ ] `StructuralFrame` (35 methods) - Structural frame/weldment
- [ ] `Cable` / `Wire` / `Bundle` - Wiring/harness features
- [ ] `Tube` (17 methods) - Tube/pipe routing
- [ ] `VirtualComponentOccurrence` - Virtual/in-place components
- [ ] `Splice` - Splice features
- [ ] `InternalComponent` - Internal component management

---

## Part 3: Draft/Drawing (draft.tlb)

### 3.1 DrawingViews Collection (11 Add methods)

- [x] `Add` (via AddPartView workaround) - via `create_drawing`, `add_assembly_drawing_view`
- [x] `AddWithConfiguration` - via `add_drawing_view_with_config`
- [x] `AddByFold` - via `add_projected_view`
- [x] `AddByAuxiliaryFold` - via `add_auxiliary_view`
- [x] `AddByDetailEnvelope` - via `add_detail_view`
- [x] `AddDraftView` - via `add_draft_view`
- [x] `AddByDraftView` - via `add_by_draft_view`
- [x] `AddAssemblyView` - via `add_assembly_drawing_view_ex`
- [x] `Align` / `Unalign` - via `align_drawing_views`

### 3.2 DrawingView Interface (59 methods, 155 properties)

Key methods:
- [x] `GetSectionCuts` / `AddSectionCut` - via `get_section_cuts`, `add_section_cut`
- [x] `Update` - via `update_drawing_view`
- [x] `Activate` - via `activate_drawing_view`
- [x] `Deactivate` - via `deactivate_drawing_view`
- [x] `Delete` - via `delete_drawing_view`
- [x] `Move` - via `move_drawing_view`
- [x] `SetOrientation` - via `set_drawing_view_orientation`
- [x] `Count` - via `get_drawing_view_count` (on collection)

Key properties:
- [x] `Scale` (get/put) - via `get_drawing_view_scale` / `set_drawing_view_scale`
- [x] `ShowHiddenEdges` (get/put) - via `show_hidden_edges`, `get_drawing_view_info`
- [x] `ShowTangentEdges` (get/put) - via `show_tangent_edges`
- [x] `DisplayMode` (get/put) - via `set_drawing_view_display_mode`, `get_drawing_view_info`
- [x] `ModelLink` (get) - via `get_drawing_view_model_link`
- [x] `Dimensions` (get) - via `get_drawing_view_dimensions`

### 3.3 Sheet Interface (23 methods, 63 properties)

- [x] `AddDrawingView` (via DrawingViews) - via drawing tools
- [x] Activate - via `activate_sheet`
- [x] Delete - via `delete_sheet`
- [x] Rename - via `rename_sheet`
- [x] `Sections` (get) - via `get_sheet_sections`
- [x] `Dimensions` (get) - via `get_sheet_dimensions`
- [x] `Balloons` (get) - via `get_sheet_balloons`
- [x] `TextBoxes` (get) - via `get_sheet_text_boxes`
- [x] `DrawingObjects` (get) - via `get_sheet_drawing_objects`

### 3.4 Annotations & Dimensions

- [x] `add_dimension` - Basic linear dimension
- [x] `add_angular_dimension` - Angular dimension
- [x] `add_radial_dimension` - Radial dimension
- [x] `add_diameter_dimension` - Diameter dimension
- [x] `add_ordinate_dimension` - Ordinate dimensioning
- [x] `add_balloon` - Balloon annotation
- [x] `add_note` - Free-standing text
- [x] `add_leader` - Leader annotation
- [x] `add_text_box` - Text box
- [x] `add_center_mark` - Center mark on circular features
- [x] `add_centerline` - Centerline annotation
- [x] `add_surface_finish_symbol` - Surface finish (GD&T)
- [x] `add_weld_symbol` - Welding annotations
- [x] `add_geometric_tolerance` - Geometric tolerance (FCF)
- [x] **Parts list** - via `create_parts_list` (BOM table on drawing)
- [ ] **Hole table** (HoleTable2 interface, 86 members) - Hole call-outs
- [ ] **Bend table** (DraftBendTable, 43 members) - Sheet metal bend table

### 3.5 DraftPrintUtility (10 methods)
- [x] `PrintOut` - via `print_drawing`
- [x] `SetPrinter` - via `set_printer`
- [x] `GetPrinter` - via `get_printer`
- [x] Paper size controls - via `set_paper_size`

---

## Part 4: Framework (framewrk.tlb)

### 4.1 Application Interface (81 methods)

- [x] Connect / GetActiveObject - via `connect_to_solidedge`
- [x] Quit - via `quit_application`
- [x] `DoIdle` - via `do_idle`
- [x] `StartCommand` - via `start_command`
- [x] `AbortCommand` - via `abort_command`
- [x] `GetGlobalParameter` / `SetGlobalParameter` - via `get_global_parameter`, `set_global_parameter`
- [ ] `GetModelessTaskEventSource` - Event handling (infeasible: requires event sink)
- [x] `Activate` - via `activate_application`

Key Properties:
- [x] `ActiveDocument` - via document tools
- [x] `ActiveEnvironment` - via `get_active_environment`
- [x] `StatusBar` (get/put) - via `get_status_bar`, `set_status_bar`
- [x] `DisplayAlerts` (get/put) - covered via `set_performance_mode(display_alerts=...)`
- [x] `Visible` (get/put) - via `get_visible`, `set_visible`
- [ ] `SensorEvents` - Sensor monitoring (infeasible: requires event sink)

### 4.2 View Interface (60 methods)

- [x] `Fit` - via `zoom_fit`
- [x] `SetRenderMode` - via `set_display_mode`
- [x] `ApplyNamedView` - via `set_view`
- [x] `SaveAsImage` - via `capture_screenshot`
- [x] `GetCamera` / `SetCamera` - via `get_camera`, `set_camera`
- [x] `RotateCamera` - via `rotate_view`
- [x] `PanCamera` - via `pan_view`
- [x] `ZoomCamera` - via `zoom_view`
- [x] `TransformModelToDC` / `TransformDCToModel` - via `transform_model_to_screen`, `transform_screen_to_model`
- [x] `Update` - via `refresh_view`
- [x] `BeginCameraDynamics` / `EndCameraDynamics` - via `begin_camera_dynamics`, `end_camera_dynamics`

### 4.3 Variables Interface (13 methods)

- [x] `Item` / iteration - via `get_variables`, `get_variable`
- [x] `Edit` / value setting - via `set_variable`
- [x] `Add` - via `add_variable`
- [x] `AddFromClipboard` - via `add_variable_from_clipboard`
- [x] `PutName` / `GetName` - via `rename_variable`
- [x] `Translate` - via `translate_variable`
- [x] `Query` - via `query_variables`
- [x] `GetFormula` / `Formula` (put) - via `get_variable_formula`, `set_variable_formula`
- [x] `GetDisplayName` / `GetSystemName` - via `get_variable_names`
- [x] `CopyToClipboard` - via `copy_variable_to_clipboard`

### 4.4 SelectSet Interface (13 methods)

- [x] `RemoveAll` - via `clear_select_set`
- [x] `Item` / `Count` - via `get_select_set`
- [x] `Add` - via `select_add`
- [x] `Remove` - via `select_remove`
- [x] `AddAll` - via `select_all`
- [x] `Copy` / `Cut` / `Delete` - via `select_copy`, `select_cut`, `select_delete`
- [x] `SuspendDisplay` / `ResumeDisplay` / `RefreshDisplay` - via `select_suspend_display`, `select_resume_display`, `select_refresh_display`

### 4.5 Other Framework Interfaces

#### MatTable (59 methods)
- [x] Density setting - via `set_material_density`
- [x] `ApplyMaterial` - via `set_material`
- [x] `GetMaterialList` - via `get_material_list`
- [x] `GetMatPropValue` - via `get_material_property`
- [x] Material library access - via `get_material_library`
- [x] Material set by name - via `set_material_by_name`

#### Layer Interface (15 methods)
- [x] `Add` / `SetActive` - via `add_layer`, `activate_layer`
- [x] `Visible` / `Selectable` - via `set_layer_properties`, `get_layers`
- [x] `Delete` - via `delete_layer`

#### FaceStyle Interface (49 methods, 69 properties)
- [x] Basic body color - via `set_body_color`, `set_face_color`
- [x] `Texture` - via `set_face_texture`
- [x] `Transparency` - via `set_body_opacity`
- [x] `Reflectivity` - via `set_body_reflectivity`

---

## Part 5: Additional Type Libraries

### 5.1 Geometry Type Library (geometry.tlb) - 68 interfaces

Key interfaces for precise geometry queries:
- [x] `Body` - via `get_body_extreme_point`, `get_body_shells`, `get_body_vertices`, `get_faces_by_ray`, `is_point_inside_body`
- [x] `Face` - via `get_face_normal`, `get_face_geometry`, `get_face_loops`, `get_face_curvature`
- [x] `Edge` - via `get_edge_endpoints`, `get_edge_length`, `get_edge_tangent`, `get_edge_geometry`, `get_edge_curvature`
- [x] `Vertex` - via `get_vertex_point`
- [ ] `Curve2d` / `Curve3d` - Parametric curve access (low-level, covered by edge/face geometry tools)
- [x] `BSplineCurve` / `BSplineSurface` - via `get_bspline_curve_info`, `get_bspline_surface_info`
- [x] `Loop` - via `get_face_loops`
- [x] `Shell` - via `get_shell_info`, `get_body_shells`, `is_point_inside_body`

### 5.2 FrameworkSupport (fwksupp.tlb) - 228 interfaces

2D geometry and annotation primitives:
- [x] `Lines2d` / `Circles2d` / `Arcs2d` - via `get_lines2d`, `get_circles2d`, `get_arcs2d`
- [x] `Dimensions` - via `add_distance_dimension`, `add_length_dimension`, `add_radius_dimension_2d`, `add_angle_dimension_2d`
- [x] `SmartFrames2d` - via `add_smart_frame`, `add_smart_frame_by_origin`
- [x] `Symbols` - via `add_symbol`, `get_symbols`
- [ ] `BalloonTypes` / `BalloonNotes` - Balloon annotations (not standalone collections; balloons already covered via `add_balloon`)
- [x] `PMI` - via `get_pmi_info`, `set_pmi_visibility`

---

## Known Limitations

1. **Assembly constraints** require face/edge geometry selection which is complex to automate via COM
2. **Feature patterns** (non-Ex variants): `model.Patterns.AddByRectangular` / `AddByCircular` require SAFEARRAY(IDispatch) marshaling that fails in late binding. The `Ex` variants (`AddByRectangularEx`, `AddByCircularEx`, `AddDuplicate`, `AddByFill`, `AddPatternByTable`) are now implemented and work correctly.
3. **Shell/Thinwalls** requires face selection for open faces, not automatable via COM
4. **Cutout via models.Add*Cutout** does NOT work - must use collection-level methods (ExtrudedCutouts.AddFiniteMulti)
5. **Mirror copy**: `Add` / `AddSync` create feature tree entry but no geometry via COM (partially broken). `AddSyncEx` is now implemented via `create_mirror_sync_ex`, though it may have the same geometry computation limitation.
6. **Extrude thin wall/infinite** via models.AddExtrudedProtrusionWithThinWall has unknown extra params

---

## Part 6: Key Constants Reference

The `constant.tlb` has 745 enums. Most important for new implementations:

| Enum | Members | Used For |
|------|---------|----------|
| FeaturePropertyConstants | 258 | Every feature operation (directions, extents, types) |
| FeatureTypeConstants | 129 | Feature type identification |
| ObjectType | 130 | COM object type identification |
| ApplicationGlobalConstants | 497 | App settings and command IDs |
| RefPlaneCreationMethodConstants | 39 | Reference plane creation |
| DimConstraintTypeConstants | 71 | Dimension types |
| KeyPointType | 54 | Keypoint selection types |
| seLocateFilterConstants | 77 | Selection filters |
| DrawingViewTypeConstants | 44 | Drawing view types |
| ImportOptionConstants | 53 | File import options |

**All 14,575 enum values are available in `reference/typelib_dump.json`** for lookup.

---

### Structure
```
src/solidedge_mcp/
├── server.py              <-- Thin entry point (< 100 lines)
├── tools/                 <-- Modular tool definitions
│   ├── __init__.py        <-- register_tools(mcp) function
│   ├── connection.py
│   ├── documents.py
│   ├── sketching.py
│   ├── features.py
│   ├── assembly.py
│   ├── query.py
│   ├── export.py
│   └── diagnostics.py
└── backends/              <-- EXISTING: Core logic (unchanged)
```
