# Solid Edge Type Library Implementation Map

Generated: 2026-02-18 | Source: 40 type libraries, 2,240 interfaces, 21,237 methods
Current: 110 composite MCP tools + 52 MCP resources = 162 endpoints

This document maps every actionable COM API surface from the Solid Edge type libraries
against our current MCP tool coverage. It identifies gaps and prioritizes what to implement next.

## Coverage Summary

|     Category          | Sections | Complete | Partial | Not Started | Methods (impl/total) |
|-----------------------|----------|----------|---------|-------------|----------------------|
| **Part Features**     |    52    |    38    |    11   |      3      |      197 / 201       |
| **Assembly**          |    11    |     9    |     1   |      1      |       66 /  73       |
| **Draft/Drawing**     |     5    |     3    |     2   |      0      |       57 /  58       |
| **Framework/App**     |     7    |     7    |     0   |      0      |       62 /  62       |
| **Geometry/Topology** |     2    |     0    |     2   |      0      |       15 /  19       |
| **Total**             | **77**   |  **57**  |  **16** |    **4**    | **397 / 413 (96%)**  |

**110 composite MCP tools + 52 MCP resources** registered. Tools use match/case dispatch
to consolidate related methods into single composites (e.g. `create_extrude(method=...)` covers
11 extrusion methods). 52 read-only endpoints use MCP Resources with `solidedge://` URIs.

## Tool Count by Category (107 tools + 52 resources)

| Category                  | Tools | Description |
|:--------------------------|:-----:|:---|
| **Connection/App**        | 7     | manage_connection, app_command, app_config, convert_by_file_path, arrange_windows, get_active_command, run_macro |
| **Documents**             | 7     | create_document, open_document, close_document, save_document, undo_redo, activate_document, import_file |
| **Sketching**             | 5     | manage_sketch, draw, sketch_modify, sketch_constraint, sketch_project |
| **Features (Part)**       | 49    | 45 composites (extrude, revolve, cutouts, loft, sweep, helix, rounds, chamfers, holes, ref planes, patterns, mirror, surfaces, sheet metal, body ops, simplify, manage) + 4 standalone |
| **Query/Analysis**        | 14    | measure, manage_variable/property/material, set_appearance, manage_layer, select_set, edit_feature_extent, manage_feature_tree, query_edge/face/body/bspline, recompute |
| **Export/Drawing**        | 14    | export_file, add_drawing_view, manage_drawing_view, add_annotation, add_2d_dimension, camera_control, display_control, manage_sheet, print_control, query_sheet, manage_annotation_data, add_smart_frame, draft_config, create_table |
| **Assembly**              | 12    | add_assembly_component, manage_component, query_component, set_component_appearance, transform_component, add_assembly_constraint, add_assembly_relation, manage_relation, assembly_feature, virtual_component, structural_frame, wiring |
| **Diagnostics**           | 2     | diagnose_api, diagnose_feature |

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
- [x] `Add` (full params with guide curves) - via `create_loft_with_guides`

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
- ~~`AddByCurve` / `AddByCurveSync` - Pattern along curve (SAFEARRAY(VT_DISPATCH) — same marshaling issue as non-Ex variants)~~
- [x] `AddByFill` - via `create_pattern_by_fill`
- [x] `AddPatternByTable` - via `create_pattern_by_table`
- [x] `AddPatternByTableSync` - via `create_pattern_by_table_sync`
- [x] `AddDuplicate` - via `create_pattern_duplicate`
- ~~`RecognizeAndCreatePatterns` - Auto-recognize patterns (SAFEARRAY(Hole*) + in/out arrays — infeasible in late binding)~~
- [x] `AddByRectangularEx` - via `create_pattern_rectangular_ex`
- [x] `AddByCircularEx` - via `create_pattern_circular_ex`
- [x] `AddByFillEx` / `AddByCurveEx` - via `create_pattern_by_fill_ex`, `create_pattern_by_curve_ex`

#### UserDefinedPatterns Collection
- [x] `Add` - via `create_user_defined_pattern` (pattern from accumulated profiles)

**Note**: The non-Ex feature pattern methods (`AddByRectangular`, `AddByCircular`, `AddByCurve`)
are known broken due to SAFEARRAY(VT_DISPATCH) marshaling issues in late binding. The `Ex` variants
(`AddByRectangularEx`, `AddByCircularEx`, `AddByCurveEx`) work correctly.
`RecognizeAndCreatePatterns` is infeasible (SAFEARRAY(Hole*) + in/out params).

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

#### BlueSurfs Collection (21 methods)
- [x] `Add` - via `create_bounded_surface` (basic bounded surface with end caps/periodic)
- [ ] Full BlueSurf editing (21 methods) - SAFEARRAY(VT_DISPATCH) params for guide curves

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
- [x] `GetToFaceOffsetData` / `SetToFaceOffsetData` - via `get_to_face_offset`, `set_to_face_offset`
- [x] `GetDirection1Treatment` / `ApplyDirection1Treatment` - via `get_direction1_treatment`, `apply_direction1_treatment`
- [x] `ConvertToCutout` / `ConvertToProtrusion` - via `convert_feature_type`
- [x] `GetStatusEx` - via `get_feature_status`
- [x] `GetTopologyParents` - via `get_feature_parents`

### 1.11 Part Document Methods
- [x] `Recompute` (document-level) - via `recompute(scope="document")`
- [x] `UserPhysicalProperties` (get) - via `query_body(property="user_physical_properties")`

**Impact**: Feature editing methods enable parametric editing of existing features, not just creation.

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
- [x] `AddTube` - via `add_tube`
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
- [x] `Mirror` - via `mirror_component`
- [x] `MakeWritable` - via `assembly_make_writable`
- [x] `IsTube` - via `assembly_is_tube`
- [x] `GetTube` - via `get_tube`
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
- [x] `Add` - via `assembly_create_extruded_cutout`

#### AssemblyFeaturesHole (23 methods)
- [x] `Add` - via `assembly_create_hole`

#### AssemblyFeaturesRevolvedCutout (21 methods)
- [x] `Add` - via `assembly_create_revolved_cutout`

#### AssemblyFeaturesExtrudedProtrusion (21 methods)
- [x] `Add` - via `assembly_create_extruded_protrusion`

#### AssemblyFeaturesRevolvedProtrusion (19 methods)
- [x] `Add` - via `assembly_create_revolved_protrusion`

#### AssemblyFeaturesMirror (13 methods)
- [x] `Add` - via `assembly_create_mirror`

#### AssemblyFeaturesPattern (13 methods)
- [x] `Add` - via `assembly_create_pattern`

#### AssemblyFeaturesSweptProtrusion
- [x] `Add` - via `assembly_create_swept_protrusion`

#### AssemblyFeatures Recompute
- [x] `Recompute` - via `assembly_recompute_features`

### 2.5 Specialized Assembly

- [x] `StructuralFrames.Add` / `AddByOrientation` - via `structural_frame`
- [x] `Cables.Add` / `Wires.Add` / `Bundles.Add` - via `wiring` (SAFEARRAY - may need runtime verification)
- [x] `Occurrences.AddTube` / `Occurrence.GetTube` - via `add_assembly_component(method="tube")`, `query_component(property="tube")`
- [x] `VirtualComponentOccurrences.Add` / `AddAsPreDefined` / `AddBIDM` - via `virtual_component`
- [x] `Splices.Add` - via `wiring(type="splice")`
- [ ] `InternalComponent` - No Add method found in type library

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
- ~~**Hole table** (HoleTable2 interface, 86 members) - No Add method in collection — not programmable via COM~~
- [x] **Bend table** - via `create_bend_table` (DraftBendTables.Add)

### 3.5 DraftPrintUtility (10 methods)
- [x] `PrintOut` - via `print_drawing`
- [x] `PrintDocument` (full params) - via `print_control(action="print_full")`
- [x] `SetPrinter` - via `set_printer`
- [x] `GetPrinter` - via `get_printer`
- [x] Paper size controls - via `set_paper_size`

### 3.6 Draft Document Methods
- [x] `GetGlobalParameter` / `SetGlobalParameter` - via `draft_config(action="get_global"/"set_global")`
- [x] `UpdateAllViews` - via `manage_drawing_view(action="update_all")`
- [x] `GetSymbolFileOrigin` / `SetSymbolFileOrigin` - via `draft_config(action="get_origin"/"set_origin")`

---

## Part 4: Framework (framewrk.tlb)

### 4.1 Application Interface (81 methods)

- [x] Connect / GetActiveObject - via `connect_to_solidedge`
- [x] Quit - via `quit_application`
- [x] `DoIdle` - via `do_idle`
- [x] `StartCommand` - via `start_command`
- [x] `AbortCommand` - via `abort_command`
- [x] `GetGlobalParameter` / `SetGlobalParameter` - via `get_global_parameter`, `set_global_parameter`
- ~~`GetModelessTaskEventSource` - Event handling (infeasible: requires event sink)~~
- [x] `Activate` - via `activate_application`
- [x] `RunMacro` - via `run_macro`
- [x] `ConvertByFilePath` - via `convert_by_file_path`
- [x] `ArrangeWindows` - via `arrange_windows`
- [x] `ActiveCommand` (get) - via `get_active_command`
- [x] `GetDefaultTemplatePath` / `SetDefaultTemplatePath` - via `app_config(property="get_template"/"set_template")`

Key Properties:
- [x] `ActiveDocument` - via document tools
- [x] `ActiveEnvironment` - via `get_active_environment`
- [x] `StatusBar` (get/put) - via `get_status_bar`, `set_status_bar`
- [x] `DisplayAlerts` (get/put) - covered via `set_performance_mode(display_alerts=...)`
- [x] `Visible` (get/put) - via `get_visible`, `set_visible`
- ~~`SensorEvents` - Sensor monitoring (infeasible: requires event sink)~~

### 4.2 Documents Collection
- [x] `Add` - via `create_document`
- [x] `Open` - via `open_document(method="foreground")`
- [x] `OpenInBackground` - via `open_document(method="background")`
- [x] `OpenWithTemplate` - via `open_document(method="with_template")`
- [x] `OpenWithFileOpenDialog` - via `open_document(method="dialog")`

### 4.3 View Interface (60 methods)

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

### 4.4 Variables Interface (13 methods)

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

### 4.5 SelectSet Interface (13 methods)

- [x] `RemoveAll` - via `clear_select_set`
- [x] `Item` / `Count` - via `get_select_set`
- [x] `Add` - via `select_add`
- [x] `Remove` - via `select_remove`
- [x] `AddAll` - via `select_all`
- [x] `Copy` / `Cut` / `Delete` - via `select_copy`, `select_cut`, `select_delete`
- [x] `SuspendDisplay` / `ResumeDisplay` / `RefreshDisplay` - via `select_suspend_display`, `select_resume_display`, `select_refresh_display`

### 4.6 Other Framework Interfaces

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
- ~~`Curve2d` / `Curve3d` - Parametric curve access (low-level, already covered by edge/face geometry tools)~~
- [x] `BSplineCurve` / `BSplineSurface` - via `get_bspline_curve_info`, `get_bspline_surface_info`
- [x] `Loop` - via `get_face_loops`
- [x] `Shell` - via `get_shell_info`, `get_body_shells`, `is_point_inside_body`

### 5.2 FrameworkSupport (fwksupp.tlb) - 228 interfaces

2D geometry and annotation primitives:
- [x] `Lines2d` / `Circles2d` / `Arcs2d` - via `get_lines2d`, `get_circles2d`, `get_arcs2d`
- [x] `Dimensions` - via `add_distance_dimension`, `add_length_dimension`, `add_radius_dimension_2d`, `add_angle_dimension_2d`
- [x] `SmartFrames2d` - via `add_smart_frame`, `add_smart_frame_by_origin`
- [x] `Symbols` - via `add_symbol`, `get_symbols`
- ~~`BalloonTypes` / `BalloonNotes` - Balloon annotations (already covered via `add_balloon`)~~
- [x] `PMI` - via `get_pmi_info`, `set_pmi_visibility`

---

## Known Limitations

1. **Assembly constraints** require face/edge geometry selection which is complex to automate via COM
2. **Feature patterns** (non-Ex variants): `AddByRectangular`, `AddByCircular`, `AddByCurve` require SAFEARRAY(IDispatch) that fails in late binding. Use `Ex` variants instead.
3. **RecognizeAndCreatePatterns**: SAFEARRAY(Hole*) + in/out array params — infeasible in late binding
4. **Shell/Thinwalls** requires face selection for open faces, not automatable via COM
5. **Cutout via models.Add*Cutout** does NOT work - must use collection-level methods (ExtrudedCutouts.AddFiniteMulti)
6. **Mirror copy**: `Add` / `AddSync` create feature tree entry but no geometry via COM (partially broken). `AddSyncEx` is implemented but may have same limitation.
7. **Extrude thin wall/infinite** via models.AddExtrudedProtrusionWithThinWall has unknown extra params
8. **Event sinks** (GetModelessTaskEventSource, SensorEvents): require COM event sink setup, not feasible in MCP tool model
9. **BlueSurf full editing** (21 methods): Multiple SAFEARRAY(VT_DISPATCH) params for guide curves — basic creation works but full editing interface infeasible

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
