# Solid Edge Type Library Implementation Map

Generated: 2026-02-13 | Source: 40 type libraries, 2,240 interfaces, 21,237 methods
Current: 299 MCP tools implemented

This document maps every actionable COM API surface from the Solid Edge type libraries
against our current MCP tool coverage. It identifies gaps and prioritizes what to implement next.

## Coverage Summary

|     Category      | Sections | Complete | Partial | Not Started | Methods (impl/total) |
|-------------------|----------|----------|---------|-------------|----------------------|
| **Part Features** |    52    |    15    |    28   |      9      |       74 / 181       |
| **Assembly**      |    11    |     0    |     3   |      8      |       14 /  60       |
| **Draft/Drawing** |     5    |     0    |     4   |      1      |       16 /  49       |
| **Framework/App** |     7    |     0    |     6   |      1      |       44 /  53       |
| **Total**         | **75**   |  **15**  |  **41** |   **19**    | **148 / 343 (43%)**  |

**299 MCP tools** registered (many tools cover multiple methods or provide capabilities
beyond what the type library tracks, e.g. primitives, view controls, export formats).

## Tool Count by Category (299 total)

| Category                  | Count | Tools |
|:--------------------------|:-----:|:---|
| **Connection/Application**| 17    | Connect, disconnect, app info, quit, is_connected, process_info, install_info, start_command, set_performance_mode, do_idle, activate, abort_command, active_environment, status_bar (get/set), visible (get/set) |
| **Document Management**   | 13    | Create (part, assembly, sheet metal, draft), open, save, close, list, activate, undo, redo |
| **Sketching**             | 24    | Lines, circles, arcs (multiple), rects, polygons, ellipses, splines, points, constraints (9 types), fillet, chamfer, mirror, construction, hide profile, project_edge, include_edge |
| **Basic Primitives**      | 10    | Box (3 variants), cylinder, sphere, box cutout (3 variants), cylinder cutout, sphere cutout |
| **Extrusions**            | 4     | Finite, infinite, thin-wall, extruded surface |
| **Revolves**              | 5     | Basic, finite, sync, thin-wall |
| **Cutouts**               | 9     | Extruded finite/through-all/through-next, revolved, normal/normal-through-all, lofted, swept, helix |
| **Rounds/Chamfers/Holes** | 9     | Round (all/face), variable round, blend, chamfer (equal/unequal/angle/face), hole, hole through-all |
| **Reference Planes**      | 5     | Offset, normal-to-curve, angle, 3-points, mid-plane |
| **Loft**                  | 2     | Basic, thin-wall |
| **Sweep**                 | 2     | Basic, thin-wall |
| **Helix/Spiral**          | 4     | Basic, sync, thin-wall variants |
| **Construction Surfaces** | 3     | Revolved surface, lofted surface, swept surface |
| **Sheet Metal**           | 10    | Base flange/tab, lofted flange, web network, advanced variants, emboss, flange |
| **Body Operations**       | 11    | Add body, thicken, mesh, tag, construction, delete faces (2), delete holes (2), delete blends |
| **Simplification**        | 4     | Auto-simplify, enclosure, duplicate |
| **View/Display**          | 11    | Orientation, zoom, display mode, background color, get/set camera, rotate, pan, zoom factor, refresh |
| **Variables**             | 8     | Get all, get by name, set value, add variable, query/search, get formula, rename, get names (display+system) |
| **Custom Properties**     | 3     | Get all, set/create, delete |
| **Body Topology**         | 3     | Body faces, body edges, face info |
| **Performance**           | 1     | Recompute (set_performance_mode moved to Connection/Application) |
| **Query/Analysis**        | 28    | Mass properties, bounding box, features, measurements, facet data, solid bodies, modeling mode, face/edge info, colors, angles, volume, delete feature, material table, feature dimensions, material list/set/property, feature status, feature profiles, vertex count |
| **Feature Management**    | 6     | Suppress, unsuppress, face rotate (2), draft angle, convert feature type |
| **Export**                | 10    | STEP, STL, IGES, PDF, DXF, flat DXF, Parasolid, JT, drawing, screenshot |
| **Assembly**              | 28    | Place, list, constraints, patterns, suppress, BOM, structured BOM, interference, bbox, relations, doc tree, replace, delete, visibility, color, transform, count, move, rotate, is_subassembly, display_name, occurrence_document, sub_occurrences |
| **Draft/Drawing**         | 16    | Sheets (add, activate, delete, rename), views (add, count, get/set scale, delete, update), annotations (dimension, balloon, note, leader, text box), parts list |
| **Part Features**         | 10    | Dimple, etch, rib, lip, drawn cutout, bead, louver, gusset, thread, slot, split |
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
- [ ] `AddFromTo` - Extrude from face to face
- [ ] `AddThroughNext` - Extrude to next face
- [ ] `AddFiniteByKeyPoint` - Extrude to keypoint
- [ ] `AddFromToMulti` - From-to with multiple profiles
- [ ] `AddThroughNextMulti` - Through-next with multiple profiles

#### RevolvedProtrusions Collection (12 Add methods)
- [x] `AddFinite` / `AddFiniteMulti` - via `create_revolve`, `create_revolve_finite`
- [x] `AddFiniteSync` - via `create_revolve_sync`, `create_revolve_finite_sync`
- [x] `AddWithThinWall` - via `create_revolve_thin_wall`
- [ ] `AddFiniteByKeyPoint` - Revolve to keypoint
- [ ] `AddFiniteByKeyPointSync` - Sync variant
- [ ] `Add` (full params with treatment) - Full revolve with draft/crown treatment

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
- [ ] `AddFromTo` - Helix between faces
- [ ] `AddFromToWithThinWall` - Thin wall variant
- [ ] `AddFromToSync` - Sync from-to
- [ ] `AddFromToSyncWithThinWall` - Sync thin wall from-to

### 1.2 Cutout Features

#### ExtrudedCutouts Collection (17 Add methods)
- [x] `AddFiniteMulti` - via `create_extruded_cutout`
- [x] `AddThroughAllMulti` - via `create_extruded_cutout_through_all`
- [ ] `AddFromTo` / `AddFromToMulti` - Cut from face to face
- [x] `AddThroughNextMulti` - via `create_extruded_cutout_through_next`
- [ ] `AddThroughNext` - Single-profile variant
- [ ] `AddFiniteByKeyPoint` / `AddFiniteByKeyPointMulti` - Cut to keypoint
- [ ] `AddFiniteMultiBody` - Multi-body finite cutout
- [ ] `AddFromToMultiBody` - Multi-body from-to cutout
- [ ] `AddThroughAllMultiBody` - Multi-body through-all cutout

#### RevolvedCutouts Collection (17 Add methods)
- [x] `AddFiniteMulti` - via `create_revolved_cutout`
- [ ] `AddFiniteSync` / `AddFiniteMultiSync` - Sync revolve cut
- [ ] `AddFiniteByKeyPoint` / `AddFiniteByKeyPointMulti` - Cut to keypoint
- [ ] `AddFiniteMultiBody` - Multi-body revolve cut
- [ ] `Add` / `AddSync` (full params) - Full revolve cut with treatment

#### NormalCutouts Collection (6 Add methods)
- [x] `AddFiniteMulti` - via `create_normal_cutout`
- [ ] `AddFromToMulti` - Normal cut from-to
- [ ] `AddThroughNextMulti` - Normal cut through-next
- [x] `AddThroughAllMulti` - via `create_normal_cutout_through_all`
- [ ] `AddFiniteByKeyPointMulti` - Normal cut to keypoint

#### LoftedCutouts Collection (3 Add methods)
- [x] `AddSimple` - via `create_lofted_cutout`
- [ ] `Add` (full params) - Lofted cut with guide curves

#### SweptCutouts Collection (3 Add methods)
- [x] `Add` - via `create_swept_cutout`
- [ ] `AddMultiBody` - Multi-body swept cutout

#### HelixCutouts Collection (5 Add methods)
- [x] `AddFinite` - via `create_helix_cutout`
- [ ] `AddFiniteSync` - Sync helix cutout
- [ ] `AddFromTo` - Helix cutout between faces
- [ ] `AddFromToSync` - Sync variant

### 1.3 Rounds, Chamfers, Holes

#### Rounds Collection (5 Add methods)
- [x] `Add` (constant radius) - via `create_round`, `create_round_on_face`
- [x] `AddVariable` - via `create_variable_round`
- [ ] `AddBlend` - **Blend (face-to-face fillet)**
- [ ] `AddSurfaceBlend` - Surface blend

#### Chamfers Collection (5 Add methods)
- [x] `AddEqualSetback` - via `create_chamfer`, `create_chamfer_on_face`
- [x] `AddUnequalSetback` - via `create_chamfer_unequal`
- [x] `AddSetbackAngle` - via `create_chamfer_angle`

#### Holes Collection (14 Add methods)
- [x] `AddFinite` (via circular cutout workaround) - via `create_hole`
- [ ] `AddFromTo` - Hole from face to face
- [ ] `AddThroughNext` - Hole to next face
- [x] `AddThroughAll` - via `create_hole_through_all` (circular cutout through-all)
- [ ] `AddSync` - Synchronous hole
- [ ] `AddMultiBody` - Multi-body hole
- [ ] `AddFiniteEx` / `AddFromToEx` / `AddThroughNextEx` / `AddThroughAllEx` - Extended variants
- [ ] `AddSyncEx` / `AddSyncMultiBody` - Sync multi-body

### 1.4 Patterns

#### Patterns Collection (18 Add methods)
- ~~`Add` / `AddSync` - Feature pattern (SAFEARRAY marshaling broken)~~
- ~~`AddByRectangular` / `AddByCircular` - Rectangular/circular pattern (broken)~~
- [ ] `AddByCurve` / `AddByCurveSync` - Pattern along curve
- [ ] `AddByFill` - Fill pattern
- [ ] `AddPatternByTable` / `AddPatternByTableSync` - Table-driven pattern
- [ ] `AddDuplicate` - Duplicate pattern
- [ ] `RecognizeAndCreatePatterns` - Auto-recognize patterns
- [ ] `AddByRectangularEx` / `AddByCircularEx` / `AddByFillEx` / `AddByCurveEx` - Extended variants

**Note**: Feature patterns are known broken due to SAFEARRAY(VT_DISPATCH) marshaling issues
in late binding. The `Ex` variants may use different parameter types - worth investigating.

### 1.5 Reference Planes

#### RefPlanes Collection (14 Add methods)
- [x] `AddParallelByDistance` - via `create_ref_plane_by_offset`
- [x] `AddNormalToCurve` - via `create_ref_plane_normal_to_curve`
- [x] `AddAngularByAngle` - via `create_ref_plane_by_angle`
- [x] `AddBy3Points` - via `create_ref_plane_by_3_points`
- [ ] `AddNormalToCurveAtDistance` - Normal to curve at distance
- [ ] `AddNormalToCurveAtArcLengthRatio` - Normal at arc ratio
- [ ] `AddNormalToCurveAtDistanceAlongCurve` - Normal at distance along curve
- [ ] `AddNormalToCurveAtKeyPoint` - Normal at keypoint
- [ ] `AddParallelByTangent` - Tangent parallel plane
- [ ] `AddTangentToCylinderOrConeAtAngle` - Tangent to cylinder
- [ ] `AddTangentToCylinderOrConeAtKeyPoint` - Tangent to cylinder at keypoint
- [ ] `AddTangentToCurvedSurfaceAtKeyPoint` - Tangent to surface
- [x] `AddMidPlane` - via `create_ref_plane_midplane`

### 1.6 Surface Features

#### ExtrudedSurfaces Collection (6 Add methods)
- [x] `AddFinite` - via `create_extruded_surface`
- [ ] `AddFromTo` - Surface from-to
- [ ] `AddFiniteByKeyPoint` - Surface to keypoint
- [ ] `Add` (full params with treatment) - Surface with draft/crown
- [ ] `AddByCurves` - **Surface from curves**

#### RevolvedSurfaces Collection (7 Add methods)
- [x] `AddFinite` - via `create_revolved_surface`
- [ ] `AddFiniteSync` - Sync variant
- [ ] `AddFiniteByKeyPoint` - To keypoint
- [ ] `Add` / `AddSync` (full params) - Full revolve surface

#### LoftedSurfaces Collection (3 Add methods)
- [x] `Add` - via `create_lofted_surface`
- [ ] `Add2` - Extended lofted surface

#### SweptSurfaces Collection (3 Add methods)
- [x] `Add` - via `create_swept_surface`
- [ ] `AddEx` - Extended swept surface

#### BlueSurf Interface (21 methods)
- [ ] **BlueSurf (bounded surface)** - Advanced surface creation, not exposed

### 1.7 Other Part Features

#### Threads Collection (3 methods)
- [x] `Add` - via `create_thread`
- [ ] `AddEx` - Extended thread (more parameters)

#### Ribs Collection (2 methods)
- [x] `Add` - via `create_rib`

#### Slots Collection (6 methods)
- [x] `Add` - via `create_slot`
- [ ] `AddEx` - Extended slot
- [ ] `AddSync` - Sync slot
- [ ] `AddMultiBody` / `AddSyncMultiBody` - Multi-body

#### Blends Collection (4 methods)
- [x] `Add` - via `create_blend`
- [ ] `AddVariable` - Variable blend
- [ ] `AddSurfaceBlend` - Surface blend

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
- [ ] `AddEx` - Extended dimple

#### DrawnCutouts Collection (3 methods)
- [x] `Add` - via `create_drawn_cutout`
- [ ] `AddEx` - Extended drawn cutout

#### EmbossFeatures Collection (2 methods)
- [x] `Add` - via `create_emboss`

#### Lips Collection (2 methods)
- [x] `Add` - via `create_lip`

#### Louvers Collection (3 methods)
- [x] `Add` - via `create_louver`
- [ ] `AddSync` - Sync louver

#### Gussets Collection (3 methods)
- [x] `AddByProfile` / `AddByBend` - via `create_gusset`

#### Thickens Collection (3 methods)
- [x] `Add` - via `thicken_surface`
- [ ] `AddSync` - Sync thicken

#### MirrorCopies Collection (4 methods)
- ~~`Add` / `AddSync` - Mirror copy (partially broken - feature tree entry but no geometry)~~
- [ ] `AddSyncEx` - Extended sync mirror (worth trying)

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
- [ ] `AddByMatchFace` - Flange matching adjacent face
- [ ] `AddSync` - Sync flange
- [ ] `AddFlangeByFace` - Flange from face
- [ ] `AddByBendDeductionOrBendAllowance` - With bend calc
- [ ] `AddByMatchFaceAndBendDeductionOrBendAllowance`
- [ ] `AddFlangeByFaceAndBendDeductionOrBendAllowance`
- [ ] `AddSyncByBendDeductionOrBendAllowance`

#### ContourFlanges Collection (8 Add methods)
- [x] `Add` (via Models.AddBaseContourFlange) - via `create_base_flange`
- [x] `AddByBendDeductionOrBendAllowance` - via `create_base_contour_flange_advanced`
- [ ] `AddEx` / `Add3` - Extended contour flange
- [ ] `AddSync` / `AddSyncEx` - Sync variants
- [ ] `AddSyncByBendDeductionOrBendAllowance`

#### Bends Collection (3 methods)
- [ ] `Add` - **Add bend to sheet metal (NOT IMPLEMENTED)**
- [ ] `AddByBendDeductionOrBendAllowance`

#### MultiEdgeFlanges (not a collection, but a feature type)
- [ ] **Multi-edge flange** - Flange on multiple edges simultaneously

#### Jogs Collection
- [ ] **Jog feature** - Offset step in sheet metal

#### Hems Collection
- [ ] **Hem feature** - Folded edge on sheet metal

#### CloseCorners Collection
- [ ] **Close corner** - Close gaps at sheet metal corners

#### ConvertPartToSM
- [ ] **Convert solid part to sheet metal** - Major capability

### 1.9 Profile/Sketch Interface (26 methods)

- [x] `End` (close sketch) - via `close_sketch`
- [x] `SetAxisOfRevolution` - via `set_axis_of_revolution`
- [x] `ToggleConstruction` - via `draw_construction_line` (internal)
- [x] `IsConstructionElement` - used internally
- [x] `ProjectEdge` - via `project_edge`
- [x] `IncludeEdge` - via `include_edge`
- [ ] `ProjectSilhouetteEdges` - Project silhouette onto sketch
- [ ] `ProjectRefPlane` - Project ref plane intersection
- [ ] `IncludeRegionFaces` - Include face region
- [ ] `Offset2d` - Offset profile curves
- [ ] `ChainLocate` - Locate connected chain of edges
- [ ] `ConvertToCurve` - Convert element to curve
- [ ] `CleanGeometry2d` - Clean up sketch geometry
- [ ] `Paste` - Paste clipboard into sketch
- [ ] `OrderedGeometry` - Get ordered geometry
- [ ] `GetMatrix` - Get sketch coordinate system

### 1.10 Feature Editing (Common Methods on Feature Objects)

These methods exist on nearly every individual feature object (ExtrudedProtrusion,
RevolvedCutout, Round, Chamfer, Hole, etc.) and allow editing after creation:

- [x] `GetDimensions` - via `get_feature_dimensions`
- [x] `GetProfiles` - via `get_feature_profiles`; `SetProfiles` not implemented
- [ ] `GetDirection1Extent` / `ApplyDirection1Extent` - Edit extent
- [ ] `GetDirection2Extent` / `ApplyDirection2Extent` - Edit 2nd direction
- [ ] `GetFromFaceOffsetData` / `SetFromFaceOffsetData` - Edit offsets
- [ ] `GetThinWallOptions` / `SetThinWallOptions` - Edit thin wall params
- [ ] `GetBodyArray` / `SetBodyArray` - Multi-body targeting
- [x] `ConvertToCutout` / `ConvertToProtrusion` - via `convert_feature_type`
- [x] `GetStatusEx` - via `get_feature_status`
- [ ] `GetTopologyParents` - Query parent geometry

**Impact**: These would enable parametric editing of existing features, not just creation.

---

## Part 2: Assembly (assembly.tlb)

### 2.1 Occurrences Collection (13 Add methods)

- [x] `AddByFilename` - via `place_component`
- [x] `AddWithMatrix` - via `place_component` (with position)
- [ ] `AddWithTransform` - Place with Euler angles
- [ ] `AddFamilyByFilename` - **Place family-of-parts member**
- [ ] `AddFamilyWithTransform` - Family with transform
- [ ] `AddFamilyWithMatrix` - Family with matrix
- [ ] `AddByTemplate` - Place with template
- [ ] `AddAsAdjustablePart` - Place as adjustable part
- [ ] `AddTube` - **Create tube/pipe between segments**
- [ ] `ReorderOccurrence` - Reorder in tree
- [ ] `GetOccurrence` - Get by internal ID

### 2.2 Occurrence Interface (66 methods, 73 properties)

- [x] `Delete` - via `delete_component`
- [x] `GetTransform` / `GetMatrix` - via `get_component_transform`
- [x] `PutMatrix` - via `update_component_position`
- [x] `Replace` - via `replace_component`
- [x] `Select` (implicit) - via selection tools
- [ ] `PutTransform` - **Set position by Euler angles**
- [ ] `PutOrigin` - **Set origin position only**
- [x] `Move` - via `occurrence_move`
- [x] `Rotate` - via `occurrence_rotate`
- [ ] `Mirror` - **Mirror occurrence across plane**
- [ ] `MakeWritable` - Make in-context editable
- [ ] `IsTube` / `GetTube` - Query tube info
- [ ] `SwapFamilyMember` - **Swap family member**
- [ ] `MakeAdjustablePart` / `GetAdjustablePart` - Adjustable parts
- [ ] `GetFaceStyle2` - Get appearance

Key Properties NOT exposed:
- [x] `Visible` (get/put) - via `set_component_visibility`
- [x] `Subassembly` (get) - via `is_subassembly`
- [x] `SubOccurrences` (get) - via `get_sub_occurrences`
- [ ] `Bodies` (get) - Access body geometry
- [x] `OccurrenceDocument` (get) - via `get_occurrence_document`
- [x] `DisplayName` (get) - via `get_component_display_name`
- [ ] `Style` (get/put) - Face style override

### 2.3 Assembly Relations (6 relation types)

**Relations3d Collection**:
- [x] Iterate and read relations - via `get_assembly_relations`
- [ ] `AddPlanar` - **Create planar mate programmatically**
- [ ] `AddAxial` - **Create axial mate**
- [ ] `AddAngular` - **Create angular constraint**
- [ ] `AddPoint` - **Create point constraint**
- [ ] `AddTangent` - **Create tangent constraint**
- [ ] `AddGear` - **Create gear relationship**

**Individual Relation Editing** (PlanarRelation3d, AxialRelation3d, etc.):
- [ ] `Offset` (get/put) - **Edit relation offset distance**
- [ ] `Angle` (get/put) - **Edit angular constraint angle**
- [ ] `NormalsAligned` (get/put) - Flip alignment
- [ ] `Suppress` (get/put) - **Suppress/unsuppress constraint**
- [ ] `Delete` - Delete constraint
- [ ] `GetGeometry1` / `GetGeometry2` - Query constraint geometry
- [ ] `RatioValue1` / `RatioValue2` - Gear ratio

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
- [ ] `AddWithConfiguration` - **View with named configuration**
- [ ] `AddByFold` - **Folded (projected) view from parent**
- [ ] `AddByAuxiliaryFold` - Auxiliary view
- [ ] `AddByDetailEnvelope` - **Detail view with circle**
- [ ] `AddDraftView` - Empty 2D draft view
- [ ] `AddByDraftView` - Copy draft view
- [ ] `AddAssemblyView` - Assembly view with explode/snapshot options
- [ ] `Align` / `Unalign` - Align drawing views

### 3.2 DrawingView Interface (59 methods, 155 properties)

Key methods:
- [ ] `GetSectionCuts` / `AddSectionCut` - **Section/cross-section views**
- [x] `Update` - via `update_drawing_view`
- [ ] `Activate` / `Deactivate` - Activate for editing
- [x] `Delete` - via `delete_drawing_view`
- [ ] `Move` - Reposition view on sheet
- [ ] `SetOrientation` - Change orientation
- [x] `Count` - via `get_drawing_view_count` (on collection)

Key properties:
- [x] `Scale` (get/put) - via `get_drawing_view_scale` / `set_drawing_view_scale`
- [ ] `ShowHiddenEdges` (get/put) - Hidden lines
- [ ] `ShowTangentEdges` (get/put) - Tangent edges
- [ ] `DisplayMode` (get/put) - Shaded/wireframe
- [ ] `ModelLink` (get) - Model reference
- [ ] `Dimensions` (get) - Access dimensions in view

### 3.3 Sheet Interface (23 methods, 63 properties)

- [x] `AddDrawingView` (via DrawingViews) - via drawing tools
- [x] Activate - via `activate_sheet`
- [x] Delete - via `delete_sheet`
- [x] Rename - via `rename_sheet`
- [ ] `Sections` (get) - Section view collection
- [ ] `Dimensions` (get) - **Dimension collection on sheet**
- [ ] `Balloons` (get) - **Balloon collection**
- [ ] `TextBoxes` (get) - Text boxes on sheet
- [ ] `DrawingObjects` (get) - All objects on sheet

### 3.4 Annotations & Dimensions

- [x] `add_dimension` - Basic linear dimension
- [x] `add_balloon` - Balloon annotation
- [x] `add_note` - Free-standing text
- [x] `add_leader` - Leader annotation
- [x] `add_text_box` - Text box
- [ ] **Angular dimension** - Between lines/edges
- [ ] **Radial/diameter dimension** - On arcs/circles
- [ ] **Ordinate dimension** - Ordinate dimensioning
- [ ] **Surface finish symbol** - GD&T
- [ ] **Weld symbol** - Welding annotations
- [ ] **Geometric tolerance** (FCF) - GD&T frames
- [ ] **Center mark** / **Centerline** - On circular features
- [x] **Parts list** - via `create_parts_list` (BOM table on drawing)
- [ ] **Hole table** (HoleTable2 interface, 86 members) - Hole call-outs
- [ ] **Bend table** (DraftBendTable, 43 members) - Sheet metal bend table

### 3.5 DraftPrintUtility (10 methods)
- [ ] `PrintOut` - Print drawing
- [ ] `SetPrinter` - Set printer
- [ ] `GetPrinter` - Get printer
- [ ] Paper size / orientation controls

---

## Part 4: Framework (framewrk.tlb)

### 4.1 Application Interface (81 methods)

- [x] Connect / GetActiveObject - via `connect_to_solidedge`
- [x] Quit - via `quit_application`
- [x] `DoIdle` - via `do_idle`
- [x] `StartCommand` - via `start_command`
- [x] `AbortCommand` - via `abort_command`
- [ ] `GetGlobalParameter` / `SetGlobalParameter` - Global settings
- [ ] `GetModelessTaskEventSource` - Event handling
- [x] `Activate` - via `activate_application`

Key Properties:
- [x] `ActiveDocument` - via document tools
- [x] `ActiveEnvironment` - via `get_active_environment`
- [x] `StatusBar` (get/put) - via `get_status_bar`, `set_status_bar`
- [ ] `DisplayAlerts` (get/put) - Alert suppression (have `set_performance_mode`)
- [x] `Visible` (get/put) - via `get_visible`, `set_visible`
- [ ] `SensorEvents` - Sensor monitoring

### 4.2 View Interface (60 methods)

- [x] `Fit` - via `zoom_fit`
- [x] `SetRenderMode` - via `set_display_mode`
- [x] `ApplyNamedView` - via `set_view`
- [x] `SaveAsImage` - via `capture_screenshot`
- [x] `GetCamera` / `SetCamera` - via `get_camera`, `set_camera`
- [x] `RotateCamera` - via `rotate_view`
- [x] `PanCamera` - via `pan_view`
- [x] `ZoomCamera` - via `zoom_view`
- [ ] `TransformModelToDC` / `TransformDCToModel` - **3D-to-screen coordinate mapping**
- [x] `Update` - via `refresh_view`
- [ ] `BeginCameraDynamics` / `EndCameraDynamics` - Smooth camera transitions

### 4.3 Variables Interface (13 methods)

- [x] `Item` / iteration - via `get_variables`, `get_variable`
- [x] `Edit` / value setting - via `set_variable`
- [x] `Add` - via `add_variable`
- [ ] `AddFromClipboard` - Variable from clipboard
- [x] `PutName` / `GetName` - via `rename_variable`
- [ ] `Translate` - Get variable by system name
- [x] `Query` - via `query_variables`
- [x] `GetFormula` - via `get_variable_formula`
- [x] `GetDisplayName` / `GetSystemName` - via `get_variable_names`
- [ ] `CopyToClipboard` - Copy variable value

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
- [ ] Material library access (full)

#### Layer Interface (15 methods)
- [ ] `Add` / `Delete` / `SetActive` - **Layer management**
- [ ] `Visible` / `Selectable` - Layer properties

#### FaceStyle Interface (49 methods, 69 properties)
- [x] Basic body color - via `set_body_color`, `set_face_color`
- [ ] `Texture` - Apply texture
- [ ] `Transparency` - Set transparency
- [ ] `Reflectivity` - Material appearance

---

## Part 5: Additional Type Libraries

### 5.1 Geometry Type Library (geometry.tlb) - 68 interfaces

Key interfaces for precise geometry queries:
- [ ] `Body` - Full body geometry access
- [ ] `Face` - Face geometry (normal, area, loops)
- [ ] `Edge` - Edge geometry (length, curvature, endpoints)
- [ ] `Vertex` - Vertex coordinates
- [ ] `Curve2d` / `Curve3d` - Parametric curve access
- [ ] `BSplineCurve` / `BSplineSurface` - NURBS data
- [ ] `Loop` - Face boundary loops
- [ ] `Shell` - Shell topology

### 5.2 FrameworkSupport (fwksupp.tlb) - 228 interfaces

2D geometry and annotation primitives:
- [ ] `Lines2d` / `Circles2d` / `Arcs2d` - 2D geometry collections
- [ ] `Dimensions` - Dimension management
- [ ] `SmartFrames2d` - Smart frame annotations
- [ ] `Symbols` - Symbol placement
- [ ] `BalloonTypes` / `BalloonNotes` - Balloon annotations
- [ ] `PMI` - Product Manufacturing Information

---

## Known Limitations

1. **Assembly constraints** require face/edge geometry selection which is complex to automate via COM
2. **Feature patterns** (model.Patterns) require SAFEARRAY(IDispatch) marshaling that fails in late binding
3. **Shell/Thinwalls** requires face selection for open faces, not automatable via COM
4. **Cutout via models.Add*Cutout** does NOT work - must use collection-level methods (ExtrudedCutouts.AddFiniteMulti)
5. **Mirror copy** creates feature tree entry but no geometry via COM (partially broken)
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