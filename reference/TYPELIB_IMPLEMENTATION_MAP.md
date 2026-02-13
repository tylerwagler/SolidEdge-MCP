# Solid Edge Type Library Implementation Map

Generated: 2026-02-12 | Source: 40 type libraries, 2,240 interfaces, 21,237 methods
Current: 254 MCP tools implemented

This document maps every actionable COM API surface from the Solid Edge type libraries
against our current MCP tool coverage. It identifies gaps and prioritizes what to implement next.

## Coverage Summary

|       Category        |       Available       |  Implemented  | Coverage  | Priority  |
|-----------------------|-----------------------|---------------|-----------|-----------|
|   **Part Features**   |    42     collections |       41      |    98%    |   HIGH    |
|     **Assembly**      |    18  key interfaces |       14      |    78%    |   HIGH    |
|   **Draft/Drawing**   |    15  key interfaces |        9      |    60%    |  MEDIUM   |
|   **Framework/App**   |    12  key interfaces |       12      |   100%    |    LOW    |
|  **Query/Topology**   |    25+        methods |       24      |    96%    |  MEDIUM   |
|   **Export/Import**   |    12         formats |       10      |    83%    |    LOW    |
|---------------------------------------------------------------------------------------|
|      **Total **       |   124           items |      110      |    89%    |    --     |
|---------------------------------------------------------------------------------------|

## Tool Count by Category (252 total)

| Category                  | Count | Tools |
|:--------------------------|:-----:|:---|
| **Connection**            | 7     | Connect, disconnect, app info, quit, is_connected, process_info, start_command |
| **Document Management**   | 13    | Create (part, assembly, sheet metal, draft), open, save, close, list, activate, undo, redo |
| **Sketching**             | 24    | Lines, circles, arcs (multiple), rects, polygons, ellipses, splines, points, constraints (9 types), fillet, chamfer, mirror, construction, hide profile, project_edge, include_edge |
| **Basic Primitives**      | 8     | Box (3 variants), cylinder, sphere, box cutout, cylinder cutout, sphere cutout |
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
| **Body Operations**       | 9     | Add body, thicken, mesh, tag, construction, delete holes, delete blends |
| **Simplification**        | 4     | Auto-simplify, enclosure, duplicate |
| **View/Display**          | 7     | Orientation, zoom, display mode, background color, get/set camera |
| **Variables**             | 5     | Get all, get by name, set value, add variable, query/search |
| **Custom Properties**     | 3     | Get all, set/create, delete |
| **Body Topology**         | 3     | Body faces, body edges, face info |
| **Performance**           | 2     | Set performance mode, recompute |
| **Query/Analysis**        | 25    | Mass properties, bounding box, features, measurements, facet data, solid bodies, modeling mode, face/edge info, colors, angles, volume, delete feature, material table, feature dimensions, material list/set/property |
| **Feature Management**    | 6     | Suppress, unsuppress, face rotate (2), draft angle, convert feature type |
| **Export**                | 10    | STEP, STL, IGES, PDF, DXF, flat DXF, Parasolid, JT, drawing, screenshot |
| **Assembly**              | 24    | Place, list, constraints, patterns, suppress, BOM, structured BOM, interference, bbox, relations, doc tree, replace, delete, visibility, color, transform, count, move, rotate |
| **Draft/Drawing**         | 11    | Sheets (add, activate, delete, rename), views, annotations (dimension, balloon, note, leader, text box), parts list |
| **Part Features**         | 10    | Dimple, etch, rib, lip, drawn cutout, bead, louver, gusset, thread, slot, split |
| **Diagnostics**           | 2     | API and feature inspection |
| **Select Set**            | 3     | Get selection, clear selection, add to selection |

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
- [ ] `AddNoHeal` - Delete without healing

#### DeleteHoles Collection (3 methods)
- [x] `Add` - via `create_delete_hole`
- [ ] `AddByFace` - Delete holes by face

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
- [ ] `AddCutoutByCenter` - Box cutout by center
- [x] `AddCutoutByTwoPoints` - via `create_box_cutout`
- [ ] `AddCutoutByThreePoints` - Box cutout by 3 points

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
- [ ] `GetProfiles` / `SetProfiles` - Get/set sketch profiles
- [ ] `GetDirection1Extent` / `ApplyDirection1Extent` - Edit extent
- [ ] `GetDirection2Extent` / `ApplyDirection2Extent` - Edit 2nd direction
- [ ] `GetFromFaceOffsetData` / `SetFromFaceOffsetData` - Edit offsets
- [ ] `GetThinWallOptions` / `SetThinWallOptions` - Edit thin wall params
- [ ] `GetBodyArray` / `SetBodyArray` - Multi-body targeting
- [x] `ConvertToCutout` / `ConvertToProtrusion` - via `convert_feature_type`
- [ ] `GetStatusEx` - Detailed feature status
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
- [ ] `Visible` (get/put) - **Show/hide occurrence** (have `set_component_visibility`)
- [ ] `Subassembly` (get) - Check if subassembly
- [ ] `SubOccurrences` (get) - Access nested components
- [ ] `Bodies` (get) - Access body geometry
- [ ] `OccurrenceDocument` (get) - Open document reference
- [ ] `DisplayName` (get) - Display name
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

Key methods NOT implemented:
- [ ] `GetSectionCuts` / `AddSectionCut` - **Section/cross-section views**
- [ ] `Update` - Force view update
- [ ] `Activate` / `Deactivate` - Activate for editing
- [ ] `Delete` - Remove view
- [ ] `Move` - Reposition view on sheet
- [ ] `SetOrientation` - Change orientation

Key properties NOT exposed:
- [ ] `Scale` (get/put) - **View scale**
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
- [ ] `AbortCommand` - Abort running command
- [ ] `GetGlobalParameter` / `SetGlobalParameter` - Global settings
- [ ] `GetModelessTaskEventSource` - Event handling
- [ ] `Activate` - Bring SE to foreground

Key Properties:
- [x] `ActiveDocument` - via document tools
- [ ] `ActiveEnvironment` - Current environment name
- [ ] `StatusBar` - Status bar text
- [ ] `DisplayAlerts` (get/put) - Alert suppression (have `set_performance_mode`)
- [ ] `Visible` (get/put) - Show/hide SE window
- [ ] `SensorEvents` - Sensor monitoring

### 4.2 View Interface (60 methods)

- [x] `Fit` - via `zoom_fit`
- [x] `SetRenderMode` - via `set_display_mode`
- [x] `ApplyNamedView` - via `set_view`
- [x] `SaveAsImage` - via `capture_screenshot`
- [x] `GetCamera` / `SetCamera` - via `get_camera`, `set_camera`
- [ ] `RotateCamera` - Orbit view
- [ ] `PanCamera` - Pan view
- [ ] `ZoomCamera` - Zoom by factor
- [ ] `TransformModelToDC` / `TransformDCToModel` - **3D-to-screen coordinate mapping**
- [ ] `Update` - Force view refresh
- [ ] `BeginCameraDynamics` / `EndCameraDynamics` - Smooth camera transitions

### 4.3 Variables Interface (13 methods)

- [x] `Item` / iteration - via `get_variables`, `get_variable`
- [x] `Edit` / value setting - via `set_variable`
- [x] `Add` - via `add_variable`
- [ ] `AddFromClipboard` - Variable from clipboard
- [ ] `PutName` / `GetName` - Rename variable
- [ ] `Translate` - Get variable by system name
- [x] `Query` - via `query_variables`
- [ ] `GetFormula` - Get variable formula
- [ ] `GetDisplayName` / `GetSystemName` - Display vs system names
- [ ] `CopyToClipboard` - Copy variable value

### 4.4 SelectSet Interface (13 methods)

- [x] `RemoveAll` - via `clear_select_set`
- [x] `Item` / `Count` - via `get_select_set`
- [x] `Add` - via `select_add`
- [ ] `Remove` - Remove specific item
- [ ] `AddAll` - Select everything
- [ ] `Copy` / `Cut` / `Delete` - Clipboard operations on selection
- [ ] `SuspendDisplay` / `ResumeDisplay` / `RefreshDisplay` - Selection display control

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

## Part 6: Priority Implementation Roadmap

### Tier 1: HIGH IMPACT, LIKELY FEASIBLE (14 tools) - ✅ 14/14 COMPLETE

These fill critical gaps and use proven API patterns:

| # | Tool | API Method | Status |
|---|------|-----------|--------|
| 1 | `create_swept_cutout` | SweptCutouts.Add | ✅ Implemented |
| 2 | `create_helix_cutout` | HelixCutouts.AddFinite | ✅ Implemented |
| 3 | `create_variable_round` | Rounds.AddVariable | ✅ Implemented |
| 4 | `create_blend` | Blends.Add | ✅ Implemented |
| 5 | `create_emboss` | EmbossFeatures.Add | ✅ Implemented |
| 6 | `create_ref_plane_by_angle` | RefPlanes.AddAngularByAngle | ✅ Implemented |
| 7 | `create_ref_plane_by_3_points` | RefPlanes.AddBy3Points | ✅ Implemented |
| 8 | `create_ref_plane_midplane` | RefPlanes.AddMidPlane | ✅ Implemented |
| 9 | `create_hole_through_all` | ExtrudedCutouts.AddThroughAllMulti | ✅ Implemented |
| 10 | `create_flange` | Flanges.Add | ✅ Implemented |
| 11 | `add_variable` | Variables.Add | ✅ Implemented |
| 12 | `select_add` | SelectSet.Add | ✅ Implemented |
| 13 | `create_box_cutout` | BoxFeatures.AddCutoutByTwoPoints | ✅ Implemented |
| 14 | `create_cylinder_cutout` | CylinderFeatures.AddCutoutByCenterAndRadius | ✅ Implemented |

**Bonus**: `create_sphere_cutout` (SphereFeatures.AddCutoutByCenterAndRadius) also implemented.

### Tier 2: MEDIUM IMPACT, LIKELY FEASIBLE (16 tools) - ✅ 10/16 IMPLEMENTED

| # | Tool | API Method | Status |
|---|------|-----------|--------|
| 15 | `create_extruded_cutout_from_to` | ExtrudedCutouts.AddFromToMulti | ⏳ Pending (requires face/plane selection) |
| 16 | `create_extruded_cutout_through_next` | ExtrudedCutouts.AddThroughNextMulti | ✅ Implemented |
| 17 | `create_normal_cutout_through_all` | NormalCutouts.AddThroughAllMulti | ✅ Implemented |
| 18 | `create_delete_hole` | DeleteHoles.Add | ✅ Implemented |
| 19 | `create_delete_blend` | DeleteBlends.Add | ✅ Implemented |
| 20 | `create_revolved_surface` | RevolvedSurfaces.AddFinite | ✅ Implemented |
| 21 | `create_lofted_surface` | LoftedSurfaces.Add | ✅ Implemented |
| 22 | `create_swept_surface` | SweptSurfaces.Add | ✅ Implemented |
| 23 | `create_bend` | Bends.Add | ⏳ Pending (14 params, complex SM) |
| 24 | `create_jog` | (Jog feature) | ⏳ Pending (complex SM) |
| 25 | `create_hem` | (Hem feature) | ⏳ Pending (complex SM) |
| 26 | `create_close_corner` | CloseCorners.Add | ⏳ Pending (complex SM) |
| 27 | `create_folded_view` | DrawingViews.AddByFold | ⏳ Pending (needs parent DrawingView) |
| 28 | `create_detail_view` | DrawingViews.AddByDetailEnvelope | ⏳ Pending (needs parent DrawingView) |
| 29 | `get_camera` / `set_camera` | View.GetCamera/SetCamera | ✅ Implemented |
| 30 | `query_variables` | Variables.Query | ✅ Implemented |

### Tier 3: MEDIUM IMPACT, MORE COMPLEX (12 tools) - ✅ 9/12 IMPLEMENTED (11 tools)

| # | Tool | API Method | Status |
|---|------|-----------|--------|
| 31 | `project_edge` | Profile.ProjectEdge | ✅ Implemented |
| 32 | `include_edge` | Profile.IncludeEdge | ✅ Implemented |
| 33 | `get_feature_dimensions` | Feature.GetDimensions | ✅ Implemented |
| 34 | `convert_feature_type` | Feature.ConvertToCutout/Protrusion | ✅ Implemented |
| 35 | `create_assembly_cutout` | AssemblyFeaturesExtrudedCutout | ⏳ Complex face selection needed |
| 36 | `create_assembly_hole` | AssemblyFeaturesHole | ⏳ Complex face selection needed |
| 37 | `create_parts_list` | PartsList (draft) | ✅ Implemented |
| 38 | `start_command` | Application.StartCommand | ✅ Implemented |
| 39 | `occurrence_move` | Occurrence.Move | ✅ Implemented |
| 40 | `occurrence_rotate` | Occurrence.Rotate | ✅ Implemented |
| 41 | `set_material` | MatTable.ApplyMaterial + GetMaterialList + GetMatPropValue | ✅ 3 tools (set_material, get_material_list, get_material_property) |
| 42 | `convert_to_sheet_metal` | ConvertPartToSM | ⏳ Complex - needs face selection |

### Tier 4: LOW PRIORITY / EXPERIMENTAL (future)

- Feature patterns (Ex variants - investigate if SAFEARRAY issue is avoided)
- Assembly feature patterns/mirrors
- Full relation creation with face selection (AddPlanar, AddAxial, etc.)
- Wiring/harness features (Cable, Wire, Bundle)
- Structural frames
- DraftPrintUtility
- Advanced surface operations (BlueSurf, SurfaceByBoundary, IntersectSurface)
- FEA/Simulation interfaces (Study, Load, Constraint, etc.)

---

## Part 7: Key Constants Reference

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

## Appendix: Interface Method Signatures for Tier 1 Items

### SweptCutouts.Add
```
Add(NumCurves: VT_I4, TraceCurves: VT_VARIANT, TraceCurveTypes: VT_VARIANT,
    NumSections: VT_I4, CrossSections: VT_VARIANT, CrossSectionTypes: VT_VARIANT,
    Origins: VT_VARIANT, SegmentMaps: VT_VARIANT, MaterialSide: FeaturePropertyConstants,
    StartExtentType: FeaturePropertyConstants, StartExtentDistance: VT_R8,
    StartSurfaceOrRefPlane: VT_DISPATCH, EndExtentType: FeaturePropertyConstants,
    EndExtentDistance: VT_R8, EndSurfaceOrRefPlane: VT_DISPATCH) -> SweptCutout*
```

### Rounds.AddVariable
```
AddVariable(NumberOfEdgeSets: VT_I4, EdgeSetArray: SAFEARRAY(VT_DISPATCH)*,
            RadiusArray: SAFEARRAY(VT_R8)*) -> Round*
```

### Blends.Add
```
Add(NumberOfEdgeSets: VT_I4, EdgeSetArray: SAFEARRAY(VT_DISPATCH)*,
    RadiusArray: SAFEARRAY(VT_R8)*) -> Blend*
```

### RefPlanes.AddAngularByAngle
```
AddAngularByAngle(ParentPlane: VT_DISPATCH, Angle: VT_R8,
                  NormalSide: FeaturePropertyConstants,
                  [Edge: VT_VARIANT]) -> RefPlane*
```

### RefPlanes.AddBy3Points
```
AddBy3Points(Point1X: VT_R8, Point1Y: VT_R8, Point1Z: VT_R8,
             Point2X: VT_R8, Point2Y: VT_R8, Point2Z: VT_R8,
             Point3X: VT_R8, Point3Y: VT_R8, Point3Z: VT_R8) -> RefPlane*
```

### RefPlanes.AddMidPlane
```
AddMidPlane(Plane1: VT_DISPATCH, Plane2: VT_DISPATCH) -> RefPlane*
```

### Holes.AddThroughAll
```
AddThroughAll(Profile: VT_DISPATCH, ProfilePlaneSide: FeaturePropertyConstants,
              HoleData: HoleData*) -> Hole*
```

### Flanges.Add
```
Add(Profile: VT_DISPATCH, ThicknessSide: FeaturePropertyConstants,
    Direction: FeaturePropertyConstants, DepthType: FeaturePropertyConstants,
    Depth: VT_R8, ...) -> Flange*
```

### Variables.Add
```
Add(pName: VT_BSTR, pFormula: VT_BSTR, [UnitsType: VT_VARIANT]) -> variable*
```

### SelectSet.Add
```
Add(Dispatch: VT_DISPATCH) -> VT_VOID
```

### BoxFeatures.AddCutoutByTwoPoints
```
AddCutoutByTwoPoints(x1: VT_R8, y1: VT_R8, z1: VT_R8,
                     x2: VT_R8, y2: VT_R8, z2: VT_R8,
                     dAngle: VT_R8, dDepth: VT_R8,
                     pPlane: VT_DISPATCH, ExtentSide: FeaturePropertyConstants,
                     vbKeyPointExtent: VT_BOOL, pKeyPointObj: VT_DISPATCH,
                     pKeyPointFlags: KeyPointExtentConstants*) -> Model*
```

### EmbossFeatures.Add
```
Add(Profile: VT_DISPATCH, FaceToEmboss: VT_DISPATCH,
    EmbossType: FeaturePropertyConstants, ...) -> EmbossFeature*
```

---

## Appendix: Future Server Refactoring

### Problem
`server.py` is a monolith of ~4300 lines containing all 252 `@mcp.tool()` wrappers. This is a significant context sink for LLMs and hard to maintain.

### Proposed Solution
1. **Modularize** tool definitions into `src/solidedge_mcp/tools/` (connection.py, documents.py, sketching.py, etc.)
2. **Dynamic registration** via `mcp.add_tool(func)` in `tools/__init__.py`
3. **Thin entry point**: `server.py` becomes <100 lines (init FastMCP + call `register_tools`)

### Target Structure
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

### Notes
- `mcp.add_tool` method confirmed available on FastMCP
- Existing unit tests verify `backends` logic and should pass without modification
- A verification script will confirm all tools register correctly after refactoring
