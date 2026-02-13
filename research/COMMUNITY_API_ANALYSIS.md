# Solid Edge API Coverage Analysis

**Generated:** 2026-02-12 (Updated)
**Source:** [SolidEdgeCommunity GitHub](https://github.com/SolidEdgeCommunity) repositories
**Purpose:** Identify all available Solid Edge COM API operations vs. our MCP server implementation

---

## Summary

| Metric | Count |
|--------|-------|
| **Our implemented MCP tools** | 218 |
| **API operations found in community repos** | 250+ |
| **Operations we're missing** | ~20 |
| **High-priority gaps remaining** | ~3 |
| **Community repos analyzed** | 10 |

---

## Table of Contents

1. [Connection & Application](#1-connection--application)
2. [Document Management](#2-document-management)
3. [Sketching & 2D Geometry](#3-sketching--2d-geometry)
4. [2D Constraints (Relations2d)](#4-2d-constraints-relations2d)
5. [Extrusions](#5-extrusions)
6. [Revolves](#6-revolves)
7. [Cutouts](#7-cutouts)
8. [Loft & Sweep](#8-loft--sweep)
9. [Rounds, Chamfers, Fillets](#9-rounds-chamfers-fillets)
10. [Holes & Patterns](#10-holes--patterns)
11. [Advanced Part Features](#11-advanced-part-features)
12. [Reference Planes](#12-reference-planes)
13. [Primitives](#13-primitives)
14. [Sheet Metal](#14-sheet-metal)
15. [Assembly](#15-assembly)
16. [Draft / Drawing](#16-draft--drawing)
17. [Query & Analysis](#17-query--analysis)
18. [Export & Import](#18-export--import)
19. [View & Display](#19-view--display)
20. [Selection & UI](#20-selection--ui)
21. [Properties & Metadata](#21-properties--metadata)
22. [Variables & Dimensions](#22-variables--dimensions)
23. [Events](#23-events)
24. [File Reading (Without Solid Edge)](#24-file-reading-without-solid-edge)
25. [Performance & Batch](#25-performance--batch)
26. [Feature Management](#26-feature-management)
27. [Body & Topology Query](#27-body--topology-query)
28. [Structural Frames & Tubes](#28-structural-frames--tubes)
29. [Priority Recommendations](#29-priority-recommendations)

---

## 1. Connection & Application

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Connect to running instance | `Marshal.GetActiveObject("SolidEdge.Application")` | YES | Working | Samples |
| Start new instance | `Dispatch("SolidEdge.Application")` | YES | Working | Samples |
| Get app info (version, path) | `Application.Version`, `Application.ProcessID` | YES | Working | SDK |
| Quit application | `Application.Quit()` | YES | Working | Interop |
| Disconnect (release COM) | Release COM references | YES | Working | — |
| Check connection status | Connection state check | YES | Working | — |
| Get process info (PID, hWnd) | `Application.ProcessID`, `Application.hWnd` | YES | Working | SDK |
| Get installation info | `SEInstallData.GetInstalledPath/Language/Version` | YES | Working | SDK |

## 2. Document Management

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Create part document | `Documents.Add("SolidEdge.PartDocument")` | YES | Working | Samples |
| Create assembly document | `Documents.Add("SolidEdge.AssemblyDocument")` | YES | Working | Samples |
| Create draft document | `Documents.Add("SolidEdge.DraftDocument")` | YES | Working | Samples |
| Create sheet metal document | `Documents.Add("SolidEdge.SheetMetalDocument")` | YES | Working | Samples |
| Create weldment document | `Documents.Add("SolidEdge.WeldmentDocument")` | YES | Working | SDK |
| Open document | `Documents.Open(filename)` | YES | Working | Samples |
| Import file (auto type detection) | `Documents.Open(filename)` | YES | Working | — |
| Open in background (no window) | `Documents.Open(filename, 0x8)` | YES | Working | SDK |
| Save document | `Document.Save()` | YES | Working | Samples |
| Save as | `Document.SaveAs(filename)` | YES | Working | Samples |
| Close document | `Document.Close()` | YES | Working | Samples |
| Close all documents | Iterate + close each doc | YES | Working | Samples |
| List documents | Documents collection iteration | YES | Working | — |
| Get active document type | `Application.ActiveDocument` type detection | YES | Working | — |
| Get document count | `Application.Documents.Count` | YES | Working | — |
| Activate document | `Document.Activate()` | YES | Working | — |
| Undo / Redo | `Document.Undo()` / `Document.Redo()` | YES | Working | — |
| Recompute document/model | `Document.Recompute()` | YES | Working | Samples |
| Get/set modeling mode | `PartDocument.ModelingMode` | YES | Working | Samples |

## 3. Sketching & 2D Geometry

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Create sketch on named plane | `ProfileSets.Add() + Profiles.Add(refPlane)` | YES | Working | Samples |
| Create sketch on plane by index | Same, with `RefPlanes.Item(index)` | YES | Working | — |
| Draw line | `Lines2d.AddBy2Points(x1,y1,x2,y2)` | YES | Working | Samples |
| Draw construction line | `Lines2d.AddBy2Points + ToggleConstruction` | YES | Working | — |
| Draw circle | `Circles2d.AddByCenterRadius(x,y,r)` | YES | Working | Samples |
| Draw arc (center/start/end) | `Arcs2d.AddByCenterStartEnd(...)` | YES | Working | Samples |
| Draw arc (start/center/end) | `Arcs2d.AddByStartCenterEnd(...)` | YES | Working | Interop |
| Draw rectangle | Lines2d (4 lines) | YES | Working | — |
| Draw polygon | Lines2d (n lines) | YES | Working | Samples |
| Draw ellipse | `Ellipses2d.AddByCenterMajorMinor(...)` | YES | Working | Interop |
| Draw B-spline | `BSplineCurves2d.AddByPoints(order, size, array)` | YES | Working | Samples |
| Draw point | `Points2d.Add(x, y)` | YES | Working | — |
| Sketch fillet | `Arcs2d.AddByFillet(line1, line2, radius)` | YES | Working | — |
| Sketch chamfer | `Lines2d.AddByChamfer(line1, line2, d1, d2)` | YES | Working | — |
| Sketch offset | `Profile.OffsetProfile(distance)` | YES | Working | — |
| Sketch mirror | Mirror geometry across axis | YES | Working | — |
| Get sketch info | Profile element counts | YES | Working | — |
| Get sketch constraints | `Relations2d` iteration | YES | Working | — |
| Mirror B-spline | `BSplineCurve2d.Mirror(x1,y1,x2,y2,copy)` | YES | Working | Samples |
| Draw circle by 2 points | `Circles2d.AddBy2Points(...)` | YES | Working | Interop |
| Draw circle by 3 points | `Circles2d.AddBy3Points(...)` | YES | Working | Interop |
| Set axis of revolution | `Profile.SetAxisOfRevolution(axis)` | YES | Working | Samples |
| Toggle construction geometry | `Profile.ToggleConstruction(element)` | YES (internal) | Working | Samples |
| Close/validate sketch | `Profile.End(flags)` | YES | Working | Samples |
| Hide profile | `Profile.Visible = false` | YES | Working | Samples |

## 4. 2D Constraints (Relations2d)

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Keypoint constraint | `Relations2d.AddKeypoint(obj1, kp1, obj2, kp2)` | YES | Working | Samples |
| Equal constraint | `Relations2d.AddEqual(obj1, obj2)` | YES | Working | Samples |
| Parallel constraint | `Relations2d.AddParallel(obj1, obj2)` | YES | Working | Interop |
| Perpendicular constraint | `Relations2d.AddPerpendicular(obj1, obj2)` | YES | Working | Interop |
| Concentric constraint | `Relations2d.AddConcentric(obj1, obj2)` | YES | Working | Interop |
| Tangent constraint | `Relations2d.AddTangent(obj1, obj2)` | YES | Working | Interop |
| Horizontal constraint | `Relations2d.AddHorizontal(obj)` | YES | Working | Interop |
| Vertical constraint | `Relations2d.AddVertical(obj)` | YES | Working | Interop |

**Note:** Elements are referenced as `[type, index]` pairs (e.g., `["line", 1]`). The `add_constraint` tool supports all 7 constraint types, and `add_keypoint_constraint` handles endpoint-to-endpoint connections.

## 5. Extrusions

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Finite extrusion | `Models.AddFiniteExtrudedProtrusion(...)` | YES | Working | Samples |
| Through-all extrusion | `ExtrudedProtrusions.AddThroughAll(...)` | YES (untested) | Untested | Interop |
| Infinite extrusion | `Models.AddExtrudedProtrusion(...)` | YES (untested) | Untested | — |
| Thin-wall extrusion | `Models.AddExtrudedProtrusionWithThinWall(...)` | YES (untested) | Broken | — |
| Extruded surface | `Constructions.ExtrudedSurfaces.Add(...)` | YES | Working | Samples |

## 6. Revolves

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Finite revolve | `Models.AddFiniteRevolvedProtrusion(...)` | YES | Working | Samples |
| Full revolve (360) | `Models.AddFiniteRevolvedProtrusion(..., 2*pi)` | YES | Working | — |
| Sync revolve | `Models.AddRevolvedProtrusionSync(...)` | YES (untested) | Untested | — |
| Thin-wall revolve | `Models.AddRevolvedProtrusionWithThinWall(...)` | YES (untested) | Untested | — |

## 7. Cutouts

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Finite extruded cutout | `ExtrudedCutouts.AddFiniteMulti(...)` | YES | Working | Samples |
| Through-all extruded cutout | `ExtrudedCutouts.AddThroughAllMulti(...)` | YES | Working | Samples |
| Revolved cutout | `RevolvedCutouts.AddFiniteMulti(...)` | YES | Working | Samples |
| Normal cutout | `NormalCutouts.AddFiniteMulti(...)` | YES | Working | MEMORY |
| Lofted cutout | `LoftedCutouts.AddSimple(...)` | YES | Working | Samples |
| Drawn cutout | `DrawnCutouts.Add(...)` | YES | Working | MEMORY |

## 8. Loft & Sweep

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Loft (initial feature) | `Models.AddLoftedProtrusion(...)` | YES | Working | Samples |
| Loft (on existing body) | `LoftedProtrusions.AddSimple(...)` | YES | Working | Samples |
| Loft thin-wall | `Models.AddLoftedProtrusionWithThinWall(...)` | YES (untested) | Untested | — |
| Sweep | `Models.AddSweptProtrusion(...)` | YES | Working | Samples |
| Sweep thin-wall | `Models.AddSweptProtrusionWithThinWall(...)` | YES (untested) | Untested | — |
| Ref plane normal to curve | `RefPlanes.AddNormalToCurve(...)` | YES | Working | Samples |

## 9. Rounds, Chamfers, Fillets

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Round (fillet) - multiple edges | `Rounds.Add(edgeCount, edgeArray, radiusArray)` | YES | Working | Samples |
| Round on specific face | `Rounds.Add` with face-specific edges | YES | Working | — |
| Chamfer - equal setback | `Chamfers.AddEqualSetback(count, edges, distance)` | YES | Working | Samples |
| Chamfer on specific face | `Chamfers.AddEqualSetback` with face-specific edges | YES | Working | — |
| Chamfer - unequal setback | `Chamfers.AddUnequalSetback(face, count, edges, d1, d2)` | YES | Working | Samples |
| Chamfer - setback + angle | `Chamfers.AddSetbackAngle(face, count, edges, dist, angle)` | YES | Working | Samples |
| Blend | `Blends.Add(...)` / `AddSurfaceBlend` / `AddVariable` | NO | Available | MEMORY |

## 10. Holes & Patterns

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Create hole (via circular cutout) | `ExtrudedCutouts.AddFiniteMulti` (circular profile) | YES | Working | Samples |
| Place finite hole (HoleData API) | `Holes.AddFinite(profile, side, depth, holeData)` | YES | Working | Samples |
| User-defined pattern | `UserDefinedPatterns.AddByProfiles(...)` | NO | **Available** | Samples |
| Rectangular pattern | `Patterns.AddByRectangular(...)` | NO | Broken (SAFEARRAY) | MEMORY |
| Circular pattern | `Patterns.AddByCircular(...)` | NO | Broken (SAFEARRAY) | MEMORY |

## 11. Advanced Part Features

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Mirror copy | `MirrorCopies.Add(plane, count, features)` | YES | Working | MEMORY |
| Rib | `Ribs.Add(profile, extensionType, thicknessType, side, thickness)` | YES | Working | MEMORY |
| Slot | `Slots.Add(profile, slotType, endType, width, ..., 22 params)` | YES | Working | Samples |
| Thread | `Threads.Add(face, pitch)` | YES | Working | MEMORY |
| Split | `Splits.Add(profile, side)` | YES | Working | MEMORY |
| Lip | `Lips.Add(...)` | YES | Working | MEMORY |
| Delete faces | `DeleteFaces.Add(faceSetToDelete)` | YES | Working | MEMORY |
| Thicken feature | `Thickens.Add(side, thickness, faceCount, facesArray)` | YES (untested) | Untested | Samples |
| Face rotate (by geometry) | `FaceRotates.Add(face, ByGeometry, ..., edge, ..., angle)` | YES | Working | Samples |
| Face rotate (by points) | `FaceRotates.Add(face, ByPoints, ..., pt1, pt2, ..., angle)` | YES | Working | Samples |
| Draft angle | `Drafts.Add(plane, numSets, faceArray, draftAngleArray, side)` | YES | Working | MEMORY |
| Emboss | `EmbossFeatures.Add(...)` | NO | Available | MEMORY |
| Shell (thin wall) | Shell feature (requires face selection) | NO | Complex | — |
| Suppress feature | `Feature.Suppress()` | YES | Working | — |
| Unsuppress feature | `Feature.Unsuppress()` | YES | Working | — |
| Move to synchronous | `MoveOrderedFeaturesToSynchronous` | NO | Available | Samples |
| Heal and optimize body | `HealAndOptimizeBody` | NO | Available | Samples |
| Place feature library | Feature library placement | NO | Available | Samples |

## 12. Reference Planes

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Offset parallel plane | `RefPlanes.AddParallelByDistance(parent, dist, side)` | YES | Working | Samples |
| List reference planes | `RefPlanes` iteration | YES | Working | — |
| Normal to curve | `RefPlanes.AddNormalToCurve(curve, end, plane, pivot, ...)` | YES | Working | Samples |
| Get top plane | `RefPlanes.Item(1)` / `GetTopPlane()` | YES (internal) | Working | SDK |
| Get front plane | `RefPlanes.Item(2)` / `GetFrontPlane()` | YES (internal) | Working | SDK |
| Get right plane | `RefPlanes.Item(3)` / `GetRightPlane()` | YES (internal) | Working | SDK |
| Parallel plane (alternate) | `RefPlanes.AddParallel(...)` | NO | Available | Interop |

## 13. Primitives

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Box by center | `Models.AddBoxByCenter(...)` | YES | Working | — |
| Box by two points | `Models.AddBoxByTwoPoints(...)` | YES | Working | — |
| Box by three points | `Models.AddBoxByThreePoints(...)` | YES | Working | — |
| Cylinder | `Models.AddCylinderByCenterAndRadius(...)` | YES | Working | — |
| Sphere | `Models.AddSphereByCenterAndRadius(...)` | YES | Working | — |

## 14. Sheet Metal

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Create sheet metal doc | `Documents.Add("SolidEdge.SheetMetalDocument")` | YES | Working | Samples |
| Base tab | `Tabs.Add(...)` | YES (untested) | Untested | Samples |
| Base flange | `Models.AddBaseContourFlange(...)` | YES (untested) | Untested | — |
| Lofted flange | `Models.AddLoftedFlange(...)` | YES (untested) | Untested | Samples |
| Web network | `Models.AddWebNetwork(...)` | YES (untested) | Untested | — |
| Dimple | `Dimples.Add(...)` | YES | Working | Samples |
| Etch | `Etches.Add(...)` | YES | Working | Samples |
| Louver | `Louvers.Add(...)` | YES | Working | MEMORY |
| Gusset | `Gussets.Add(...)` | YES | Working | MEMORY |
| Bead | `Beads.Add(...)` | YES | Working | MEMORY |
| Drawn cutout | `DrawnCutouts.Add(...)` | YES | Working | MEMORY |
| SM extruded cutout | `ExtrudedCutouts.Add(...)` (23 params) | YES (via general cutout) | Working | Samples |
| SM holes with pattern | `Holes.AddFinite + UserDefinedPatterns.AddByProfiles` | NO | **Available** | Samples |
| Save flat DXF | `SheetMetalDocument.SaveAsFlatDXF(...)` | YES | Working | Samples |

## 15. Assembly

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Place by filename | `Occurrences.AddByFilename(path)` | YES | Working | Samples |
| Place with matrix | `Occurrences.AddWithMatrix(path, matrix)` | YES | Working | Samples |
| Place with transform | `Occurrences.AddWithTransform(path, ox,oy,oz,ax,ay,az)` | NO | Available | Samples |
| List occurrences | Occurrences iteration | YES | Working | Samples |
| Get occurrence count | `Occurrences.Count` | YES | Working | — |
| Get transform | `Occurrence.GetTransform(...)` | YES | Working | Samples |
| Get 4x4 matrix | `Occurrence.GetMatrix(...)` | YES | Working | Samples |
| Get range box | `Occurrence.GetRangeBox(min, max)` | YES | Working | Samples |
| Check interference | `AssemblyDocument.CheckInterference(...)` | YES | Working | Samples |
| Report BOM (flat) | BOM traversal via Occurrences | YES | Working | Samples |
| Report BOM (structured) | Recursive `SubOccurrences` traversal | YES | Working | Samples |
| Report document tree | Recursive `SubOccurrences` traversal | YES | Working | Samples |
| Relations3d query | `Relations3d` iteration (15+ constraint types) | YES | Working | Samples |
| Replace component | `Occurrence.Replace(newFilePath)` | YES | Working | — |
| Set component color | `Occurrence.SetColor(...)` | YES | Working | — |
| Set component visibility | `Occurrence.Visible = bool` | YES | Working | — |
| Delete component | `Occurrence.Delete()` | YES | Working | — |
| Ground component | `Occurrence.Ground/Unground` | YES | Working | — |
| Suppress component | `Occurrence.Suppress/Unsuppress` | YES | Working | — |
| Pattern component | Matrix-based copies | YES | Working | — |
| Configurations | `Configurations.Add/Apply/Item` | NO | **Available** | Samples |
| Derived configurations | `Configurations.AddDerivedConfig(...)` | NO | Available | Samples |
| Delete ground relations | `Relations3d.Delete()` | NO | Available | Samples |
| Detect under-constrained | Under-constrained occurrence detection | NO | Available | Samples |
| Structural frames | `StructuralFrames.Add(partFile, numPaths, pathArray)` | NO | Available | Samples |
| Report tubes | `Occurrence.IsTube()` / `GetTube()` | NO | Available | Samples |
| Tube bend table | `Tube.BendTable(...)` | NO | Available | Samples |

## 16. Draft / Drawing

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Create draft document | `Documents.Add("SolidEdge.DraftDocument")` | YES | Working | Samples |
| Link 3D model | `ModelLinks.Add(filename)` | YES | Working | Samples |
| Add part view | `DrawingViews.AddPartView(link, orient, scale, x, y)` | YES | Working | Samples |
| Add assembly view | `DrawingViews.AddAssemblyView(link, orient, scale, x, y, type)` | YES | Working | Samples |
| Add sheet | `Sheets.AddSheet()` | YES | Working | Samples |
| Activate sheet | `Sheet.Activate()` | YES | Working | — |
| Get sheet info | `Sheets` iteration (Name, Index, Number) | YES | Working | — |
| Rename sheet | `Sheet.Name = newName` | YES | Working | — |
| Delete sheet | `Sheet.Delete()` | YES | Working | — |
| Create balloon | `Balloons.Add(x, y, z)` | YES | Working | Samples |
| Create leader | `Leaders.Add(x1,y1,z1,x2,y2,z2)` | YES | Working | Samples |
| Create text box | `TextBoxes.Add(x, y, z)` | YES | Working | Samples |
| Create note | `TextBoxes.Add(x, y, z)` + text | YES | Working | — |
| Create dimension | `Dimensions.AddLength(x1,y1,z1,x2,y2,z2,dx,dy,dz)` | YES | Working | Samples |
| Report drawing views | `DrawingViews` iteration | NO | Available | Samples |
| Report sections | `Sections` iteration | NO | Available | Samples |
| Update drawing views | `DrawingView.Update()` | NO | Available | Samples |
| Convert views to 2D | ConvertDrawingViewsTo2D | NO | Available | Samples |
| Draw 2D lines in draft | `Lines2d.AddBy2Points(...)` (in draft sheet context) | NO | Available | Samples |
| Draw 2D circles in draft | `Circles2d.AddByCenterRadius(...)` (in draft) | NO | Available | Samples |
| Batch print | Batch printing of sheets | NO | Available | Samples |
| Copy parts lists | `CopyPartsListsToClipboard` | NO | Available | Samples |
| Export to Excel | `WritePartsListsToExcel` | NO | Available | Samples |
| Export sheet as EMF | `Sheet.SaveAsEnhancedMetafile(filename)` | NO | **Available** | Samples |
| Delete all drawing objects | Mass deletion of drawing objects | NO | Available | Samples |
| Convert property text | Property text conversion | NO | Available | Samples |

## 17. Query & Analysis

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Mass properties (with density) | `Model.ComputePhysicalProperties(density, accuracy, ...)` | YES | Working | Samples |
| Mass properties (cached) | `Model.GetPhysicalProperties(...)` | NO | Available | Samples |
| Bounding box | `Body.GetRange(min, max)` | YES | Working | Samples |
| Body volume | `Body.Volume` | YES | Working | Samples |
| Surface area | `Body.SurfaceArea` | YES | Working | — |
| Face area | `Face.Area` | YES | Working | — |
| Face count | `Body.Faces(igQueryAll).Count` | YES | Working | — |
| Edge info | `Face.Edges` iteration | YES | Working | — |
| Edge count | `Body.Edges` count | YES | Working | — |
| Center of gravity | `Variables.CoMX/Y/Z` or computed | YES | Working | — |
| Moments of inertia | `ComputePhysicalPropertiesWithSpecifiedDensity` | YES | Working | — |
| Body color | `Body.Style.ForegroundColor` | YES | Working | — |
| Set face color | `Face.SetColor(r, g, b)` | YES | Working | — |
| Set body color | `Body.SetColor(r, g, b)` | YES | Working | — |
| Measure distance | Math calculation | YES | Working | — |
| Measure angle | Math (dot product, acos) | YES | Working | — |
| Material table | `Variables` query for material props | YES | Working | — |
| List features (edgebar) | `DesignEdgebarFeatures` iteration | YES | Working | Samples |
| Feature count | `Models.Count` | YES | Working | — |
| Feature info | Feature property access | YES | Working | — |
| Document properties | Document property access | YES | Working | Samples |
| List reference planes | `RefPlanes` iteration | YES | Working | — |
| Body facet data (mesh) | `Body.GetFacetData(tolerance, ...)` | YES | Working | Samples |
| Report solid bodies | `Models` + `Body.IsSolid`, `Shells.IsClosed` | YES | Working | Samples |
| Body shell info | `Body.Shells.Item(i).IsClosed/IsVoid` | NO | Available | Samples |
| Construction bodies | `Constructions` collection | NO | Available | Samples |
| Report variables | `Variables.Query(criteria, nameBy, varType)` | NO | **Available** | Samples |

## 18. Export & Import

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Export STEP | `Document.SaveAs(...)` | YES | Working | — |
| Export STL | `Document.SaveAs(...)` | YES | Working | — |
| Export IGES | `Document.SaveAs(...)` | YES | Working | — |
| Export PDF | `Document.SaveAs(...)` | YES | Working | — |
| Export DXF | `Document.SaveAs(...)` | YES | Working | — |
| Export Parasolid | `Document.SaveAs(...)` | YES | Working | — |
| Export JT (simple) | `Document.SaveAs(...)` | YES | Working | — |
| Export flat DXF (sheet metal) | `SheetMetalDocument.SaveAsFlatDXF(...)` | YES | Working | Samples |
| Import file (auto-detect type) | `Documents.Open(filename)` | YES | Working | — |
| Export JT (advanced, 17 params) | `Document.SaveAsJT(name, ...)` (17 params) | NO | **Available** | Samples |
| Export EMF (draft sheet) | `Sheet.SaveAsEnhancedMetafile(filename)` | NO | **Available** | Samples |
| Export BMP/JPG/PNG/TIF (from EMF) | `Metafile.Save(filename, imageFormat)` | NO | Available | Samples |
| Create drawing | Draft document + views | YES | Working | — |
| Capture screenshot | `View.SaveAsImage(path, w, h)` | YES | Working | — |

## 19. View & Display

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Set view orientation | `View.ApplyNamedView(name)` | YES | Working | — |
| Zoom fit | `View.Fit()` | YES | Working | — |
| Zoom to selection | `View.Fit()` | YES | Working | — |
| Set render mode | `View.SetRenderMode(mode)` | YES | Working | — |
| Set background color | `View.SetBackgroundColor(color)` | YES | Working | — |

## 20. Selection & UI

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Get active select set | `Application.ActiveSelectSet` | YES | Working | Samples |
| Clear select set | `SelectSet.RemoveAll()` | YES | Working | Samples |
| Add to select set | `SelectSet.Add(object)` | NO | Available | Samples |
| Report select set | SelectSet iteration | NO | Available | Samples |
| Start command | `Application.StartCommand(constant)` | NO | **Available** | Samples |
| Report add-ins | `AddIns` collection iteration | NO | Available | Samples |
| Report environments | `Environments` iteration | NO | Available | Samples |

## 21. Properties & Metadata

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Get file properties | `Document.Properties` | YES | Working | Samples |
| Get custom properties | `PropertySets["Custom"]` | YES | Working | Samples |
| Add custom property | `Properties.Add(name, value)` | YES | Working | Samples |
| Update custom property | `Property.Value = newValue` | YES | Working | SDK |
| Delete custom property | `Property.Delete()` | YES | Working | Samples |
| Set document property | `Property.Value = newValue` | YES | Working | — |
| Get summary info | `Document.SummaryInfo` | NO | Available | SDK |
| Get created version | `Document.CreatedVersion` | NO | Available | SDK |
| Get last saved version | `Document.LastSavedVersion` | NO | Available | SDK |

## 22. Variables & Dimensions

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Get variables collection | `Document.Variables` | YES | Working | Samples |
| Get variable value | `Variable.DisplayName`, `Variable.Value` | YES | Working | Samples |
| Set variable value | `Variable.Value = newValue` | YES | Working | Samples |
| Report all variables | Full variables iteration | YES | Working | Samples |
| Query variables | `Variables.Query(criteria, nameBy, varType)` | NO | **Available** | Samples |
| Units of measure | Unit conversion API | NO | Available | Samples |
| Global parameters | `Application.GetGlobalParameter(ref)` | NO | Available | SDK |

## 23. Events

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Application events | `ISEApplicationEvents_Event` | NO | N/A (MCP) | AddIn |
| Window events | `ISEApplicationWindowEvents_Event` | NO | N/A (MCP) | AddIn |
| Mouse events | `ISEMouseEvents` | NO | N/A (MCP) | Samples |
| File UI events | `ISEFileUIEvents_Event` | NO | N/A (MCP) | AddIn |
| Shortcut menu events | `ISEShortCutMenuEvents_Event` | NO | N/A (MCP) | AddIn |
| Feature library events | `ISEFeatureLibraryEvents_Event` | NO | N/A (MCP) | AddIn |

**Note:** Events are generally N/A for an MCP server (request/response model), but could be used for monitoring/polling if needed.

## 24. File Reading (Without Solid Edge)

From `SolidEdge.Community.Reader` - reads OLE compound storage files directly.

| API Operation | Reader Method | We Have? | Status | Source |
|---|---|---|---|---|
| Open any SE file | `SolidEdgeDocument.Open(path)` | NO | **Available** | Reader |
| Detect document type | `SolidEdgeDocument.DocumentType` | NO | **Available** | Reader |
| Get created version | `SolidEdgeDocument.CreatedVersion` | NO | **Available** | Reader |
| Get last saved version | `SolidEdgeDocument.LastSavedVersion` | NO | **Available** | Reader |
| Get document status | `SolidEdgeDocument.Status` | NO | **Available** | Reader |
| Extract thumbnail | `SolidEdgeDocument.GetThumbnail()` | NO | **Available** | Reader |
| Get material | `MechanicalModelingPropertySet.Material` | NO | **Available** | Reader |
| Get document number | `ProjectInformationPropertySet.DocumentNumber` | NO | **Available** | Reader |
| Get revision | `ProjectInformationPropertySet.Revision` | NO | **Available** | Reader |
| Get project name | `ProjectInformationPropertySet.ProjectName` | NO | **Available** | Reader |
| Get author/user | `ExtendedSummaryInformationPropertySet.Username` | NO | **Available** | Reader |
| Get custom properties | `CompoundFile.CustomProperties` | NO | **Available** | Reader |
| Get all property sets | `CompoundFile.PropertySets` | NO | **Available** | Reader |
| Get draft sheet info | `DraftDocument.Sheets` + `SheetPaperSizeInfo` | NO | **Available** | Reader |
| Export sheet as EMF | `DraftDocument.GetEmfMemoryStream(sheet)` | NO | **Available** | Reader |

**Note:** The Reader library is C# / .NET. To use from Python, we would need to either:
1. Port the OLE compound storage parsing to Python (using `olefile` package)
2. Create a C# subprocess/service
3. Use Python's `win32com` to read OLE properties directly

## 25. Performance & Batch

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Delay compute | `Application.DelayCompute = true/false` | YES | Working | Samples |
| Disable alerts | `Application.DisplayAlerts = false` | YES | Working | Samples |
| Non-interactive mode | `Application.Interactive = false` | YES | Working | Samples |
| Disable screen updates | `Application.ScreenUpdating = false` | YES | Working | Samples |
| Do idle | `Application.DoIdle()` | YES | Working | Samples |

## 26. Feature Management

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| List edgebar features | `DesignEdgebarFeatures` iteration | YES | Working | Samples |
| Feature name/type | `Feature.Name`, `Feature.Type` | YES | Working | — |
| Rename feature | `Feature.Name = newName` | YES | Working | — |
| Delete feature | `Feature.Delete()` | YES | Working | — |
| Suppress feature | `Feature.Suppress()` | YES | Working | — |
| Unsuppress feature | `Feature.Unsuppress()` | YES | Working | — |
| Suppress feature (in family) | `FamilyMember.AddSuppressedFeature(feat)` | NO | Available | Samples |
| Family members | `FamilyMembers.Add(name)` / `.Apply()` | NO | Available | Samples |

## 27. Body & Topology Query

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Get faces | `Body.Faces[igQueryAll]` | YES | Working | Samples |
| Get edges | `Body.Edges[igQueryAll]` | YES (via faces) | Working | Samples |
| Get face edges | `Face.Edges` | YES | Working | Samples |
| Get face vertices | `Face.Vertices` | YES | Working | Samples |
| Get face info (area, normal) | `Face.Area`, topology | YES | Working | — |
| Get edge info (length, type) | `Edge.Length`, `Edge.Type` | YES | Working | — |
| Get solid bodies | `Models` + `Body.IsSolid` | YES | Working | Samples |
| Get facet data (tessellation) | `Body.GetFacetData(tol, ...)` | YES | Working | Samples |
| Feature faces | `Feature.Faces[igQueryAll]` | NO | **Available** | Samples |
| Feature edges | `Feature.Edges[igQueryAll]` | NO | **Available** | Samples |
| Body is solid | `Body.IsSolid` | YES (internal) | Working | Samples |
| Shell is closed | `Shell.IsClosed` | NO | Available | Samples |
| Shell is void | `Shell.IsVoid` | NO | Available | Samples |

## 28. Structural Frames & Tubes

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Create structural frame | `StructuralFrames.Add(partFile, numPaths, pathArray)` | NO | Available | Samples |
| Create line segments | `LineSegments.Add(startPoint, endPoint)` | NO | Available | Samples |
| Detect tubes | `Occurrence.IsTube()` | NO | Available | Samples |
| Get tube | `Occurrence.GetTube()` | NO | Available | Samples |
| Tube bend table | `Tube.BendTable(...)` | NO | Available | Samples |

---

## 29. Priority Recommendations

### Tier 1-2: DONE - All high-priority operations implemented

All Tier 1 and Tier 2 items are now fully implemented with MCP tool wrappers:

- Rounds, chamfers, holes, cutouts (normal, lofted, drawn, revolved)
- Variables read/write, custom properties CRUD
- Performance flags, recompute, modeling mode
- Face/edge topology query, facet data, solid bodies
- Assembly BOM (flat + structured), document tree, interference check, relations3d
- Draft annotations (dimension, balloon, leader, text box, note)
- Sheet management (add, activate, rename, delete)
- Sheet metal features (dimple, etch, louver, gusset, bead)
- Feature management (suppress, unsuppress, rename, delete)
- Select set, flat DXF export, view background

### Tier 3: MEDIUM PRIORITY - Remaining useful additions

| Operation | Effort | Impact |
|---|---|---|
| **Emboss** (`EmbossFeatures.Add`) | Medium | Low |
| **Shell** (thin wall, complex face selection) | High | Medium |
| **Export JT advanced** (17 params) | Low | Medium |
| **Export EMF** (draft sheet) | Low | Medium |
| **Assembly configurations** | Medium | Medium |
| **Blend** (`Blends.Add`) | Medium | Low |

### Tier 4: LOW PRIORITY / COMPLEX

| Operation | Effort | Impact |
|---|---|---|
| Structural frames | High | Low |
| Feature patterns (SAFEARRAY issue) | High | Medium |
| AddIn framework | N/A | N/A (different use case) |
| File reader (without SE) | High | Medium |
| Event system | N/A | N/A (MCP model) |
| Batch print | Medium | Low |
| Parts list to Excel | High | Low |

---

## Community Repository Inventory

| Repository | Files | Language | Key Content |
|---|---|---|---|
| **Samples** | 1737 | VB.NET / C# | 85+ automation samples across Part, Assembly, Draft, SheetMetal, General |
| **SolidEdge.Community** | 72 | C# | Extension methods for COM API (95+ methods, typed accessors) |
| **SDK** | 108 | C# | AddIn framework, event management, utilities |
| **Interop.SolidEdge** | 49 | C# | Prebuilt COM interop assemblies (10 type libraries) |
| **SolidEdge.Community.AddIn** | 115 | C# | AddIn lifecycle, ribbon, edge bar, view overlay framework |
| **SolidEdge.Community.Reader** | 87 | C# | Read SE files without Solid Edge (OLE compound storage parser) |
| **Templates** | 91 | VB.NET / C# | Visual Studio AddIn project templates |
| **SEU14** | 494 | VB.NET / C# | Developer conference hands-on samples |
| **CodeProject** | 111 | C++ / C# | Article source code for SE AddIn development |
| **SolidEdge.Community.InstallInfo** | 53 | C# | Installation data reader |

### Type Libraries Referenced

| Library | GUID | Namespace |
|---|---|---|
| Solid Edge Part | `{8A7EFA42-F000-11D1-BDFC-080036B4D502}` | `SolidEdgePart` |
| Solid Edge Assembly | varies | `SolidEdgeAssembly` |
| Solid Edge Draft | `{3E2B3BDC-F0B9-11D1-BDFD-080036B4D502}` | `SolidEdgeDraft` |
| Solid Edge Constants | `{C467A6F5-27ED-11D2-BE30-080036B4D502}` | `SolidEdgeConstants` |
| Solid Edge Framework | varies | `SolidEdgeFramework` |
| Solid Edge Framework Support | varies | `SolidEdgeFrameworkSupport` |
| Solid Edge Geometry | varies | `SolidEdgeGeometry` |
| Solid Edge File Properties | varies | `SolidEdgeFileProperties` |
| Revision Manager | varies | `RevisionManager` |
| Install Data | varies | `SEInstallDataLib` |

---

## Coverage Heatmap

```
Category                    Implemented    Available    Coverage
─────────────────────────────────────────────────────────────
Connection                  8              8            100%
Documents                   19             19           100%
Sketching 2D                25             25           100%
Constraints                 8              8            100%
Extrusions                  5              5            100%
Revolves                    5              5            100%
Cutouts                     6              6            100%
Loft & Sweep                6              6            100%
Rounds/Chamfers             6              7            86%
Holes & Patterns            2              5            40%
Advanced Features           15             18           83%
Reference Planes            3              4            75%
Primitives                  5              5            100%
Sheet Metal                 12             14           86%
Assembly                    21             28           75%
Draft/Drawing               14             26           54%
Query & Analysis            24             27           89%
Export & Import              12             14           86%
View & Display              5              5            100%
Selection & UI              2              7            29%
Properties                  6              9            67%
Variables                   4              7            57%
Performance                 5              5            100%
Feature Management          6              8            75%
Topology Query              8              13           62%
─────────────────────────────────────────────────────────────
TOTAL                       218            ~250         87%
```

---

## Next Steps

1. **Improve constraint system** - Requires topology queries to reference edges/faces by index
2. **Add remaining Tier 3 features** - Draft angle, face rotate, emboss, shell, configurations
3. **Fix untested tools** - Validate thin-wall extrude, helix, sheet metal base operations
4. **Export EMF** - Useful for draft sheet image export
5. **Assembly configurations** - Important for design variants
