# Solid Edge API Coverage Analysis

**Generated:** 2026-02-12
**Source:** [SolidEdgeCommunity GitHub](https://github.com/SolidEdgeCommunity) repositories
**Purpose:** Identify all available Solid Edge COM API operations vs. our MCP server implementation

---

## Summary

| Metric | Count |
|--------|-------|
| **Our implemented MCP tools** | 118 |
| **API operations found in community repos** | 250+ |
| **Operations we're missing** | ~58 |
| **High-priority gaps (Tier 2)** | ~5 |
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
| Quit application | `Application.Quit()` | NO | Available | Interop |
| Get installation path | SEInstallData.dll | NO | Available | SDK |
| Get installed language | `SEInstallData.GetInstalledLanguage()` | NO | Available | SDK |
| Get process handle | `Application.ProcessID` → `Process` | NO | Available | SDK |
| Get window handle | `Application.hWnd` | NO | Available | SDK |

## 2. Document Management

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Create part document | `Documents.Add("SolidEdge.PartDocument")` | YES | Working | Samples |
| Create assembly document | `Documents.Add("SolidEdge.AssemblyDocument")` | YES | Working | Samples |
| Create draft document | `Documents.Add("SolidEdge.DraftDocument")` | YES (via create_drawing) | Working | Samples |
| Create sheet metal document | `Documents.Add("SolidEdge.SheetMetalDocument")` | YES | Working | Samples |
| Create weldment document | `Documents.Add("SolidEdge.WeldmentDocument")` | NO | Available | SDK |
| Open document | `Documents.Open(filename)` | YES | Working | Samples |
| Open in background (no window) | `Documents.Open(filename, 0x8)` | NO | Available | SDK |
| Save document | `Document.Save()` | YES | Working | Samples |
| Save as | `Document.SaveAs(filename)` | YES | Working | Samples |
| Close document | `Document.Close()` | YES | Working | Samples |
| Close all documents | `Documents.CloseDocument(...)` per doc | NO | Available | Samples |
| List documents | Documents collection iteration | YES | Working | — |
| Get active document | `Application.ActiveDocument` | YES (internal) | Working | SDK |
| Recompute document | `Document.Recompute()` | YES | Working | Samples |
| Recompute model | `Model.Recompute()` | YES | Working | Samples |
| Toggle modeling mode | `PartDocument.ModelingMode` | NO | **Available** | Samples |
| Get/set modeling mode (Ordered/Sync) | `seModelingModeOrdered/Synchronous` | NO | **Available** | Samples |

## 3. Sketching & 2D Geometry

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Create sketch on named plane | `ProfileSets.Add() + Profiles.Add(refPlane)` | YES | Working | Samples |
| Create sketch on plane by index | Same, with `RefPlanes.Item(index)` | YES | Working | — |
| Draw line | `Lines2d.AddBy2Points(x1,y1,x2,y2)` | YES | Working | Samples |
| Draw circle | `Circles2d.AddByCenterRadius(x,y,r)` | YES | Working | Samples |
| Draw arc (center/start/end) | `Arcs2d.AddByCenterStartEnd(...)` | YES | Working | Samples |
| Draw arc (start/center/end) | `Arcs2d.AddByStartCenterEnd(...)` | NO | Available | Interop |
| Draw rectangle | Lines2d (4 lines) | YES | Working | — |
| Draw polygon | Lines2d (n lines) | YES | Working | Samples |
| Draw ellipse | `Ellipses2d.AddByCenterMajorMinor(...)` | YES | Working | Interop |
| Draw B-spline | `BSplineCurves2d.AddByPoints(order, size, array)` | YES | Working | Samples |
| Mirror B-spline | `BSplineCurve2d.Mirror(x1,y1,x2,y2,copy)` | NO | Available | Samples |
| Draw circle by 2 points | `Circles2d.AddBy2Points(...)` | NO | Available | Interop |
| Draw circle by 3 points | `Circles2d.AddBy3Points(...)` | NO | Available | Interop |
| Set axis of revolution | `Profile.SetAxisOfRevolution(axis)` | YES | Working | Samples |
| Toggle construction geometry | `Profile.ToggleConstruction(element)` | YES (internal) | Working | Samples |
| Close/validate sketch | `Profile.End(flags)` | YES | Working | Samples |
| Hide profile | `Profile.Visible = false` | NO | Available | Samples |

## 4. 2D Constraints (Relations2d)

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Keypoint constraint | `Relations2d.AddKeypoint(obj1, kp1, obj2, kp2)` | Stub | Needs work | Samples |
| Equal constraint | `Relations2d.AddEqual(obj1, obj2)` | Stub | Needs work | Samples |
| Parallel constraint | `Relations2d.AddParallel(obj1, obj2)` | Stub | Needs work | Interop |
| Perpendicular constraint | `Relations2d.AddPerpendicular(obj1, obj2)` | Stub | Needs work | Interop |
| Concentric constraint | `Relations2d.AddConcentric(obj1, obj2)` | Stub | Needs work | Interop |
| Tangent constraint | `Relations2d.AddTangent(obj1, obj2)` | Stub | Needs work | Interop |
| Horizontal constraint | `Relations2d.AddHorizontal(obj)` | Stub | Needs work | Interop |
| Vertical constraint | `Relations2d.AddVertical(obj)` | Stub | Needs work | Interop |

**Note:** Constraints are stubbed in our `add_constraint` tool. The challenge is referencing specific sketch elements. Could be improved by tracking element indices during sketch creation.

## 5. Extrusions

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Finite extrusion | `Models.AddFiniteExtrudedProtrusion(...)` | YES | Working | Samples |
| Through-all extrusion | `ExtrudedProtrusions.AddThroughAll(...)` | YES (untested) | Untested | Interop |
| Infinite extrusion | `Models.AddExtrudedProtrusion(...)` | YES (untested) | Untested | — |
| Thin-wall extrusion | `Models.AddExtrudedProtrusionWithThinWall(...)` | YES (untested) | Broken | — |
| Extruded surface | `Constructions.ExtrudedSurfaces.Add(...)` | NO | **Available** | Samples |

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
| Revolved cutout | `RevolvedCutouts.AddFiniteMulti(...)` | YES | Implemented | Samples |
| Normal cutout | `NormalCutouts.AddFiniteMulti(...)` | YES | **Implemented** | MEMORY |
| Lofted cutout | `LoftedCutouts.AddSimple(...)` | YES | **Implemented** | Samples |
| Drawn cutout | `DrawnCutouts.Add(...)` | NO | Available | MEMORY |

## 8. Loft & Sweep

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Loft (initial feature) | `Models.AddLoftedProtrusion(...)` | YES | Working | Samples |
| Loft (on existing body) | `LoftedProtrusions.AddSimple(...)` | YES | Working | Samples |
| Loft thin-wall | `Models.AddLoftedProtrusionWithThinWall(...)` | YES (untested) | Untested | — |
| Sweep | `Models.AddSweptProtrusion(...)` | YES | Working | Samples |
| Sweep thin-wall | `Models.AddSweptProtrusionWithThinWall(...)` | YES (untested) | Untested | — |
| Ref plane normal to curve | `RefPlanes.AddNormalToCurve(...)` | NO | **Available** | Samples |

## 9. Rounds, Chamfers, Fillets

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Round (fillet) - multiple edges | `Rounds.Add(edgeCount, edgeArray, radiusArray)` | YES | **Working** | Samples |
| Chamfer - equal setback | `Chamfers.AddEqualSetback(count, edges, distance)` | YES | **Working** | Samples |
| Chamfer - unequal setback | `Chamfers.AddUnequalSetback(face, count, edges, d1, d2)` | NO | **Available** | Samples |
| Chamfer - setback + angle | `Chamfers.AddSetbackAngle(face, count, edges, dist, angle)` | NO | **Available** | Samples |
| Blend | `Blends.Add(...)` / `AddSurfaceBlend` / `AddVariable` | NO | Available | MEMORY |

## 10. Holes & Patterns

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Create hole (via circular cutout) | `ExtrudedCutouts.AddFiniteMulti` (circular profile) | YES | **Working** | Samples |
| Place finite hole (HoleData API) | `Holes.AddFinite(profile, side, depth, holeData)` | NO | Available | Samples |
| User-defined pattern | `UserDefinedPatterns.AddByProfiles(...)` | NO | **Available** | Samples |
| Rectangular pattern | `Patterns.AddByRectangular(...)` | NO | Broken (SAFEARRAY) | MEMORY |
| Circular pattern | `Patterns.AddByCircular(...)` | NO | Broken (SAFEARRAY) | MEMORY |

## 11. Advanced Part Features

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Mirror copy | `MirrorCopies.Add(plane, count, features)` | YES | Implemented | MEMORY |
| Face rotate (by geometry) | `FaceRotates.Add(face, ByGeometry, ..., edge, ..., angle)` | NO | **Available** | Samples |
| Face rotate (by points) | `FaceRotates.Add(face, ByPoints, ..., pt1, pt2, ..., angle)` | NO | **Available** | Samples |
| Thicken feature | `Thickens.Add(side, thickness, faceCount, facesArray)` | YES (untested) | Untested | Samples |
| Delete faces | `DeleteFaces.Add(faceSetToDelete)` | NO | Available | MEMORY |
| Draft angle | `Drafts.Add(plane, numSets, faceArray, draftAngleArray, side)` | NO | **Available** | MEMORY |
| Rib | `Ribs.Add(profile, extensionType, thicknessType, side, thickness)` | NO | **Available** | MEMORY |
| Slot | `Slots.Add(profile, slotType, endType, width, ..., 22 params)` | NO | **Available** | Samples |
| Thread | `Threads.Add(holeData, numCylinders, cylinderArray, endArray)` | NO | Available | MEMORY |
| Emboss | `EmbossFeatures.Add(...)` | NO | Available | MEMORY |
| Lip | `Lips.Add(...)` | NO | Available | MEMORY |
| Shell (thin wall) | Shell feature (requires face selection) | NO | Complex | — |
| Suppress feature | `FamilyMembers + AddSuppressedFeature` | NO | **Available** | Samples |
| Move to synchronous | `MoveOrderedFeaturesToSynchronous` | NO | Available | Samples |
| Heal and optimize body | `HealAndOptimizeBody` | NO | Available | Samples |
| Place feature library | Feature library placement | NO | Available | Samples |

## 12. Reference Planes

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Offset parallel plane | `RefPlanes.AddParallelByDistance(parent, dist, side)` | YES | Working | Samples |
| Normal to curve | `RefPlanes.AddNormalToCurve(curve, end, plane, pivot, ...)` | NO | **Available** | Samples |
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
| Create sheet metal doc | `Documents.Add("SolidEdge.SheetMetalDocument")` | NO | **Available** | Samples |
| Base tab | `Tabs.Add(...)` | YES (untested) | Untested | Samples |
| Base flange | `Models.AddBaseContourFlange(...)` | YES (untested) | Untested | — |
| Lofted flange | `Models.AddLoftedFlange(...)` | YES (untested) | Untested | Samples |
| Web network | `Models.AddWebNetwork(...)` | YES (untested) | Untested | — |
| Dimple | `Dimples.Add(...)` | NO | **Available** | Samples |
| Etch | `Etches.Add(...)` | NO | **Available** | Samples |
| Louver | `Louvers.Add(...)` | NO | Available | MEMORY |
| Gusset | `Gussets.Add(...)` | NO | Available | MEMORY |
| Bead | `Beads.Add(...)` | NO | Available | MEMORY |
| SM extruded cutout | `ExtrudedCutouts.Add(...)` (23 params) | YES (via general cutout) | Working | Samples |
| SM holes with pattern | `Holes.AddFinite + UserDefinedPatterns.AddByProfiles` | NO | **Available** | Samples |
| Save flat DXF | `SheetMetalDocument.SaveAsFlatDXF(...)` | NO | **Available** | Samples |

## 15. Assembly

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Place by filename | `Occurrences.AddByFilename(path)` | YES | Working | Samples |
| Place with matrix | `Occurrences.AddWithMatrix(path, matrix)` | YES | Working | Samples |
| Place with transform | `Occurrences.AddWithTransform(path, ox,oy,oz,ax,ay,az)` | NO | Available | Samples |
| List occurrences | Occurrences iteration | YES | Working | Samples |
| Get transform | `Occurrence.GetTransform(...)` | YES | Working | Samples |
| Get 4x4 matrix | `Occurrence.GetMatrix(...)` | YES | Working | Samples |
| Get range box | `Occurrence.GetRangeBox(min, max)` | NO | **Available** | Samples |
| Check interference | `AssemblyDocument.CheckInterference(...)` | NO | **Available** | Samples |
| Report BOM | BOM traversal via Occurrences | NO | **Available** | Samples |
| Report document tree | Recursive `SubOccurrences` traversal | NO | **Available** | Samples |
| Configurations | `Configurations.Add/Apply/Item` | NO | **Available** | Samples |
| Derived configurations | `Configurations.AddDerivedConfig(...)` | NO | Available | Samples |
| Relations3d query | `Relations3d` iteration (15+ constraint types) | NO | **Available** | Samples |
| Delete ground relations | `Relations3d.Delete()` | NO | Available | Samples |
| Detect under-constrained | Under-constrained occurrence detection | NO | Available | Samples |
| Structural frames | `StructuralFrames.Add(partFile, numPaths, pathArray)` | NO | Available | Samples |
| Report tubes | `Occurrence.IsTube()` / `GetTube()` | NO | Available | Samples |
| Tube bend table | `Tube.BendTable(...)` | NO | Available | Samples |
| Suppress component | `Occurrence.Suppress/Unsuppress` | YES | Working | — |
| Pattern component | Matrix-based copies | YES | Working | — |

## 16. Draft / Drawing

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Create draft document | `Documents.Add("SolidEdge.DraftDocument")` | YES | Working | Samples |
| Link 3D model | `ModelLinks.Add(filename)` | YES | Working | Samples |
| Add part view | `DrawingViews.AddPartView(link, orient, scale, x, y)` | YES | Working | Samples |
| Add assembly view | `DrawingViews.AddAssemblyView(link, orient, scale, x, y, type)` | NO | **Available** | Samples |
| Add sheet | `Sheets.AddSheet()` / `Sheet.Activate()` | NO | **Available** | Samples |
| Report sheets | `Sheets` iteration (Name, Index, Number) | NO | Available | Samples |
| Report drawing views | `DrawingViews` iteration | NO | Available | Samples |
| Report sections | `Sections` iteration | NO | Available | Samples |
| Update drawing views | `DrawingView.Update()` | NO | Available | Samples |
| Convert views to 2D | ConvertDrawingViewsTo2D | NO | Available | Samples |
| Create balloon | `Balloons.Add(x, y, z)` | NO | **Available** | Samples |
| Create leader | `Leaders.Add(x1,y1,z1,x2,y2,z2)` | NO | Available | Samples |
| Create text box | `TextBoxes.Add(x, y, z)` | NO | Available | Samples |
| Draw 2D lines in draft | `Lines2d.AddBy2Points(...)` (in draft sheet context) | NO | Available | Samples |
| Draw 2D circles in draft | `Circles2d.AddByCenterRadius(...)` (in draft) | NO | Available | Samples |
| Create dimensions | Linear/angular dimensions | NO | Available | Samples |
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
| Body volume | `Body.Volume` | NO | Available | Samples |
| List features (edgebar) | `DesignEdgebarFeatures` iteration | YES | Working | Samples |
| Feature count | `Models.Count` | YES | Working | — |
| Document properties | Document property access | YES | Working | Samples |
| Measure distance | Math calculation | YES | Working | — |
| Body facet data (mesh) | `Body.GetFacetData(tolerance, ...)` | NO | **Available** | Samples |
| Report solid bodies | `Models` + `Body.IsSolid`, `Shells.IsClosed` | NO | **Available** | Samples |
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
| Export JT (advanced, 17 params) | `Document.SaveAsJT(name, ...)` (17 params) | NO | **Available** | Samples |
| Export flat DXF (sheet metal) | `SheetMetalDocument.SaveAsFlatDXF(...)` | NO | **Available** | Samples |
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

## 20. Selection & UI

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Get active select set | `Application.ActiveSelectSet` | NO | **Available** | Samples |
| Add to select set | `SelectSet.Add(object)` | NO | **Available** | Samples |
| Clear select set | `SelectSet.RemoveAll()` | NO | **Available** | Samples |
| Report select set | SelectSet iteration | NO | Available | Samples |
| Start command | `Application.StartCommand(constant)` | NO | **Available** | Samples |
| Report add-ins | `AddIns` collection iteration | NO | Available | Samples |
| Report environments | `Environments` iteration | NO | Available | Samples |

## 21. Properties & Metadata

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Get file properties | `Document.Properties` | YES (basic) | Working | Samples |
| Get custom properties | `PropertySets["Custom"]` | YES | Working | Samples |
| Add custom property | `Properties.Add(name, value)` | YES | Working | Samples |
| Update custom property | `Property.Value = newValue` | YES | Working | SDK |
| Delete custom property | `Property.Delete()` | YES | Working | Samples |
| Get summary info | `Document.SummaryInfo` | NO | Available | SDK |
| Get created version | `Document.CreatedVersion` | NO | Available | SDK |
| Get last saved version | `Document.LastSavedVersion` | NO | Available | SDK |

## 22. Variables & Dimensions

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Get variables collection | `Document.Variables` | YES | Working | Samples |
| Query variables | `Variables.Query(criteria, nameBy, varType)` | NO | **Available** | Samples |
| Get variable value | `Variable.DisplayName`, `Variable.Value` | YES | Working | Samples |
| Set variable value | `Variable.Value = newValue` | YES | Working | Samples |
| Report all variables | Full variables iteration | YES | Working | Samples |
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
| Do idle | `Application.DoIdle()` | NO | Available | Samples |

**Note:** These performance flags are critical for batch operations. HIGH PRIORITY.

## 26. Feature Management

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| List edgebar features | `DesignEdgebarFeatures` iteration | YES | Working | Samples |
| Feature name/type | `Feature.Name`, `Feature.Type` | YES | Working | — |
| Suppress feature (in family) | `FamilyMember.AddSuppressedFeature(feat)` | NO | Available | Samples |
| Family members | `FamilyMembers.Add(name)` / `.Apply()` | NO | Available | Samples |
| Delete thicken feature | `Thicken.Delete()` | NO | Available | Samples |

## 27. Body & Topology Query

| API Operation | COM Method | We Have? | Status | Source |
|---|---|---|---|---|
| Get faces | `Body.Faces[igQueryAll]` | YES | Working | Samples |
| Get edges | `Body.Edges[igQueryAll]` | YES (via faces) | Working | Samples |
| Get face edges | `Face.Edges` | YES | Working | Samples |
| Get face vertices | `Face.Vertices` | YES | Working | Samples |
| Feature faces | `Feature.Faces[igQueryAll]` | NO | **Available** | Samples |
| Feature edges | `Feature.Edges[igQueryAll]` | NO | **Available** | Samples |
| Body is solid | `Body.IsSolid` | NO | Available | Samples |
| Shell is closed | `Shell.IsClosed` | NO | Available | Samples |
| Shell is void | `Shell.IsVoid` | NO | Available | Samples |
| Get facet data (tessellation) | `Body.GetFacetData(tol, ...)` | NO | **Available** | Samples |

**Note:** Topology queries are a prerequisite for making rounds, chamfers, holes, and constraints work properly. CRITICAL for feature selection.

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

### Tier 1: ~~HIGH PRIORITY~~ DONE - All implemented

These operations are now fully implemented with MCP tool wrappers:

| Operation | Effort | Impact | Status |
|---|---|---|---|
| **Rounds/fillets** (`Rounds.Add`) | Low | Very High | **DONE** (create_round) |
| **Chamfers** (`Chamfers.AddEqualSetback`) | Low | Very High | **DONE** (create_chamfer) |
| **Holes** (via circular ExtrudedCutout) | Low | Very High | **DONE** (create_hole) |
| **Normal cutout** (`NormalCutouts.AddFiniteMulti`) | Low | Medium | **DONE** (create_normal_cutout) |
| **Lofted cutout** (`LoftedCutouts.AddSimple`) | Medium | Medium | **DONE** (create_lofted_cutout) |

### Tier 2: ~~HIGH PRIORITY~~ DONE - Nearly all implemented

| Operation | Effort | Impact | Status |
|---|---|---|---|
| **Variables read/write** | Low | Very High | **DONE** (get_variables, get_variable, set_variable) |
| **Custom properties** (CRUD) | Low | Very High | **DONE** (get/set/delete_custom_property) |
| **Performance flags** (DelayCompute, ScreenUpdating) | Low | High | **DONE** (set_performance_mode) |
| **Recompute document/model** | Low | High | **DONE** (recompute) |
| **Create sheet metal document** | Low | High | **DONE** (create_sheet_metal_document) |
| **Face/edge topology query** | Medium | Very High | **DONE** (get_body_faces, get_body_edges, get_face_info) |
| **Body facet data** (mesh export) | Medium | High | |
| **Interference check** (assembly) | Medium | High | **DONE** (check_interference) |
| **Report BOM** (assembly) | Medium | High | **DONE** (get_bom) |
| **Draft: Add assembly view** | Low | High | **DONE** (add_assembly_drawing_view) |
| **Draft: Add sheet** | Low | Medium | **DONE** (add_draft_sheet) |
| **Occurrence bounding box** | Low | Medium | **DONE** (get_occurrence_bounding_box) |

### Tier 3: MEDIUM PRIORITY - Useful additions

| Operation | Effort | Impact |
|---|---|---|
| **Draft angle** (`Drafts.Add`) | Medium | Medium |
| **Rib** (`Ribs.Add`) | Medium | Medium |
| **Slot** (`Slots.Add`, 22 params) | High | Medium |
| **Face rotate** | Medium | Low |
| **Thicken** (finish untested) | Low | Medium |
| **Export flat DXF** (sheet metal) | Low | High |
| **Export JT advanced** (17 params) | Low | Medium |
| **Export EMF** (draft sheet) | Low | Medium |
| **Assembly configurations** | Medium | Medium |
| **Document tree traversal** | Medium | Medium |
| **Relations3d query** | Medium | Medium |
| **Ref plane normal to curve** | Low | Medium |
| **Dimple, Etch** (sheet metal) | Medium | Medium |
| **Extruded surfaces** (construction) | Medium | Low |

### Tier 4: LOW PRIORITY / COMPLEX

| Operation | Effort | Impact |
|---|---|---|
| Structural frames | High | Low |
| Feature patterns (SAFEARRAY issue) | High | Medium |
| AddIn framework | N/A | N/A (different use case) |
| File reader (without SE) | High | Medium |
| Event system | N/A | N/A (MCP model) |
| Shell feature | High (face selection) | Medium |
| Thread feature | Medium | Low |
| Emboss/Lip features | Medium | Low |
| Balloon/Leader annotations | Medium | Low |
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
Connection                  2              8            25%
Documents                   7              17           41%
Sketching 2D                11             17           65%
Constraints                 1 (stub)       8            13%
Extrusions                  3              5            60%
Revolves                    5              5            100%
Cutouts                     5              6            83%
Loft & Sweep                4              6            67%
Rounds/Chamfers             2              5            40%
Holes & Patterns            1              5            20%
Advanced Features           2              16           13% ← GAP
Reference Planes            1              3            33%
Primitives                  5              5            100%
Sheet Metal                 8              13           62%
Assembly                    11             20           55%
Draft/Drawing               3              25           12% ← GAP
Query & Analysis            6              13           46%
Export                       9              14           64%
View & Display              4              4            100%
Selection & UI              0              6            0%  ← GAP
Properties                  1              8            13% ← GAP
Variables                   0              7            0%  ← GAP
Performance                 0              5            0%  ← GAP
Topology Query              0              10           0%  ← GAP
─────────────────────────────────────────────────────────────
TOTAL                       101            ~250         40%
```

---

## Next Steps

1. **Immediately add Tier 1 tools** - Rounds, chamfers, and holes are already solved; just need MCP wrappers
2. **Add Tier 2 tools** - Variables, properties, performance flags, topology queries
3. **Fix untested tools** - Validate thin-wall extrude, helix, sheet metal operations
4. **Improve constraint system** - Requires topology queries to reference edges/faces
5. **Enhance assembly tools** - BOM, interference, configurations
6. **Expand draft capabilities** - More view types, annotations, dimensions
