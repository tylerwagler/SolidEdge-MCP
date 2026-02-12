# Solid Edge MCP - Implementation Status

Last Updated: 2026-02-12

## MCP Server Status: **OPERATIONAL**

**118 MCP tools** are now registered and ready to use!

## Quick Summary

| Category | Implemented | Notes |
|----------|-------------|-------|
| **Connection** | 2 | Connect, app info |
| **Document Management** | 8 | Create (part, assembly, sheet metal), open, save, close, list |
| **Sketching** | 11 | Lines, circles, arcs, rects, polygons, ellipses, splines, constraints, axis |
| **Basic Primitives** | 5 | Box (3 variants), cylinder, sphere |
| **Extrusions** | 3 | Finite, infinite, thin-wall |
| **Revolves** | 5 | Basic, finite, sync, thin-wall |
| **Cutouts** | 5 | Extruded finite, through-all, revolved, normal, lofted |
| **Rounds/Chamfers/Holes** | 3 | Round (fillet), chamfer, hole |
| **Reference Planes** | 1 | Offset parallel plane |
| **Loft** | 2 | Basic, thin-wall |
| **Sweep** | 2 | Basic, thin-wall |
| **Helix/Spiral** | 4 | Basic, sync, thin-wall variants |
| **Sheet Metal** | 8 | Base flange/tab, lofted flange, web network |
| **Body Operations** | 7 | Add body, thicken, mesh, tag, construction |
| **Simplification** | 4 | Auto-simplify, enclosure, duplicate |
| **View/Display** | 4 | Orientation, zoom, display mode |
| **Variables** | 3 | Get all, get by name, set value |
| **Custom Properties** | 3 | Get all, set/create, delete |
| **Body Topology** | 3 | Body faces, body edges, face info |
| **Performance** | 2 | Set performance mode, recompute |
| **Query/Analysis** | 6 | Mass properties, bounding box, features, measurements |
| **Export** | 9 | STEP, STL, IGES, PDF, DXF, Parasolid, JT, drawing, screenshot |
| **Assembly** | 14 | Place, list, constraints, patterns, suppress, BOM, interference, bbox |
| **Draft/Drawing** | 2 | Add sheet, assembly drawing view |
| **Diagnostics** | 2 | API and feature inspection |
| **TOTAL** | **118** | |

---

## Tool Categories

### 1. Connection & Application (2)
| Tool | API Method | Status |
|------|-----------|--------|
| connect_to_solidedge | GetActiveObject/Dispatch | Working |
| get_application_info | Application properties | Working |

### 2. Document Management (8)
| Tool | API Method | Status |
|------|-----------|--------|
| create_part_document | Documents.Add | Working |
| create_assembly_document | Documents.Add | Working |
| create_sheet_metal_document | Documents.Add("SolidEdge.SheetMetalDocument") | Implemented |
| open_document | Documents.Open | Working |
| save_document | Document.Save/SaveAs | Working |
| close_document | Document.Close | Working |
| list_documents | Documents collection | Working |

### 3. Sketching & 2D Geometry (11)
| Tool | API Method | Status |
|------|-----------|--------|
| create_sketch | ProfileSets.Add/Profiles.Add | Working |
| create_sketch_on_plane | ProfileSets.Add with plane index | Working |
| draw_line | Lines2d.AddBy2Points | Working |
| draw_circle | Circles2d.AddByCenterRadius | Working |
| draw_rectangle | Lines2d (4 lines) | Working |
| draw_arc | Arcs2d.AddByCenterStartEnd | Working |
| draw_polygon | Lines2d (n lines) | Working |
| draw_ellipse | Ellipses2d.AddByCenter | Working |
| draw_spline | BSplineCurves2d.AddByPoints | Working |
| set_axis_of_revolution | SetAxisOfRevolution | Working |
| close_sketch | Profile.End | Working |
| add_constraint | Relations2d | Stub (needs element refs) |

### 4. Primitives (5)
| Tool | API Method | Status |
|------|-----------|--------|
| create_box_by_center | Models.AddBoxByCenter | Working |
| create_box_by_two_points | Models.AddBoxByTwoPoints | Working |
| create_box_by_three_points | Models.AddBoxByThreePoints | Working |
| create_cylinder | Models.AddCylinderByCenterAndRadius | Working |
| create_sphere | Models.AddSphereByCenterAndRadius | Working |

### 5. Extrusions (3)
| Tool | API Method | Status |
|------|-----------|--------|
| create_extrude | Models.AddFiniteExtrudedProtrusion | Working |
| create_extrude_infinite | Models.AddExtrudedProtrusion | Untested |
| create_extrude_thin_wall | Models.AddExtrudedProtrusionWithThinWall | Untested |

### 6. Revolves (5)
| Tool | API Method | Status |
|------|-----------|--------|
| create_revolve | Models.AddFiniteRevolvedProtrusion | Working |
| create_revolve_finite | Models.AddFiniteRevolvedProtrusion | Working |
| create_revolve_sync | Models.AddRevolvedProtrusionSync | Untested |
| create_revolve_finite_sync | Models.AddFiniteRevolvedProtrusionSync | Untested |
| create_revolve_thin_wall | Models.AddRevolvedProtrusionWithThinWall | Untested |

### 7. Cutouts (5)
| Tool | API Method | Status |
|------|-----------|--------|
| create_extruded_cutout | ExtrudedCutouts.AddFiniteMulti | **Working** |
| create_extruded_cutout_through_all | ExtrudedCutouts.AddThroughAllMulti | **Working** |
| create_revolved_cutout | RevolvedCutouts.AddFiniteMulti | Implemented |
| create_normal_cutout | NormalCutouts.AddFiniteMulti | Implemented |
| create_lofted_cutout | LoftedCutouts.AddSimple | Implemented |

### 7b. Rounds, Chamfers & Holes (3)
| Tool | API Method | Status |
|------|-----------|--------|
| create_round | Rounds.Add | **Working** |
| create_chamfer | Chamfers.AddEqualSetback | **Working** |
| create_hole | ExtrudedCutouts.AddFiniteMulti (circular) | **Working** |

### 8. Reference Planes - NEW! (1)
| Tool | API Method | Status |
|------|-----------|--------|
| create_ref_plane_by_offset | RefPlanes.AddParallelByDistance | **Working** |

### 9. Loft (2)
| Tool | API Method | Status |
|------|-----------|--------|
| create_loft | LoftedProtrusions.AddSimple / Models.AddLoftedProtrusion | Working |
| create_loft_thin_wall | Models.AddLoftedProtrusionWithThinWall | Untested |

### 10. Sweep (2)
| Tool | API Method | Status |
|------|-----------|--------|
| create_sweep | Models.AddSweptProtrusion | Working |
| create_sweep_thin_wall | Models.AddSweptProtrusionWithThinWall | Untested |

### 11. Helix/Spiral (4)
| Tool | API Method | Status |
|------|-----------|--------|
| create_helix | Models.AddFiniteBaseHelix | Untested |
| create_helix_sync | Models.AddFiniteBaseHelixSync | Untested |
| create_helix_thin_wall | Models.AddFiniteBaseHelixWithThinWall | Untested |
| create_helix_sync_thin_wall | Models.AddFiniteBaseHelixSyncWithThinWall | Untested |

### 12. Sheet Metal (8)
| Tool | API Method | Status |
|------|-----------|--------|
| create_base_flange | Models.AddBaseContourFlange | Untested |
| create_base_tab | Models.AddBaseTab | Untested |
| create_lofted_flange | Models.AddLoftedFlange | Untested |
| create_web_network | Models.AddWebNetwork | Untested |
| create_base_contour_flange_advanced | Models.AddBaseContourFlangeBy... | Untested |
| create_base_tab_multi_profile | Models.AddBaseTabWithMultipleProfiles | Untested |
| create_lofted_flange_advanced | Models.AddLoftedFlangeBy... | Untested |
| create_lofted_flange_ex | Models.AddLoftedFlangeEx | Untested |

### 13. Body Operations (7)
| Tool | API Method | Status |
|------|-----------|--------|
| add_body | Models.AddBody | Untested |
| thicken_surface | Models.AddThickenFeature | Untested |
| add_body_by_mesh | Models.AddBodyByMeshFacets | Untested |
| add_body_feature | Models.AddBodyFeature | Untested |
| add_by_construction | Models.AddByConstruction | Untested |
| add_body_by_tag | Models.AddBodyByTag | Untested |

### 14. Simplification (4)
| Tool | API Method | Status |
|------|-----------|--------|
| auto_simplify | Models.AddAutoSimplify | Untested |
| simplify_enclosure | Models.AddSimplifyEnclosure | Untested |
| simplify_duplicate | Models.AddSimplifyDuplicate | Untested |
| local_simplify_enclosure | Models.AddLocalSimplifyEnclosure | Untested |

### 15. View & Display (4)
| Tool | API Method | Status |
|------|-----------|--------|
| set_view | View.ApplyNamedView | Working |
| zoom_fit | View.Fit | Working |
| zoom_to_selection | View.Fit | Working |
| set_display_mode | View.SetRenderMode | Working |

### 16. Query & Analysis (6)
| Tool | API Method | Status |
|------|-----------|--------|
| get_mass_properties | Model.ComputePhysicalProperties... | Working |
| get_bounding_box | Body.GetRange | Working |
| list_features | DesignEdgebarFeatures | Working |
| get_feature_count | Models.Count | Working |
| get_document_properties | Document properties | Working |
| measure_distance | Math calculation | Working |

### 17. Export (9)
| Tool | API Method | Status |
|------|-----------|--------|
| export_step | Document.SaveAs | Working |
| export_stl | Document.SaveAs | Working |
| export_iges | Document.SaveAs | Working |
| export_pdf | Document.SaveAs | Working |
| export_dxf | Document.SaveAs | Working |
| export_parasolid | Document.SaveAs | Working |
| export_jt | Document.SaveAs | Working |
| create_drawing | DraftDocument + DrawingViews | Working |
| capture_screenshot | View.SaveAsImage | Working |

### 17b. Draft/Drawing (2)
| Tool | API Method | Status |
|------|-----------|--------|
| add_draft_sheet | Sheets.AddSheet | Implemented |
| add_assembly_drawing_view | DrawingViews.AddAssemblyView | Implemented |

### 18. Assembly (14)
| Tool | API Method | Status |
|------|-----------|--------|
| place_component | Occurrences.AddByFilename/AddWithMatrix | Working |
| list_assembly_components | Occurrences iteration | Working |
| get_component_info | Occurrence properties + GetTransform | Working |
| update_component_position | Occurrence.SetMatrix | Working |
| pattern_component | Occurrences.AddWithMatrix (copies) | Working |
| suppress_component | Occurrence.Suppress/Unsuppress | Working |
| get_occurrence_bounding_box | Occurrence.GetRangeBox | Implemented |
| get_bom | Occurrences iteration + dedup | Implemented |
| check_interference | AssemblyDocument.CheckInterference | Implemented |
| create_mate | Relations3d | Stub (needs face selection) |
| add_align_constraint | Relations3d | Stub (needs face selection) |
| add_angle_constraint | Relations3d | Stub (needs face selection) |
| add_planar_align_constraint | Relations3d | Stub (needs face selection) |
| add_axial_align_constraint | Relations3d | Stub (needs face selection) |

### 19. Diagnostics (2)
| Tool | API Method | Status |
|------|-----------|--------|
| diagnose_api | Runtime introspection | Working |
| diagnose_feature | Feature property inspection | Working |

### 20. Variables (3)
| Tool | API Method | Status |
|------|-----------|--------|
| get_variables | Document.Variables iteration | Implemented |
| get_variable | Variable.DisplayName match | Implemented |
| set_variable | Variable.Value = newValue | Implemented |

### 21. Custom Properties (3)
| Tool | API Method | Status |
|------|-----------|--------|
| get_custom_properties | Document.Properties iteration | Implemented |
| set_custom_property | Property.Value / Properties.Add | Implemented |
| delete_custom_property | Property.Delete | Implemented |

### 22. Body Topology (3)
| Tool | API Method | Status |
|------|-----------|--------|
| get_body_faces | Body.Faces(igQueryAll) | Implemented |
| get_body_edges | Face.Edges iteration | Implemented |
| get_face_info | Face properties (Type, Area, Edges) | Implemented |

### 23. Performance (2)
| Tool | API Method | Status |
|------|-----------|--------|
| set_performance_mode | App.DelayCompute/ScreenUpdating/etc | Implemented |
| recompute | Model.Recompute + Document.Recompute | Implemented |

---

## Known Limitations

1. **Assembly constraints** require face/edge geometry selection which is complex to automate via COM
2. **Feature patterns** (model.Patterns) require SAFEARRAY(IDispatch) marshaling that fails in late binding
3. **Shell/Thinwalls** requires face selection for open faces, not automatable via COM
4. **Sketch constraints** (add_constraint) needs specific element references, partially stubbed
5. **Cutout via models.Add*Cutout** does NOT work - must use collection-level methods (ExtrudedCutouts.AddFiniteMulti)

## Available but Not Yet Implemented

Feature collections accessible on Model object that could be wrapped as tools:
- **MirrorCopies** - Add(PatternPlane, NumFeatures, FeatureArray), AddSync
- **Drafts** - Add(DraftPlane, NumFaceSets, FaceSetArray, DraftAngleArray, DraftSide)
- **Threads** - Add(HoleData, NumCylinders, CylinderArray, CylinderEndArray)
- **Ribs** - Add(RibProfile, ExtensionType, ThicknessType, MaterialSide, ThicknessSide, Thickness)
- **Slots** - Add (22 params, complex)
- **Blends** - Add, AddSurfaceBlend, AddVariable
- **DeleteFaces** - Add(FaceSetToDelete)
- **Dimples, Gussets, Louvers, Beads, Lips** - Various sheet metal features
