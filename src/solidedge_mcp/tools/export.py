"""Export, drawing, and view tools for Solid Edge MCP."""

from solidedge_mcp.managers import export_manager, view_manager

# ================================================================
# Group 47: export_file (8 → 1)
# ================================================================


def export_file(
    format: str = "step",
    file_path: str = "",
    ini_file_path: str = "",
    width: int = 800,
    height: int = 600,
) -> dict:
    """Export the active document to a file.

    format: 'step' | 'stl' | 'iges' | 'pdf' | 'dxf'
            | 'parasolid' | 'jt' | 'flat_dxf'
            | 'prc' | 'plmxml' | 'image'

    Args:
        format: Export file format.
        file_path: Output file path.
        ini_file_path: Path to PLMXML INI config file (plmxml).
        width: Image width in pixels (image).
        height: Image height in pixels (image).
    """
    match format:
        case "step":
            return export_manager.export_step(file_path)
        case "stl":
            return export_manager.export_stl(file_path)
        case "iges":
            return export_manager.export_iges(file_path)
        case "pdf":
            return export_manager.export_pdf(file_path)
        case "dxf":
            return export_manager.export_dxf(file_path)
        case "parasolid":
            return export_manager.export_parasolid(file_path)
        case "jt":
            return export_manager.export_jt(file_path)
        case "flat_dxf":
            return export_manager.export_flat_dxf(file_path)
        case "prc":
            return export_manager.export_to_prc(file_path)
        case "plmxml":
            return export_manager.export_to_plmxml(file_path, ini_file_path)
        case "image":
            return export_manager.capture_screenshot(file_path, width, height)
        case _:
            return {"error": f"Unknown format: {format}"}


# ================================================================
# Group 48: add_drawing_view (8 → 1)
# ================================================================


def add_drawing_view(
    type: str = "assembly",
    x: float = 0.15,
    y: float = 0.15,
    orientation: str = "Isometric",
    scale: float = 1.0,
    config: str | None = None,
    configuration: str = "Default",
    parent_view_index: int = 0,
    fold_direction: str = "Up",
    center_x: float = 0.0,
    center_y: float = 0.0,
    radius: float = 0.01,
    source_view_index: int = 0,
    section_type: int = 0,
) -> dict:
    """Add a drawing view to the active draft.

    type: 'assembly' | 'assembly_ex' | 'with_config'
          | 'projected' | 'detail' | 'auxiliary'
          | 'draft' | 'by_draft_view' | 'section'

    Args:
        type: Drawing view type to add.
        x: X position for the view (meters).
        y: Y position for the view (meters).
        orientation: View orientation (assembly types).
        scale: View scale factor.
        config: Optional configuration (assembly_ex).
        configuration: Configuration name (with_config).
        parent_view_index: Parent view 0-based index
            (projected, detail, auxiliary, section).
        fold_direction: 'Up'|'Down'|'Left'|'Right'
            (projected, auxiliary).
        center_x: Detail view center X (detail).
        center_y: Detail view center Y (detail).
        radius: Detail view radius (detail).
        source_view_index: Source view index (by_draft_view).
        section_type: Section cut type: 0=standard, 1=revolved
            (section).
    """
    match type:
        case "assembly":
            return export_manager.add_assembly_drawing_view(x, y, orientation, scale)
        case "assembly_ex":
            return export_manager.add_assembly_drawing_view_ex(x, y, orientation, scale, config)
        case "with_config":
            return export_manager.add_drawing_view_with_config(
                x, y, orientation, scale, configuration
            )
        case "projected":
            return export_manager.add_projected_view(parent_view_index, fold_direction, x, y)
        case "detail":
            return export_manager.add_detail_view(
                parent_view_index,
                center_x,
                center_y,
                radius,
                x,
                y,
                scale,
            )
        case "auxiliary":
            return export_manager.add_auxiliary_view(parent_view_index, x, y, fold_direction)
        case "draft":
            return export_manager.add_draft_view(x, y)
        case "by_draft_view":
            return export_manager.add_by_draft_view(source_view_index, x, y, scale)
        case "section":
            return export_manager.add_section_cut(parent_view_index, x, y, section_type)
        case _:
            return {"error": f"Unknown type: {type}"}


# ================================================================
# Group 49: manage_drawing_view (12 → 1)
# ================================================================


def manage_drawing_view(
    action: str,
    view_index: int = 0,
    view_index2: int = 0,
    scale: float = 1.0,
    x: float = 0.0,
    y: float = 0.0,
    show: bool = True,
    mode: str = "Wireframe",
    orientation: str = "Front",
    align: bool = True,
    force_update: bool = True,
) -> dict:
    """Manage an existing drawing view.

    action: 'get_model_link' | 'show_tangent_edges'
            | 'set_scale' | 'delete' | 'update' | 'move'
            | 'show_hidden_edges' | 'set_display_mode'
            | 'set_orientation' | 'activate' | 'deactivate'
            | 'get_dimensions' | 'align' | 'update_all'

    Args:
        action: Management action to perform.
        view_index: 0-based drawing view index.
        view_index2: Second view index (align).
        scale: Scale factor (set_scale).
        x: X position in meters (move).
        y: Y position in meters (move).
        show: Visibility flag (show_tangent_edges,
            show_hidden_edges).
        mode: Display mode string (set_display_mode):
            'Wireframe'|'HiddenEdgesVisible'|'Shaded'
            |'ShadedWithEdges'.
        orientation: View orientation (set_orientation):
            'Front'|'Top'|'Right'|'Back'|'Bottom'|'Left'
            |'Isometric'.
        align: Align or unalign views (align).
        force_update: Force update all views (update_all).
    """
    match action:
        case "get_model_link":
            return export_manager.get_drawing_view_model_link(view_index)
        case "show_tangent_edges":
            return export_manager.show_tangent_edges(view_index, show)
        case "set_scale":
            return export_manager.set_drawing_view_scale(view_index, scale)
        case "delete":
            return export_manager.delete_drawing_view(view_index)
        case "update":
            return export_manager.update_drawing_view(view_index)
        case "move":
            return export_manager.move_drawing_view(view_index, x, y)
        case "show_hidden_edges":
            return export_manager.show_hidden_edges(view_index, show)
        case "set_display_mode":
            return export_manager.set_drawing_view_display_mode(view_index, mode)
        case "set_orientation":
            return export_manager.set_drawing_view_orientation(view_index, orientation)
        case "activate":
            return export_manager.activate_drawing_view(view_index)
        case "deactivate":
            return export_manager.deactivate_drawing_view(view_index)
        case "get_dimensions":
            return export_manager.get_drawing_view_dimensions(view_index)
        case "align":
            return export_manager.align_drawing_views(view_index, view_index2, align)
        case "update_all":
            return export_manager.update_all_views(force_update)
        case _:
            return {"error": f"Unknown action: {action}"}


# ================================================================
# Group 50: add_annotation (14 → 1)
# ================================================================


def add_annotation(
    type: str,
    x: float = 0.0,
    y: float = 0.0,
    x1: float = 0.0,
    y1: float = 0.0,
    x2: float = 0.0,
    y2: float = 0.0,
    x3: float = 0.0,
    y3: float = 0.0,
    text: str = "",
    height: float = 0.005,
    leader_x: float | None = None,
    leader_y: float | None = None,
    dim_x: float | None = None,
    dim_y: float | None = None,
    center_x: float = 0.0,
    center_y: float = 0.0,
    point_x: float = 0.0,
    point_y: float = 0.0,
    origin_x: float = 0.0,
    origin_y: float = 0.0,
    symbol_type: str = "machined",
    weld_type: str = "fillet",
    tolerance_text: str = "",
) -> dict:
    """Add an annotation to the active draft.

    type: 'text_box' | 'leader' | 'dimension' | 'balloon'
          | 'note' | 'angular_dimension'
          | 'radial_dimension' | 'diameter_dimension'
          | 'ordinate_dimension' | 'center_mark'
          | 'centerline' | 'surface_finish'
          | 'weld_symbol' | 'geometric_tolerance'

    Args:
        type: Annotation type to add.
        x: X position (text_box, balloon, note, center_mark,
            surface_finish, weld_symbol, geometric_tolerance).
        y: Y position (text_box, balloon, note, center_mark,
            surface_finish, weld_symbol, geometric_tolerance).
        x1: Start X (leader, dimension, centerline).
        y1: Start Y (leader, dimension, centerline).
        x2: End X (leader, dimension, centerline).
        y2: End Y (leader, dimension, centerline).
        x3: Third point X (angular_dimension).
        y3: Third point Y (angular_dimension).
        text: Text content (text_box, leader, balloon, note).
        height: Text height (text_box, note).
        leader_x: Leader X (balloon).
        leader_y: Leader Y (balloon).
        dim_x: Dimension placement X (dimension,
            angular/radial/diameter/ordinate).
        dim_y: Dimension placement Y (dimension,
            angular/radial/diameter/ordinate).
        center_x: Center X (radial_dimension,
            diameter_dimension).
        center_y: Center Y (radial_dimension,
            diameter_dimension).
        point_x: Point X (radial_dimension,
            diameter_dimension).
        point_y: Point Y (radial_dimension,
            diameter_dimension).
        origin_x: Datum origin X (ordinate_dimension).
        origin_y: Datum origin Y (ordinate_dimension).
        symbol_type: Surface finish type:
            'machined'|'any'|'prohibited'.
        weld_type: Weld type:
            'fillet'|'groove'|'plug'|'spot'|'seam'.
        tolerance_text: GD&T text (geometric_tolerance).
    """
    match type:
        case "text_box":
            return export_manager.add_text_box(x, y, text, height)
        case "leader":
            return export_manager.add_leader(x1, y1, x2, y2, text)
        case "dimension":
            return export_manager.add_dimension(x1, y1, x2, y2, dim_x, dim_y)
        case "balloon":
            return export_manager.add_balloon(x, y, text, leader_x, leader_y)
        case "note":
            return export_manager.add_note(x, y, text, height)
        case "angular_dimension":
            return export_manager.add_angular_dimension(x1, y1, x2, y2, x3, y3, dim_x, dim_y)
        case "radial_dimension":
            return export_manager.add_radial_dimension(
                center_x,
                center_y,
                point_x,
                point_y,
                dim_x,
                dim_y,
            )
        case "diameter_dimension":
            return export_manager.add_diameter_dimension(
                center_x,
                center_y,
                point_x,
                point_y,
                dim_x,
                dim_y,
            )
        case "ordinate_dimension":
            return export_manager.add_ordinate_dimension(
                origin_x,
                origin_y,
                point_x,
                point_y,
                dim_x,
                dim_y,
            )
        case "center_mark":
            return export_manager.add_center_mark(x, y)
        case "centerline":
            return export_manager.add_centerline(x1, y1, x2, y2)
        case "surface_finish":
            return export_manager.add_surface_finish_symbol(x, y, symbol_type)
        case "weld_symbol":
            return export_manager.add_weld_symbol(x, y, weld_type)
        case "geometric_tolerance":
            return export_manager.add_geometric_tolerance(x, y, tolerance_text)
        case _:
            return {"error": f"Unknown type: {type}"}


# ================================================================
# Group 51: add_2d_dimension (4 → 1)
# ================================================================


def add_2d_dimension(
    type: str = "distance",
    x1: float = 0.0,
    y1: float = 0.0,
    x2: float = 0.0,
    y2: float = 0.0,
    x3: float = 0.0,
    y3: float = 0.0,
    object_index: int = 0,
    object_type: str = "circle",
) -> dict:
    """Add a 2D dimension on the active draft sheet.

    type: 'distance' | 'length' | 'radius' | 'angle'

    Args:
        type: Dimension type to add.
        x1: Start X (distance, angle).
        y1: Start Y (distance, angle).
        x2: End X (distance, angle).
        y2: End Y (distance, angle).
        x3: Third point X (angle). Vertex is (x2, y2).
        y3: Third point Y (angle).
        object_index: 0-based index into Lines2d (length)
            or Circles2d/Arcs2d (radius).
        object_type: 'circle' or 'arc' (radius).
    """
    match type:
        case "distance":
            return export_manager.add_distance_dimension(x1, y1, x2, y2)
        case "length":
            return export_manager.add_length_dimension(object_index)
        case "radius":
            return export_manager.add_radius_dimension_2d(object_index, object_type)
        case "angle":
            return export_manager.add_angle_dimension_2d(x1, y1, x2, y2, x3, y3)
        case _:
            return {"error": f"Unknown type: {type}"}


# ================================================================
# Group 52: camera_control (10 → 1)
# ================================================================


def camera_control(
    action: str,
    view: str = "Iso",
    angle: float = 0.0,
    center_x: float = 0.0,
    center_y: float = 0.0,
    center_z: float = 0.0,
    axis_x: float = 0.0,
    axis_y: float = 1.0,
    axis_z: float = 0.0,
    dx: int = 0,
    dy: int = 0,
    factor: float = 1.0,
    eye_x: float = 0.0,
    eye_y: float = 0.0,
    eye_z: float = 1.0,
    target_x: float = 0.0,
    target_y: float = 0.0,
    target_z: float = 0.0,
    up_x: float = 0.0,
    up_y: float = 1.0,
    up_z: float = 0.0,
    perspective: bool = False,
    scale_or_angle: float = 1.0,
) -> dict:
    """Control the 3D camera/view.

    action: 'set_orientation' | 'zoom_fit'
            | 'zoom_to_selection' | 'rotate' | 'pan'
            | 'zoom' | 'refresh' | 'set_camera'
            | 'begin_dynamics' | 'end_dynamics'

    Args:
        action: Camera action to perform.
        view: Named view (set_orientation):
            'Iso'|'Top'|'Front'|'Right'|'Left'
            |'Back'|'Bottom'.
        angle: Rotation angle in radians (rotate).
        center_x: Rotation center X in meters (rotate).
        center_y: Rotation center Y in meters (rotate).
        center_z: Rotation center Z in meters (rotate).
        axis_x: Rotation axis X component (rotate).
        axis_y: Rotation axis Y component (rotate).
        axis_z: Rotation axis Z component (rotate).
        dx: Horizontal pan in pixels (pan).
        dy: Vertical pan in pixels (pan).
        factor: Zoom factor >1=in, <1=out (zoom).
        eye_x: Camera eye X (set_camera).
        eye_y: Camera eye Y (set_camera).
        eye_z: Camera eye Z (set_camera).
        target_x: Camera target X (set_camera).
        target_y: Camera target Y (set_camera).
        target_z: Camera target Z (set_camera).
        up_x: Camera up vector X (set_camera).
        up_y: Camera up vector Y (set_camera).
        up_z: Camera up vector Z (set_camera).
        perspective: Use perspective projection (set_camera).
        scale_or_angle: Ortho scale or perspective angle
            (set_camera).
    """
    match action:
        case "set_orientation":
            return view_manager.set_view(view)
        case "zoom_fit":
            return view_manager.zoom_fit()
        case "zoom_to_selection":
            return view_manager.zoom_to_selection()
        case "rotate":
            return view_manager.rotate_camera(
                angle,
                center_x,
                center_y,
                center_z,
                axis_x,
                axis_y,
                axis_z,
            )
        case "pan":
            return view_manager.pan_camera(dx, dy)
        case "zoom":
            return view_manager.zoom_camera(factor)
        case "refresh":
            return view_manager.refresh_view()
        case "set_camera":
            return view_manager.set_camera(
                eye_x,
                eye_y,
                eye_z,
                target_x,
                target_y,
                target_z,
                up_x,
                up_y,
                up_z,
                perspective,
                scale_or_angle,
            )
        case "begin_dynamics":
            return view_manager.begin_camera_dynamics()
        case "end_dynamics":
            return view_manager.end_camera_dynamics()
        case _:
            return {"error": f"Unknown action: {action}"}


# ================================================================
# Group 53: display_control (4 → 1)
# ================================================================


def display_control(
    action: str,
    mode: str = "Shaded",
    red: int = 0,
    green: int = 0,
    blue: int = 0,
    x: float = 0.0,
    y: float = 0.0,
    z: float = 0.0,
    screen_x: int = 0,
    screen_y: int = 0,
    face_index: int = 0,
    texture_name: str = "",
) -> dict:
    """Control display settings and coordinate transforms.

    action: 'set_mode' | 'set_background'
            | 'model_to_screen' | 'screen_to_model'
            | 'set_texture'

    Args:
        action: Display action to perform.
        mode: Display mode (set_mode):
            'Shaded'|'ShadedWithEdges'|'Wireframe'
            |'HiddenEdgesVisible'.
        red: Red component 0-255 (set_background).
        green: Green component 0-255 (set_background).
        blue: Blue component 0-255 (set_background).
        x: Model X coordinate (model_to_screen).
        y: Model Y coordinate (model_to_screen).
        z: Model Z coordinate (model_to_screen).
        screen_x: Screen X pixel (screen_to_model).
        screen_y: Screen Y pixel (screen_to_model).
        face_index: 0-based face index (set_texture).
        texture_name: Texture name to apply (set_texture).
    """
    match action:
        case "set_mode":
            return view_manager.set_display_mode(mode)
        case "set_background":
            return view_manager.set_view_background(red, green, blue)
        case "model_to_screen":
            return view_manager.transform_model_to_screen(x, y, z)
        case "screen_to_model":
            return view_manager.transform_screen_to_model(screen_x, screen_y)
        case "set_texture":
            return export_manager.set_face_texture(face_index, texture_name)
        case _:
            return {"error": f"Unknown action: {action}"}


# ================================================================
# Group 54: manage_sheet (3 → 1)
# ================================================================


def manage_sheet(
    action: str,
    sheet_index: int = 0,
    new_name: str = "",
    template: str | None = None,
    views: list[str] | None = None,
) -> dict:
    """Manage draft sheets.

    action: 'activate' | 'rename' | 'delete'
            | 'create_drawing' | 'add'

    Args:
        action: Sheet management action.
        sheet_index: 0-based sheet index.
        new_name: New sheet name (rename).
        template: Drawing template path (create_drawing).
        views: List of view orientations (create_drawing).
    """
    match action:
        case "activate":
            return export_manager.activate_sheet(sheet_index)
        case "rename":
            return export_manager.rename_sheet(sheet_index, new_name)
        case "delete":
            return export_manager.delete_sheet(sheet_index)
        case "create_drawing":
            return export_manager.create_drawing(template, views)
        case "add":
            return export_manager.add_draft_sheet()
        case _:
            return {"error": f"Unknown action: {action}"}


# ================================================================
# Group 55: print_control (4 → 1)
# ================================================================


def print_control(
    action: str,
    copies: int = 1,
    all_sheets: bool = True,
    printer_name: str = "",
    width: float = 0.0,
    height: float = 0.0,
    orientation: str = "Landscape",
    num_copies: int = 1,
    print_orientation: int | None = None,
    paper_size: int | None = None,
    scale: float | None = None,
    print_to_file: bool = False,
    output_file_name: str | None = None,
    print_range: int | None = None,
    sheets: str | None = None,
    color_as_black: bool = False,
    collate: bool = True,
) -> dict:
    """Control printing for the active draft.

    action: 'print' | 'set_printer' | 'get_printer'
            | 'set_paper_size' | 'print_full'

    Args:
        action: Print action to perform.
        copies: Number of copies (print).
        all_sheets: Print all sheets (print).
        printer_name: Printer name (set_printer, print_full).
        width: Paper width in meters (set_paper_size).
        height: Paper height in meters (set_paper_size).
        orientation: 'Landscape' or 'Portrait'
            (set_paper_size).
        num_copies: Number of copies (print_full).
        print_orientation: Orientation constant (print_full).
        paper_size: Paper size constant (print_full).
        scale: Print scale (print_full).
        print_to_file: Print to file (print_full).
        output_file_name: Output file path (print_full).
        print_range: Print range constant (print_full).
        sheets: Sheet range string (print_full).
        color_as_black: Print color as black (print_full).
        collate: Collate copies (print_full).
    """
    match action:
        case "print":
            return export_manager.print_drawing(copies, all_sheets)
        case "set_printer":
            return export_manager.set_printer(printer_name)
        case "get_printer":
            return export_manager.get_printer()
        case "set_paper_size":
            return export_manager.set_paper_size(width, height, orientation)
        case "print_full":
            return export_manager.print_document(
                printer=printer_name or None,
                num_copies=num_copies,
                orientation=print_orientation,
                paper_size=paper_size,
                scale=scale,
                print_to_file=print_to_file,
                output_file_name=output_file_name,
                print_range=print_range,
                sheets=sheets,
                color_as_black=color_as_black,
                collate=collate,
            )
        case _:
            return {"error": f"Unknown action: {action}"}


# ================================================================
# Group 56: query_sheet (9 → 1)
# ================================================================


def query_sheet(
    type: str,
    view_index: int = 0,
) -> dict:
    """Query sheet collections on the active draft.

    type: 'dimensions' | 'balloons' | 'text_boxes'
          | 'drawing_objects' | 'sections'
          | 'lines2d' | 'circles2d' | 'arcs2d'
          | 'section_cuts'

    Args:
        type: Collection type to query.
        view_index: Drawing view index (section_cuts only).
    """
    match type:
        case "dimensions":
            return export_manager.get_sheet_dimensions()
        case "balloons":
            return export_manager.get_sheet_balloons()
        case "text_boxes":
            return export_manager.get_sheet_text_boxes()
        case "drawing_objects":
            return export_manager.get_sheet_drawing_objects()
        case "sections":
            return export_manager.get_sheet_sections()
        case "lines2d":
            return export_manager.get_lines2d()
        case "circles2d":
            return export_manager.get_circles2d()
        case "arcs2d":
            return export_manager.get_arcs2d()
        case "section_cuts":
            return export_manager.get_section_cuts(view_index)
        case _:
            return {"error": f"Unknown type: {type}"}


# ================================================================
# Group 57: manage_annotation_data (4 → 1)
# ================================================================


def manage_annotation_data(
    action: str,
    file_path: str = "",
    x: float = 0.0,
    y: float = 0.0,
    insertion_type: int = 0,
    show: bool = True,
    show_dimensions: bool = True,
    show_annotations: bool = True,
) -> dict:
    """Manage symbols and PMI annotation data.

    action: 'add_symbol' | 'get_symbols'
            | 'get_pmi' | 'set_pmi_visibility'

    Args:
        action: Annotation data action.
        file_path: Symbol file path (add_symbol).
        x: Symbol X position (add_symbol).
        y: Symbol Y position (add_symbol).
        insertion_type: Symbol insertion type (add_symbol).
        show: Show PMI (set_pmi_visibility).
        show_dimensions: Show PMI dimensions
            (set_pmi_visibility).
        show_annotations: Show PMI annotations
            (set_pmi_visibility).
    """
    match action:
        case "add_symbol":
            return export_manager.add_symbol(file_path, x, y, insertion_type)
        case "get_symbols":
            return export_manager.get_symbols()
        case "get_pmi":
            return export_manager.get_pmi_info()
        case "set_pmi_visibility":
            return export_manager.set_pmi_visibility(show, show_dimensions, show_annotations)
        case _:
            return {"error": f"Unknown action: {action}"}


# ================================================================
# Group 58: add_smart_frame (2 → 1)
# ================================================================


def add_smart_frame(
    method: str = "two_point",
    style_name: str = "",
    x1: float = 0.0,
    y1: float = 0.0,
    x2: float = 0.0,
    y2: float = 0.0,
    x: float = 0.0,
    y: float = 0.0,
    top: float = 0.0,
    bottom: float = 0.0,
    left: float = 0.0,
    right: float = 0.0,
) -> dict:
    """Add a smart frame (title block/border) to the sheet.

    method: 'two_point' | 'by_origin'

    Args:
        method: Smart frame creation method.
        style_name: Frame style name.
        x1: First corner X (two_point).
        y1: First corner Y (two_point).
        x2: Second corner X (two_point).
        y2: Second corner Y (two_point).
        x: Origin X (by_origin).
        y: Origin Y (by_origin).
        top: Top margin (by_origin).
        bottom: Bottom margin (by_origin).
        left: Left margin (by_origin).
        right: Right margin (by_origin).
    """
    match method:
        case "two_point":
            return export_manager.add_smart_frame(style_name, x1, y1, x2, y2)
        case "by_origin":
            return export_manager.add_smart_frame_by_origin(
                style_name, x, y, top, bottom, left, right
            )
        case _:
            return {"error": f"Unknown method: {method}"}


# ================================================================
# Group 59: draft_config (4 → 1)
# ================================================================


def draft_config(
    action: str,
    parameter: int = 0,
    value: float = 0.0,
    x: float = 0.0,
    y: float = 0.0,
) -> dict:
    """Manage draft document configuration.

    action: 'get_global' | 'set_global'
            | 'get_origin' | 'set_origin'

    Args:
        action: Draft config action.
        parameter: Global parameter ID (get_global, set_global).
        value: Parameter value to set (set_global).
        x: Symbol file origin X (set_origin).
        y: Symbol file origin Y (set_origin).
    """
    match action:
        case "get_global":
            return export_manager.get_draft_global_parameter(parameter)
        case "set_global":
            return export_manager.set_draft_global_parameter(parameter, value)
        case "get_origin":
            return export_manager.get_symbol_file_origin()
        case "set_origin":
            return export_manager.set_symbol_file_origin(x, y)
        case _:
            return {"error": f"Unknown action: {action}"}


# ================================================================
# Group 60: create_table (2 → 1)
# ================================================================


def create_table(
    type: str = "parts_list",
    auto_balloon: bool = True,
    x: float = 0.15,
    y: float = 0.25,
    view_index: int = 0,
    saved_settings: str = "",
) -> dict:
    """Create a table on the active draft sheet.

    type: 'parts_list' | 'bend'

    Args:
        type: Table type to create.
        auto_balloon: Auto-create balloons (parts_list, bend).
        x: Table X position (parts_list).
        y: Table Y position (parts_list).
        view_index: 0-based drawing view index (bend).
        saved_settings: Saved settings name (bend).
    """
    match type:
        case "parts_list":
            return export_manager.create_parts_list(auto_balloon, x, y)
        case "bend":
            return export_manager.create_bend_table(view_index, saved_settings, auto_balloon)
        case _:
            return {"error": f"Unknown type: {type}"}


# ================================================================
# Registration
# ================================================================


def register(mcp):
    """Register export, drawing, and view tools."""
    mcp.tool()(export_file)
    mcp.tool()(add_drawing_view)
    mcp.tool()(manage_drawing_view)
    mcp.tool()(add_annotation)
    mcp.tool()(add_2d_dimension)
    mcp.tool()(camera_control)
    mcp.tool()(display_control)
    mcp.tool()(manage_sheet)
    mcp.tool()(print_control)
    mcp.tool()(query_sheet)
    mcp.tool()(manage_annotation_data)
    mcp.tool()(add_smart_frame)
    mcp.tool()(draft_config)
    mcp.tool()(create_table)
