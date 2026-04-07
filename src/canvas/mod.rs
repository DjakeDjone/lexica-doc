mod editor_input;
mod page_layout;
mod palette;

use std::collections::hash_map::Entry;
use std::sync::Arc;

use eframe::egui::{
    self,
    epaint::CornerRadius,
    text_selection::visuals::{paint_text_cursor, paint_text_selection},
    Align2, Color32, FontFamily, FontId, Id, Rect, Sense, Stroke, StrokeKind,
};

use crate::{
    app::{CanvasState, ImageResizeDrag, ResizeHandle, ThemeMode},
    document::{
        text_format, CharacterStyle, DocumentImage, DocumentState, ImageRendering,
        ParagraphAlignment, OBJECT_REPLACEMENT_CHAR,
    },
    layout::{
        centered_page_rect, document_points_to_pixels, document_points_to_screen_points,
        page_content_rect,
    },
};

use editor_input::{apply_viewport_input, handle_keyboard_input, handle_pointer_interaction};
use page_layout::layout_page_stack;
use palette::canvas_palette;

struct DocumentLayout {
    galley: Arc<egui::Galley>,
    list_markers: Vec<ListMarkerLayout>,
    images: Vec<ImageLayout>,
    manual_page_break_rows: Vec<usize>,
}

struct ListMarkerLayout {
    row_index: usize,
    text: String,
    x: f32,
    font_id: FontId,
    color: Color32,
}

struct ImageLayout {
    row_index: usize,
    size: egui::Vec2,
    image: DocumentImage,
}

pub fn paint_document_canvas(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    theme_mode: ThemeMode,
) {
    let palette = canvas_palette(theme_mode);
    let viewport = ui.available_rect_before_wrap();
    let editor_id = Id::new("document_canvas");
    let response = ui.interact(viewport, editor_id, Sense::click_and_drag());
    let painter = ui.painter_at(viewport);
    let pixels_per_point = ui.ctx().pixels_per_point();
    apply_viewport_input(ui, &response, canvas);

    painter.rect_filled(viewport, CornerRadius::ZERO, palette.canvas_bg);

    let base_page_rect =
        centered_page_rect(viewport, document.page_size, canvas.zoom, egui::Vec2::ZERO);
    let content_size = page_content_rect(base_page_rect, document.margins, canvas.zoom).size();
    let mut document_layout = layout_document(ui, document, canvas, content_size.x);
    let page_layout = layout_page_stack(
        viewport,
        document,
        canvas,
        &document_layout.galley,
        &document_layout.manual_page_break_rows,
    );

    handle_image_interaction(ui, &response, canvas, document);

    handle_pointer_interaction(
        ui,
        &response,
        &page_layout,
        &document_layout.galley,
        canvas,
        document,
    );

    let has_focus = ui.memory(|mem| mem.has_focus(editor_id));
    if has_focus && handle_keyboard_input(ui, document, canvas, &document_layout.galley) {
        document_layout = layout_document(ui, document, canvas, content_size.x);
    }

    if has_focus && !canvas.selection.is_empty() {
        paint_text_selection(
            &mut document_layout.galley,
            ui.visuals(),
            &canvas.selection,
            None,
        );
    }

    let mut new_image_rects: Vec<(usize, Rect)> = Vec::new();

    for page in &page_layout.pages {
        let shadow_offset = egui::vec2(
            document_points_to_screen_points(6.0, canvas.zoom),
            document_points_to_screen_points(8.0, canvas.zoom),
        );
        painter.rect_filled(
            page.page_rect.translate(shadow_offset),
            CornerRadius::same(2),
            palette.page_shadow,
        );
        painter.rect_filled(page.page_rect, CornerRadius::same(2), palette.page_bg);
        painter.rect_stroke(
            page.page_rect,
            CornerRadius::same(2),
            Stroke::new(
                1.0,
                if has_focus {
                    palette.page_focus
                } else {
                    palette.page_border
                },
            ),
            StrokeKind::Outside,
        );

        let visible_content_rect = Rect::from_min_size(
            page.content_rect.min,
            egui::vec2(
                page.content_rect.width(),
                (page.end_y - page.start_y)
                    .max(0.0)
                    .min(page.content_rect.height()),
            ),
        );
        let galley_origin = page.content_rect.min - egui::vec2(0.0, page.start_y);
        painter.with_clip_rect(visible_content_rect).galley(
            galley_origin,
            document_layout.galley.clone(),
            Color32::BLACK,
        );

        let clipped_painter = painter.with_clip_rect(visible_content_rect);
        for marker in &document_layout.list_markers {
            let Some(row) = document_layout.galley.rows.get(marker.row_index) else {
                continue;
            };
            let marker_y = row.pos.y;
            if marker_y < page.start_y || marker_y > page.end_y {
                continue;
            }

            let marker_pos = egui::pos2(
                page.content_rect.left() + marker.x,
                page.content_rect.top() + marker_y - page.start_y,
            );
            clipped_painter.text(
                marker_pos,
                Align2::RIGHT_TOP,
                &marker.text,
                marker.font_id.clone(),
                marker.color,
            );
        }

        for image in &document_layout.images {
            let Some(row) = document_layout.galley.rows.get(image.row_index) else {
                continue;
            };
            let image_y = row.pos.y;
            if image_y < page.start_y || image_y > page.end_y {
                continue;
            }

            let image_rect = Rect::from_min_size(
                egui::pos2(
                    page.content_rect.left() + row.pos.x,
                    page.content_rect.top() + image_y - page.start_y,
                ),
                image.size,
            );

            if let Some(texture) = texture_for_image(ui.ctx(), canvas, &image.image) {
                let tint = Color32::from_white_alpha(
                    (image.image.opacity * 255.0).round().clamp(0.0, 255.0) as u8,
                );
                clipped_painter.image(
                    texture.id(),
                    image_rect,
                    Rect::from_min_max(egui::Pos2::ZERO, egui::pos2(1.0, 1.0)),
                    tint,
                );
            } else {
                clipped_painter.rect_filled(image_rect, CornerRadius::same(4), palette.footer_bg);
                clipped_painter.rect_stroke(
                    image_rect,
                    CornerRadius::same(4),
                    Stroke::new(1.0, palette.footer_stroke),
                    StrokeKind::Outside,
                );
                clipped_painter.text(
                    image_rect.center(),
                    Align2::CENTER_CENTER,
                    &image.image.alt_text,
                    FontId::new(12.0, FontFamily::Proportional),
                    palette.footer_text,
                );
            }

            new_image_rects.push((image.image.id, image_rect));
        }
    }

    canvas.image_rects = new_image_rects;

    // Draw selection border + handles with unclipped painter so they aren't cut at page margins
    if let Some((_, selected_rect)) = canvas
        .image_rects
        .iter()
        .find(|(id, _)| Some(*id) == canvas.selected_image_id)
    {
        paint_image_selection(&painter, *selected_rect);
    }

    if has_focus && canvas.selected_image_id.is_none() {
        if let Some(caret_rect) =
            page_layout.caret_rect(&document_layout.galley, canvas.selection.primary)
        {
            paint_text_cursor(
                ui,
                &painter,
                caret_rect,
                ui.input(|i| i.time) - canvas.last_interaction_time,
            );
        }
    }

    let page_pixels = (
        document_points_to_pixels(
            document.page_size.width_points,
            pixels_per_point,
            canvas.zoom,
        ),
        document_points_to_pixels(
            document.page_size.height_points,
            pixels_per_point,
            canvas.zoom,
        ),
    );
    let footer = format!(
        "{:.0} x {:.0} px  |  {} pages  |  y {:.0}",
        page_pixels.0,
        page_pixels.1,
        page_layout.pages.len(),
        canvas.pan.y
    );
    let footer_galley = painter.layout_no_wrap(
        footer,
        FontId::new(11.0, FontFamily::Monospace),
        palette.footer_text,
    );
    let footer_rect = Rect::from_min_size(
        egui::pos2(
            viewport.left() + 22.0,
            viewport.bottom() - footer_galley.size().y - 24.0,
        ),
        footer_galley.size() + egui::vec2(20.0, 14.0),
    );
    painter.rect_filled(footer_rect, CornerRadius::same(3), palette.footer_bg);
    painter.rect_stroke(
        footer_rect,
        CornerRadius::same(3),
        Stroke::new(1.0, palette.footer_stroke),
        StrokeKind::Outside,
    );
    painter.galley(
        egui::pos2(footer_rect.left() + 10.0, footer_rect.top() + 7.0),
        footer_galley,
        palette.footer_text,
    );
}

fn resize_handle_rects(image_rect: Rect) -> [(ResizeHandle, Rect); 8] {
    const H: f32 = 5.0;
    let sq = |x: f32, y: f32| {
        Rect::from_center_size(egui::pos2(x, y), egui::vec2(H * 2.0, H * 2.0))
    };
    let cx = image_rect.center().x;
    let cy = image_rect.center().y;
    [
        (ResizeHandle::NW, sq(image_rect.left(), image_rect.top())),
        (ResizeHandle::N, sq(cx, image_rect.top())),
        (ResizeHandle::NE, sq(image_rect.right(), image_rect.top())),
        (ResizeHandle::E, sq(image_rect.right(), cy)),
        (ResizeHandle::SE, sq(image_rect.right(), image_rect.bottom())),
        (ResizeHandle::S, sq(cx, image_rect.bottom())),
        (ResizeHandle::SW, sq(image_rect.left(), image_rect.bottom())),
        (ResizeHandle::W, sq(image_rect.left(), cy)),
    ]
}

fn paint_image_selection(painter: &egui::Painter, image_rect: Rect) {
    const SELECTION_COLOR: Color32 = Color32::from_rgb(54, 116, 206);
    painter.rect_stroke(
        image_rect,
        CornerRadius::ZERO,
        Stroke::new(2.0, SELECTION_COLOR),
        StrokeKind::Outside,
    );
    for (_, handle_rect) in &resize_handle_rects(image_rect) {
        painter.rect_filled(*handle_rect, CornerRadius::ZERO, Color32::WHITE);
        painter.rect_stroke(
            *handle_rect,
            CornerRadius::ZERO,
            Stroke::new(1.5, SELECTION_COLOR),
            StrokeKind::Outside,
        );
    }
}

fn handle_image_interaction(
    ui: &mut egui::Ui,
    response: &egui::Response,
    canvas: &mut CanvasState,
    document: &mut DocumentState,
) {
    // Change cursor icon when hovering over a resize handle
    if let Some(hover_pos) = ui.ctx().pointer_hover_pos() {
        let mut cursor_icon = None;
        'hover: for &(_, image_rect) in &canvas.image_rects {
            for &(handle, handle_rect) in &resize_handle_rects(image_rect) {
                if handle_rect.expand(3.0).contains(hover_pos) {
                    cursor_icon = Some(match handle {
                        ResizeHandle::NW | ResizeHandle::SE => egui::CursorIcon::ResizeNwSe,
                        ResizeHandle::NE | ResizeHandle::SW => egui::CursorIcon::ResizeNeSw,
                        ResizeHandle::N | ResizeHandle::S => egui::CursorIcon::ResizeSouth,
                        ResizeHandle::E | ResizeHandle::W => egui::CursorIcon::ResizeEast,
                    });
                    break 'hover;
                }
            }
        }
        if let Some(icon) = cursor_icon {
            ui.ctx().set_cursor_icon(icon);
        }
    }

    // Finalize any active resize drag when mouse released
    if !response.dragged() {
        canvas.resize_drag = None;
    }

    // Continue active resize drag
    if response.dragged() {
        let drag_data = canvas.resize_drag.as_ref().map(|d| {
            (
                d.image_id,
                d.handle,
                d.start_ptr,
                d.start_width_points,
                d.start_height_points,
            )
        });
        if let Some((image_id, handle, start_ptr, start_w, start_h)) = drag_data {
            if let Some(pointer_pos) = response.interact_pointer_pos() {
                let zoom = canvas.zoom;
                let dx = (pointer_pos.x - start_ptr.x) / zoom;
                let dy = (pointer_pos.y - start_ptr.y) / zoom;
                let shift = ui.input(|i| i.modifiers.shift);
                let aspect = start_h / start_w.max(1.0);
                let (new_w, new_h) = match handle {
                    ResizeHandle::SE => {
                        if shift { let w = start_w + dx; (w, w * aspect) }
                        else { (start_w + dx, start_h + dy) }
                    }
                    ResizeHandle::SW => {
                        if shift { let w = start_w - dx; (w, w * aspect) }
                        else { (start_w - dx, start_h + dy) }
                    }
                    ResizeHandle::NE => {
                        if shift { let w = start_w + dx; (w, w * aspect) }
                        else { (start_w + dx, start_h - dy) }
                    }
                    ResizeHandle::NW => {
                        if shift { let w = start_w - dx; (w, w * aspect) }
                        else { (start_w - dx, start_h - dy) }
                    }
                    ResizeHandle::E => (start_w + dx, start_h),
                    ResizeHandle::W => (start_w - dx, start_h),
                    ResizeHandle::S => (start_w, start_h + dy),
                    ResizeHandle::N => (start_w, start_h - dy),
                };
                document.resize_image_by_id(image_id, new_w.max(24.0), new_h.max(24.0));
            }
            return;
        }
    }

    let Some(pointer_pos) = response.interact_pointer_pos() else {
        return;
    };

    // Start a new resize drag when dragging from a handle
    if response.drag_started() {
        let mut found = None;
        'find: for &(image_id, image_rect) in &canvas.image_rects {
            for &(handle, handle_rect) in &resize_handle_rects(image_rect) {
                if handle_rect.expand(3.0).contains(pointer_pos) {
                    found = Some((image_id, handle));
                    break 'find;
                }
            }
        }
        if let Some((image_id, handle)) = found {
            let size = document
                .paragraph_images
                .iter()
                .flatten()
                .find(|img| img.id == image_id)
                .map(|img| (img.width_points, img.height_points));
            if let Some((w, h)) = size {
                canvas.resize_drag = Some(ImageResizeDrag {
                    image_id,
                    handle,
                    start_ptr: pointer_pos,
                    start_width_points: w,
                    start_height_points: h,
                });
            }
            return;
        }
    }

    // Click on image body → select it; click elsewhere → deselect
    if response.clicked() {
        let hit = canvas
            .image_rects
            .iter()
            .find(|(_, rect)| rect.contains(pointer_pos))
            .map(|(id, _)| *id);
        canvas.selected_image_id = hit;
    }
}

fn layout_document(
    ui: &mut egui::Ui,
    document: &DocumentState,
    canvas: &CanvasState,
    wrap_width: f32,
) -> DocumentLayout {
    let marker_gutter = document_points_to_screen_points(24.0, canvas.zoom);
    let marker_gap = document_points_to_screen_points(6.0, canvas.zoom);
    let default_style = CharacterStyle::default();
    let painter = ui.painter();

    let mut paragraph_galleys = Vec::new();
    let mut list_markers = Vec::new();
    let mut images = Vec::new();
    let mut manual_page_break_rows = Vec::new();
    let mut row_index = 0usize;

    for paragraph in document.paragraphs() {
        let indent = if paragraph.list_marker.is_some() {
            marker_gutter
        } else {
            0.0
        };
        let paragraph_wrap_width = (wrap_width - indent).max(1.0);
        let mut job = egui::epaint::text::LayoutJob::default();
        job.wrap.max_width = paragraph_wrap_width;
        job.break_on_newline = true;
        job.halign = egui::Align::LEFT;
        job.justify = paragraph.style.alignment == ParagraphAlignment::Justify;

        let has_visible_text = paragraph
            .runs
            .iter()
            .any(|run| run.text.chars().any(|ch| ch != OBJECT_REPLACEMENT_CHAR));

        if paragraph.runs.is_empty() {
            job.append("", 0.0, text_format(default_style, canvas.zoom));
        } else {
            for run in &paragraph.runs {
                append_run_with_placeholders(
                    &mut job,
                    run,
                    canvas.zoom,
                    paragraph.image.is_some() && !has_visible_text,
                );
            }
        }

        let marker_style = paragraph
            .runs
            .first()
            .map(|run| run.style)
            .unwrap_or(default_style);
        let mut paragraph_galley = painter.layout_job(job);

        if paragraph.style.page_break_before && row_index > 0 {
            manual_page_break_rows.push(row_index);
        }

        if let Some(image) = paragraph.image.clone().filter(|_| !has_visible_text) {
            let display_size = image_display_size(&image, paragraph_wrap_width, canvas.zoom);
            if reserve_block_image_space(&mut paragraph_galley, display_size) {
                images.push(ImageLayout {
                    row_index,
                    size: display_size,
                    image,
                });
            }
        }
        align_paragraph_galley(
            &mut paragraph_galley,
            indent,
            paragraph_wrap_width,
            paragraph.style.alignment,
        );

        if let Some(marker_text) = paragraph.list_marker {
            list_markers.push(ListMarkerLayout {
                row_index,
                text: marker_text,
                x: indent - marker_gap,
                font_id: text_format(marker_style, canvas.zoom).font_id,
                color: marker_style.text_color,
            });
        }

        row_index += paragraph_galley.rows.len();
        paragraph_galleys.push(paragraph_galley);
    }

    let plain_text = document.plain_text();
    let mut merged_job = egui::epaint::text::LayoutJob::default();
    merged_job.wrap.max_width = wrap_width;
    merged_job.break_on_newline = true;
    merged_job.append(&plain_text, 0.0, text_format(default_style, canvas.zoom));

    DocumentLayout {
        galley: Arc::new(egui::Galley::concat(
            Arc::new(merged_job),
            &paragraph_galleys,
            painter.pixels_per_point(),
        )),
        list_markers,
        images,
        manual_page_break_rows,
    }
}

fn append_run_with_placeholders(
    job: &mut egui::epaint::text::LayoutJob,
    run: &crate::document::TextRun,
    zoom: f32,
    keep_visible_placeholder: bool,
) {
    let mut segment = String::new();
    for ch in run.text.chars() {
        if ch == OBJECT_REPLACEMENT_CHAR {
            if !segment.is_empty() {
                job.append(&segment, 0.0, text_format(run.style, zoom));
                segment.clear();
            }

            let mut placeholder_style = run.style;
            if !keep_visible_placeholder {
                placeholder_style.text_color = Color32::TRANSPARENT;
            }
            job.append(
                &OBJECT_REPLACEMENT_CHAR.to_string(),
                0.0,
                text_format(placeholder_style, zoom),
            );
        } else {
            segment.push(ch);
        }
    }

    if !segment.is_empty() {
        job.append(&segment, 0.0, text_format(run.style, zoom));
    }
}

fn reserve_block_image_space(galley: &mut Arc<egui::Galley>, size: egui::Vec2) -> bool {
    let galley = Arc::make_mut(galley);
    let Some(placed_row) = galley.rows.first_mut() else {
        return false;
    };
    let row = Arc::make_mut(&mut placed_row.row);
    if row.glyphs.len() != 1 || row.glyphs[0].chr != OBJECT_REPLACEMENT_CHAR {
        return false;
    }

    row.glyphs[0].advance_width = size.x;
    row.glyphs[0].line_height = size.y;
    row.glyphs[0].font_ascent = size.y;
    row.glyphs[0].font_height = size.y;
    row.glyphs[0].font_face_ascent = size.y;
    row.glyphs[0].font_face_height = size.y;
    row.size = size;
    row.visuals = Default::default();

    galley.rect = Rect::from_min_size(galley.rect.min, size);
    galley.mesh_bounds = Rect::NOTHING;
    true
}

fn image_display_size(image: &DocumentImage, wrap_width: f32, zoom: f32) -> egui::Vec2 {
    let width = document_points_to_screen_points(image.width_points.max(24.0), zoom);
    let height = document_points_to_screen_points(image.height_points.max(24.0), zoom);
    if width <= wrap_width {
        return egui::vec2(width, height);
    }

    let scale = wrap_width / width;
    egui::vec2(wrap_width, (height * scale).max(24.0))
}

fn texture_for_image<'a>(
    ctx: &egui::Context,
    canvas: &'a mut CanvasState,
    image: &DocumentImage,
) -> Option<&'a egui::TextureHandle> {
    // Encode rendering mode into the cache key so smooth/crisp use separate textures
    let cache_key = image.id * 2 + if image.rendering == ImageRendering::Crisp { 1 } else { 0 };
    let tex_options = if image.rendering == ImageRendering::Crisp {
        egui::TextureOptions::NEAREST
    } else {
        egui::TextureOptions::LINEAR
    };
    match canvas.image_textures.entry(cache_key) {
        Entry::Occupied(entry) => Some(entry.into_mut()),
        Entry::Vacant(entry) => {
            let decoded = ::image::load_from_memory(&image.bytes).ok()?;
            let rgba = decoded.to_rgba8();
            let size = [rgba.width() as usize, rgba.height() as usize];
            let color_image =
                egui::ColorImage::from_rgba_unmultiplied(size, rgba.as_raw().as_slice());
            let texture = ctx.load_texture(
                format!("doc-image-{}-{}", image.id, cache_key & 1),
                color_image,
                tex_options,
            );
            Some(entry.insert(texture))
        }
    }
}

fn align_paragraph_galley(
    galley: &mut Arc<egui::Galley>,
    indent: f32,
    wrap_width: f32,
    alignment: ParagraphAlignment,
) {
    let target_offsets: Vec<f32> = galley
        .rows
        .iter()
        .enumerate()
        .map(|(index, row)| {
            let base_offset = match alignment {
                ParagraphAlignment::Left | ParagraphAlignment::Justify => 0.0,
                ParagraphAlignment::Center => ((wrap_width - row.size.x) * 0.5).max(0.0),
                ParagraphAlignment::Right => (wrap_width - row.size.x).max(0.0),
            };
            let current_x = galley.rows[index].pos.x;
            indent + base_offset - current_x
        })
        .collect();

    if target_offsets
        .iter()
        .all(|delta| delta.abs() <= f32::EPSILON)
    {
        return;
    }

    let galley = Arc::make_mut(galley);
    let mut min_rect = galley.rect.min;
    let mut max_rect = galley.rect.max;
    let mut mesh_bounds = egui::Rect::NOTHING;

    for (row, delta) in galley.rows.iter_mut().zip(target_offsets) {
        row.pos.x += delta;
        let row_rect = row.rect();
        min_rect.x = min_rect.x.min(row_rect.min.x);
        min_rect.y = min_rect.y.min(row_rect.min.y);
        max_rect.x = max_rect.x.max(row_rect.max.x);
        max_rect.y = max_rect.y.max(row_rect.max.y);
        mesh_bounds |= row.visuals.mesh_bounds.translate(row.pos.to_vec2());
    }

    galley.rect = egui::Rect::from_min_max(min_rect, max_rect);
    galley.mesh_bounds = mesh_bounds;
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{
        app::CanvasState,
        document::{
            CharacterStyle, DocumentImage, DocumentState, ListKind, PageMargins, PageSize,
            ParagraphAlignment, ParagraphStyle, TextRun, OBJECT_REPLACEMENT_CHAR,
        },
    };

    /// Run `layout_document` inside a headless egui context and return
    /// the layout result for assertion.
    fn run_headless_layout(
        document: &DocumentState,
        canvas: &CanvasState,
        wrap_width: f32,
    ) -> DocumentLayout {
        let ctx = egui::Context::default();
        let mut layout = None;

        let _ = ctx.run_ui(egui::RawInput::default(), |ui| {
            layout = Some(layout_document(ui, document, canvas, wrap_width));
        });

        layout.expect("layout_document should have been called inside the egui frame")
    }

    fn make_document(
        runs: Vec<TextRun>,
        paragraph_styles: Vec<ParagraphStyle>,
        paragraph_images: Vec<Option<DocumentImage>>,
    ) -> DocumentState {
        DocumentState {
            title: "Test".to_owned(),
            runs,
            paragraph_styles,
            paragraph_images,
            page_size: PageSize::a4(),
            margins: PageMargins::standard(),
        }
    }

    #[test]
    fn single_paragraph_produces_at_least_one_row() {
        let document = make_document(
            vec![TextRun {
                text: "Hello world".to_owned(),
                style: CharacterStyle::default(),
            }],
            vec![ParagraphStyle::default()],
            vec![None],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 600.0);

        assert!(
            !layout.galley.rows.is_empty(),
            "galley should have at least one row"
        );
        assert!(
            layout.manual_page_break_rows.is_empty(),
            "no manual page breaks expected"
        );
        assert!(layout.images.is_empty(), "no images expected");
    }

    #[test]
    fn long_text_wraps_into_multiple_rows() {
        let long_text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. ".repeat(20);
        let document = make_document(
            vec![TextRun {
                text: long_text,
                style: CharacterStyle::default(),
            }],
            vec![ParagraphStyle::default()],
            vec![None],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 300.0);

        assert!(
            layout.galley.rows.len() > 1,
            "long text at narrow width should produce multiple rows, got {}",
            layout.galley.rows.len()
        );
    }

    #[test]
    fn manual_page_break_is_recorded() {
        let document = make_document(
            vec![TextRun {
                text: "First paragraph\nSecond paragraph".to_owned(),
                style: CharacterStyle::default(),
            }],
            vec![
                ParagraphStyle::default(),
                ParagraphStyle {
                    page_break_before: true,
                    ..ParagraphStyle::default()
                },
            ],
            vec![None, None],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 600.0);

        assert!(
            !layout.manual_page_break_rows.is_empty(),
            "should record a manual page break"
        );
        // The page break should be at the row index where the second paragraph starts.
        assert!(
            layout.manual_page_break_rows[0] > 0,
            "page break row index should be > 0"
        );
    }

    #[test]
    fn block_image_paragraph_produces_image_layout() {
        let document = make_document(
            vec![TextRun {
                text: format!("{OBJECT_REPLACEMENT_CHAR}"),
                style: CharacterStyle::default(),
            }],
            vec![ParagraphStyle::default()],
            vec![Some(DocumentImage {
                id: 1,
                bytes: vec![],
                alt_text: "test image".to_owned(),
                width_points: 200.0,
                height_points: 100.0,
            })],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 600.0);

        assert_eq!(layout.images.len(), 1, "should have one image layout");
        assert_eq!(layout.images[0].row_index, 0);
        assert!(
            layout.images[0].size.x > 0.0 && layout.images[0].size.y > 0.0,
            "image size should be positive"
        );
    }

    #[test]
    fn centered_alignment_offsets_row_positions() {
        let document = make_document(
            vec![TextRun {
                text: "Short".to_owned(),
                style: CharacterStyle::default(),
            }],
            vec![ParagraphStyle {
                alignment: ParagraphAlignment::Center,
                ..ParagraphStyle::default()
            }],
            vec![None],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 600.0);

        assert!(!layout.galley.rows.is_empty());
        let row = &layout.galley.rows[0];
        // Centered text on a 600px wrap should have a positive x offset.
        assert!(
            row.pos.x > 0.0,
            "centered text should be offset from left edge, got x={}",
            row.pos.x
        );
    }

    #[test]
    fn multiple_paragraphs_with_varying_styles() {
        let document = make_document(
            vec![
                TextRun {
                    text: "Left aligned\nCenter aligned\nRight aligned".to_owned(),
                    style: CharacterStyle::default(),
                },
            ],
            vec![
                ParagraphStyle {
                    alignment: ParagraphAlignment::Left,
                    ..ParagraphStyle::default()
                },
                ParagraphStyle {
                    alignment: ParagraphAlignment::Center,
                    ..ParagraphStyle::default()
                },
                ParagraphStyle {
                    alignment: ParagraphAlignment::Right,
                    ..ParagraphStyle::default()
                },
            ],
            vec![None, None, None],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 600.0);

        // Should have at least 3 rows (one per paragraph).
        assert!(
            layout.galley.rows.len() >= 3,
            "expected at least 3 rows for 3 paragraphs, got {}",
            layout.galley.rows.len()
        );
    }

    #[test]
    fn image_with_page_break_in_multi_paragraph_document() {
        let document = make_document(
            vec![TextRun {
                text: format!("Intro paragraph\n{OBJECT_REPLACEMENT_CHAR}\nClosing paragraph"),
                style: CharacterStyle::default(),
            }],
            vec![
                ParagraphStyle::default(),
                ParagraphStyle {
                    page_break_before: true,
                    ..ParagraphStyle::default()
                },
                ParagraphStyle::default(),
            ],
            vec![
                None,
                Some(DocumentImage {
                    id: 2,
                    bytes: vec![],
                    alt_text: "diagram".to_owned(),
                    width_points: 400.0,
                    height_points: 300.0,
                }),
                None,
            ],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 600.0);

        assert_eq!(layout.images.len(), 1, "should have one image");
        assert!(
            !layout.manual_page_break_rows.is_empty(),
            "should have a manual page break"
        );
        // Image row should match the page break row.
        assert_eq!(
            layout.images[0].row_index, layout.manual_page_break_rows[0],
            "image should be on the page-break row"
        );
    }

    #[test]
    fn list_markers_are_produced_for_bullet_paragraphs() {
        let document = make_document(
            vec![TextRun {
                text: "Item one\nItem two".to_owned(),
                style: CharacterStyle::default(),
            }],
            vec![
                ParagraphStyle {
                    list_kind: ListKind::Bullet,
                    ..ParagraphStyle::default()
                },
                ParagraphStyle {
                    list_kind: ListKind::Bullet,
                    ..ParagraphStyle::default()
                },
            ],
            vec![None, None],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 600.0);

        assert_eq!(
            layout.list_markers.len(),
            2,
            "should have two list markers"
        );
        assert_eq!(layout.list_markers[0].text, "•");
        assert_eq!(layout.list_markers[1].text, "•");
    }
}
