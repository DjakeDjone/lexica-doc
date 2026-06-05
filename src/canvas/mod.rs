mod editor_input;
mod image;
mod page_layout;
mod palette;
mod table;

use std::sync::Arc;

use eframe::egui::{
    self,
    epaint::text::cursor::CCursor,
    epaint::CornerRadius,
    text_selection::visuals::{paint_text_cursor, paint_text_selection},
    text_selection::CCursorRange,
    Align2, Color32, EventFilter, FontFamily, FontId, Id, Rect, Sense, Stroke, StrokeKind,
};

use crate::{
    app::{
        ActiveHeaderFooter, CanvasState, ChangeHistory, ResizeHandle, TableResizeDrag,
        TableResizeHandleRect, TableResizeKind, ThemeMode,
    },
    document::{
        text_format, CharacterStyle, DocumentImage, DocumentState, DocumentTable, HeaderFooterKind,
        LineSpacingKind, ParagraphAlignment, TextRun, WrapMode, OBJECT_REPLACEMENT_CHAR,
    },
    grammar::GrammarError,
    layout::{
        centered_page_rect, document_points_to_pixels, document_points_to_screen_points,
        fit_page_zoom, section_page_content_rect,
    },
    ui::squiggles::{paint_grammar_squiggles, ReplacementAction, SquigglePageSlice},
};

use editor_input::{apply_viewport_input, handle_keyboard_input, handle_pointer_interaction};
use image::{
    handle_image_interaction, image_body_hit, image_display_size, image_handle_hit,
    paint_image_on_page, paint_image_selection,
};
use page_layout::{layout_page_stack, PageLayout};
use palette::canvas_palette;
use table::{paint_table, table_cell_text_galley, table_row_heights_screen, TablePaintParams};

struct DocumentLayout {
    galley: Arc<egui::Galley>,
    list_markers: Vec<ListMarkerLayout>,
    images: Vec<ImageLayout>,
    tables: Vec<TableLayout>,
    manual_page_break_rows: Vec<usize>,
    paragraph_start_rows: Vec<usize>,
}

#[allow(dead_code)]
struct TableLayout {
    row_index: usize,
    size: egui::Vec2,
    table: DocumentTable,
}

struct ActiveSideWrapFlow {
    pending_top_height: f32,
    remaining_height: f32,
    text_start_x: f32,
    text_width: f32,
}

struct TightWrapZone {
    row_index: usize,
    top: f32,
    bottom: f32,
    text_start_x: f32,
    text_width: f32,
}

struct ListMarkerLayout {
    row_index: usize,
    text: String,
    x: f32,
    font_id: FontId,
    color: Color32,
}

pub(super) struct ImageLayout {
    row_index: usize,
    size: egui::Vec2,
    offset: egui::Vec2,
    image: DocumentImage,
}

#[derive(Clone, Copy, Debug, Default)]
pub struct CanvasOutput {
    pub text_changed: bool,
}

pub fn paint_document_canvas(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    theme_mode: ThemeMode,
    history: &mut ChangeHistory,
    grammar_errors: &[GrammarError],
) -> CanvasOutput {
    let mut output = CanvasOutput::default();
    let palette = canvas_palette(theme_mode);
    let viewport = ui.available_rect_before_wrap();
    let editor_id = Id::new("document_canvas");
    let response = ui.interact(viewport, editor_id, Sense::click_and_drag());
    let painter = ui.painter_at(viewport);
    let pixels_per_point = ui.ctx().pixels_per_point();
    apply_viewport_input(ui, &response, canvas);
    if canvas.zoom_mode == crate::app::ZoomMode::FitPage {
        canvas.zoom = fit_page_zoom(viewport, document.default_page_setup().page_size);
    }

    painter.rect_filled(viewport, CornerRadius::ZERO, palette.canvas_bg);

    let default_setup = document.default_page_setup();
    let base_page_rect = centered_page_rect(
        viewport,
        default_setup.page_size,
        canvas.zoom,
        egui::Vec2::ZERO,
    );
    let content_size =
        section_page_content_rect(base_page_rect, default_setup, 14.0, 14.0, canvas.zoom).size();
    let mut document_layout = layout_document(ui, document, canvas, content_size.x);
    let page_layout = layout_page_stack(
        viewport,
        document,
        canvas,
        &document_layout.galley,
        &document_layout.manual_page_break_rows,
        &document_layout.paragraph_start_rows,
    );

    if canvas.active_header_footer.is_some()
        && ui.input(|input| input.key_pressed(egui::Key::Escape))
    {
        canvas.active_header_footer = None;
    }
    if response.double_clicked() {
        if let Some(pointer_pos) = response.interact_pointer_pos() {
            if let Some(active) = header_footer_hit(&page_layout, document, canvas, pointer_pos) {
                canvas.active_header_footer = Some(active);
                canvas.active_header_footer_cursor =
                    runs_total_chars(active_header_footer_runs(document, active));
                canvas.active_header_footer_selection =
                    CCursorRange::one(CCursor::new(canvas.active_header_footer_cursor));
                canvas.active_table_cell = None;
                canvas.selected_image_id = None;
                response.request_focus();
            } else if canvas.active_header_footer.is_some()
                && page_layout
                    .pages
                    .iter()
                    .any(|page| page.content_rect.contains(pointer_pos))
            {
                canvas.active_header_footer = None;
            }
        }
    }

    let has_focus = ui.memory(|mem| mem.has_focus(editor_id));
    if has_focus {
        ui.memory_mut(|mem| {
            mem.set_focus_lock_filter(
                editor_id,
                EventFilter {
                    tab: true,
                    horizontal_arrows: true,
                    vertical_arrows: true,
                    escape: false,
                },
            );
        });
    }
    if has_focus
        && canvas.active_header_footer.is_none()
        && handle_keyboard_input(ui, document, canvas, &document_layout.galley, history)
    {
        output.text_changed = true;
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
    let mut new_table_cell_rects: Vec<(usize, usize, usize, Rect)> = Vec::new();
    let mut new_table_cell_content_rects: Vec<(usize, usize, usize, Rect)> = Vec::new();
    let mut new_table_resize_handles: Vec<TableResizeHandleRect> = Vec::new();

    for (page_index, page) in page_layout.pages.iter().enumerate() {
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

        paint_page_header_footer(
            &painter,
            document,
            canvas,
            page.page_rect,
            page.section_id,
            page.page_index_within_section,
            page.section_page_count,
            page_index + 1,
            page_layout.pages.len(),
            canvas.active_header_footer,
            palette.footer_text,
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
        let zoom = canvas.zoom;

        // Helper: compute the screen rect for an image layout entry on this page.
        let image_screen_rect = |image: &ImageLayout| -> Option<Rect> {
            let row = document_layout.galley.rows.get(image.row_index)?;
            let image_y = row.pos.y;
            if image_y < page.start_y || image_y > page.end_y {
                return None;
            }
            Some(Rect::from_min_size(
                egui::pos2(
                    page.content_rect.left()
                        + row.pos.x
                        + image.offset.x
                        + document_points_to_screen_points(image.image.offset_x_points(), zoom),
                    page.content_rect.top() + image_y - page.start_y
                        + image.offset.y
                        + document_points_to_screen_points(image.image.offset_y_points(), zoom),
                ),
                image.size,
            ))
        };

        let page_clipped_painter = painter.with_clip_rect(page.page_rect);

        // Layer 1: Paint behind-text images (below everything)
        for image in &document_layout.images {
            if image.image.wrap_mode != WrapMode::BehindText {
                continue;
            }
            let Some(image_rect) = image_screen_rect(image) else {
                continue;
            };
            paint_image_on_page(
                ui,
                canvas,
                &page_clipped_painter,
                image,
                image_rect,
                &palette,
                1.0,
            );
            new_image_rects.push((image.image.id, image_rect));
        }

        // Layer 2: Paint text galley
        painter.with_clip_rect(visible_content_rect).galley(
            galley_origin,
            document_layout.galley.clone(),
            Color32::BLACK,
        );

        // Tables
        for table_layout in &document_layout.tables {
            let Some(row) = document_layout.galley.rows.get(table_layout.row_index) else {
                continue;
            };
            let table_y = row.pos.y;
            if table_y < page.start_y || table_y > page.end_y {
                continue;
            }

            let table_origin = egui::pos2(
                page.content_rect.left() + row.pos.x,
                page.content_rect.top() + table_y - page.start_y,
            );

            let active_cell = canvas.active_table_cell;
            let geometry = paint_table(
                ui,
                canvas,
                &page_clipped_painter,
                &table_layout.table,
                TablePaintParams {
                    origin: table_origin,
                    zoom,
                    active_cell,
                    time: ui.input(|i| i.time),
                },
            );
            new_table_cell_rects.extend(geometry.cell_rects);
            new_table_cell_content_rects.extend(geometry.cell_content_rects);
            new_table_resize_handles.extend(geometry.resize_handles);
        }

        // List markers
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

        // Layer 3: Paint normal images (not behind-text, not in-front-of-text) sorted by z-index
        let mut normal_images: Vec<&ImageLayout> = document_layout
            .images
            .iter()
            .filter(|img| !img.image.wrap_mode.is_no_text_displacement())
            .collect();
        normal_images.sort_by_key(|img| img.image.z_index);

        for image in &normal_images {
            let Some(image_rect) = image_screen_rect(image) else {
                continue;
            };
            paint_image_on_page(
                ui,
                canvas,
                &page_clipped_painter,
                image,
                image_rect,
                &palette,
                1.0,
            );
            new_image_rects.push((image.image.id, image_rect));
        }

        // Layer 4: Paint in-front-of-text images (above everything)
        let mut front_images: Vec<&ImageLayout> = document_layout
            .images
            .iter()
            .filter(|img| img.image.wrap_mode == WrapMode::InFrontOfText)
            .collect();
        front_images.sort_by_key(|img| img.image.z_index);

        for image in &front_images {
            let Some(image_rect) = image_screen_rect(image) else {
                continue;
            };
            paint_image_on_page(
                ui,
                canvas,
                &page_clipped_painter,
                image,
                image_rect,
                &palette,
                1.0,
            );
            new_image_rects.push((image.image.id, image_rect));
        }
    }

    if paint_active_header_footer_editor(
        ui,
        document,
        canvas,
        history,
        &page_layout,
        &response,
        editor_id,
    ) {
        output.text_changed = true;
    }

    let squiggle_pages: Vec<SquigglePageSlice> = page_layout
        .pages
        .iter()
        .map(|page| SquigglePageSlice {
            content_rect: page.content_rect,
            start_y: page.start_y,
            end_y: page.end_y,
        })
        .collect();
    let pending_replacement = paint_grammar_squiggles(
        ui,
        &painter,
        &document_layout.galley,
        &squiggle_pages,
        grammar_errors,
    );

    canvas.image_rects = new_image_rects;
    canvas.table_cell_rects = new_table_cell_rects;
    canvas.table_cell_content_rects = new_table_cell_content_rects;
    canvas.table_resize_handles = new_table_resize_handles;

    let (table_pointer_captured, table_document_changed) = if canvas.active_header_footer.is_some()
    {
        (true, false)
    } else {
        handle_table_interaction(ui, &response, canvas, document, history)
    };
    output.text_changed |= table_document_changed;

    let (image_pointer_captured, image_document_changed) = if table_pointer_captured {
        (true, false)
    } else {
        handle_image_interaction(ui, &response, canvas, document, history)
    };
    output.text_changed |= image_document_changed;

    if !table_pointer_captured && !image_pointer_captured && canvas.active_header_footer.is_none() {
        handle_pointer_interaction(
            ui,
            &response,
            &page_layout,
            &document_layout.galley,
            canvas,
            document,
        );
    }
    update_canvas_hover_cursor(ui, &response, canvas, &page_layout);

    // Draw ghost image if dragging
    if let Some(move_drag) = &canvas.move_drag {
        if move_drag.current_ptr != move_drag.start_ptr {
            let offset = move_drag.current_ptr - move_drag.start_ptr;
            let ghost_rect = move_drag.start_rect.translate(offset);
            if let Some(image_layout) = document_layout
                .images
                .iter()
                .find(|i| i.image.id == move_drag.image_id)
            {
                paint_image_on_page(
                    ui,
                    canvas,
                    &painter,
                    image_layout,
                    ghost_rect,
                    &palette,
                    0.5,
                );
            }
        }
    }

    // Draw selection border + handles with unclipped painter so they aren't cut at page margins
    if let Some((_, selected_rect)) = canvas
        .image_rects
        .iter()
        .find(|(id, _)| Some(*id) == canvas.selected_image_id)
    {
        paint_image_selection(&painter, *selected_rect);
    }

    if has_focus
        && canvas.selected_image_id.is_none()
        && canvas.active_table_cell.is_none()
        && canvas.active_header_footer.is_none()
    {
        if let Some(caret_rect) = page_layout.caret_rect(
            &document_layout.galley,
            canvas.selection.primary,
            caret_height(canvas.active_style, canvas.zoom),
        ) {
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

    if let Some(replacement) = pending_replacement {
        if apply_grammar_replacement(document, canvas, history, ui, replacement) {
            output.text_changed = true;
        }
    }

    output
}

fn paint_page_header_footer(
    painter: &egui::Painter,
    document: &DocumentState,
    canvas: &CanvasState,
    page_rect: Rect,
    section_id: crate::document::SectionId,
    page_index_within_section: usize,
    section_page_count: usize,
    page_number: usize,
    page_count: usize,
    active_header_footer: Option<ActiveHeaderFooter>,
    color: Color32,
) {
    let setup = document
        .section_by_id(section_id)
        .map(|section| section.page_setup)
        .unwrap_or_else(|| document.default_page_setup());
    let horizontal_margin =
        document_points_to_screen_points(setup.margins.left_points.max(18.0), canvas.zoom);
    let text_width = (page_rect.width() - horizontal_margin * 2.0).max(1.0);

    let header_variant = document.header_footer_variant_for_page(
        section_id,
        page_index_within_section,
        HeaderFooterKind::Header,
    );
    let header_story =
        document.resolve_header_footer_slot(section_id, HeaderFooterKind::Header, header_variant);
    let header_runs = rendered_header_footer_runs(
        document,
        &header_story.story.runs,
        section_id,
        page_index_within_section,
        page_number,
        page_count,
        section_page_count,
    );
    if !runs_plain_text(&header_runs).trim().is_empty()
        && active_header_footer
            != Some(ActiveHeaderFooter {
                kind: HeaderFooterKind::Header,
                section_id,
                variant: header_variant,
                page_number,
            })
    {
        let font_size = header_footer_base_font_size(&header_runs, canvas.zoom);
        let y = page_rect.top()
            + document_points_to_screen_points(setup.header_from_top_points, canvas.zoom);
        paint_tab_aligned_margin_runs(
            painter,
            &header_runs,
            canvas.zoom,
            color,
            Rect::from_min_size(
                egui::pos2(page_rect.left() + horizontal_margin, y),
                egui::vec2(text_width, font_size),
            ),
            None,
        );
    }

    let footer_variant = document.header_footer_variant_for_page(
        section_id,
        page_index_within_section,
        HeaderFooterKind::Footer,
    );
    let footer_story =
        document.resolve_header_footer_slot(section_id, HeaderFooterKind::Footer, footer_variant);
    let footer_runs = rendered_header_footer_runs(
        document,
        &footer_story.story.runs,
        section_id,
        page_index_within_section,
        page_number,
        page_count,
        section_page_count,
    );
    if !runs_plain_text(&footer_runs).trim().is_empty()
        && active_header_footer
            != Some(ActiveHeaderFooter {
                kind: HeaderFooterKind::Footer,
                section_id,
                variant: footer_variant,
                page_number,
            })
    {
        let font_size = header_footer_base_font_size(&footer_runs, canvas.zoom);
        let y = page_rect.bottom()
            - document_points_to_screen_points(setup.footer_from_bottom_points, canvas.zoom)
            - font_size;
        paint_tab_aligned_margin_runs(
            painter,
            &footer_runs,
            canvas.zoom,
            color,
            Rect::from_min_size(
                egui::pos2(page_rect.left() + horizontal_margin, y),
                egui::vec2(text_width, font_size),
            ),
            None,
        );
    }
}

fn rendered_header_footer_runs(
    document: &DocumentState,
    runs: &[TextRun],
    section_id: crate::document::SectionId,
    page_index_within_section: usize,
    page_number: usize,
    page_count: usize,
    section_page_count: usize,
) -> Vec<TextRun> {
    runs.iter()
        .map(|run| TextRun {
            text: document.render_page_field_for_section_page(
                &run.text,
                section_id,
                page_index_within_section,
                page_number.saturating_sub(1),
                page_count,
                section_page_count,
            ),
            style: run.style,
        })
        .collect()
}

fn paint_tab_aligned_margin_runs(
    painter: &egui::Painter,
    runs: &[TextRun],
    zoom: f32,
    fallback_color: Color32,
    rect: Rect,
    selection: Option<std::ops::Range<usize>>,
) {
    let segments = split_runs_for_header_tabs(runs);
    for slot in 0..3 {
        let Some(segment) = segments.get(slot) else {
            continue;
        };
        if segment.runs.is_empty() {
            continue;
        }
        let segment_width = measure_runs_width(painter, &segment.runs, zoom);
        let mut x = match slot {
            0 => rect.left(),
            1 => rect.center().x - segment_width * 0.5,
            _ => rect.right() - segment_width,
        };
        for piece in &segment.runs {
            let Some(piece_range) = piece.range.clone() else {
                continue;
            };
            if let Some(selection) = &selection {
                let start = piece_range.start.max(selection.start);
                let end = piece_range.end.min(selection.end);
                if start < end {
                    let before = slice_run_text_chars(&piece.text, 0..start - piece_range.start);
                    let selected = slice_run_text_chars(
                        &piece.text,
                        start - piece_range.start..end - piece_range.start,
                    );
                    let selected_x = x + measure_text_width(painter, &before, piece.style, zoom);
                    let selected_width = measure_text_width(painter, &selected, piece.style, zoom);
                    painter.rect_filled(
                        Rect::from_min_size(
                            egui::pos2(selected_x, rect.top()),
                            egui::vec2(
                                selected_width,
                                header_footer_line_height(piece.runs_style(), zoom),
                            ),
                        ),
                        CornerRadius::ZERO,
                        Color32::from_rgba_unmultiplied(80, 135, 230, 80),
                    );
                }
            }
            paint_run_text(
                painter,
                &piece.text,
                piece.style,
                zoom,
                egui::pos2(x, rect.top()),
                fallback_color,
            );
            x += measure_text_width(painter, &piece.text, piece.style, zoom);
        }
    }
}

#[derive(Clone)]
struct HeaderRunPiece {
    text: String,
    style: CharacterStyle,
    range: Option<std::ops::Range<usize>>,
}

impl HeaderRunPiece {
    fn runs_style(&self) -> CharacterStyle {
        self.style
    }
}

struct HeaderSegment {
    runs: Vec<HeaderRunPiece>,
    end: usize,
}

fn split_runs_for_header_tabs(runs: &[TextRun]) -> Vec<HeaderSegment> {
    let mut segments = vec![HeaderSegment {
        runs: Vec::new(),
        end: 0,
    }];
    let mut slot = 0usize;
    let mut char_index = 0usize;

    for run in runs {
        let mut text = String::new();
        let mut piece_start = char_index;
        for ch in run.text.chars() {
            if ch == '\t' && slot < 2 {
                if !text.is_empty() {
                    segments[slot].runs.push(HeaderRunPiece {
                        text: std::mem::take(&mut text),
                        style: run.style,
                        range: Some(piece_start..char_index),
                    });
                }
                segments[slot].end = char_index;
                slot += 1;
                char_index += 1;
                piece_start = char_index;
                segments.push(HeaderSegment {
                    runs: Vec::new(),
                    end: char_index,
                });
            } else {
                text.push(if ch == '\t' { ' ' } else { ch });
                char_index += 1;
            }
        }
        if !text.is_empty() {
            segments[slot].runs.push(HeaderRunPiece {
                text,
                style: run.style,
                range: Some(piece_start..char_index),
            });
        }
    }
    if let Some(segment) = segments.get_mut(slot) {
        segment.end = char_index;
    }
    segments
}

fn measure_runs_width(painter: &egui::Painter, runs: &[HeaderRunPiece], zoom: f32) -> f32 {
    runs.iter()
        .map(|run| measure_text_width(painter, &run.text, run.style, zoom))
        .sum()
}

fn measure_text_width(
    painter: &egui::Painter,
    text: &str,
    style: CharacterStyle,
    zoom: f32,
) -> f32 {
    if text.is_empty() {
        return 0.0;
    }
    header_footer_text_galley(painter, text, style, zoom, Color32::BLACK)
        .size()
        .x
}

fn paint_run_text(
    painter: &egui::Painter,
    text: &str,
    style: CharacterStyle,
    zoom: f32,
    pos: egui::Pos2,
    fallback_color: Color32,
) {
    if text.is_empty() {
        return;
    }
    painter.galley(
        pos,
        header_footer_text_galley(painter, text, style, zoom, fallback_color),
        fallback_color,
    );
}

fn header_footer_text_galley(
    painter: &egui::Painter,
    text: &str,
    style: CharacterStyle,
    zoom: f32,
    fallback_color: Color32,
) -> std::sync::Arc<egui::Galley> {
    let mut format = text_format(style, zoom);
    if format.color == Color32::TRANSPARENT {
        format.color = fallback_color;
    }
    let mut job = egui::epaint::text::LayoutJob::default();
    job.append(text, 0.0, format);
    painter.layout_job(job)
}

fn header_footer_line_height(style: CharacterStyle, zoom: f32) -> f32 {
    (style.font_size_points * zoom).max(1.0) * 1.2
}

fn header_footer_base_font_size(runs: &[TextRun], zoom: f32) -> f32 {
    runs.iter()
        .find(|run| !run.text.is_empty())
        .map(|run| run.style.font_size_points * zoom)
        .unwrap_or(9.0 * zoom)
        .clamp(7.0, 28.0)
}

fn header_footer_hit(
    page_layout: &PageLayout,
    document: &DocumentState,
    canvas: &CanvasState,
    pointer_pos: egui::Pos2,
) -> Option<ActiveHeaderFooter> {
    page_layout
        .pages
        .iter()
        .enumerate()
        .find_map(|(index, page)| {
            let page_number = index + 1;
            let header_variant = document.header_footer_variant_for_page(
                page.section_id,
                page.page_index_within_section,
                HeaderFooterKind::Header,
            );
            let footer_variant = document.header_footer_variant_for_page(
                page.section_id,
                page.page_index_within_section,
                HeaderFooterKind::Footer,
            );
            if page_header_rect(page.page_rect, document, canvas, page.section_id)
                .contains(pointer_pos)
            {
                Some(ActiveHeaderFooter {
                    kind: HeaderFooterKind::Header,
                    section_id: page.section_id,
                    variant: header_variant,
                    page_number,
                })
            } else if page_footer_rect(page.page_rect, document, canvas, page.section_id)
                .contains(pointer_pos)
            {
                Some(ActiveHeaderFooter {
                    kind: HeaderFooterKind::Footer,
                    section_id: page.section_id,
                    variant: footer_variant,
                    page_number,
                })
            } else {
                None
            }
        })
}

fn paint_active_header_footer_editor(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    page_layout: &PageLayout,
    response: &egui::Response,
    editor_id: Id,
) -> bool {
    let Some(active) = canvas.active_header_footer else {
        return false;
    };
    let Some(page) = page_layout.pages.get(active.page_number.saturating_sub(1)) else {
        canvas.active_header_footer = None;
        return false;
    };

    let margin_rect = match active.kind {
        HeaderFooterKind::Header => {
            page_header_rect(page.page_rect, document, canvas, active.section_id)
        }
        HeaderFooterKind::Footer => {
            page_footer_rect(page.page_rect, document, canvas, active.section_id)
        }
    };
    let horizontal_margin = document_points_to_screen_points(
        document
            .section_by_id(active.section_id)
            .map(|section| section.page_setup.margins.left_points)
            .unwrap_or_else(|| document.default_page_setup().margins.left_points)
            .max(18.0),
        canvas.zoom,
    );
    let editor_height = document_points_to_screen_points(20.0, canvas.zoom).clamp(18.0, 28.0);
    let editor_rect = Rect::from_center_size(
        margin_rect.center(),
        egui::vec2(
            (margin_rect.width() - horizontal_margin * 2.0).max(80.0),
            editor_height,
        ),
    );

    let guide_y = match active.kind {
        HeaderFooterKind::Header => margin_rect.bottom(),
        HeaderFooterKind::Footer => margin_rect.top(),
    };
    ui.painter().line_segment(
        [
            egui::pos2(editor_rect.left(), guide_y),
            egui::pos2(editor_rect.right(), guide_y),
        ],
        Stroke::new(1.0, Color32::from_rgba_unmultiplied(120, 130, 145, 110)),
    );

    let active_runs = active_header_footer_runs(document, active).to_vec();

    if let Some(pointer_pos) = response.interact_pointer_pos() {
        let press_origin = ui.input(|i| i.pointer.press_origin());
        let interact_header = editor_rect.expand(20.0).contains(pointer_pos)
            || press_origin.map_or(false, |origin| editor_rect.expand(20.0).contains(origin));

        if interact_header {
            if response.clicked() || response.drag_started() {
                response.request_focus();
            }
            let cursor = CCursor::new(header_footer_cursor_from_pos(
                ui.painter(),
                &active_runs,
                canvas.zoom,
                editor_rect,
                pointer_pos.x,
            ));
            if response.drag_started() {
                if ui.input(|input| input.modifiers.shift) {
                    canvas.active_header_footer_selection.primary = cursor;
                } else {
                    canvas.active_header_footer_selection = CCursorRange::one(cursor);
                }
                canvas.active_header_footer_cursor = cursor.index;
                canvas.active_header_footer_selection.h_pos = None;
                canvas.last_interaction_time = ui.input(|input| input.time);
            } else if response.dragged() {
                canvas.active_header_footer_selection.primary = cursor;
                canvas.active_header_footer_cursor = cursor.index;
                canvas.active_header_footer_selection.h_pos = None;
                canvas.last_interaction_time = ui.input(|input| input.time);
            } else if response.clicked() {
                if ui.input(|input| input.modifiers.shift) {
                    canvas.active_header_footer_selection.primary = cursor;
                } else {
                    canvas.active_header_footer_selection = CCursorRange::one(cursor);
                }
                canvas.active_header_footer_cursor = cursor.index;
                canvas.active_header_footer_selection.h_pos = None;
                canvas.last_interaction_time = ui.input(|input| input.time);
            }
        }
    }

    let has_focus = ui.memory(|memory| memory.has_focus(editor_id));

    let mut edited_runs = active_runs;
    let total_chars = runs_total_chars(&edited_runs);
    canvas.active_header_footer_cursor = canvas.active_header_footer_cursor.min(total_chars);
    canvas.active_header_footer_selection.primary.index = canvas
        .active_header_footer_selection
        .primary
        .index
        .min(total_chars);
    canvas.active_header_footer_selection.secondary.index = canvas
        .active_header_footer_selection
        .secondary
        .index
        .min(total_chars);
    let before_runs = edited_runs.clone();
    let changed = if has_focus {
        handle_header_footer_keyboard_input(ui, &mut edited_runs, canvas, history, document)
    } else {
        false
    };
    if runs_plain_text(&edited_runs).trim().is_empty() {
        let hint = match active.kind {
            HeaderFooterKind::Header => "Header",
            HeaderFooterKind::Footer => "Footer",
        };
        let inherited = document
            .resolve_header_footer_slot(active.section_id, active.kind, active.variant)
            .inherited;
        let section_label = format!(
            "{hint} - Section {}{}",
            active.section_id,
            if inherited { " (Same as Previous)" } else { "" }
        );
        ui.painter().text(
            editor_rect.left_top(),
            Align2::LEFT_TOP,
            section_label,
            FontId::new(editor_height * 0.72, FontFamily::Proportional),
            Color32::from_rgba_premultiplied(96, 104, 118, 140),
        );
    } else {
        let selection = has_focus
            .then(|| canvas.active_header_footer_selection.as_sorted_char_range())
            .filter(|range| range.start < range.end);
        paint_tab_aligned_margin_runs(
            ui.painter(),
            &edited_runs,
            canvas.zoom,
            Color32::from_rgb(36, 39, 46),
            editor_rect,
            selection,
        );
    }

    if has_focus {
        let cursor_pos = header_footer_cursor_pos(
            ui.painter(),
            &edited_runs,
            canvas.zoom,
            editor_rect,
            canvas.active_header_footer_cursor,
        );
        let time = ui.input(|input| input.time) - canvas.last_interaction_time;
        paint_text_cursor(
            ui,
            ui.painter(),
            Rect::from_min_size(cursor_pos, egui::vec2(1.5, editor_height * 0.85)),
            time,
        );
    }

    if changed && edited_runs != before_runs {
        normalize_header_footer_runs(&mut edited_runs);
        *active_header_footer_runs_mut(document, active) = edited_runs;
        document.sync_compat_from_first_section();
        canvas.active_style = header_footer_style_at(
            active_header_footer_runs(document, active),
            canvas.active_header_footer_selection.primary.index,
        );
        true
    } else {
        false
    }
}

fn handle_header_footer_keyboard_input(
    ui: &egui::Ui,
    runs: &mut Vec<TextRun>,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    document: &DocumentState,
) -> bool {
    let mut changed = false;
    let events = ui.input(|input| input.events.clone());
    for event in events {
        match event {
            egui::Event::Copy => {
                copy_header_footer_selection(ui, runs, canvas);
            }
            egui::Event::Cut => {
                if copy_header_footer_selection(ui, runs, canvas) {
                    history.checkpoint(document, ui.input(|input| input.time));
                    let selected = canvas.active_header_footer_selection.as_sorted_char_range();
                    delete_header_footer_range(runs, selected.clone());
                    canvas.active_header_footer_selection =
                        CCursorRange::one(CCursor::new(selected.start));
                    canvas.active_header_footer_cursor = selected.start;
                    changed = true;
                }
            }
            egui::Event::Text(inserted) if !inserted.is_empty() => {
                history.checkpoint_coalesced(document, ui.input(|input| input.time));
                let selected = canvas.active_header_footer_selection.as_sorted_char_range();
                let next = replace_header_footer_range_with_text(
                    runs,
                    selected,
                    &inserted,
                    canvas.active_style,
                );
                canvas.active_header_footer_selection = CCursorRange::one(CCursor::new(next));
                canvas.active_header_footer_cursor = next;
                changed = true;
            }
            egui::Event::Paste(pasted) => {
                history.checkpoint(document, ui.input(|input| input.time));
                let selected = canvas.active_header_footer_selection.as_sorted_char_range();
                let next = replace_header_footer_range_with_text(
                    runs,
                    selected,
                    &pasted,
                    canvas.active_style,
                );
                canvas.active_header_footer_selection = CCursorRange::one(CCursor::new(next));
                canvas.active_header_footer_cursor = next;
                changed = true;
            }
            egui::Event::Key {
                key: egui::Key::Tab,
                pressed: true,
                modifiers,
                ..
            } => {
                history.checkpoint(document, ui.input(|input| input.time));
                let next_cursor = if modifiers.shift {
                    remove_previous_header_footer_tab(
                        runs,
                        canvas.active_header_footer_selection.as_sorted_char_range(),
                    )
                } else {
                    insert_header_footer_tab(
                        runs,
                        canvas.active_header_footer_selection.as_sorted_char_range(),
                        canvas.active_style,
                    )
                };
                if let Some(next_cursor) = next_cursor {
                    canvas.active_header_footer_selection =
                        CCursorRange::one(CCursor::new(next_cursor));
                    canvas.active_header_footer_cursor = next_cursor;
                    changed = true;
                }
            }
            egui::Event::Key {
                key: egui::Key::Backspace,
                pressed: true,
                ..
            } => {
                let selected = canvas.active_header_footer_selection.as_sorted_char_range();
                if selected.start < selected.end {
                    history.checkpoint_coalesced(document, ui.input(|input| input.time));
                    delete_header_footer_range(runs, selected.clone());
                    canvas.active_header_footer_selection =
                        CCursorRange::one(CCursor::new(selected.start));
                    canvas.active_header_footer_cursor = selected.start;
                    changed = true;
                } else if canvas.active_header_footer_cursor > 0 {
                    history.checkpoint_coalesced(document, ui.input(|input| input.time));
                    let start = canvas.active_header_footer_cursor - 1;
                    delete_header_footer_range(runs, start..canvas.active_header_footer_cursor);
                    canvas.active_header_footer_selection = CCursorRange::one(CCursor::new(start));
                    canvas.active_header_footer_cursor = start;
                    changed = true;
                }
            }
            egui::Event::Key {
                key: egui::Key::Delete,
                pressed: true,
                ..
            } => {
                let selected = canvas.active_header_footer_selection.as_sorted_char_range();
                if selected.start < selected.end {
                    history.checkpoint_coalesced(document, ui.input(|input| input.time));
                    delete_header_footer_range(runs, selected.clone());
                    canvas.active_header_footer_selection =
                        CCursorRange::one(CCursor::new(selected.start));
                    canvas.active_header_footer_cursor = selected.start;
                    changed = true;
                } else {
                    let total_chars = runs_total_chars(runs);
                    if canvas.active_header_footer_cursor < total_chars {
                        history.checkpoint_coalesced(document, ui.input(|input| input.time));
                        delete_header_footer_range(
                            runs,
                            canvas.active_header_footer_cursor
                                ..canvas.active_header_footer_cursor + 1,
                        );
                        changed = true;
                    }
                }
            }
            egui::Event::Key {
                key,
                pressed: true,
                modifiers,
                ..
            } => {
                if handle_header_footer_shortcut_key(
                    ui, document, canvas, history, runs, key, modifiers,
                ) {
                    changed = true;
                    continue;
                }
                match key {
                    egui::Key::ArrowLeft => {
                        let next = canvas.active_header_footer_cursor.saturating_sub(1);
                        set_header_footer_cursor(canvas, next, modifiers.shift);
                    }
                    egui::Key::ArrowRight => {
                        let next =
                            (canvas.active_header_footer_cursor + 1).min(runs_total_chars(runs));
                        set_header_footer_cursor(canvas, next, modifiers.shift);
                    }
                    egui::Key::Home => {
                        set_header_footer_cursor(canvas, 0, modifiers.shift);
                    }
                    egui::Key::End => {
                        set_header_footer_cursor(canvas, runs_total_chars(runs), modifiers.shift);
                    }
                    _ => {}
                }
            }
            _ => {}
        }
    }
    changed
}

fn handle_header_footer_shortcut_key(
    ui: &egui::Ui,
    document: &DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    runs: &mut Vec<TextRun>,
    key: egui::Key,
    modifiers: egui::Modifiers,
) -> bool {
    if !modifiers.command {
        return false;
    }

    match key {
        egui::Key::A => {
            let total_chars = runs_total_chars(runs);
            canvas.active_header_footer_selection =
                CCursorRange::two(CCursor::new(0), CCursor::new(total_chars));
            canvas.active_header_footer_cursor = total_chars;
            false
        }
        egui::Key::B => {
            history.checkpoint(document, ui.input(|input| input.time));
            let next = !canvas.active_style.bold;
            apply_header_footer_style_change(runs, canvas, |style| style.bold = next)
        }
        egui::Key::I => {
            history.checkpoint(document, ui.input(|input| input.time));
            let next = !canvas.active_style.italic;
            apply_header_footer_style_change(runs, canvas, |style| style.italic = next)
        }
        egui::Key::U => {
            history.checkpoint(document, ui.input(|input| input.time));
            let next = !canvas.active_style.underline;
            apply_header_footer_style_change(runs, canvas, |style| style.underline = next)
        }
        _ => false,
    }
}

fn apply_header_footer_style_change(
    runs: &mut Vec<TextRun>,
    canvas: &mut CanvasState,
    mutate: impl Fn(&mut CharacterStyle) + Copy,
) -> bool {
    let selected = canvas.active_header_footer_selection.as_sorted_char_range();
    let changed = if selected.start < selected.end {
        apply_style_to_header_footer_range(runs, selected, mutate);
        true
    } else {
        false
    };
    mutate(&mut canvas.active_style);
    changed
}

fn set_header_footer_cursor(canvas: &mut CanvasState, cursor: usize, extend_selection: bool) {
    canvas.active_header_footer_cursor = cursor;
    if extend_selection {
        canvas.active_header_footer_selection.primary = CCursor::new(cursor);
    } else {
        canvas.active_header_footer_selection = CCursorRange::one(CCursor::new(cursor));
    }
}

fn copy_header_footer_selection(ui: &egui::Ui, runs: &[TextRun], canvas: &CanvasState) -> bool {
    let selected = canvas.active_header_footer_selection.as_sorted_char_range();
    if selected.start >= selected.end {
        return false;
    }
    ui.copy_text(selected_header_footer_text(runs, selected));
    true
}

fn selected_header_footer_text(runs: &[TextRun], range: std::ops::Range<usize>) -> String {
    runs_plain_text(runs)
        .chars()
        .skip(range.start)
        .take(range.end.saturating_sub(range.start))
        .collect()
}

fn insert_header_footer_tab(
    runs: &mut Vec<TextRun>,
    range: std::ops::Range<usize>,
    style: CharacterStyle,
) -> Option<usize> {
    let plain = runs_plain_text(runs);
    let tab_count_before_cursor = plain
        .chars()
        .take(range.start)
        .filter(|ch| *ch == '\t')
        .count();
    if tab_count_before_cursor >= 2 {
        return None;
    }

    Some(replace_header_footer_range_with_text(
        runs, range, "\t", style,
    ))
}

fn remove_previous_header_footer_tab(
    runs: &mut Vec<TextRun>,
    range: std::ops::Range<usize>,
) -> Option<usize> {
    if range.start < range.end {
        delete_header_footer_range(runs, range.clone());
        return Some(range.start);
    }

    let plain = runs_plain_text(runs);
    let previous_tab = plain
        .chars()
        .take(range.start)
        .enumerate()
        .filter_map(|(index, ch)| (ch == '\t').then_some(index))
        .last()?;
    delete_header_footer_range(runs, previous_tab..previous_tab + 1);
    Some(previous_tab)
}

fn char_to_byte_index_for_header_footer(text: &str, char_index: usize) -> usize {
    text.char_indices()
        .nth(char_index)
        .map(|(byte_index, _)| byte_index)
        .unwrap_or(text.len())
}

fn runs_plain_text(runs: &[TextRun]) -> String {
    runs.iter().map(|run| run.text.as_str()).collect()
}

fn runs_total_chars(runs: &[TextRun]) -> usize {
    runs.iter().map(|run| run.text.chars().count()).sum()
}

fn slice_run_text_chars(text: &str, range: std::ops::Range<usize>) -> String {
    text.chars()
        .skip(range.start)
        .take(range.end.saturating_sub(range.start))
        .collect()
}

fn header_footer_style_at(runs: &[TextRun], char_index: usize) -> CharacterStyle {
    let total = runs_total_chars(runs);
    let target = char_index.min(total);
    let mut offset = 0usize;
    for run in runs {
        let len = run.text.chars().count();
        if target < offset + len {
            return run.style;
        }
        offset += len;
    }
    runs.last().map(|run| run.style).unwrap_or_default()
}

fn normalize_header_footer_runs(runs: &mut Vec<TextRun>) {
    runs.retain(|run| !run.text.is_empty());
    let mut normalized: Vec<TextRun> = Vec::with_capacity(runs.len().max(1));
    for run in runs.drain(..) {
        if let Some(last) = normalized.last_mut() {
            if last.style == run.style {
                last.text.push_str(&run.text);
                continue;
            }
        }
        normalized.push(run);
    }
    if normalized.is_empty() {
        normalized.push(TextRun {
            text: String::new(),
            style: CharacterStyle::default(),
        });
    }
    *runs = normalized;
}

fn split_header_footer_runs_at(runs: &mut Vec<TextRun>, char_index: usize) {
    if char_index == 0 || char_index >= runs_total_chars(runs) {
        return;
    }

    let mut offset = 0usize;
    for idx in 0..runs.len() {
        let len = runs[idx].text.chars().count();
        if char_index > offset && char_index < offset + len {
            let local = char_index - offset;
            let byte_index = char_to_byte_index_for_header_footer(&runs[idx].text, local);
            let right = runs[idx].text.split_off(byte_index);
            let style = runs[idx].style;
            runs.insert(idx + 1, TextRun { text: right, style });
            break;
        }
        offset += len;
    }
}

fn replace_header_footer_range_with_text(
    runs: &mut Vec<TextRun>,
    range: std::ops::Range<usize>,
    text: &str,
    style: CharacterStyle,
) -> usize {
    let start = range.start.min(runs_total_chars(runs));
    let end = range.end.min(runs_total_chars(runs));
    delete_header_footer_range(runs, start..end);
    insert_header_footer_text(runs, start, text, style);
    start + text.chars().count()
}

fn insert_header_footer_text(
    runs: &mut Vec<TextRun>,
    char_index: usize,
    text: &str,
    style: CharacterStyle,
) {
    if text.is_empty() {
        return;
    }
    let insertion_index = char_index.min(runs_total_chars(runs));
    split_header_footer_runs_at(runs, insertion_index);

    let mut offset = 0usize;
    let mut target = runs.len();
    for (idx, run) in runs.iter().enumerate() {
        if offset == insertion_index {
            target = idx;
            break;
        }
        offset += run.text.chars().count();
    }
    runs.insert(
        target,
        TextRun {
            text: text.to_owned(),
            style,
        },
    );
    normalize_header_footer_runs(runs);
}

fn delete_header_footer_range(runs: &mut Vec<TextRun>, range: std::ops::Range<usize>) {
    if range.start >= range.end {
        return;
    }
    let start = range.start.min(runs_total_chars(runs));
    let end = range.end.min(runs_total_chars(runs));
    split_header_footer_runs_at(runs, start);
    split_header_footer_runs_at(runs, end);

    let mut offset = 0usize;
    runs.retain(|run| {
        let len = run.text.chars().count();
        let keep = offset + len <= start || offset >= end;
        offset += len;
        keep
    });
    normalize_header_footer_runs(runs);
}

fn apply_style_to_header_footer_range(
    runs: &mut Vec<TextRun>,
    range: std::ops::Range<usize>,
    mutate: impl Fn(&mut CharacterStyle) + Copy,
) {
    let start = range.start.min(runs_total_chars(runs));
    let end = range.end.min(runs_total_chars(runs));
    if start >= end {
        return;
    }
    split_header_footer_runs_at(runs, start);
    split_header_footer_runs_at(runs, end);

    let mut offset = 0usize;
    for run in runs.iter_mut() {
        let len = run.text.chars().count();
        if offset >= start && offset + len <= end {
            mutate(&mut run.style);
        }
        offset += len;
    }
    normalize_header_footer_runs(runs);
}

fn header_footer_cursor_pos(
    painter: &egui::Painter,
    runs: &[TextRun],
    zoom: f32,
    rect: Rect,
    cursor: usize,
) -> egui::Pos2 {
    let segments = split_runs_for_header_tabs(runs);
    let mut slot = 0usize;
    for (idx, segment) in segments.iter().enumerate() {
        if cursor <= segment.end || idx + 1 == segments.len() {
            slot = idx;
            break;
        }
    }
    let segment = segments.get(slot);
    let full_width = segment
        .map(|segment| measure_runs_width(painter, &segment.runs, zoom))
        .unwrap_or(0.0);
    let prefix_width = segment
        .map(|segment| measure_segment_prefix_width(painter, segment, cursor, zoom))
        .unwrap_or(0.0);
    let segment_left = match slot {
        0 => rect.left(),
        1 => rect.center().x - full_width * 0.5,
        _ => rect.right() - full_width,
    };
    egui::pos2(segment_left + prefix_width, rect.top())
}

fn header_footer_cursor_from_pos(
    painter: &egui::Painter,
    runs: &[TextRun],
    zoom: f32,
    rect: Rect,
    x: f32,
) -> usize {
    let total_chars = runs_total_chars(runs);
    (0..=total_chars)
        .min_by(|left, right| {
            let left_x = header_footer_cursor_pos(painter, runs, zoom, rect, *left).x;
            let right_x = header_footer_cursor_pos(painter, runs, zoom, rect, *right).x;
            (left_x - x)
                .abs()
                .partial_cmp(&(right_x - x).abs())
                .unwrap_or(std::cmp::Ordering::Equal)
        })
        .unwrap_or(total_chars)
}

fn measure_segment_prefix_width(
    painter: &egui::Painter,
    segment: &HeaderSegment,
    cursor: usize,
    zoom: f32,
) -> f32 {
    let mut width = 0.0;
    for piece in &segment.runs {
        let Some(range) = piece.range.clone() else {
            continue;
        };
        if cursor >= range.end {
            width += measure_text_width(painter, &piece.text, piece.style, zoom);
        } else if cursor > range.start {
            let prefix = slice_run_text_chars(&piece.text, 0..cursor - range.start);
            width += measure_text_width(painter, &prefix, piece.style, zoom);
            break;
        } else {
            break;
        }
    }
    width
}

fn active_header_footer_runs(document: &DocumentState, active: ActiveHeaderFooter) -> &[TextRun] {
    document
        .resolve_header_footer_slot(active.section_id, active.kind, active.variant)
        .story
        .runs
        .as_slice()
}

fn active_header_footer_runs_mut(
    document: &mut DocumentState,
    active: ActiveHeaderFooter,
) -> &mut Vec<TextRun> {
    &mut document
        .header_footer_story_mut_materialized(active.section_id, active.kind, active.variant)
        .expect("active header/footer section should exist")
        .runs
}

fn page_header_rect(
    page_rect: Rect,
    document: &DocumentState,
    canvas: &CanvasState,
    section_id: crate::document::SectionId,
) -> Rect {
    let setup = document
        .section_by_id(section_id)
        .map(|section| section.page_setup)
        .unwrap_or_else(|| document.default_page_setup());
    let height = document_points_to_screen_points(setup.margins.top_points, canvas.zoom)
        .clamp(18.0, page_rect.height() * 0.25);
    Rect::from_min_size(page_rect.min, egui::vec2(page_rect.width(), height))
}

fn page_footer_rect(
    page_rect: Rect,
    document: &DocumentState,
    canvas: &CanvasState,
    section_id: crate::document::SectionId,
) -> Rect {
    let setup = document
        .section_by_id(section_id)
        .map(|section| section.page_setup)
        .unwrap_or_else(|| document.default_page_setup());
    let height = document_points_to_screen_points(setup.margins.bottom_points, canvas.zoom)
        .clamp(18.0, page_rect.height() * 0.25);
    Rect::from_min_max(
        egui::pos2(page_rect.left(), page_rect.bottom() - height),
        page_rect.max,
    )
}

fn update_canvas_hover_cursor(
    ui: &egui::Ui,
    response: &egui::Response,
    canvas: &CanvasState,
    page_layout: &PageLayout,
) {
    if response.dragged() {
        return;
    }

    let Some(hover_pos) = ui.ctx().pointer_hover_pos() else {
        return;
    };
    if !response.rect.contains(hover_pos) {
        return;
    }

    let cursor_icon = if let Some(handle) = table_resize_handle_hit(canvas, hover_pos) {
        match handle.kind {
            TableResizeKind::Column { .. } => egui::CursorIcon::ResizeEast,
            TableResizeKind::Row { .. } => egui::CursorIcon::ResizeSouth,
        }
    } else if table_cell_hit(canvas, hover_pos).is_some() {
        egui::CursorIcon::Text
    } else if let Some((_, handle)) = image_handle_hit(canvas, hover_pos, 6.0) {
        match handle {
            ResizeHandle::NW | ResizeHandle::SE => egui::CursorIcon::ResizeNwSe,
            ResizeHandle::NE | ResizeHandle::SW => egui::CursorIcon::ResizeNeSw,
            ResizeHandle::N | ResizeHandle::S => egui::CursorIcon::ResizeSouth,
            ResizeHandle::E | ResizeHandle::W => egui::CursorIcon::ResizeEast,
        }
    } else if image_body_hit(canvas, hover_pos).is_some() {
        egui::CursorIcon::Grab
    } else if page_layout.document_pos(hover_pos).is_some() {
        egui::CursorIcon::Text
    } else {
        return;
    };

    ui.ctx().set_cursor_icon(cursor_icon);
}

fn caret_height(style: CharacterStyle, zoom: f32) -> f32 {
    style.font_size_points.max(1.0) * zoom * 1.25
}

fn apply_grammar_replacement(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    ui: &egui::Ui,
    replacement: ReplacementAction,
) -> bool {
    let text = document.plain_text();
    if replacement.byte_start >= text.len() || replacement.byte_start > replacement.byte_end {
        return false;
    }

    let start_char = byte_to_char_index(&text, replacement.byte_start);
    let end_char = byte_to_char_index(&text, replacement.byte_end).max(start_char);
    let style = document.style_at(start_char);

    let now = ui.input(|i| i.time);
    history.checkpoint(document, now);
    document.delete_range(start_char..end_char);
    document.insert_text(start_char, &replacement.replacement, style);

    let cursor_char = start_char + replacement.replacement.chars().count();
    canvas.selection = egui::text_selection::CCursorRange::one(CCursor::new(cursor_char));
    canvas.active_style = document.typing_style_at(cursor_char);
    canvas.active_paragraph_style = document.paragraph_style_at(cursor_char);
    canvas.last_interaction_time = now;
    true
}

fn byte_to_char_index(text: &str, byte_offset: usize) -> usize {
    let clamped = byte_offset.min(text.len());
    let mut count = 0usize;
    for (idx, _) in text.char_indices() {
        if idx >= clamped {
            break;
        }
        count += 1;
    }
    if clamped == text.len() {
        text.chars().count()
    } else {
        count
    }
}

fn table_cell_hit(canvas: &CanvasState, pointer_pos: egui::Pos2) -> Option<(usize, usize, usize)> {
    canvas
        .table_cell_rects
        .iter()
        .rev()
        .find(|(_, _, _, rect)| rect.contains(pointer_pos))
        .map(|(table_id, row, col, _)| (*table_id, *row, *col))
}

fn table_cell_content_rect(
    canvas: &CanvasState,
    cell: (usize, usize, usize),
) -> Option<egui::Rect> {
    canvas
        .table_cell_content_rects
        .iter()
        .find(|(table_id, row, col, _)| (*table_id, *row, *col) == cell)
        .map(|(_, _, _, rect)| *rect)
}

fn table_cell_cursor_from_pointer(
    ui: &egui::Ui,
    canvas: &CanvasState,
    document: &DocumentState,
    cell_ref: (usize, usize, usize),
    pointer_pos: egui::Pos2,
) -> Option<CCursor> {
    let content_rect = table_cell_content_rect(canvas, cell_ref)?;
    let cell = document
        .table_by_id(cell_ref.0)?
        .rows
        .get(cell_ref.1)?
        .get(cell_ref.2)?;
    let galley = table_cell_text_galley(
        ui.painter(),
        cell,
        content_rect.width().max(1.0),
        canvas.zoom,
    );
    Some(galley.cursor_from_pos(pointer_pos - content_rect.min))
}

fn table_resize_handle_hit(
    canvas: &CanvasState,
    pointer_pos: egui::Pos2,
) -> Option<TableResizeHandleRect> {
    canvas
        .table_resize_handles
        .iter()
        .rev()
        .find(|handle| handle.rect.contains(pointer_pos))
        .copied()
}

fn handle_table_interaction(
    ui: &mut egui::Ui,
    response: &egui::Response,
    canvas: &mut CanvasState,
    document: &mut DocumentState,
    history: &mut ChangeHistory,
) -> (bool, bool) {
    const MIN_SIZE_POINTS: f32 = 18.0;
    let mut document_changed = false;

    if let Some(hover_pos) = ui.ctx().pointer_hover_pos() {
        if let Some(handle) = table_resize_handle_hit(canvas, hover_pos) {
            ui.ctx().set_cursor_icon(match handle.kind {
                TableResizeKind::Column { .. } => egui::CursorIcon::ResizeEast,
                TableResizeKind::Row { .. } => egui::CursorIcon::ResizeSouth,
            });
        } else if table_cell_hit(canvas, hover_pos).is_some() {
            ui.ctx().set_cursor_icon(egui::CursorIcon::Text);
        }
    }

    if !response.dragged() {
        canvas.table_resize_drag = None;
    }

    if response.dragged() {
        if let Some(drag) = canvas.table_resize_drag.as_ref() {
            if let Some(pointer_pos) = response.interact_pointer_pos() {
                let zoom = canvas.zoom.max(0.01);
                match drag.kind {
                    TableResizeKind::Column { left_col } => {
                        let delta = (pointer_pos.x - drag.start_ptr.x) / zoom;
                        let total = drag.first_points + drag.second_points;
                        let first = (drag.first_points + delta)
                            .clamp(MIN_SIZE_POINTS, total - MIN_SIZE_POINTS);
                        let second = total - first;
                        document.resize_table_column_pair(drag.table_id, left_col, first, second);
                    }
                    TableResizeKind::Row { top_row } => {
                        let delta = (pointer_pos.y - drag.start_ptr.y) / zoom;
                        let total = drag.first_points + drag.second_points;
                        let first = (drag.first_points + delta).clamp(12.0, total - 12.0);
                        let second = total - first;
                        document.resize_table_row_pair(drag.table_id, top_row, first, second);
                    }
                }
                document_changed = true;
            }
            return (true, document_changed);
        }

        if let Some(cell) = canvas.active_table_cell {
            if let Some(pointer_pos) = response.interact_pointer_pos() {
                if let Some(cursor) =
                    table_cell_cursor_from_pointer(ui, canvas, document, cell, pointer_pos)
                {
                    canvas.table_cell_selection.primary = cursor;
                    canvas.last_interaction_time = ui.input(|i| i.time);
                    return (true, document_changed);
                }
            }
        }
    }

    let Some(pointer_pos) = response.interact_pointer_pos() else {
        return (false, document_changed);
    };

    if response.drag_started() {
        if let Some(handle) = table_resize_handle_hit(canvas, pointer_pos) {
            if let Some(table) = document.table_by_id(handle.table_id) {
                let dimensions = match handle.kind {
                    TableResizeKind::Column { left_col } => {
                        if left_col + 1 >= table.col_widths_points.len() {
                            None
                        } else {
                            Some((
                                table.col_widths_points[left_col],
                                table.col_widths_points[left_col + 1],
                            ))
                        }
                    }
                    TableResizeKind::Row { top_row } => {
                        if top_row + 1 >= table.row_heights_points.len() {
                            None
                        } else {
                            Some((
                                table.row_heights_points[top_row],
                                table.row_heights_points[top_row + 1],
                            ))
                        }
                    }
                };
                if let Some((first_points, second_points)) = dimensions {
                    history.checkpoint(document, ui.input(|i| i.time));
                    canvas.table_resize_drag = Some(TableResizeDrag {
                        table_id: handle.table_id,
                        kind: handle.kind,
                        start_ptr: pointer_pos,
                        first_points,
                        second_points,
                    });
                    canvas.active_table_cell = None;
                    canvas.selected_image_id = None;
                    return (true, document_changed);
                }
            }
        }

        if let Some(cell) = table_cell_hit(canvas, pointer_pos) {
            response.request_focus();
            canvas.active_table_cell = Some(cell);
            if let Some(cursor) =
                table_cell_cursor_from_pointer(ui, canvas, document, cell, pointer_pos)
            {
                if ui.input(|i| i.modifiers.shift) {
                    canvas.table_cell_selection.primary = cursor;
                } else {
                    canvas.table_cell_selection = CCursorRange::one(cursor);
                }
                if let Some(style) =
                    document.table_cell_style_at(cell.0, cell.1, cell.2, cursor.index)
                {
                    canvas.active_style = style;
                }
            }
            canvas.selected_image_id = None;
            canvas.resize_drag = None;
            canvas.move_drag = None;
            canvas.last_interaction_time = ui.input(|i| i.time);
            return (true, document_changed);
        }
    }

    if response.clicked() {
        if let Some(cell) = table_cell_hit(canvas, pointer_pos) {
            response.request_focus();
            canvas.active_table_cell = Some(cell);
            if let Some(cursor) =
                table_cell_cursor_from_pointer(ui, canvas, document, cell, pointer_pos)
            {
                if ui.input(|i| i.modifiers.shift) {
                    canvas.table_cell_selection.primary = cursor;
                } else {
                    canvas.table_cell_selection = CCursorRange::one(cursor);
                }
                if let Some(style) =
                    document.table_cell_style_at(cell.0, cell.1, cell.2, cursor.index)
                {
                    canvas.active_style = style;
                }
            }
            canvas.selected_image_id = None;
            canvas.resize_drag = None;
            canvas.move_drag = None;
            canvas.last_interaction_time = ui.input(|i| i.time);
            return (true, document_changed);
        }
        canvas.active_table_cell = None;
        canvas.table_cell_selection = CCursorRange::default();
    }

    (false, document_changed)
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
    let mut tables = Vec::new();
    let mut manual_page_break_rows = Vec::new();
    let mut paragraph_start_rows = Vec::new();
    let mut paragraph_spacing_ranges = Vec::new();
    let mut row_index = 0usize;
    let mut side_wrap_flow: Option<ActiveSideWrapFlow> = None;

    for paragraph in document.paragraphs() {
        paragraph_start_rows.push(row_index);
        if side_wrap_flow
            .as_ref()
            .is_some_and(|flow| flow.remaining_height <= 0.0)
        {
            side_wrap_flow = None;
        }

        let base_indent = if paragraph.list_marker.is_some() {
            marker_gutter
        } else {
            0.0
        };
        let (indent, paragraph_wrap_width) = side_wrap_flow
            .as_ref()
            .filter(|flow| flow.pending_top_height <= 0.0)
            .map_or_else(
                || (base_indent, (wrap_width - base_indent).max(1.0)),
                |flow| {
                    let start_x = flow.text_start_x.max(base_indent).clamp(0.0, wrap_width);
                    let end_x = (flow.text_start_x + flow.text_width).clamp(start_x, wrap_width);
                    (start_x, (end_x - start_x).max(1.0))
                },
            );
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

        let mut side_wrap_image_spec: Option<(f32, f32, f32, f32, f32)> = None;
        if let Some(table) = paragraph.table.clone().filter(|_| !has_visible_text) {
            let table_width =
                document_points_to_screen_points(table.total_width_points(), canvas.zoom);
            let table_height: f32 = table_row_heights_screen(painter, &table, canvas.zoom)
                .into_iter()
                .sum();
            let table_size = egui::vec2(table_width.min(paragraph_wrap_width), table_height);
            if reserve_block_image_space(&mut paragraph_galley, table_size) {
                tables.push(TableLayout {
                    row_index,
                    size: table_size,
                    table,
                });
            }
        } else if let Some(image) = paragraph.image.clone().filter(|_| !has_visible_text) {
            let wrap_mode = image.wrap_mode;
            let image_offset_x_points = image.offset_x_points();
            let image_offset_y_points = image.offset_y_points();
            let display_size = image_display_size(&image, paragraph_wrap_width, canvas.zoom);
            let reservation =
                block_image_reservation(wrap_mode, display_size, paragraph_wrap_width, canvas.zoom);
            if reserve_block_image_space(&mut paragraph_galley, reservation.row_size) {
                images.push(ImageLayout {
                    row_index,
                    size: display_size,
                    offset: reservation.image_offset,
                    image,
                });
            }

            if wrap_uses_side_flow(wrap_mode) {
                let pad = side_wrap_pad(wrap_mode, canvas.zoom);
                side_wrap_image_spec = Some((
                    display_size.x,
                    display_size.y,
                    reservation.image_offset.x
                        + document_points_to_screen_points(image_offset_x_points, canvas.zoom),
                    reservation.image_offset.y
                        + document_points_to_screen_points(image_offset_y_points, canvas.zoom),
                    pad,
                ));
            }
        }
        align_paragraph_galley(
            &mut paragraph_galley,
            indent,
            paragraph_wrap_width,
            paragraph.style.alignment,
        );
        apply_line_spacing(
            &mut paragraph_galley,
            paragraph.style.line_spacing,
            canvas.zoom,
        );

        if let Some((image_width, image_height, image_offset_x, image_offset_y, pad)) =
            side_wrap_image_spec
        {
            let image_row_x = paragraph_galley
                .rows
                .first()
                .map(|row| row.pos.x)
                .unwrap_or(indent);
            let image_left = image_row_x + image_offset_x;
            let image_right = image_left + image_width;
            let left_width = (image_left - pad).max(0.0);
            let right_start = (image_right + pad).clamp(0.0, wrap_width);
            let right_width = (wrap_width - right_start).max(0.0);
            let min_side_width = document_points_to_screen_points(72.0, canvas.zoom);
            let (text_start_x, text_width) = if right_width >= left_width {
                (right_start, right_width)
            } else {
                (0.0, left_width)
            };

            let zone_top = image_offset_y - pad;
            let zone_bottom = image_offset_y + image_height + pad;
            let pending_top_height = zone_top.max(0.0);
            let remaining_height = (zone_bottom - pending_top_height).max(0.0);

            if text_width >= min_side_width && remaining_height > 0.0 {
                side_wrap_flow = Some(ActiveSideWrapFlow {
                    pending_top_height,
                    remaining_height,
                    text_start_x,
                    text_width,
                });
            } else {
                side_wrap_flow = None;
            }
        }

        if let Some(marker_text) = paragraph.list_marker {
            list_markers.push(ListMarkerLayout {
                row_index,
                text: marker_text,
                x: indent - marker_gap,
                font_id: text_format(marker_style, canvas.zoom).font_id,
                color: marker_style.text_color,
            });
        }

        if let Some(flow) = side_wrap_flow.as_mut() {
            let paragraph_height = paragraph_galley.rect.height().max(0.0);
            if flow.pending_top_height > 0.0 {
                let consumed_top = paragraph_height.min(flow.pending_top_height);
                flow.pending_top_height -= consumed_top;
                flow.remaining_height -= (paragraph_height - consumed_top).max(0.0);
            } else {
                flow.remaining_height -= paragraph_height;
            }
        }

        let paragraph_row_count = paragraph_galley.rows.len();
        let paragraph_spacing_top = document_points_to_screen_points(
            f32::from(paragraph.style.spacing_before_points),
            canvas.zoom,
        );
        let paragraph_spacing_bottom = document_points_to_screen_points(
            f32::from(paragraph.style.spacing_after_points),
            canvas.zoom,
        );
        if paragraph_row_count > 0 {
            paragraph_spacing_ranges.push(ParagraphSpacingRange {
                row_start: row_index,
                row_end: row_index + paragraph_row_count,
                top: paragraph_spacing_top,
                bottom: paragraph_spacing_bottom,
            });
        }

        row_index += paragraph_row_count;
        paragraph_galleys.push(paragraph_galley);
    }

    let plain_text = document.plain_text();
    let mut merged_job = egui::epaint::text::LayoutJob::default();
    merged_job.wrap.max_width = wrap_width;
    merged_job.break_on_newline = true;
    merged_job.append(&plain_text, 0.0, text_format(default_style, canvas.zoom));

    let mut galley = Arc::new(egui::Galley::concat(
        Arc::new(merged_job),
        &paragraph_galleys,
        painter.pixels_per_point(),
    ));
    apply_paragraph_row_spacing(&mut galley, &paragraph_spacing_ranges);
    apply_tight_wrap_row_offsets(&mut galley, &images, canvas.zoom, wrap_width);

    DocumentLayout {
        galley,
        list_markers,
        images,
        tables,
        manual_page_break_rows,
        paragraph_start_rows,
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

fn apply_line_spacing(
    galley: &mut Arc<egui::Galley>,
    line_spacing: crate::document::LineSpacing,
    zoom: f32,
) {
    if galley.rows.len() < 2 {
        return;
    }

    let galley = Arc::make_mut(galley);
    let original_rect = galley.rect;
    let mut mesh_bounds = egui::Rect::NOTHING;
    let mut cumulative_shift = 0.0;

    for row_index in 1..galley.rows.len() {
        let previous_row_height = galley.rows[row_index - 1].row.height();
        let desired_advance = match line_spacing.kind {
            LineSpacingKind::AutoMultiplier => previous_row_height * line_spacing.value.max(0.0),
            LineSpacingKind::AtLeastPoints => previous_row_height.max(
                document_points_to_screen_points(line_spacing.value.max(0.0), zoom),
            ),
            LineSpacingKind::ExactPoints => {
                document_points_to_screen_points(line_spacing.value.max(0.0), zoom)
            }
        };
        cumulative_shift += desired_advance - previous_row_height;
        galley.rows[row_index].pos.y += cumulative_shift;
    }

    for row in &galley.rows {
        mesh_bounds |= row.visuals.mesh_bounds.translate(row.pos.to_vec2());
    }

    galley.rect = egui::Rect::from_min_max(
        original_rect.min,
        egui::pos2(original_rect.max.x, original_rect.max.y + cumulative_shift),
    );
    galley.mesh_bounds = mesh_bounds;
}

struct BlockImageReservation {
    row_size: egui::Vec2,
    image_offset: egui::Vec2,
}

fn block_image_reservation(
    wrap_mode: WrapMode,
    image_size: egui::Vec2,
    wrap_width: f32,
    zoom: f32,
) -> BlockImageReservation {
    let square_pad = square_wrap_pad(zoom);
    let tight_pad = tight_wrap_pad(zoom);

    match wrap_mode {
        WrapMode::Inline => BlockImageReservation {
            row_size: image_size,
            image_offset: egui::Vec2::ZERO,
        },
        WrapMode::Square => {
            let row_width = (image_size.x + square_pad * 2.0).min(wrap_width);
            let row_height = (square_pad * 2.0).max(tight_wrap_row_height(zoom));
            BlockImageReservation {
                row_size: egui::vec2(row_width, row_height),
                image_offset: egui::vec2(square_pad, square_pad),
            }
        }
        WrapMode::Tight => {
            let row_width = (image_size.x + tight_pad * 2.0).min(wrap_width);
            let row_height = tight_wrap_row_height(zoom);
            BlockImageReservation {
                row_size: egui::vec2(row_width, row_height),
                image_offset: egui::vec2(tight_pad, 0.0),
            }
        }
        WrapMode::Through | WrapMode::BehindText | WrapMode::InFrontOfText => {
            BlockImageReservation {
                row_size: egui::Vec2::ZERO,
                image_offset: egui::Vec2::ZERO,
            }
        }
        WrapMode::TopAndBottom => BlockImageReservation {
            row_size: egui::vec2(wrap_width, image_size.y),
            image_offset: egui::vec2(((wrap_width - image_size.x) * 0.5).max(0.0), 0.0),
        },
    }
}

fn tight_wrap_pad(zoom: f32) -> f32 {
    document_points_to_screen_points(4.0, zoom)
}

fn square_wrap_pad(zoom: f32) -> f32 {
    document_points_to_screen_points(12.0, zoom)
}

fn side_wrap_pad(wrap_mode: WrapMode, zoom: f32) -> f32 {
    match wrap_mode {
        WrapMode::Square => square_wrap_pad(zoom),
        WrapMode::Tight | WrapMode::Through => tight_wrap_pad(zoom),
        _ => 0.0,
    }
}

fn wrap_uses_side_flow(wrap_mode: WrapMode) -> bool {
    matches!(
        wrap_mode,
        WrapMode::Square | WrapMode::Tight | WrapMode::Through
    )
}

fn tight_wrap_row_height(zoom: f32) -> f32 {
    document_points_to_screen_points(14.0, zoom).max(1.0)
}

fn apply_tight_wrap_row_offsets(
    galley: &mut Arc<egui::Galley>,
    images: &[ImageLayout],
    zoom: f32,
    wrap_width: f32,
) {
    let zones = tight_wrap_zones(galley, images, zoom, wrap_width);
    if zones.is_empty() {
        return;
    }

    let galley = Arc::make_mut(galley);
    let mut min_rect = galley.rect.min;
    let mut max_rect = galley.rect.max;
    let mut mesh_bounds = egui::Rect::NOTHING;
    let mut cumulative_shift = 0.0;

    for (index, row) in galley.rows.iter_mut().enumerate() {
        let row_text = row.row.text();
        if !row_text.chars().all(|ch| ch == OBJECT_REPLACEMENT_CHAR) {
            row.pos.y += cumulative_shift;

            loop {
                let Some(zone) = zones.iter().find(|zone| {
                    zone.row_index != index && row.max_y() > zone.top && row.min_y() < zone.bottom
                }) else {
                    break;
                };

                // Check if the row's actual glyph content fits in the side column.
                // Use the intrinsic row width (row.row.size.x) to decide if the row
                // can fit beside the image, and reposition it into the side column.
                let row_content_width = row.row.size.x;
                let min_side_width = document_points_to_screen_points(36.0, zoom);
                if row_content_width <= zone.text_width && zone.text_width >= min_side_width {
                    row.pos.x = zone.text_start_x;
                    break;
                }

                let offset_y = zone.bottom - row.pos.y;
                if offset_y <= 0.0 {
                    break;
                }
                row.pos.y += offset_y;
                cumulative_shift += offset_y;
            }
        } else if cumulative_shift > 0.0 {
            row.pos.y += cumulative_shift;
        }

        let row_rect = row.rect();
        min_rect.x = min_rect.x.min(row_rect.min.x);
        min_rect.y = min_rect.y.min(row_rect.min.y);
        max_rect.x = max_rect.x.max(row_rect.max.x);
        max_rect.y = max_rect.y.max(row_rect.max.y);
        mesh_bounds |= row.visuals.mesh_bounds.translate(row.pos.to_vec2());
    }

    galley.rect = Rect::from_min_max(min_rect, max_rect);
    galley.mesh_bounds = mesh_bounds;
}

fn tight_wrap_zones(
    galley: &egui::Galley,
    images: &[ImageLayout],
    zoom: f32,
    wrap_width: f32,
) -> Vec<TightWrapZone> {
    let mut zones = Vec::new();

    for image in images
        .iter()
        .filter(|image| wrap_uses_side_flow(image.image.wrap_mode))
    {
        let Some(row) = galley.rows.get(image.row_index) else {
            continue;
        };

        let image_left = row.pos.x
            + image.offset.x
            + document_points_to_screen_points(image.image.offset_x_points(), zoom);
        let image_top = row.pos.y
            + image.offset.y
            + document_points_to_screen_points(image.image.offset_y_points(), zoom);
        let pad = side_wrap_pad(image.image.wrap_mode, zoom);

        let image_right = image_left + image.size.x;
        let image_bottom = image_top + image.size.y;
        let left_width = (image_left - pad).max(0.0);
        let right_start = (image_right + pad).max(0.0);
        let right_width = (wrap_width - right_start).max(0.0);
        let (text_start_x, text_width) = if right_width >= left_width {
            (right_start, right_width)
        } else {
            (0.0, left_width)
        };

        zones.push(TightWrapZone {
            row_index: image.row_index,
            top: image_top - pad,
            bottom: image_bottom + pad,
            text_start_x,
            text_width,
        });
    }

    zones
}

fn reserve_block_image_space(galley: &mut Arc<egui::Galley>, row_size: egui::Vec2) -> bool {
    let galley = Arc::make_mut(galley);
    let Some(placed_row) = galley.rows.first_mut() else {
        return false;
    };
    let row = Arc::make_mut(&mut placed_row.row);
    if row.glyphs.len() != 1 || row.glyphs[0].chr != OBJECT_REPLACEMENT_CHAR {
        return false;
    }

    row.glyphs[0].advance_width = row_size.x;
    row.glyphs[0].line_height = row_size.y;
    row.glyphs[0].font_ascent = row_size.y;
    row.glyphs[0].font_height = row_size.y;
    row.glyphs[0].font_face_ascent = row_size.y;
    row.glyphs[0].font_face_height = row_size.y;
    row.size = row_size;
    row.visuals = Default::default();

    galley.rect = Rect::from_min_size(galley.rect.min, row_size);
    galley.mesh_bounds = Rect::NOTHING;
    true
}

struct ParagraphSpacingRange {
    row_start: usize,
    row_end: usize,
    top: f32,
    bottom: f32,
}

fn apply_paragraph_row_spacing(
    galley: &mut Arc<egui::Galley>,
    paragraph_spacing_ranges: &[ParagraphSpacingRange],
) {
    if paragraph_spacing_ranges.is_empty() {
        return;
    }

    let galley = Arc::make_mut(galley);
    let original_rect = galley.rect;
    let mut mesh_bounds = egui::Rect::NOTHING;
    let mut cumulative_shift = 0.0;
    let row_count = galley.rows.len();

    for range in paragraph_spacing_ranges {
        let paragraph_shift = cumulative_shift + range.top;
        let start = range.row_start.min(row_count);
        let end = range.row_end.min(row_count);
        for row in galley.rows[start..end].iter_mut() {
            row.pos.y += paragraph_shift;
        }
        cumulative_shift += range.top + range.bottom;
    }

    for row in &galley.rows {
        mesh_bounds |= row.visuals.mesh_bounds.translate(row.pos.to_vec2());
    }

    galley.rect = egui::Rect::from_min_max(
        original_rect.min,
        egui::pos2(original_rect.max.x, original_rect.max.y + cumulative_shift),
    );
    galley.mesh_bounds = mesh_bounds;
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
        app::{CanvasState, ZoomMode},
        document::{
            CharacterStyle, DocumentImage, DocumentState, ImageLayoutMode, ImageRendering,
            LineSpacing, LineSpacingKind, ListKind, PageMargins, PageSize, ParagraphAlignment,
            ParagraphStyle, TextRun, WrapMode, OBJECT_REPLACEMENT_CHAR,
        },
        layout::fit_page_zoom,
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
        let paragraph_count = paragraph_styles.len();
        DocumentState {
            title: "Test".to_owned(),
            runs,
            paragraph_styles,
            paragraph_images,
            paragraph_tables: vec![None; paragraph_count],
            page_size: PageSize::a4(),
            margins: PageMargins::standard(),
            header_text: String::new(),
            footer_text: String::new(),
            first_page_header_text: String::new(),
            first_page_footer_text: String::new(),
            even_page_header_text: String::new(),
            even_page_footer_text: String::new(),
            header_runs: vec![TextRun {
                text: String::new(),
                style: CharacterStyle::default(),
            }],
            footer_runs: vec![TextRun {
                text: String::new(),
                style: CharacterStyle::default(),
            }],
            first_page_header_runs: vec![TextRun {
                text: String::new(),
                style: CharacterStyle::default(),
            }],
            first_page_footer_runs: vec![TextRun {
                text: String::new(),
                style: CharacterStyle::default(),
            }],
            even_page_header_runs: vec![TextRun {
                text: String::new(),
                style: CharacterStyle::default(),
            }],
            even_page_footer_runs: vec![TextRun {
                text: String::new(),
                style: CharacterStyle::default(),
            }],
            different_first_page: false,
            different_odd_even_pages: false,
            page_number_start: 1,
            sections: vec![crate::document::Section::first(
                crate::document::PageSetup::standard(),
            )],
        }
    }

    fn make_test_image(id: usize, width: f32, height: f32, wrap_mode: WrapMode) -> DocumentImage {
        DocumentImage {
            id,
            bytes: vec![],
            alt_text: "test".to_owned(),
            width_points: width,
            height_points: height,
            lock_aspect_ratio: true,
            opacity: 1.0,
            layout_mode: ImageLayoutMode::Inline,
            wrap_mode,
            rendering: ImageRendering::Smooth,
            horizontal_position: Default::default(),
            vertical_position: Default::default(),
            distance_from_text: Default::default(),
            z_index: 0,
            move_with_text: true,
            allow_overlap: false,
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
            vec![Some(make_test_image(1, 200.0, 100.0, WrapMode::Inline))],
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
    fn wide_image_display_size_can_exceed_wrap_width() {
        let document = make_document(
            vec![TextRun {
                text: format!("{OBJECT_REPLACEMENT_CHAR}"),
                style: CharacterStyle::default(),
            }],
            vec![ParagraphStyle::default()],
            vec![Some(make_test_image(31, 900.0, 300.0, WrapMode::Inline))],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 600.0);

        assert_eq!(layout.images.len(), 1);
        assert!(
            layout.images[0].size.x > 600.0,
            "wide images should keep stored size instead of being clamped to wrap width"
        );
        assert!((layout.images[0].size.x - 900.0).abs() < 1.0);
        assert!((layout.images[0].size.y - 300.0).abs() < 1.0);
    }

    #[test]
    fn resize_geometry_non_locked_handles_keep_opposite_edges_stable() {
        let cases = [
            (ResizeHandle::NW, (90.0, 65.0, 10.0, 15.0)),
            (ResizeHandle::N, (100.0, 65.0, 0.0, 15.0)),
            (ResizeHandle::NE, (110.0, 65.0, 0.0, 15.0)),
            (ResizeHandle::E, (110.0, 80.0, 0.0, 0.0)),
            (ResizeHandle::SE, (110.0, 95.0, 0.0, 0.0)),
            (ResizeHandle::S, (100.0, 95.0, 0.0, 0.0)),
            (ResizeHandle::SW, (90.0, 95.0, 10.0, 0.0)),
            (ResizeHandle::W, (90.0, 80.0, 10.0, 0.0)),
        ];

        for (handle, expected) in cases {
            let actual = image::resized_image_geometry(handle, 100.0, 80.0, 10.0, 15.0, false);
            assert_geometry_close(actual, expected, handle);
        }
    }

    #[test]
    fn resize_geometry_clamps_to_minimum_size() {
        let west = image::resized_image_geometry(ResizeHandle::W, 100.0, 80.0, 200.0, 0.0, false);
        assert_geometry_close(west, (24.0, 80.0, 76.0, 0.0), ResizeHandle::W);

        let north = image::resized_image_geometry(ResizeHandle::N, 100.0, 80.0, 0.0, 200.0, false);
        assert_geometry_close(north, (100.0, 24.0, 0.0, 56.0), ResizeHandle::N);
    }

    #[test]
    fn resize_geometry_locked_ratio_keeps_anchors_stable() {
        let nw = image::resized_image_geometry(ResizeHandle::NW, 100.0, 50.0, 20.0, 0.0, true);
        assert_geometry_close(nw, (80.0, 40.0, 20.0, 10.0), ResizeHandle::NW);

        let east = image::resized_image_geometry(ResizeHandle::E, 100.0, 50.0, 20.0, 0.0, true);
        assert_geometry_close(east, (120.0, 60.0, 0.0, -5.0), ResizeHandle::E);

        let south = image::resized_image_geometry(ResizeHandle::S, 100.0, 50.0, 0.0, 20.0, true);
        assert_geometry_close(south, (140.0, 70.0, -20.0, 0.0), ResizeHandle::S);
    }

    fn assert_geometry_close(
        actual: (f32, f32, f32, f32),
        expected: (f32, f32, f32, f32),
        handle: ResizeHandle,
    ) {
        assert!(
            (actual.0 - expected.0).abs() < 0.01
                && (actual.1 - expected.1).abs() < 0.01
                && (actual.2 - expected.2).abs() < 0.01
                && (actual.3 - expected.3).abs() < 0.01,
            "{handle:?}: expected {expected:?}, got {actual:?}"
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
            vec![TextRun {
                text: "Left aligned\nCenter aligned\nRight aligned".to_owned(),
                style: CharacterStyle::default(),
            }],
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
    fn paragraph_spacing_offsets_following_rows() {
        let document = make_document(
            vec![TextRun {
                text: "First\nSecond".to_owned(),
                style: CharacterStyle::default(),
            }],
            vec![
                ParagraphStyle {
                    spacing_after_points: 24,
                    ..ParagraphStyle::default()
                },
                ParagraphStyle::default(),
            ],
            vec![None, None],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 600.0);

        assert!(layout.galley.rows.len() >= 2);
        let first_row = &layout.galley.rows[0];
        let second_row = &layout.galley.rows[1];
        assert!(
            second_row.pos.y - first_row.pos.y > first_row.rect().height(),
            "paragraph spacing should increase the gap between rows"
        );
    }

    #[test]
    fn auto_multiplier_line_spacing_offsets_following_line() {
        let document = make_document(
            vec![TextRun {
                text: "First line in paragraph that wraps on purpose because the width is narrow"
                    .to_owned(),
                style: CharacterStyle::default(),
            }],
            vec![ParagraphStyle {
                line_spacing: LineSpacing {
                    kind: LineSpacingKind::AutoMultiplier,
                    value: 1.5,
                },
                ..ParagraphStyle::default()
            }],
            vec![None],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 240.0);

        assert!(layout.galley.rows.len() >= 2);
        let first_row = &layout.galley.rows[0];
        let second_row = &layout.galley.rows[1];
        let default_gap = first_row.row.height();
        let actual_gap = second_row.pos.y - first_row.pos.y;
        assert!(
            actual_gap > default_gap * 1.45,
            "1.5x line spacing should enlarge row advance, got actual_gap={actual_gap}, default_gap={default_gap}"
        );
    }

    #[test]
    fn exact_line_spacing_uses_requested_row_advance() {
        let document = make_document(
            vec![TextRun {
                text: "First line in paragraph that wraps on purpose because the width is narrow"
                    .to_owned(),
                style: CharacterStyle::default(),
            }],
            vec![ParagraphStyle {
                line_spacing: LineSpacing {
                    kind: LineSpacingKind::ExactPoints,
                    value: 24.0,
                },
                ..ParagraphStyle::default()
            }],
            vec![None],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 240.0);

        assert!(layout.galley.rows.len() >= 2);
        let first_row = &layout.galley.rows[0];
        let second_row = &layout.galley.rows[1];
        let actual_gap = second_row.pos.y - first_row.pos.y;
        assert!(
            (actual_gap - 24.0).abs() < 1.5,
            "exact line spacing should follow the requested advance, got {actual_gap}"
        );
    }

    #[test]
    fn fit_page_zoom_uses_manual_override_rules() {
        let viewport = egui::Rect::from_min_size(egui::Pos2::ZERO, egui::vec2(400.0, 500.0));
        let fit = fit_page_zoom(viewport, PageSize::a4());
        assert!(
            fit < 1.0,
            "fit zoom should shrink an A4 page in a small viewport"
        );

        let mut canvas = CanvasState::default();
        canvas.imported_docx_view = true;
        canvas.zoom_mode = ZoomMode::FitPage;
        canvas.zoom = fit;
        canvas.zoom_mode = ZoomMode::Manual;
        canvas.zoom = (canvas.zoom * 1.1).clamp(0.5, 3.0);
        assert_eq!(canvas.zoom_mode, ZoomMode::Manual);
        assert!(canvas.zoom > fit);
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
                Some(make_test_image(2, 400.0, 300.0, WrapMode::Inline)),
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

        assert_eq!(layout.list_markers.len(), 2, "should have two list markers");
        assert_eq!(layout.list_markers[0].text, "•");
        assert_eq!(layout.list_markers[1].text, "•");
    }

    #[test]
    fn wrap_modes_change_image_row_geometry() {
        let make_doc = |wrap_mode| {
            make_document(
                vec![TextRun {
                    text: format!("{OBJECT_REPLACEMENT_CHAR}"),
                    style: CharacterStyle::default(),
                }],
                vec![ParagraphStyle::default()],
                vec![Some(make_test_image(10, 180.0, 90.0, wrap_mode))],
            )
        };

        let canvas = CanvasState::default();
        let wrap_width = 600.0;

        let inline = run_headless_layout(&make_doc(WrapMode::Inline), &canvas, wrap_width);
        let square = run_headless_layout(&make_doc(WrapMode::Square), &canvas, wrap_width);
        let tight = run_headless_layout(&make_doc(WrapMode::Tight), &canvas, wrap_width);
        let through = run_headless_layout(&make_doc(WrapMode::Through), &canvas, wrap_width);
        let top_bottom =
            run_headless_layout(&make_doc(WrapMode::TopAndBottom), &canvas, wrap_width);

        let inline_row = &inline.galley.rows[inline.images[0].row_index];
        let square_row = &square.galley.rows[square.images[0].row_index];
        let tight_row = &tight.galley.rows[tight.images[0].row_index];
        let through_row = &through.galley.rows[through.images[0].row_index];
        let top_bottom_row = &top_bottom.galley.rows[top_bottom.images[0].row_index];

        assert!(
            square_row.size.x > tight_row.size.x,
            "square wrap should reserve more horizontal space than tight"
        );
        assert!(
            square_row.size.y > tight_row.size.y,
            "square wrap should reserve more vertical space than tight"
        );
        assert!(
            through_row.size.x <= 1.0 && through_row.size.y <= 1.0,
            "through wrap should not reserve layout space"
        );
        assert!(
            (top_bottom_row.size.x - wrap_width).abs() < 1.0,
            "top-and-bottom wrap should reserve the full row width"
        );
        assert!(
            top_bottom.images[0].offset.x > through.images[0].offset.x,
            "top-and-bottom should center image while through should anchor without horizontal reservation"
        );
        assert!(
            inline_row.size.x < top_bottom_row.size.x,
            "inline wrap should keep a tighter row than top-and-bottom"
        );
    }

    #[test]
    fn tight_wrap_chooses_side_from_image_position() {
        let make_doc = |offset_x: f32| {
            let mut img = make_test_image(21, 180.0, 90.0, WrapMode::Tight);
            img.horizontal_position.offset_points = offset_x;
            make_document(
                vec![TextRun {
                    text: format!(
                        "{OBJECT_REPLACEMENT_CHAR}\nthis paragraph should flow beside the image and expose side choice"
                    ),
                    style: CharacterStyle::default(),
                }],
                vec![ParagraphStyle::default(), ParagraphStyle::default()],
                vec![
                    Some(img),
                    None,
                ],
            )
        };

        let canvas = CanvasState::default();
        let wrap_width = 600.0;
        let left_placed = run_headless_layout(&make_doc(0.0), &canvas, wrap_width);
        let right_placed = run_headless_layout(&make_doc(220.0), &canvas, wrap_width);

        let left_row_x = left_placed.galley.rows[1].pos.x;
        let right_row_x = right_placed.galley.rows[1].pos.x;

        assert!(
            left_row_x > right_row_x,
            "left-placed image should push text to the right; got left_row_x={left_row_x}, right_row_x={right_row_x}"
        );
    }

    #[test]
    fn side_wrap_modes_place_short_rows_beside_image() {
        for wrap_mode in [WrapMode::Square, WrapMode::Tight, WrapMode::Through] {
            let document = make_document(
                vec![TextRun {
                    text: format!("{OBJECT_REPLACEMENT_CHAR}\nshort side text"),
                    style: CharacterStyle::default(),
                }],
                vec![ParagraphStyle::default(), ParagraphStyle::default()],
                vec![Some(make_test_image(41, 180.0, 90.0, wrap_mode)), None],
            );
            let canvas = CanvasState::default();
            let layout = run_headless_layout(&document, &canvas, 600.0);
            let image = &layout.images[0];
            let image_row = &layout.galley.rows[image.row_index];
            let text_row = &layout.galley.rows[image.row_index + 1];
            let image_right = image_row.pos.x + image.offset.x + image.size.x;

            assert!(
                text_row.pos.x > image_right,
                "{wrap_mode:?} should place short text beside the image: text_x={}, image_right={}, text_width={}, row_size={}",
                text_row.pos.x,
                image_right,
                text_row.size.x,
                text_row.row.size.x,
            );
        }
    }

    #[test]
    fn tight_wrap_reflows_wrappable_paragraph_into_side_column() {
        let document = make_document(
            vec![TextRun {
                text: format!(
                    "{OBJECT_REPLACEMENT_CHAR}\n{}",
                    "bon above to change bold italic underline strike through text size font family text color and highlight"
                ),
                style: CharacterStyle::default(),
            }],
            vec![ParagraphStyle::default(), ParagraphStyle::default()],
            vec![Some(make_test_image(43, 175.0, 130.0, WrapMode::Tight)), None],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 451.0);
        let image = &layout.images[0];
        let image_row = &layout.galley.rows[image.row_index];
        let image_right = image_row.pos.x + image.offset.x + image.size.x;
        let side_rows: Vec<_> = layout
            .galley
            .rows
            .iter()
            .skip(image.row_index + 1)
            .take_while(|row| row.min_y() < image_row.pos.y + image.offset.y + image.size.y)
            .collect();

        assert!(
            side_rows.len() >= 2,
            "tight wrap should reflow the paragraph into multiple side-column rows"
        );
        assert!(
            side_rows.iter().all(|row| row.pos.x > image_right),
            "all rows beside the image should start to the right of the image"
        );
    }

    #[test]
    fn side_wrap_modes_push_wide_rows_below_image() {
        let wide_word = "x".repeat(140);
        for wrap_mode in [WrapMode::Square, WrapMode::Tight, WrapMode::Through] {
            let document = make_document(
                vec![TextRun {
                    text: format!("{OBJECT_REPLACEMENT_CHAR}\n{wide_word}"),
                    style: CharacterStyle::default(),
                }],
                vec![ParagraphStyle::default(), ParagraphStyle::default()],
                vec![Some(make_test_image(42, 540.0, 90.0, wrap_mode)), None],
            );
            let canvas = CanvasState::default();
            let layout = run_headless_layout(&document, &canvas, 600.0);
            let image = &layout.images[0];
            let image_row = &layout.galley.rows[image.row_index];
            let text_row = &layout.galley.rows[image.row_index + 1];
            let image_bottom = image_row.pos.y + image.offset.y + image.size.y;

            assert!(
                text_row.pos.y >= image_bottom,
                "{wrap_mode:?} should push text below the image when no side fits: text_y={}, image_bottom={}, text_x={}, glyph_width={}, row_width={}",
                text_row.pos.y,
                image_bottom,
                text_row.pos.x,
                text_row.row.size.x,
                text_row.size.x
            );
        }
    }

    #[test]
    fn tight_wrap_vertical_offset_delays_text_wrapping() {
        let make_doc = |offset_y: f32| {
            let mut img = make_test_image(22, 180.0, 90.0, WrapMode::Tight);
            img.vertical_position.offset_points = offset_y;
            make_document(
                vec![TextRun {
                    text: format!(
                        "{OBJECT_REPLACEMENT_CHAR}\nthis paragraph should start unwrapped when image is moved down enough"
                    ),
                    style: CharacterStyle::default(),
                }],
                vec![ParagraphStyle::default(), ParagraphStyle::default()],
                vec![
                    Some(img),
                    None,
                ],
            )
        };

        let canvas = CanvasState::default();
        let wrap_width = 600.0;
        let normal = run_headless_layout(&make_doc(0.0), &canvas, wrap_width);
        let moved_down = run_headless_layout(&make_doc(220.0), &canvas, wrap_width);

        let normal_row_x = normal.galley.rows[1].pos.x;
        let moved_down_row_x = moved_down.galley.rows[1].pos.x;

        assert!(
            normal_row_x > moved_down_row_x + 40.0,
            "moving image down should reduce early side-wrapping; got normal_row_x={normal_row_x}, moved_down_row_x={moved_down_row_x}"
        );
    }
}
