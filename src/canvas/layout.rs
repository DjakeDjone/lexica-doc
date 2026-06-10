use std::sync::Arc;
use eframe::egui::{self, Color32, FontId};

use crate::app::CanvasState;
use crate::document::{
    text_format, CharacterStyle, DocumentState, DocumentTable, LineSpacingKind, ParagraphAlignment,
    TextRun, OBJECT_REPLACEMENT_CHAR,
};
use crate::layout::document_points_to_screen_points;

use super::image_wrap::{
    apply_tight_wrap_row_offsets, block_image_reservation, reserve_block_image_space,
    ActiveSideWrapFlow, ImageLayout,
};
use super::table::table_row_heights_screen;

pub struct DocumentLayout {
    pub galley: Arc<egui::Galley>,
    pub list_markers: Vec<ListMarkerLayout>,
    pub images: Vec<ImageLayout>,
    pub tables: Vec<TableLayout>,
    pub manual_page_break_rows: Vec<usize>,
    pub paragraph_start_rows: Vec<usize>,
}

pub struct TableLayout {
    pub row_index: usize,
    pub table: DocumentTable,
}

pub struct ListMarkerLayout {
    pub row_index: usize,
    pub text: String,
    pub x: f32,
    pub font_id: FontId,
    pub color: Color32,
}

pub struct ParagraphSpacingRange {
    pub row_start: usize,
    pub row_end: usize,
    pub top: f32,
    pub bottom: f32,
}

pub fn layout_document(
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

        let cursor_index = canvas.selection.primary.index;
        let ai_completion = canvas.ai_completion.as_deref();

        if paragraph.runs.is_empty() {
            if let Some(completion) = ai_completion.filter(|_| paragraph.range.start == cursor_index) {
                let ghost_style = CharacterStyle {
                    text_color: egui::Color32::GRAY.gamma_multiply(0.8),
                    highlight_color: Color32::TRANSPARENT,
                    ..default_style
                };
                job.append(completion, 0.0, text_format(ghost_style, canvas.zoom));
            } else {
                job.append("", 0.0, text_format(default_style, canvas.zoom));
            }
        } else {
            let mut current_char_index = paragraph.range.start;
            for run in &paragraph.runs {
                let run_len = run.text.chars().count();
                if let Some(completion) = ai_completion {
                    if current_char_index <= cursor_index && cursor_index <= current_char_index + run_len {
                        let split_idx = cursor_index - current_char_index;
                        let before_text: String = run.text.chars().take(split_idx).collect();
                        let after_text: String = run.text.chars().skip(split_idx).collect();

                        if !before_text.is_empty() {
                            let mut before_run = run.clone();
                            before_run.text = before_text;
                            append_run_with_placeholders(&mut job, &before_run, canvas.zoom, paragraph.image.is_some() && !has_visible_text);
                        }

                        let ghost_style = CharacterStyle {
                            text_color: egui::Color32::GRAY.gamma_multiply(0.8),
                            highlight_color: Color32::TRANSPARENT,
                            ..run.style
                        };
                        job.append(completion, 0.0, text_format(ghost_style, canvas.zoom));

                        if !after_text.is_empty() {
                            let mut after_run = run.clone();
                            after_run.text = after_text;
                            append_run_with_placeholders(&mut job, &after_run, canvas.zoom, paragraph.image.is_some() && !has_visible_text);
                        }
                    } else {
                        append_run_with_placeholders(&mut job, run, canvas.zoom, paragraph.image.is_some() && !has_visible_text);
                    }
                } else {
                    append_run_with_placeholders(&mut job, run, canvas.zoom, paragraph.image.is_some() && !has_visible_text);
                }
                current_char_index += run_len;
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
                    table,
                });
            }
        } else if let Some(image) = paragraph.image.clone().filter(|_| !has_visible_text) {
            let wrap_mode = image.wrap_mode;
            let image_offset_x_points = image.offset_x_points();
            let image_offset_y_points = image.offset_y_points();
            let display_size = super::image::image_display_size(&image, paragraph_wrap_width, canvas.zoom);
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

            if super::image_wrap::wrap_uses_side_flow(wrap_mode) {
                let pad = super::image_wrap::side_wrap_pad(wrap_mode, canvas.zoom);
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
    run: &TextRun,
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
                crate::layout::document_points_to_screen_points(line_spacing.value.max(0.0), zoom),
            ),
            LineSpacingKind::ExactPoints => {
                crate::layout::document_points_to_screen_points(line_spacing.value.max(0.0), zoom)
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

pub fn align_paragraph_galley(
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
