use std::sync::Arc;

use eframe::egui::{
    self,
    epaint::text::cursor::CCursor,
    epaint::CornerRadius,
    text_selection::visuals::{paint_text_cursor, paint_text_selection},
    Align2, Color32, FontFamily, FontId, Rect, Stroke, StrokeKind,
};

use crate::{
    app::{CanvasState, TableResizeHandleRect, TableResizeKind},
    document::{text_format, DocumentImage, DocumentTable, TableCell, OBJECT_REPLACEMENT_CHAR},
    layout::document_points_to_screen_points,
};

use super::{image_display_size, texture_for_image};

pub(super) struct TablePaintGeometry {
    pub(super) cell_rects: Vec<(usize, usize, usize, Rect)>,
    pub(super) cell_content_rects: Vec<(usize, usize, usize, Rect)>,
    pub(super) resize_handles: Vec<TableResizeHandleRect>,
}

#[derive(Clone, Copy)]
pub(super) struct TablePaintParams {
    pub(super) origin: egui::Pos2,
    pub(super) zoom: f32,
    pub(super) active_cell: Option<(usize, usize, usize)>,
    pub(super) time: f64,
}

pub(super) fn paint_table(
    ui: &mut egui::Ui,
    canvas: &mut CanvasState,
    painter: &egui::Painter,
    table: &DocumentTable,
    params: TablePaintParams,
) -> TablePaintGeometry {
    let TablePaintParams {
        origin,
        zoom,
        active_cell,
        time,
    } = params;
    let border_width = document_points_to_screen_points(table.borders.width_points, zoom);
    let border_stroke = Stroke::new(border_width.max(0.5), table.borders.color);
    let cell_padding = document_points_to_screen_points(4.0, zoom);
    let col_widths: Vec<f32> = table
        .col_widths_points
        .iter()
        .map(|w| document_points_to_screen_points(*w, zoom))
        .collect();
    let mut geometry = TablePaintGeometry {
        cell_rects: Vec::new(),
        cell_content_rects: Vec::new(),
        resize_handles: Vec::new(),
    };

    let actual_row_heights = table_row_heights_screen(painter, table, zoom);
    let total_width: f32 = col_widths.iter().sum();
    let total_height: f32 = actual_row_heights.iter().sum();

    let table_rect = Rect::from_min_size(origin, egui::vec2(total_width, total_height));
    painter.rect_filled(table_rect, CornerRadius::ZERO, Color32::WHITE);

    let mut col_x = Vec::with_capacity(col_widths.len() + 1);
    col_x.push(origin.x);
    for width in &col_widths {
        col_x.push(col_x.last().copied().unwrap_or(origin.x) + *width);
    }
    let mut row_y = Vec::with_capacity(actual_row_heights.len() + 1);
    row_y.push(origin.y);
    for height in &actual_row_heights {
        row_y.push(row_y.last().copied().unwrap_or(origin.y) + *height);
    }

    let mut covered = vec![vec![false; table.num_cols()]; table.num_rows()];
    for (row_idx, row) in table.rows.iter().enumerate() {
        let row_height = actual_row_heights[row_idx];
        let is_header = row_idx == 0;
        let y = row_y[row_idx];

        if is_header {
            let row_rect =
                Rect::from_min_size(egui::pos2(origin.x, y), egui::vec2(total_width, row_height));
            painter.rect_filled(
                row_rect,
                CornerRadius::ZERO,
                Color32::from_rgb(240, 243, 248),
            );
        } else if row_idx % 2 == 0 {
            let row_rect =
                Rect::from_min_size(egui::pos2(origin.x, y), egui::vec2(total_width, row_height));
            painter.rect_filled(
                row_rect,
                CornerRadius::ZERO,
                Color32::from_rgb(250, 251, 253),
            );
        }

        for (col_idx, cell) in row.iter().enumerate() {
            if col_idx >= table.num_cols()
                || cell.col_span == 0
                || cell.row_span == 0
                || covered[row_idx][col_idx]
            {
                continue;
            }
            let col_span = cell.col_span.max(1) as usize;
            let row_span = cell.row_span.max(1) as usize;
            let end_col = (col_idx + col_span).min(table.num_cols());
            let end_row = (row_idx + row_span).min(table.num_rows());
            for covered_row in covered.iter_mut().take(end_row).skip(row_idx) {
                for covered_cell in covered_row.iter_mut().take(end_col).skip(col_idx) {
                    *covered_cell = true;
                }
            }

            let x = col_x[col_idx];
            let col_width = col_x[end_col] - x;
            let row_height = row_y[end_row] - y;
            let cell_rect =
                Rect::from_min_size(egui::pos2(x, y), egui::vec2(col_width, row_height));
            geometry
                .cell_rects
                .push((table.id, row_idx, col_idx, cell_rect));
            painter.rect_stroke(
                cell_rect,
                CornerRadius::ZERO,
                border_stroke,
                StrokeKind::Inside,
            );

            let available_width = (col_width - cell_padding * 2.0).max(1.0);
            let text_pos = egui::pos2(x + cell_padding, y + cell_padding);
            let content_rect = Rect::from_min_size(
                text_pos,
                egui::vec2(available_width, row_height - cell_padding * 2.0),
            );
            geometry
                .cell_content_rects
                .push((table.id, row_idx, col_idx, content_rect));
            let mut galley = table_cell_text_galley(painter, cell, available_width, zoom);
            let text_height = galley.rect.height();

            if active_cell == Some((table.id, row_idx, col_idx))
                && !canvas.table_cell_selection.is_empty()
            {
                paint_text_selection(
                    &mut galley,
                    ui.visuals(),
                    &canvas.table_cell_selection,
                    None,
                );
            }
            painter.with_clip_rect(content_rect).galley(
                text_pos,
                galley.clone(),
                Color32::TRANSPARENT,
            );

            paint_table_cell_images(ui, canvas, painter, cell, content_rect, text_height, zoom);

            if active_cell == Some((table.id, row_idx, col_idx)) {
                let focus_color = Color32::from_rgb(54, 116, 206);
                painter.rect_stroke(
                    cell_rect.shrink(1.0),
                    CornerRadius::ZERO,
                    Stroke::new(2.0, focus_color),
                    StrokeKind::Inside,
                );
                if let Some(caret_rect) = table_cell_caret_rect(
                    &galley,
                    canvas.table_cell_selection.primary,
                    text_pos,
                    zoom,
                ) {
                    paint_text_cursor(
                        ui,
                        painter,
                        caret_rect.intersect(content_rect),
                        time - canvas.last_interaction_time,
                    );
                }
            }
        }
    }

    painter.rect_stroke(
        table_rect,
        CornerRadius::ZERO,
        border_stroke,
        StrokeKind::Outside,
    );

    for col in 0..col_widths.len().saturating_sub(1) {
        let x = col_x[col + 1];
        geometry.resize_handles.push(TableResizeHandleRect {
            table_id: table.id,
            kind: TableResizeKind::Column { left_col: col },
            rect: Rect::from_center_size(
                egui::pos2(x, table_rect.center().y),
                egui::vec2(8.0, table_rect.height()),
            ),
        });
    }
    for row in 0..actual_row_heights.len().saturating_sub(1) {
        let y = row_y[row + 1];
        geometry.resize_handles.push(TableResizeHandleRect {
            table_id: table.id,
            kind: TableResizeKind::Row { top_row: row },
            rect: Rect::from_center_size(
                egui::pos2(table_rect.center().x, y),
                egui::vec2(table_rect.width(), 8.0),
            ),
        });
    }

    geometry
}

fn span_sum(values: &[f32], start: usize, span: usize) -> f32 {
    values.iter().skip(start).take(span).sum()
}

fn table_cell_caret_rect(
    galley: &egui::Galley,
    cursor: CCursor,
    text_pos: egui::Pos2,
    zoom: f32,
) -> Option<Rect> {
    let mut rect = galley.pos_from_cursor(cursor).translate(text_pos.to_vec2());
    if let Some(row) = galley.rows.get(galley.layout_from_cursor(cursor).row) {
        rect.min.y = text_pos.y + row.min_y();
        rect.max.y = text_pos.y + row.max_y();
    } else {
        rect.max.y = rect.min.y + document_points_to_screen_points(14.0, zoom);
    }
    Some(rect.expand2(egui::vec2(0.75, 0.75)))
}

pub(super) fn table_cell_text_galley(
    painter: &egui::Painter,
    cell: &TableCell,
    available_width: f32,
    zoom: f32,
) -> Arc<egui::Galley> {
    let mut job = egui::epaint::text::LayoutJob::default();
    job.wrap.max_width = available_width;
    job.break_on_newline = true;

    for run in &cell.runs {
        let text: String = run
            .text
            .chars()
            .filter(|ch| *ch != OBJECT_REPLACEMENT_CHAR)
            .collect();
        if !text.is_empty() {
            job.append(&text, 0.0, text_format(run.style, zoom));
        }
    }

    painter.layout_job(job)
}

fn table_cell_image_size(image: &DocumentImage, available_width: f32, zoom: f32) -> egui::Vec2 {
    let raw = image_display_size(image, available_width, zoom);
    if raw.x <= available_width {
        return raw;
    }

    let scale = available_width / raw.x.max(1.0);
    egui::vec2(available_width, (raw.y * scale).max(1.0))
}

fn paint_table_cell_images(
    ui: &mut egui::Ui,
    canvas: &mut CanvasState,
    painter: &egui::Painter,
    cell: &TableCell,
    content_rect: Rect,
    text_height: f32,
    zoom: f32,
) {
    let mut y = content_rect.top() + text_height;
    if text_height > 0.0 && !cell.images.is_empty() {
        y += document_points_to_screen_points(3.0, zoom);
    }

    for image in &cell.images {
        let image_size = table_cell_image_size(image, content_rect.width(), zoom);
        if y + image_size.y > content_rect.bottom() {
            break;
        }
        let rect = Rect::from_min_size(egui::pos2(content_rect.left(), y), image_size);
        if let Some(texture) = texture_for_image(ui.ctx(), canvas, image) {
            let alpha = (image.opacity * 255.0).round().clamp(0.0, 255.0) as u8;
            painter.image(
                texture.id(),
                rect,
                Rect::from_min_max(egui::Pos2::ZERO, egui::pos2(1.0, 1.0)),
                Color32::from_white_alpha(alpha),
            );
        } else {
            painter.rect_stroke(
                rect,
                CornerRadius::same(2),
                Stroke::new(1.0, Color32::from_rgb(150, 150, 150)),
                StrokeKind::Inside,
            );
            painter.text(
                rect.center(),
                Align2::CENTER_CENTER,
                &image.alt_text,
                FontId::new(11.0 * zoom, FontFamily::Proportional),
                Color32::from_rgb(50, 53, 60),
            );
        }
        y += image_size.y + document_points_to_screen_points(3.0, zoom);
    }
}

pub(super) fn table_row_heights_screen(
    painter: &egui::Painter,
    table: &DocumentTable,
    zoom: f32,
) -> Vec<f32> {
    let cell_padding = document_points_to_screen_points(4.0, zoom);
    let default_row_height = document_points_to_screen_points(20.0, zoom);
    let col_widths: Vec<f32> = table
        .col_widths_points
        .iter()
        .map(|w| document_points_to_screen_points(*w, zoom))
        .collect();
    let mut row_heights: Vec<f32> = table
        .row_heights_points
        .iter()
        .map(|height| document_points_to_screen_points(*height, zoom).max(default_row_height))
        .collect();
    row_heights.resize(table.num_rows(), default_row_height);

    for (row_idx, row) in table.rows.iter().enumerate() {
        for (col_idx, cell) in row.iter().enumerate() {
            if cell.col_span == 0 || cell.row_span == 0 {
                continue;
            }
            let col_width = span_sum(&col_widths, col_idx, cell.col_span.max(1) as usize);
            let available_width = (col_width - cell_padding * 2.0).max(1.0);
            let galley = table_cell_text_galley(painter, cell, available_width, zoom);
            let mut required = galley.rect.height() + cell_padding * 2.0;
            if !cell.images.is_empty() {
                required += document_points_to_screen_points(3.0, zoom);
            }
            for image in &cell.images {
                let image_size = table_cell_image_size(image, available_width, zoom);
                required += image_size.y + document_points_to_screen_points(3.0, zoom);
            }
            let row_span = cell.row_span.max(1) as usize;
            let end_row = (row_idx + row_span).min(row_heights.len());
            let current: f32 = row_heights[row_idx..end_row].iter().sum();
            if required > current && end_row > row_idx {
                let extra_each = (required - current) / (end_row - row_idx) as f32;
                for height in &mut row_heights[row_idx..end_row] {
                    *height += extra_each;
                }
            }
        }
    }

    row_heights
}
