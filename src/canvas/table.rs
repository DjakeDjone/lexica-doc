use std::sync::Arc;

use eframe::egui::{
    self,
    epaint::text::cursor::CCursor,
    epaint::CornerRadius,
    text_selection::visuals::{paint_text_cursor, paint_text_selection},
    text_selection::CCursorRange,
    Align2, Color32, FontFamily, FontId, Rect, Stroke, StrokeKind,
};

use crate::{
    app::{CanvasState, ChangeHistory, TableResizeDrag, TableResizeHandleRect, TableResizeKind},
    document::{
        text_format, DocumentImage, DocumentState, DocumentTable, TableCell,
        OBJECT_REPLACEMENT_CHAR,
    },
    layout::document_points_to_screen_points,
};

use super::image::{image_display_size, texture_for_image};

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

    for col in 0..col_widths.len() {
        let x = col_x[col + 1];
        geometry.resize_handles.push(TableResizeHandleRect {
            table_id: table.id,
            kind: TableResizeKind::Column { left_col: col },
            start_points: table.col_widths_points[col],
            rect: Rect::from_center_size(
                egui::pos2(x, table_rect.center().y),
                egui::vec2(8.0, table_rect.height()),
            ),
        });
    }
    for row in 0..actual_row_heights.len() {
        let y = row_y[row + 1];
        geometry.resize_handles.push(TableResizeHandleRect {
            table_id: table.id,
            kind: TableResizeKind::Row { top_row: row },
            start_points: actual_row_heights[row] / zoom.max(0.01),
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

pub(super) fn table_cell_hit(
    canvas: &CanvasState,
    pointer_pos: egui::Pos2,
) -> Option<(usize, usize, usize)> {
    canvas
        .table_cell_rects
        .iter()
        .rev()
        .find(|(_, _, _, rect)| rect.contains(pointer_pos))
        .map(|(table_id, row, col, _)| (*table_id, *row, *col))
}

pub(super) fn table_cell_content_rect(
    canvas: &CanvasState,
    cell: (usize, usize, usize),
) -> Option<egui::Rect> {
    canvas
        .table_cell_content_rects
        .iter()
        .find(|(table_id, row, col, _)| (*table_id, *row, *col) == cell)
        .map(|(_, _, _, rect)| *rect)
}

pub(super) fn table_cell_cursor_from_pointer(
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

pub(super) fn table_resize_handle_hit(
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

pub(super) fn handle_table_interaction(
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

    let Some(pointer_pos) = response.interact_pointer_pos() else {
        return (false, document_changed);
    };

    if response.drag_started() {
        let start_ptr = pointer_pos - response.total_drag_delta().unwrap_or_default();
        if let Some(handle) = table_resize_handle_hit(canvas, start_ptr) {
            if let Some(table) = document.table_by_id(handle.table_id) {
                let dimensions = match handle.kind {
                    TableResizeKind::Column { left_col } => {
                        if left_col >= table.col_widths_points.len() {
                            None
                        } else if left_col + 1 == table.col_widths_points.len() {
                            Some((table.col_widths_points[left_col], 0.0))
                        } else {
                            Some((
                                table.col_widths_points[left_col],
                                table.col_widths_points[left_col + 1],
                            ))
                        }
                    }
                    TableResizeKind::Row { top_row } => {
                        if top_row >= table.row_heights_points.len() {
                            None
                        } else {
                            Some((handle.start_points, 0.0))
                        }
                    }
                };
                if let Some((first_points, second_points)) = dimensions {
                    history.checkpoint(document, ui.input(|i| i.time));
                    canvas.table_resize_drag = Some(TableResizeDrag {
                        table_id: handle.table_id,
                        kind: handle.kind,
                        start_ptr,
                        first_points,
                        second_points,
                    });
                    canvas.active_table_cell = None;
                    canvas.selected_image_id = None;
                    return (true, document_changed);
                }
            }
        }
    }

    if response.dragged() {
        if let Some(drag) = canvas.table_resize_drag.as_ref() {
            if let Some(pointer_pos) = response.interact_pointer_pos() {
                let zoom = canvas.zoom.max(0.01);
                match drag.kind {
                    TableResizeKind::Column { left_col } => {
                        let delta = (pointer_pos.x - drag.start_ptr.x) / zoom;
                        if drag.second_points == 0.0 {
                            if let Some(width) = document
                                .table_by_id_mut(drag.table_id)
                                .and_then(|table| table.col_widths_points.get_mut(left_col))
                            {
                                *width = (drag.first_points + delta).max(MIN_SIZE_POINTS);
                            }
                        } else {
                            let total = drag.first_points + drag.second_points;
                            let first = (drag.first_points + delta)
                                .clamp(MIN_SIZE_POINTS, total - MIN_SIZE_POINTS);
                            let second = total - first;
                            document.resize_table_column_pair(
                                drag.table_id,
                                left_col,
                                first,
                                second,
                            );
                        }
                    }
                    TableResizeKind::Row { top_row } => {
                        let delta = (pointer_pos.y - drag.start_ptr.y) / zoom;
                        if let Some(height) = document
                            .table_by_id_mut(drag.table_id)
                            .and_then(|table| table.row_heights_points.get_mut(top_row))
                        {
                            *height = (drag.first_points + delta).max(12.0);
                        }
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

    if response.drag_started() {
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn outer_edges_resize_the_table_width_and_height() {
        let ctx = egui::Context::default();
        let mut document = DocumentState::bootstrap();
        document.insert_table(0, 2, 2);
        let table_id = document
            .paragraph_tables
            .iter()
            .flatten()
            .next()
            .unwrap()
            .id;
        document
            .table_by_id_mut(table_id)
            .unwrap()
            .row_heights_points = vec![40.0, 40.0];
        let table = document.table_by_id(table_id).unwrap().clone();
        let original_width = table.total_width_points();
        let original_height = table.total_height_points();
        let table_origin = egui::pos2(20.0, 20.0);
        let horizontal_handle = egui::pos2(table_origin.x + original_width, table_origin.y + 10.0);
        let vertical_handle = egui::pos2(table_origin.x + 10.0, table_origin.y + original_height);
        let mut canvas = CanvasState::default();
        canvas.active_table_cell = Some((table_id, 0, 0));
        let mut history = ChangeHistory::new();
        let canvas_rect = Rect::from_min_size(egui::Pos2::ZERO, egui::vec2(700.0, 300.0));

        for (time, events) in [
            (-0.1, vec![]),
            (
                0.0,
                vec![
                    egui::Event::PointerMoved(horizontal_handle),
                    egui::Event::PointerButton {
                        pos: horizontal_handle,
                        button: egui::PointerButton::Primary,
                        pressed: true,
                        modifiers: egui::Modifiers::NONE,
                    },
                ],
            ),
            (
                0.1,
                vec![egui::Event::PointerMoved(
                    horizontal_handle + egui::vec2(12.0, 0.0),
                )],
            ),
            (
                0.2,
                vec![egui::Event::PointerMoved(
                    horizontal_handle + egui::vec2(20.0, 0.0),
                )],
            ),
            (
                0.3,
                vec![egui::Event::PointerButton {
                    pos: horizontal_handle + egui::vec2(20.0, 0.0),
                    button: egui::PointerButton::Primary,
                    pressed: false,
                    modifiers: egui::Modifiers::NONE,
                }],
            ),
            (
                0.4,
                vec![
                    egui::Event::PointerMoved(vertical_handle),
                    egui::Event::PointerButton {
                        pos: vertical_handle,
                        button: egui::PointerButton::Primary,
                        pressed: true,
                        modifiers: egui::Modifiers::NONE,
                    },
                ],
            ),
            (
                0.5,
                vec![egui::Event::PointerMoved(
                    vertical_handle + egui::vec2(0.0, 12.0),
                )],
            ),
            (
                0.6,
                vec![egui::Event::PointerMoved(
                    vertical_handle + egui::vec2(0.0, 20.0),
                )],
            ),
        ] {
            let input = egui::RawInput {
                time: Some(time),
                events,
                ..Default::default()
            };
            let _ = ctx.run_ui(input, |ui| {
                let painter = ui.painter().clone();
                let geometry = paint_table(
                    ui,
                    &mut canvas,
                    &painter,
                    &table,
                    TablePaintParams {
                        origin: table_origin,
                        zoom: 1.0,
                        active_cell: None,
                        time,
                    },
                );
                canvas.table_resize_handles = geometry.resize_handles;
                let response = ui.interact(
                    canvas_rect,
                    egui::Id::new("table_resize_test"),
                    egui::Sense::click_and_drag(),
                );
                handle_table_interaction(ui, &response, &mut canvas, &mut document, &mut history);
            });
        }

        assert!(document.table_by_id(table_id).unwrap().total_width_points() > original_width);
        assert!(
            document
                .table_by_id(table_id)
                .unwrap()
                .total_height_points()
                > original_height
        );
    }
}
