use eframe::egui;

use super::common::{ribbon_group, ribbon_info_group};
use crate::app::{
    actions::{delete_table_column, delete_table_row, insert_table_column, insert_table_row},
    palette::ThemePalette,
    CanvasState, ChangeHistory,
};
use crate::document::DocumentState;

pub(crate) fn table_format_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    history: &mut ChangeHistory,
    palette: ThemePalette,
) {
    let Some((table_id, row, col)) = canvas.active_table_cell else {
        ribbon_info_group(
            ui,
            "Table Format",
            "Click a table cell to select it.",
            palette,
        );
        return;
    };

    let Some(table) = document.table_by_id(table_id).cloned() else {
        canvas.active_table_cell = None;
        return;
    };

    ribbon_group(ui, "Rows & Columns", palette, |ui| {
        if ui.button("Row Above").clicked() {
            insert_table_row(
                document,
                table_id,
                if row == 0 { usize::MAX } else { row - 1 },
                status_message,
                history,
            );
            canvas.active_table_cell = Some((table_id, row, col));
        }
        if ui.button("Row Below").clicked() {
            insert_table_row(document, table_id, row, status_message, history);
            canvas.active_table_cell = Some((table_id, row + 1, col));
        }
        ui.separator();
        if ui.button("Column Left").clicked() {
            insert_table_column(
                document,
                table_id,
                if col == 0 { usize::MAX } else { col - 1 },
                status_message,
                history,
            );
            canvas.active_table_cell = Some((table_id, row, col));
        }
        if ui.button("Column Right").clicked() {
            insert_table_column(document, table_id, col, status_message, history);
            canvas.active_table_cell = Some((table_id, row, col + 1));
        }
        ui.separator();
        if ui.button("Delete Row").clicked() {
            delete_table_row(document, table_id, row, status_message, history);
            let next_row = row.min(
                document
                    .table_by_id(table_id)
                    .map_or(1, |t| t.num_rows())
                    .saturating_sub(1),
            );
            canvas.active_table_cell = Some((table_id, next_row, col));
        }
        if ui.button("Delete Column").clicked() {
            delete_table_column(document, table_id, col, status_message, history);
            let next_col = col.min(
                document
                    .table_by_id(table_id)
                    .map_or(1, |t| t.num_cols())
                    .saturating_sub(1),
            );
            canvas.active_table_cell = Some((table_id, row, next_col));
        }
    });

    ribbon_group(ui, "Borders", palette, |ui| {
        let mut width = table.borders.width_points;
        let resp = ui.add(
            egui::DragValue::new(&mut width)
                .speed(0.1)
                .range(0.0..=8.0)
                .fixed_decimals(2)
                .suffix(" pt"),
        );
        if resp.changed() {
            let now = ui.input(|i| i.time);
            history.checkpoint_coalesced(document, now);
            document.set_table_border_width(table_id, width);
            *status_message = format!("Table border: {:.2} pt", width);
        }
        let mut color = table.borders.color;
        if ui.color_edit_button_srgba(&mut color).changed() {
            let now = ui.input(|i| i.time);
            history.checkpoint_coalesced(document, now);
            document.set_table_border_color(table_id, color);
            *status_message = "Table border color updated".to_owned();
        }
    });

    ribbon_group(ui, "Cells", palette, |ui| {
        if ui.button("Merge Right").clicked() {
            history.checkpoint(document, ui.input(|i| i.time));
            if document.merge_table_cell_right(table_id, row, col) {
                *status_message = "Cells merged".to_owned();
            }
        }
        if ui.button("Split Cell").clicked() {
            history.checkpoint(document, ui.input(|i| i.time));
            if document.split_table_cell(table_id, row, col) {
                *status_message = "Cell split".to_owned();
            }
        }
    });
}
