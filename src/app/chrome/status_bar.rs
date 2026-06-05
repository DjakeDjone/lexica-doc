use eframe::egui;

use crate::document::{DocumentState, OBJECT_REPLACEMENT_CHAR};
use crate::grammar::GrammarStatus;
use crate::app::{CanvasState, palette::ThemePalette};

pub(crate) fn paint_status_bar(
    ui: &mut egui::Ui,
    document: &DocumentState,
    canvas: &CanvasState,
    status_message: &str,
    grammar_status: &GrammarStatus,
    grammar_issue_count: usize,
    palette: ThemePalette,
) {
    ui.horizontal(|ui| {
        let plain_text: String = document
            .plain_text()
            .chars()
            .filter(|ch| *ch != OBJECT_REPLACEMENT_CHAR)
            .collect();
        let word_count = plain_text.split_whitespace().count();
        ui.label(
            egui::RichText::new("Page 1")
                .size(11.0)
                .color(palette.text_muted),
        );
        ui.separator();
        ui.label(
            egui::RichText::new(format!("{word_count} words"))
                .size(11.0)
                .color(palette.text_muted),
        );
        ui.separator();
        ui.label(
            egui::RichText::new(status_message)
                .size(11.0)
                .color(palette.text_primary),
        );
        ui.separator();
        match grammar_status {
            GrammarStatus::Idle => {
                ui.label(
                    egui::RichText::new("Grammar idle")
                        .size(11.0)
                        .color(palette.text_muted),
                );
            }
            GrammarStatus::Checking => {
                ui.spinner();
                ui.label(
                    egui::RichText::new("Checking grammar…")
                        .size(11.0)
                        .color(palette.text_muted),
                );
                ui.ctx().request_repaint();
            }
            GrammarStatus::Done => {
                let text = if grammar_issue_count == 0 {
                    "No issues".to_owned()
                } else if grammar_issue_count == 1 {
                    "1 issue".to_owned()
                } else {
                    format!("{grammar_issue_count} issues")
                };
                ui.label(
                    egui::RichText::new(text)
                        .size(11.0)
                        .color(palette.text_muted),
                );
            }
            GrammarStatus::Unavailable(message) => {
                let short_message: String = message.chars().take(42).collect();
                ui.label(
                    egui::RichText::new("⚠")
                        .size(12.0)
                        .color(egui::Color32::from_rgb(194, 87, 0)),
                );
                ui.label(
                    egui::RichText::new(short_message)
                        .size(11.0)
                        .color(egui::Color32::from_rgb(194, 87, 0)),
                );
            }
        }
        ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
            ui.label(
                egui::RichText::new(format!("{:.0}%", canvas.zoom * 100.0))
                    .size(11.0)
                    .color(palette.text_muted),
            );
            ui.separator();
            let setup = document.default_page_setup();
            ui.label(
                egui::RichText::new(format!(
                    "{:.0} x {:.0} pt",
                    setup.page_size.width_points, setup.page_size.height_points
                ))
                .size(11.0)
                .color(palette.text_muted),
            );
        });
    });
}

pub(crate) fn today_label() -> String {
    #[cfg(target_arch = "wasm32")]
    {
        let date = js_sys::Date::new_0();
        return format!(
            "{:04}-{:02}-{:02}",
            date.get_full_year(),
            date.get_month() + 1,
            date.get_date()
        );
    }

    #[cfg(not(target_arch = "wasm32"))]
    {
        use std::time::{SystemTime, UNIX_EPOCH};

        let days = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .map(|duration| (duration.as_secs() / 86_400) as i64)
            .unwrap_or(0);
        let (year, month, day) = civil_from_days(days);
        format!("{year:04}-{month:02}-{day:02}")
    }
}

#[cfg(not(target_arch = "wasm32"))]
pub(crate) fn civil_from_days(days_since_unix_epoch: i64) -> (i32, u32, u32) {
    let z = days_since_unix_epoch + 719_468;
    let era = if z >= 0 { z } else { z - 146_096 } / 146_097;
    let doe = z - era * 146_097;
    let yoe = (doe - doe / 1460 + doe / 36_524 - doe / 146_096) / 365;
    let y = yoe + era * 400;
    let doy = doe - (365 * yoe + yoe / 4 - yoe / 100);
    let mp = (5 * doy + 2) / 153;
    let day = doy - (153 * mp + 2) / 5 + 1;
    let month = mp + if mp < 10 { 3 } else { -9 };
    let year = y + if month <= 2 { 1 } else { 0 };
    (year as i32, month as u32, day as u32)
}
