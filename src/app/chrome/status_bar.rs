use eframe::egui;

use crate::app::{palette::ThemePalette, CanvasState, ZoomMode};
use crate::document::{DocumentState, OBJECT_REPLACEMENT_CHAR};
use crate::grammar::GrammarStatus;

#[allow(clippy::too_many_arguments)]
pub(crate) fn paint_status_bar(
    ui: &mut egui::Ui,
    document: &DocumentState,
    canvas: &mut CanvasState,
    current_page: usize,
    page_count: usize,
    status_message: &str,
    grammar_status: &GrammarStatus,
    grammar_issue_count: usize,
    ai_config: &mut crate::app::settings::OllamaSettings,
    palette: ThemePalette,
) {
    ui.horizontal(|ui| {
        let plain_text: String = document
            .plain_text()
            .chars()
            .filter(|ch| *ch != OBJECT_REPLACEMENT_CHAR)
            .collect();
        let word_count = plain_text.split_whitespace().count();
        let (line, column) = cursor_line_column(document, canvas.selection.primary.index);
        let selection_label = selection_stats(document, canvas)
            .map(|stats| {
                format!(
                    " | {} selected words, {} chars",
                    stats.word_count, stats.char_count
                )
            })
            .unwrap_or_default();
        ui.label(
            egui::RichText::new(format!(
                "Page {} of {}",
                current_page.max(1),
                page_count.max(1)
            ))
            .size(11.0)
            .color(palette.text_muted),
        );
        ui.separator();
        ui.label(
            egui::RichText::new(format!(
                "{word_count} words{selection_label} | Ln {line}, Col {column}"
            ))
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
            if ui.small_button("+").on_hover_text("Zoom in").clicked() {
                canvas.zoom_mode = ZoomMode::Manual;
                canvas.zoom = (canvas.zoom + 0.1).min(3.0);
            }
            ui.label(
                egui::RichText::new(format!("{:.0}%", canvas.zoom * 100.0))
                    .size(11.0)
                    .color(palette.text_muted),
            );
            if ui.small_button("−").on_hover_text("Zoom out").clicked() {
                canvas.zoom_mode = ZoomMode::Manual;
                canvas.zoom = (canvas.zoom - 0.1).max(0.5);
            }
            if ui
                .selectable_label(canvas.zoom_mode == ZoomMode::FitPage, "Fit")
                .on_hover_text("Fit page width")
                .clicked()
            {
                canvas.zoom_mode = ZoomMode::FitPage;
            }
            ui.separator();
            let ai_label = if ai_config.enable { "AI On" } else { "AI Off" };
            if ui
                .selectable_label(
                    ai_config.enable,
                    egui::RichText::new(ai_label)
                        .size(11.0)
                        .color(palette.text_muted),
                )
                .clicked()
            {
                ai_config.enable = !ai_config.enable;
            }
        });
    });
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct SelectionStats {
    pub word_count: usize,
    pub char_count: usize,
}

pub(crate) fn selection_stats(
    document: &DocumentState,
    canvas: &CanvasState,
) -> Option<SelectionStats> {
    let range = canvas.selection.as_sorted_char_range();
    if range.start >= range.end {
        return None;
    }
    let selected: String = document
        .selected_text(range)
        .chars()
        .filter(|ch| *ch != OBJECT_REPLACEMENT_CHAR)
        .collect();
    Some(SelectionStats {
        word_count: selected.split_whitespace().count(),
        char_count: selected.chars().count(),
    })
}

pub(crate) fn cursor_line_column(document: &DocumentState, cursor: usize) -> (usize, usize) {
    let mut line = 1usize;
    let mut column = 1usize;
    for ch in document
        .plain_text()
        .chars()
        .take(cursor.min(document.total_chars()))
    {
        if ch == '\n' {
            line += 1;
            column = 1;
        } else {
            column += 1;
        }
    }
    (line, column)
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::document::{CharacterStyle, TextRun};
    use eframe::egui::{epaint::text::cursor::CCursor, text_selection::CCursorRange};

    #[test]
    fn cursor_line_column_counts_from_one() {
        let mut document = DocumentState::bootstrap();
        document.replace_with_runs(
            "Test".to_owned(),
            vec![TextRun {
                text: "one\ntwo".to_owned(),
                style: CharacterStyle::default(),
            }],
        );

        assert_eq!(cursor_line_column(&document, 0), (1, 1));
        assert_eq!(cursor_line_column(&document, 4), (2, 1));
        assert_eq!(cursor_line_column(&document, 6), (2, 3));
    }

    #[test]
    fn selection_stats_ignore_object_replacement_characters() {
        let mut document = DocumentState::bootstrap();
        document.replace_with_runs(
            "Test".to_owned(),
            vec![TextRun {
                text: format!("one {OBJECT_REPLACEMENT_CHAR} two"),
                style: CharacterStyle::default(),
            }],
        );
        let mut canvas = CanvasState::default();
        canvas.selection = CCursorRange::two(CCursor::new(0), CCursor::new(document.total_chars()));

        assert_eq!(
            selection_stats(&document, &canvas),
            Some(SelectionStats {
                word_count: 2,
                char_count: 8
            })
        );
    }

    #[test]
    fn empty_selection_has_no_selection_stats() {
        let document = DocumentState::bootstrap();
        let canvas = CanvasState::default();

        assert_eq!(selection_stats(&document, &canvas), None);
    }
}
