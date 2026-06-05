use eframe::egui;

use crate::grammar::{GrammarConfig, GrammarStatus, Language};
use crate::app::palette::ThemePalette;
use super::GrammarRibbonOutput;
use super::common::ribbon_group;

pub(crate) fn ribbon_grammer_actions_group(
    ui: &mut egui::Ui,
    grammar_status: &GrammarStatus,
    can_download_grammar: bool,
    output: &mut GrammarRibbonOutput,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Check", palette, |ui| {
        if ui.button("Check Now").clicked() {
            output.manual_check_requested = true;
        }
        if ui.button("Restart").clicked() {
            output.restart_requested = true;
        }
        if ui
            .add_enabled(can_download_grammar, egui::Button::new("Download"))
            .clicked()
        {
            output.download_requested = true;
        }

        ui.separator();
        let status_text = match grammar_status {
            GrammarStatus::Idle => "Idle".to_owned(),
            GrammarStatus::Checking => "Checking".to_owned(),
            GrammarStatus::Done => "Ready".to_owned(),
            GrammarStatus::Unavailable(message) => {
                let short: String = message.chars().take(32).collect();
                format!("Unavailable: {short}")
            }
        };
        ui.label(
            egui::RichText::new(status_text)
                .size(11.0)
                .color(palette.text_muted),
        );
    });
}

pub(crate) fn ribbon_grammer_settings_group(
    ui: &mut egui::Ui,
    grammar_config: &mut GrammarConfig,
    grammar_auto_check: &mut bool,
    output: &mut GrammarRibbonOutput,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Settings", palette, |ui| {
        if ui.checkbox(grammar_auto_check, "Auto Check").changed() {
            output.settings_changed = true;
        }

        egui::ComboBox::from_id_salt("grammar_language")
            .selected_text(match grammar_config.language {
                Language::Auto => "Auto",
                Language::EnUs => "English (US)",
                Language::DeDE => "German (DE)",
            })
            .width(140.0)
            .show_ui(ui, |ui| {
                if ui
                    .selectable_value(&mut grammar_config.language, Language::Auto, "Auto")
                    .changed()
                {
                    output.settings_changed = true;
                }
                if ui
                    .selectable_value(&mut grammar_config.language, Language::EnUs, "English (US)")
                    .changed()
                {
                    output.settings_changed = true;
                }
                if ui
                    .selectable_value(&mut grammar_config.language, Language::DeDE, "German (DE)")
                    .changed()
                {
                    output.settings_changed = true;
                }
            });

        ui.label(
            egui::RichText::new("Choose the LanguageTool input language.")
                .size(11.0)
                .color(palette.text_muted),
        );
    });
}
