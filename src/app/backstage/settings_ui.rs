use super::backstage_surface;
use crate::app::palette::ThemePalette;
use crate::app::settings::OllamaSettings;
use eframe::egui;

pub fn paint_settings(
    ui: &mut egui::Ui,
    ollama_settings: &mut OllamaSettings,
    width: f32,
    height: f32,
    palette: ThemePalette,
) {
    egui::Frame::new()
        .fill(backstage_surface(palette))
        .inner_margin(egui::Margin::symmetric(32, 24))
        .show(ui, |ui| {
            ui.set_width(width);
            ui.set_min_height(height);
            
            ui.heading(
                egui::RichText::new("Settings")
                    .size(24.0)
                    .color(palette.text_primary),
            );
            ui.add_space(24.0);

            ui.label(
                egui::RichText::new("Local AI Completions (Ollama)")
                    .size(16.0)
                    .strong()
                    .color(palette.text_primary),
            );
            ui.add_space(12.0);

            ui.horizontal(|ui| {
                ui.checkbox(&mut ollama_settings.enable, "Enable AI Completions");
            });

            ui.add_space(12.0);

            ui.add_enabled_ui(ollama_settings.enable, |ui| {
                egui::Grid::new("ollama_settings_grid")
                    .num_columns(2)
                    .spacing([16.0, 12.0])
                    .show(ui, |ui| {
                        ui.label(
                            egui::RichText::new("Ollama Endpoint")
                                .size(14.0)
                                .color(palette.text_primary),
                        );
                        ui.add(egui::TextEdit::singleline(&mut ollama_settings.endpoint).desired_width(240.0));
                        ui.end_row();

                        ui.label(
                            egui::RichText::new("Ollama Model")
                                .size(14.0)
                                .color(palette.text_primary),
                        );
                        ui.add(egui::TextEdit::singleline(&mut ollama_settings.model).desired_width(160.0));
                        ui.end_row();
                    });
            });

            ui.add_space(32.0);
            
            ui.label(
                egui::RichText::new("AI completions will appear as ghost text while you type. Press Tab to accept them.")
                    .size(13.0)
                    .color(palette.text_muted),
            );
        });
}
