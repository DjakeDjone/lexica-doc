use std::path::PathBuf;
use eframe::egui;

use crate::document::DocumentState;
use crate::app::palette::ThemePalette;
use super::{
    backstage_mid_surface, backstage_two_line_row, file_name_with_extension,
    BackstageLocation, BackstageOutput, BackstageSection, BackstageState, SaveFormat,
};

pub fn paint_backstage_locations(
    ui: &mut egui::Ui,
    state: &mut BackstageState,
    output: &mut BackstageOutput,
    width: f32,
    height: f32,
    palette: ThemePalette,
) {
    egui::Frame::new()
        .fill(backstage_mid_surface(palette))
        .inner_margin(egui::Margin::symmetric(18, 22))
        .stroke(egui::Stroke::new(1.0, palette.border))
        .show(ui, |ui| {
            ui.set_width(width);
            ui.set_min_height(height);
            if state.section == BackstageSection::Open {
                paint_backstage_open_locations(ui, output, width, palette);
                return;
            }

            ui.heading(
                egui::RichText::new("Save As")
                    .size(28.0)
                    .color(palette.text_primary),
            );
            ui.add_space(20.0);
            for location in BackstageLocation::ALL {
                let selected = state.location == location;
                match location {
                    BackstageLocation::Browse => {
                        if backstage_two_line_row(
                            ui,
                            location.label(),
                            "Open the system Save As dialog",
                            selected,
                            true,
                            palette,
                        )
                        .clicked()
                        {
                            state.location = BackstageLocation::Browse;
                            output.save_as_requested = true;
                        }
                    }
                    BackstageLocation::ThisPc => {
                        if backstage_two_line_row(
                            ui,
                            location.label(),
                            &location_subtitle(location, &state.local_dir),
                            selected,
                            true,
                            palette,
                        )
                        .clicked()
                        {
                            state.location = location;
                        }
                    }
                }
                ui.add_space(4.0);
            }
        });
}

pub fn paint_backstage_open_locations(
    ui: &mut egui::Ui,
    output: &mut BackstageOutput,
    width: f32,
    palette: ThemePalette,
) {
    ui.heading(
        egui::RichText::new("Open")
            .size(28.0)
            .color(palette.text_primary),
    );
    ui.add_space(20.0);
    let _ = backstage_two_line_row(
        ui,
        "Recent",
        "Recently opened and saved files",
        true,
        true,
        palette,
    );
    ui.add_space(4.0);
    if backstage_two_line_row(
        ui,
        "Browse",
        "Open the system file dialog",
        false,
        true,
        palette,
    )
    .clicked()
    {
        output.open_requested = true;
    }
    let _ = width;
}

pub fn paint_folder_contents(
    ui: &mut egui::Ui,
    state: &mut BackstageState,
    document: &DocumentState,
    width: f32,
    palette: ThemePalette,
) {
    super::recent::folder_header(ui, width, palette);

    #[cfg(target_arch = "wasm32")]
    {
        ui.label(
            egui::RichText::new("Local folder browsing is unavailable in the web build.")
                .size(12.0)
                .color(palette.text_muted),
        );
        let _ = (state, document);
    }

    #[cfg(not(target_arch = "wasm32"))]
    {
        if state.local_dir.is_none() {
            state.local_dir = std::env::current_dir().ok();
        }
        let Some(dir) = state.local_dir.clone() else {
            ui.label(
                egui::RichText::new("No local folder is available.")
                    .size(12.0)
                    .color(palette.text_muted),
            );
            return;
        };

        egui::ScrollArea::vertical()
            .auto_shrink([false, false])
            .max_height((ui.available_height() - 8.0).max(120.0))
            .show(ui, |ui| {
                if let Some(parent) = dir.parent() {
                    if super::recent::folder_row(ui, "..", "Parent folder", true, width, palette).clicked() {
                        state.local_dir = Some(parent.to_path_buf());
                    }
                }

                let mut entries = super::recent::folder_entries(&dir);
                if entries.is_empty() {
                    ui.label(
                        egui::RichText::new("This folder is empty.")
                            .size(12.0)
                            .color(palette.text_muted),
                    );
                }
                entries.truncate(80);
                for entry in entries {
                    if super::recent::folder_row(
                        ui,
                        &entry.name,
                        &entry.modified,
                        entry.is_dir,
                        width,
                        palette,
                    )
                    .clicked()
                    {
                        if entry.is_dir {
                            state.local_dir = Some(entry.path);
                        } else {
                            state.file_name = entry.name;
                            if let Some(format) = state
                                .file_name
                                .rsplit_once('.')
                                .and_then(|(_, extension)| SaveFormat::from_extension(extension))
                            {
                                state.format = format;
                            } else {
                                state.file_name = file_name_with_extension(
                                    &state.file_name,
                                    state.format.extension(),
                                );
                            }
                        }
                    }
                }
            });
        let _ = document;
    }
}

fn location_subtitle(location: BackstageLocation, local_dir: &Option<PathBuf>) -> String {
    match location {
        BackstageLocation::ThisPc => local_dir
            .as_ref()
            .and_then(|path| path.file_name().and_then(|name| name.to_str()))
            .map(|name| format!("Local folders - {name}"))
            .unwrap_or_else(|| "Local folders".to_owned()),
        BackstageLocation::Browse => "Open the system Save As dialog".to_owned(),
    }
}
