use std::path::{Path, PathBuf};
#[cfg(not(target_arch = "wasm32"))]
use std::{
    fs,
    time::{Duration, SystemTime},
};

use eframe::egui;

use crate::document::DocumentState;

use super::palette::ThemePalette;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum BackstageSection {
    Open,
    Save,
    SaveAs,
}

impl BackstageSection {
    const ALL: [Self; 3] = [Self::Open, Self::Save, Self::SaveAs];

    const fn label(self) -> &'static str {
        match self {
            Self::Open => "Open",
            Self::Save => "Save",
            Self::SaveAs => "Save As",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum BackstageLocation {
    ThisPc,
    Browse,
}

impl BackstageLocation {
    const ALL: [Self; 2] = [Self::ThisPc, Self::Browse];

    const fn label(self) -> &'static str {
        match self {
            Self::ThisPc => "This PC",
            Self::Browse => "Browse",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(super) enum SaveFormat {
    Text,
    Markdown,
    Html,
    Odt,
    Pdf,
}

impl SaveFormat {
    const ALL: [Self; 5] = [Self::Text, Self::Markdown, Self::Html, Self::Odt, Self::Pdf];

    const fn label(self) -> &'static str {
        match self {
            Self::Text => "Plain Text (*.txt)",
            Self::Markdown => "Markdown (*.md)",
            Self::Html => "Web Page (*.html)",
            Self::Odt => "OpenDocument Text (*.odt)",
            Self::Pdf => "PDF (*.pdf)",
        }
    }

    pub(super) const fn extension(self) -> &'static str {
        match self {
            Self::Text => "txt",
            Self::Markdown => "md",
            Self::Html => "html",
            Self::Odt => "odt",
            Self::Pdf => "pdf",
        }
    }

    fn from_extension(extension: &str) -> Option<Self> {
        match extension
            .trim_start_matches('.')
            .to_ascii_lowercase()
            .as_str()
        {
            "txt" => Some(Self::Text),
            "md" | "markdown" => Some(Self::Markdown),
            "html" | "htm" => Some(Self::Html),
            "odt" => Some(Self::Odt),
            "pdf" => Some(Self::Pdf),
            _ => None,
        }
    }
}

pub(super) struct BackstageState {
    pub visible: bool,
    pub section: BackstageSection,
    pub location: BackstageLocation,
    pub file_name: String,
    pub format: SaveFormat,
    pub local_dir: Option<PathBuf>,
}

impl BackstageState {
    pub fn open_save_as(&mut self, document: &DocumentState, current_path: &Option<PathBuf>) {
        self.visible = true;
        self.section = BackstageSection::SaveAs;
        self.location = BackstageLocation::ThisPc;
        if let Some(path) = current_path {
            if let Some(parent) = path.parent() {
                self.local_dir = Some(parent.to_path_buf());
            }
            if let Some(name) = path.file_name().and_then(|name| name.to_str()) {
                self.file_name = name.to_owned();
            }
            if let Some(format) = path
                .extension()
                .and_then(|extension| extension.to_str())
                .and_then(SaveFormat::from_extension)
            {
                self.format = format;
            }
        } else {
            self.format = SaveFormat::Html;
            self.file_name = file_name_with_extension(&document.title, self.format.extension());
            if self.local_dir.is_none() {
                self.local_dir = std::env::current_dir().ok();
            }
        }
    }
}

impl Default for BackstageState {
    fn default() -> Self {
        Self {
            visible: false,
            section: BackstageSection::SaveAs,
            location: BackstageLocation::ThisPc,
            file_name: "document.html".to_owned(),
            format: SaveFormat::Html,
            local_dir: std::env::current_dir().ok(),
        }
    }
}

#[derive(Default)]
pub(super) struct BackstageOutput {
    pub close_requested: bool,
    pub open_requested: bool,
    pub save_requested: bool,
    pub save_as_requested: bool,
    pub recent_open_requested: Option<PathBuf>,
}

pub(super) fn paint_backstage(
    ui: &mut egui::Ui,
    state: &mut BackstageState,
    document: &DocumentState,
    current_path: &Option<PathBuf>,
    recent_files: &[PathBuf],
    palette: ThemePalette,
) -> BackstageOutput {
    let mut output = BackstageOutput::default();
    let height = ui.available_height();
    let width = ui.available_width();
    let nav_width = if width < 700.0 { 150.0 } else { 190.0 };
    let location_width = if width < 700.0 { 210.0 } else { 280.0 };
    let detail_width = (width - nav_width - location_width).max(360.0);

    let full_rect = ui.available_rect_before_wrap();
    let nav_rect = egui::Rect::from_min_size(full_rect.min, egui::vec2(nav_width, height));
    let locations_rect = egui::Rect::from_min_size(
        egui::pos2(nav_rect.right(), full_rect.top()),
        egui::vec2(location_width, height),
    );
    let details_rect = egui::Rect::from_min_size(
        egui::pos2(locations_rect.right(), full_rect.top()),
        egui::vec2(detail_width, height),
    );

    ui.painter()
        .rect_filled(full_rect, 0.0, backstage_surface(palette));

    ui.scope_builder(egui::UiBuilder::new().max_rect(nav_rect), |ui| {
        paint_backstage_nav(ui, state, &mut output, nav_width, height, palette);
    });
    ui.scope_builder(egui::UiBuilder::new().max_rect(locations_rect), |ui| {
        paint_backstage_locations(ui, state, &mut output, location_width, height, palette);
    });
    ui.scope_builder(egui::UiBuilder::new().max_rect(details_rect), |ui| {
        paint_backstage_details(
            ui,
            state,
            document,
            current_path,
            recent_files,
            &mut output,
            detail_width,
            height,
            palette,
        );
    });

    ui.advance_cursor_after_rect(full_rect);

    output
}

fn paint_backstage_nav(
    ui: &mut egui::Ui,
    state: &mut BackstageState,
    output: &mut BackstageOutput,
    width: f32,
    height: f32,
    palette: ThemePalette,
) {
    let nav_bg = backstage_nav_surface(palette);
    let nav_fg = palette.title_fg;
    let nav_muted = palette.title_muted;
    let active_bg = backstage_surface(palette);
    let active_fg = palette.tab_active_fg;
    let nav_style = BackstageNavStyle {
        selected_bg: active_bg,
        selected_fg: active_fg,
        normal_fg: nav_fg,
        palette,
    };

    egui::Frame::new()
        .fill(nav_bg)
        .inner_margin(egui::Margin::same(0))
        .show(ui, |ui| {
            ui.set_width(width);
            ui.set_min_height(height);
            ui.vertical(|ui| {
                if backstage_back_button(ui, width, nav_fg, palette).clicked() {
                    output.close_requested = true;
                }

                for section in BackstageSection::ALL {
                    let active = state.section == section;
                    let response = backstage_nav_row(ui, section.label(), width, active, nav_style);
                    if response.clicked() {
                        match section {
                            BackstageSection::Save => output.save_requested = true,
                            BackstageSection::Open => {
                                state.section = BackstageSection::Open;
                            }
                            BackstageSection::SaveAs => {
                                state.section = BackstageSection::SaveAs;
                            }
                        }
                    }
                }
                ui.add_space((ui.available_height() - 28.0).max(0.0));
                centered_nav_hint(ui, width, "Esc returns", nav_muted);
            });
        });
}

fn paint_backstage_locations(
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

fn paint_backstage_open_locations(
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

#[allow(clippy::too_many_arguments)]
fn paint_backstage_details(
    ui: &mut egui::Ui,
    state: &mut BackstageState,
    document: &DocumentState,
    current_path: &Option<PathBuf>,
    recent_files: &[PathBuf],
    output: &mut BackstageOutput,
    width: f32,
    height: f32,
    palette: ThemePalette,
) {
    egui::Frame::new()
        .fill(backstage_surface(palette))
        .inner_margin(egui::Margin::symmetric(26, 24))
        .show(ui, |ui| {
            ui.set_width(width);
            ui.set_min_height(height);
            egui::ScrollArea::vertical()
                .auto_shrink([false, false])
                .show(ui, |ui| {
                    ui.set_width((width - 52.0).max(300.0));

                    let breadcrumb = match state.location {
                        BackstageLocation::ThisPc | BackstageLocation::Browse => {
                            current_dir_label(&state.local_dir)
                        }
                    };
                    if state.section == BackstageSection::Open {
                        ui.label(
                            egui::RichText::new("Recent files")
                                .size(14.0)
                                .strong()
                                .color(palette.text_primary),
                        );
                        ui.add_space(18.0);
                        paint_recent_files(ui, recent_files, output, width - 52.0, palette);
                        return;
                    }

                    ui.label(
                        egui::RichText::new(breadcrumb)
                            .size(14.0)
                            .strong()
                            .color(palette.text_primary),
                    );
                    if let Some(path) = current_path {
                        ui.label(
                            egui::RichText::new(format!("Current file: {}", path.display()))
                                .size(11.0)
                                .color(palette.text_muted),
                        );
                    }
                    ui.add_space(18.0);

                    ui.horizontal(|ui| {
                        ui.label(
                            egui::RichText::new("File name:")
                                .size(13.0)
                                .color(palette.text_primary),
                        );
                        ui.add(
                            egui::TextEdit::singleline(&mut state.file_name)
                                .desired_width((width - 220.0).max(180.0)),
                        );
                    });
                    ui.add_space(8.0);
                    ui.horizontal(|ui| {
                        ui.label(
                            egui::RichText::new("Save as type:")
                                .size(13.0)
                                .color(palette.text_primary),
                        );
                        let previous_format = state.format;
                        egui::ComboBox::from_id_salt("backstage_file_type")
                            .selected_text(state.format.label())
                            .width((width - 230.0).max(180.0))
                            .show_ui(ui, |ui| {
                                for format in SaveFormat::ALL {
                                    ui.selectable_value(&mut state.format, format, format.label());
                                }
                            });
                        if state.format != previous_format {
                            state.file_name = file_name_with_extension(
                                &state.file_name,
                                state.format.extension(),
                            );
                        }
                    });
                    ui.add_space(12.0);
                    ui.horizontal(|ui| {
                        if ui.button("More options…").clicked() {
                            output.save_as_requested = true;
                        }
                        if ui
                            .add(
                                egui::Button::new(
                                    egui::RichText::new("Save")
                                        .size(14.0)
                                        .strong()
                                        .color(egui::Color32::WHITE),
                                )
                                .fill(egui::Color32::from_rgb(43, 87, 154))
                                .min_size(egui::vec2(96.0, 30.0)),
                            )
                            .clicked()
                        {
                            output.save_as_requested = true;
                        }
                    });
                    ui.add_space(18.0);
                    ui.separator();
                    ui.add_space(10.0);

                    paint_folder_contents(ui, state, document, width - 52.0, palette);
                });
        });
}

fn paint_folder_contents(
    ui: &mut egui::Ui,
    state: &mut BackstageState,
    document: &DocumentState,
    width: f32,
    palette: ThemePalette,
) {
    folder_header(ui, width, palette);

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
                    if folder_row(ui, "..", "Parent folder", true, width, palette).clicked() {
                        state.local_dir = Some(parent.to_path_buf());
                    }
                }

                let mut entries = folder_entries(&dir);
                if entries.is_empty() {
                    ui.label(
                        egui::RichText::new("This folder is empty.")
                            .size(12.0)
                            .color(palette.text_muted),
                    );
                }
                entries.truncate(80);
                for entry in entries {
                    if folder_row(
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

fn paint_recent_files(
    ui: &mut egui::Ui,
    recent_files: &[PathBuf],
    output: &mut BackstageOutput,
    width: f32,
    palette: ThemePalette,
) {
    folder_header(ui, width, palette);
    if recent_files.is_empty() {
        ui.add_space(12.0);
        ui.label(
            egui::RichText::new("No recent files yet")
                .size(12.0)
                .color(palette.text_muted),
        );
        return;
    }

    let width = width.min(ui.available_width()).max(360.0);
    egui::ScrollArea::vertical()
        .auto_shrink([false, false])
        .max_height((ui.available_height() - 8.0).max(120.0))
        .show(ui, |ui| {
            for path in recent_files {
                let name = path
                    .file_name()
                    .and_then(|name| name.to_str())
                    .unwrap_or("document");
                let detail = path.parent().map_or_else(
                    || path.display().to_string(),
                    |parent| parent.display().to_string(),
                );
                if recent_file_row(ui, name, &detail, width, palette).clicked() {
                    output.recent_open_requested = Some(path.clone());
                }
            }
        });
}

#[derive(Clone, Copy)]
struct BackstageNavStyle {
    selected_bg: egui::Color32,
    selected_fg: egui::Color32,
    normal_fg: egui::Color32,
    palette: ThemePalette,
}

fn backstage_nav_row(
    ui: &mut egui::Ui,
    label: &str,
    width: f32,
    selected: bool,
    style: BackstageNavStyle,
) -> egui::Response {
    let (rect, response) = ui.allocate_exact_size(egui::vec2(width, 36.0), egui::Sense::click());
    let fill = if selected {
        style.selected_bg
    } else if response.hovered() {
        style.palette.accent.gamma_multiply(0.18)
    } else {
        egui::Color32::TRANSPARENT
    };
    ui.painter().rect_filled(rect, 0.0, fill);
    ui.painter().text(
        rect.left_center() + egui::vec2(24.0, 0.0),
        egui::Align2::LEFT_CENTER,
        label,
        egui::FontId::proportional(15.0),
        if selected {
            style.selected_fg
        } else {
            style.normal_fg
        },
    );
    response
}

fn backstage_back_button(
    ui: &mut egui::Ui,
    width: f32,
    color: egui::Color32,
    palette: ThemePalette,
) -> egui::Response {
    let (rect, response) = ui.allocate_exact_size(egui::vec2(width, 36.0), egui::Sense::click());
    let fill = if response.hovered() {
        palette.accent.gamma_multiply(0.18)
    } else {
        egui::Color32::TRANSPARENT
    };
    ui.painter().rect_filled(rect, 0.0, fill);

    let center = rect.left_center() + egui::vec2(24.0, 0.0);
    let stroke = egui::Stroke::new(1.8, color);
    ui.painter().line_segment(
        [
            center + egui::vec2(6.0, -7.0),
            center + egui::vec2(-3.0, 0.0),
        ],
        stroke,
    );
    ui.painter().line_segment(
        [
            center + egui::vec2(-3.0, 0.0),
            center + egui::vec2(6.0, 7.0),
        ],
        stroke,
    );
    ui.painter().line_segment(
        [
            center + egui::vec2(-2.0, 0.0),
            center + egui::vec2(15.0, 0.0),
        ],
        stroke,
    );

    response.on_hover_text("Back to document")
}

fn centered_nav_hint(ui: &mut egui::Ui, width: f32, text: &str, color: egui::Color32) {
    let (rect, _) = ui.allocate_exact_size(egui::vec2(width, 20.0), egui::Sense::hover());
    ui.painter().text(
        rect.center(),
        egui::Align2::CENTER_CENTER,
        text,
        egui::FontId::proportional(11.0),
        color,
    );
}

fn backstage_two_line_row(
    ui: &mut egui::Ui,
    title: &str,
    subtitle: &str,
    selected: bool,
    enabled: bool,
    palette: ThemePalette,
) -> egui::Response {
    let width = ui.available_width().max(180.0);
    let sense = if enabled {
        egui::Sense::click()
    } else {
        egui::Sense::hover()
    };
    let (rect, response) = ui.allocate_exact_size(egui::vec2(width, 54.0), sense);
    let fill = if selected {
        palette.accent.gamma_multiply(0.18)
    } else if enabled && response.hovered() {
        palette.accent.gamma_multiply(0.08)
    } else {
        egui::Color32::TRANSPARENT
    };
    ui.painter().rect_filled(rect, 0.0, fill);
    if selected {
        ui.painter().rect_stroke(
            rect,
            0.0,
            egui::Stroke::new(1.0, palette.accent),
            egui::StrokeKind::Inside,
        );
    }
    let title_color = if enabled {
        palette.text_primary
    } else {
        palette.text_muted
    };
    ui.painter().text(
        rect.left_top() + egui::vec2(12.0, 9.0),
        egui::Align2::LEFT_TOP,
        title,
        egui::FontId::proportional(14.0),
        title_color,
    );
    ui.painter().text(
        rect.left_top() + egui::vec2(12.0, 30.0),
        egui::Align2::LEFT_TOP,
        subtitle,
        egui::FontId::proportional(11.0),
        palette.text_muted,
    );
    response
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

fn recent_file_row(
    ui: &mut egui::Ui,
    name: &str,
    detail: &str,
    width: f32,
    palette: ThemePalette,
) -> egui::Response {
    let width = width.min(ui.available_width()).max(360.0);
    let (rect, response) = ui.allocate_exact_size(egui::vec2(width, 42.0), egui::Sense::click());
    let fill = if response.hovered() {
        palette.accent.gamma_multiply(0.08)
    } else {
        egui::Color32::TRANSPARENT
    };
    ui.painter().rect_filled(rect, 0.0, fill);

    let icon_rect = egui::Rect::from_min_size(
        rect.left_center() + egui::vec2(10.0, -7.0),
        egui::vec2(16.0, 14.0),
    );
    paint_file_icon(ui.painter(), icon_rect, palette);

    ui.painter().text(
        rect.left_top() + egui::vec2(34.0, 7.0),
        egui::Align2::LEFT_TOP,
        name,
        egui::FontId::proportional(12.5),
        palette.text_primary,
    );
    ui.painter().text(
        rect.left_top() + egui::vec2(34.0, 24.0),
        egui::Align2::LEFT_TOP,
        detail,
        egui::FontId::proportional(10.5),
        palette.text_muted,
    );
    response
}

#[cfg(not(target_arch = "wasm32"))]
fn folder_row(
    ui: &mut egui::Ui,
    name: &str,
    detail: &str,
    is_dir: bool,
    width: f32,
    palette: ThemePalette,
) -> egui::Response {
    let width = width.min(ui.available_width()).max(360.0);
    let (rect, response) = ui.allocate_exact_size(egui::vec2(width, 32.0), egui::Sense::click());
    let fill = if response.hovered() {
        palette.accent.gamma_multiply(0.08)
    } else {
        egui::Color32::TRANSPARENT
    };
    ui.painter().rect_filled(rect, 0.0, fill);

    let icon_rect = egui::Rect::from_min_size(
        rect.left_center() + egui::vec2(10.0, -7.0),
        egui::vec2(16.0, 14.0),
    );
    if is_dir {
        paint_folder_icon(ui.painter(), icon_rect, palette);
    } else {
        paint_file_icon(ui.painter(), icon_rect, palette);
    }

    let date_x = (rect.right() - 220.0).max(rect.left() + 280.0);
    ui.painter().text(
        rect.left_center() + egui::vec2(34.0, 0.0),
        egui::Align2::LEFT_CENTER,
        name,
        egui::FontId::proportional(12.5),
        palette.text_primary,
    );
    ui.painter().text(
        egui::pos2(date_x, rect.center().y),
        egui::Align2::LEFT_CENTER,
        detail,
        egui::FontId::proportional(12.0),
        palette.text_muted,
    );
    response
}

fn folder_header(ui: &mut egui::Ui, width: f32, palette: ThemePalette) {
    let width = width.min(ui.available_width()).max(360.0);
    let (rect, _) = ui.allocate_exact_size(egui::vec2(width, 32.0), egui::Sense::hover());
    let date_x = (rect.right() - 220.0).max(rect.left() + 280.0);
    ui.painter().text(
        rect.left_center() + egui::vec2(34.0, 0.0),
        egui::Align2::LEFT_CENTER,
        "Name",
        egui::FontId::proportional(12.0),
        palette.text_muted,
    );
    ui.painter().text(
        egui::pos2(date_x, rect.center().y),
        egui::Align2::LEFT_CENTER,
        "Date Modified",
        egui::FontId::proportional(12.0),
        palette.text_muted,
    );
    ui.painter().line_segment(
        [rect.left_bottom(), rect.right_bottom()],
        egui::Stroke::new(1.0, palette.border),
    );
}

#[cfg(not(target_arch = "wasm32"))]
fn paint_folder_icon(painter: &egui::Painter, rect: egui::Rect, palette: ThemePalette) {
    let stroke = egui::Stroke::new(1.2, palette.text_muted);
    let tab = egui::Rect::from_min_size(rect.min + egui::vec2(1.0, 0.0), egui::vec2(7.0, 4.0));
    let body = egui::Rect::from_min_max(
        rect.min + egui::vec2(1.0, 3.0),
        rect.max - egui::vec2(1.0, 1.0),
    );
    painter.rect_stroke(tab, 1.0, stroke, egui::StrokeKind::Inside);
    painter.rect_stroke(body, 1.0, stroke, egui::StrokeKind::Inside);
}

fn paint_file_icon(painter: &egui::Painter, rect: egui::Rect, palette: ThemePalette) {
    let stroke = egui::Stroke::new(1.2, palette.text_muted);
    let page = egui::Rect::from_min_max(
        rect.min + egui::vec2(3.0, 1.0),
        rect.max - egui::vec2(3.0, 1.0),
    );
    painter.rect_stroke(page, 1.0, stroke, egui::StrokeKind::Inside);
    painter.line_segment(
        [
            page.left_top() + egui::vec2(3.0, 5.0),
            page.right_top() + egui::vec2(-3.0, 5.0),
        ],
        stroke,
    );
}

fn current_dir_label(local_dir: &Option<PathBuf>) -> String {
    local_dir
        .as_ref()
        .map(|path| path.display().to_string())
        .unwrap_or_else(|| "This PC".to_owned())
}

fn file_name_with_extension(file_name: &str, extension: &str) -> String {
    let extension = extension.trim_start_matches('.');
    let path = Path::new(file_name.trim());
    let stem = path
        .file_stem()
        .and_then(|stem| stem.to_str())
        .filter(|stem| !stem.trim().is_empty())
        .unwrap_or_else(|| {
            path.file_name()
                .and_then(|name| name.to_str())
                .filter(|name| !name.trim().is_empty())
                .unwrap_or("document")
        })
        .trim();
    format!("{stem}.{extension}")
}

fn backstage_surface(palette: ThemePalette) -> egui::Color32 {
    if palette.workspace_bg.r() < 100 {
        egui::Color32::from_rgb(43, 49, 59)
    } else {
        egui::Color32::from_rgb(255, 255, 255)
    }
}

fn backstage_mid_surface(palette: ThemePalette) -> egui::Color32 {
    if palette.workspace_bg.r() < 100 {
        egui::Color32::from_rgb(49, 55, 66)
    } else {
        egui::Color32::from_rgb(246, 248, 252)
    }
}

fn backstage_nav_surface(palette: ThemePalette) -> egui::Color32 {
    if palette.workspace_bg.r() < 100 {
        palette.title_bg
    } else {
        palette.tab_bg
    }
}

#[cfg(not(target_arch = "wasm32"))]
struct FolderEntry {
    name: String,
    path: PathBuf,
    is_dir: bool,
    modified: String,
}

#[cfg(not(target_arch = "wasm32"))]
fn folder_entries(dir: &Path) -> Vec<FolderEntry> {
    let Ok(read_dir) = fs::read_dir(dir) else {
        return Vec::new();
    };
    let mut entries: Vec<_> = read_dir
        .filter_map(Result::ok)
        .filter_map(|entry| {
            let metadata = entry.metadata().ok()?;
            let name = entry.file_name().to_string_lossy().into_owned();
            if name.starts_with('.') {
                return None;
            }
            let modified = metadata
                .modified()
                .ok()
                .map(modified_label)
                .unwrap_or_else(|| "Unknown".to_owned());
            Some(FolderEntry {
                name,
                path: entry.path(),
                is_dir: metadata.is_dir(),
                modified,
            })
        })
        .collect();
    entries.sort_by(|left, right| {
        right
            .is_dir
            .cmp(&left.is_dir)
            .then_with(|| left.name.to_lowercase().cmp(&right.name.to_lowercase()))
    });
    entries
}

#[cfg(not(target_arch = "wasm32"))]
fn modified_label(modified: SystemTime) -> String {
    match SystemTime::now().duration_since(modified) {
        Ok(elapsed) if elapsed < Duration::from_secs(60) => "Just now".to_owned(),
        Ok(elapsed) if elapsed < Duration::from_secs(60 * 60) => {
            format!("{} min ago", elapsed.as_secs() / 60)
        }
        Ok(elapsed) if elapsed < Duration::from_secs(60 * 60 * 24) => {
            format!("{} hours ago", elapsed.as_secs() / 3600)
        }
        Ok(elapsed) => format!("{} days ago", elapsed.as_secs() / 86_400),
        Err(_) => "In the future".to_owned(),
    }
}
