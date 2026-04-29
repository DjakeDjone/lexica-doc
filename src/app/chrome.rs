use std::path::{Path, PathBuf};
#[cfg(not(target_arch = "wasm32"))]
use std::{
    fs,
    time::{Duration, SystemTime},
};

use eframe::egui;

use crate::document::{
    DocumentState, FontChoice, ImageLayoutMode, ImageRendering, ListKind, ParagraphAlignment,
    WrapMode, OBJECT_REPLACEMENT_CHAR,
};
use crate::grammar::{GrammarConfig, GrammarStatus, Language};

use super::{
    actions::{
        delete_table_column, delete_table_row, insert_image, insert_page_break, insert_table,
        insert_table_column, insert_table_row, open_document, reset_image_size, save_document,
        save_document_as, set_font_choice, set_font_size, set_highlight_color, set_image_opacity,
        set_image_rendering, set_image_wrap_mode, set_paragraph_alignment, set_text_color,
        sync_active_style, toggle_bold, toggle_bullet_list, toggle_italic, toggle_ordered_list,
        toggle_strikethrough, toggle_underline,
    },
    palette::{theme_switch, ThemeMode, ThemePalette},
    CanvasState, ChangeHistory, ZoomMode,
};

#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum RibbonTab {
    Home,
    Insert,
    Design,
    Layout,
    View,
    Grammer,
    Picture,
    Table,
}

impl RibbonTab {
    const ALL: [Self; 6] = [
        Self::Home,
        Self::Insert,
        Self::Design,
        Self::Layout,
        Self::View,
        Self::Grammer,
    ];

    const fn label(self) -> &'static str {
        match self {
            Self::Home => "Home",
            Self::Insert => "Insert",
            Self::Design => "Design",
            Self::Layout => "Layout",
            Self::View => "View",
            Self::Grammer => "Grammer",
            Self::Picture => "Picture Format",
            Self::Table => "Table Format",
        }
    }
}

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
    Pdf,
}

impl SaveFormat {
    const ALL: [Self; 4] = [Self::Text, Self::Markdown, Self::Html, Self::Pdf];

    const fn label(self) -> &'static str {
        match self {
            Self::Text => "Plain Text (*.txt)",
            Self::Markdown => "Markdown (*.md)",
            Self::Html => "Web Page (*.html)",
            Self::Pdf => "PDF (*.pdf)",
        }
    }

    pub(super) const fn extension(self) -> &'static str {
        match self {
            Self::Text => "txt",
            Self::Markdown => "md",
            Self::Html => "html",
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
}

#[derive(Default)]
pub(super) struct GrammarRibbonOutput {
    pub manual_check_requested: bool,
    pub restart_requested: bool,
    pub download_requested: bool,
    pub settings_changed: bool,
}

#[allow(clippy::too_many_arguments)]
pub(super) fn paint_title_bar(
    ui: &mut egui::Ui,
    document: &mut crate::document::DocumentState,
    canvas: &mut CanvasState,
    current_path: &Option<PathBuf>,
    status_message: &str,
    theme_mode: &mut ThemeMode,
    status_target: &mut String,
    history: &mut ChangeHistory,
    palette: ThemePalette,
    logo: &egui::TextureHandle,
) {
    let path_label = current_path
        .as_ref()
        .map(|path| path.to_string_lossy().into_owned())
        .unwrap_or_else(|| "Unsaved document".to_owned());

    // Render the title bar content first so buttons register their interactions
    // before the drag overlay.
    let _frame_response = egui::Frame::new()
        .inner_margin(egui::Margin::symmetric(12, 8))
        .show(ui, |ui| {
            ui.horizontal(|ui| {
                ui.add(
                    egui::Image::new(egui::load::SizedTexture::new(
                        logo.id(),
                        egui::vec2(24.0, 24.0),
                    ))
                    .sense(egui::Sense::hover()),
                );

                ui.label(
                    egui::RichText::new(format!("{} - Word", document.title))
                        .size(14.0)
                        .color(palette.title_fg),
                );
                ui.label(
                    egui::RichText::new(path_label)
                        .size(11.0)
                        .color(palette.title_muted),
                );

                // Undo / Redo buttons moved after filename/path (still left-aligned)
                let can_undo = history.can_undo();
                let can_redo = history.can_redo();
                let undo_btn =
                    egui::Button::new(egui::RichText::new("↩").size(14.0).color(if can_undo {
                        palette.title_fg
                    } else {
                        palette.title_muted
                    }))
                    .min_size(egui::vec2(24.0, 24.0))
                    .fill(egui::Color32::TRANSPARENT)
                    .stroke(egui::Stroke::NONE);
                if ui
                    .add_enabled(can_undo, undo_btn)
                    .on_hover_text("Undo (Ctrl+Z)")
                    .clicked()
                    && history.undo(document)
                {
                    canvas.image_textures.clear();
                    *status_target = "Undo".to_owned();
                }
                let redo_btn =
                    egui::Button::new(egui::RichText::new("↪").size(14.0).color(if can_redo {
                        palette.title_fg
                    } else {
                        palette.title_muted
                    }))
                    .min_size(egui::vec2(24.0, 24.0))
                    .fill(egui::Color32::TRANSPARENT)
                    .stroke(egui::Stroke::NONE);
                if ui
                    .add_enabled(can_redo, redo_btn)
                    .on_hover_text("Redo (Ctrl+Shift+Z / Ctrl+Y)")
                    .clicked()
                    && history.redo(document)
                {
                    canvas.image_textures.clear();
                    *status_target = "Redo".to_owned();
                }

                ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
                    #[cfg(not(target_arch = "wasm32"))]
                    {
                        let close_btn = egui::Button::new(
                            egui::RichText::new("🗙").size(14.0).color(palette.title_fg),
                        )
                        .min_size(egui::vec2(24.0, 24.0))
                        .fill(egui::Color32::TRANSPARENT)
                        .stroke(egui::Stroke::NONE);
                        if ui.add(close_btn).on_hover_text("Close").clicked() {
                            ui.ctx().send_viewport_cmd(egui::ViewportCommand::Close);
                        }

                        let maximized = ui.input(|i| i.viewport().maximized.unwrap_or(false));
                        let max_icon = if maximized { "🗗" } else { "🗖" };
                        let max_btn = egui::Button::new(
                            egui::RichText::new(max_icon)
                                .size(14.0)
                                .color(palette.title_fg),
                        )
                        .min_size(egui::vec2(24.0, 24.0))
                        .fill(egui::Color32::TRANSPARENT)
                        .stroke(egui::Stroke::NONE);
                        if ui
                            .add(max_btn)
                            .on_hover_text(if maximized { "Restore" } else { "Maximize" })
                            .clicked()
                        {
                            ui.ctx()
                                .send_viewport_cmd(egui::ViewportCommand::Maximized(!maximized));
                        }

                        let min_btn = egui::Button::new(
                            egui::RichText::new("🗕").size(14.0).color(palette.title_fg),
                        )
                        .min_size(egui::vec2(24.0, 24.0))
                        .fill(egui::Color32::TRANSPARENT)
                        .stroke(egui::Stroke::NONE);
                        if ui.add(min_btn).on_hover_text("Minimize").clicked() {
                            ui.ctx()
                                .send_viewport_cmd(egui::ViewportCommand::Minimized(true));
                        }

                        ui.separator();
                    }

                    if theme_switch(ui, theme_mode, palette, true) {
                        *status_target = format!("Theme switched to {}", theme_mode.label());
                    }
                    ui.separator();
                    ui.label(
                        egui::RichText::new(status_message)
                            .size(11.0)
                            .color(palette.title_muted),
                    );
                });
            });
        });

    // Window drag and double-click: handled entirely via raw pointer input.
    // We deliberately avoid ui.interact() here because ANY interaction overlay
    // on the title bar rect steals events from the buttons inside it.
    #[cfg(not(target_arch = "wasm32"))]
    let title_rect = _frame_response.response.rect;

    // Drag to move window — only when pointer is decisively dragging (past
    // threshold), the press originated inside the title bar, and no egui
    // widget has already claimed the drag (e.g. a DragValue in the ribbon).
    #[cfg(not(target_arch = "wasm32"))]
    let is_dragging = ui.input(|i| i.pointer.is_decidedly_dragging());
    #[cfg(not(target_arch = "wasm32"))]
    let press_origin = ui.input(|i| i.pointer.press_origin());
    #[cfg(not(target_arch = "wasm32"))]
    let anything_dragged = ui.ctx().dragged_id().is_some();

    #[cfg(not(target_arch = "wasm32"))]
    if is_dragging {
        if let Some(origin) = press_origin {
            if title_rect.contains(origin) && !anything_dragged {
                ui.ctx().send_viewport_cmd(egui::ViewportCommand::StartDrag);
            }
        }
    }

    // Double-click to maximize/restore.
    #[cfg(not(target_arch = "wasm32"))]
    if let Some(pos) = ui.input(|i| i.pointer.interact_pos()) {
        if title_rect.contains(pos)
            && ui.input(|i| {
                i.pointer
                    .button_double_clicked(egui::PointerButton::Primary)
            })
        {
            let maximized = ui.input(|i| i.viewport().maximized.unwrap_or(false));
            ui.ctx()
                .send_viewport_cmd(egui::ViewportCommand::Maximized(!maximized));
        }
    }
}

pub(super) fn paint_tab_row(
    ui: &mut egui::Ui,
    active_tab: &mut RibbonTab,
    selected_image_id: Option<usize>,
    active_table_cell: Option<(usize, usize, usize)>,
    palette: ThemePalette,
) -> bool {
    let mut file_requested = false;
    egui::Frame::new()
        .inner_margin(egui::Margin::symmetric(8, 0))
        .show(ui, |ui| {
            ui.horizontal(|ui| {
                let file_button = egui::Button::new(
                    egui::RichText::new("File")
                        .size(13.0)
                        .color(palette.tab_fg)
                        .strong(),
                )
                .min_size(egui::vec2(54.0, 28.0))
                .fill(egui::Color32::TRANSPARENT)
                .stroke(egui::Stroke::NONE)
                .corner_radius(0.0);
                if ui.add(file_button).clicked() {
                    file_requested = true;
                }

                for tab in RibbonTab::ALL {
                    let selected = *active_tab == tab;
                    let button =
                        egui::Button::new(egui::RichText::new(tab.label()).size(13.0).color(
                            if selected {
                                palette.tab_active_fg
                            } else {
                                palette.tab_fg
                            },
                        ))
                        .min_size(egui::vec2(64.0, 28.0))
                        .fill(if selected {
                            palette.tab_active_bg
                        } else {
                            egui::Color32::TRANSPARENT
                        })
                        .stroke(if selected {
                            egui::Stroke::new(1.0, palette.border)
                        } else {
                            egui::Stroke::NONE
                        })
                        .corner_radius(0.0);
                    if ui.add(button).clicked() {
                        *active_tab = tab;
                    }
                }

                // Contextual "Picture Format" tab — shown only when an image is selected
                if selected_image_id.is_some() {
                    ui.separator();
                    let selected = *active_tab == RibbonTab::Picture;
                    // Gold accent colours matching Word's contextual picture tab
                    let picture_accent = egui::Color32::from_rgb(176, 118, 0);
                    let fg = if selected {
                        egui::Color32::from_rgb(130, 80, 0)
                    } else {
                        egui::Color32::from_rgb(255, 238, 190)
                    };
                    let bg = if selected {
                        egui::Color32::from_rgb(255, 242, 204)
                    } else {
                        egui::Color32::TRANSPARENT
                    };
                    let button = egui::Button::new(
                        egui::RichText::new("Picture Format")
                            .size(13.0)
                            .color(fg)
                            .strong(),
                    )
                    .min_size(egui::vec2(108.0, 28.0))
                    .fill(bg)
                    .stroke(if selected {
                        egui::Stroke::new(1.0, picture_accent)
                    } else {
                        egui::Stroke::NONE
                    })
                    .corner_radius(0.0);
                    if ui.add(button).clicked() {
                        *active_tab = RibbonTab::Picture;
                    }
                }

                if active_table_cell.is_some() {
                    ui.separator();
                    let selected = *active_tab == RibbonTab::Table;
                    let table_accent = egui::Color32::from_rgb(38, 120, 96);
                    let fg = if selected {
                        egui::Color32::from_rgb(20, 88, 68)
                    } else {
                        egui::Color32::from_rgb(210, 244, 234)
                    };
                    let bg = if selected {
                        egui::Color32::from_rgb(219, 247, 239)
                    } else {
                        egui::Color32::TRANSPARENT
                    };
                    let button = egui::Button::new(
                        egui::RichText::new("Table Format")
                            .size(13.0)
                            .color(fg)
                            .strong(),
                    )
                    .min_size(egui::vec2(104.0, 28.0))
                    .fill(bg)
                    .stroke(if selected {
                        egui::Stroke::new(1.0, table_accent)
                    } else {
                        egui::Stroke::NONE
                    })
                    .corner_radius(0.0);
                    if ui.add(button).clicked() {
                        *active_tab = RibbonTab::Table;
                    }
                }
            });
        });
    file_requested
}

#[allow(clippy::too_many_arguments)]
pub(super) fn paint_ribbon(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    active_tab: &mut RibbonTab,
    status_message: &mut String,
    current_path: &mut Option<PathBuf>,
    theme_mode: &mut ThemeMode,
    history: &mut ChangeHistory,
    grammar_config: &mut GrammarConfig,
    grammar_status: &GrammarStatus,
    grammar_auto_check: &mut bool,
    can_download_grammar: bool,
    palette: ThemePalette,
) -> GrammarRibbonOutput {
    sync_active_style(document, canvas);
    let mut output = GrammarRibbonOutput::default();

    egui::Frame::new()
        .inner_margin(egui::Margin::symmetric(8, 8))
        .show(ui, |ui| {
            ui.horizontal_wrapped(|ui| match active_tab {
                RibbonTab::Home => {
                    ribbon_file_group(ui, document, canvas, status_message, current_path, history, palette);
                    ribbon_font_group(ui, document, canvas, history, palette);
                    ribbon_paragraph_group(ui, document, canvas, history, palette);
                    ribbon_color_group(ui, document, canvas, history, palette);
                    ribbon_view_group(ui, canvas, status_message, theme_mode, palette);
                }
                RibbonTab::Insert => {
                    ribbon_file_group(ui, document, canvas, status_message, current_path, history, palette);
                    ribbon_insert_group(ui, document, canvas, status_message, history, palette);
                    ribbon_info_group(
                        ui,
                        "Insert",
                        "Import supports .txt, .md, .markdown, and .docx with images.",
                        palette,
                    );
                }
                RibbonTab::Design => {
                    ribbon_font_group(ui, document, canvas, history, palette);
                    ribbon_paragraph_group(ui, document, canvas, history, palette);
                    ribbon_color_group(ui, document, canvas, history, palette);
                }
                RibbonTab::Layout => {
                    ribbon_view_group(ui, canvas, status_message, theme_mode, palette);
                    ribbon_info_group(
                        ui,
                        "Page",
                        &format!(
                            "A4 {} x {} pt, margins {} pt",
                            document.page_size.width_points as i32,
                            document.page_size.height_points as i32,
                            document.margins.top_points as i32
                        ),
                        palette,
                    );
                }
                RibbonTab::View => {
                    ribbon_view_group(ui, canvas, status_message, theme_mode, palette);
                    ribbon_info_group(
                        ui,
                        "Shortcuts",
                        "Command+S Save, Command+Shift+S Save As, Ctrl+Z Undo, Ctrl+Shift+Z / Ctrl+Y Redo, Command+B Bold, Command+I Italic, Command+U Underline",
                        palette,
                    );
                }
                RibbonTab::Grammer => {
                    ribbon_grammer_actions_group(
                        ui,
                        grammar_status,
                        can_download_grammar,
                        &mut output,
                        palette,
                    );
                    ribbon_grammer_settings_group(
                        ui,
                        grammar_config,
                        grammar_auto_check,
                        &mut output,
                        palette,
                    );
                    ribbon_info_group(
                        ui,
                        "Server",
                        &format!(
                            "JAR: {} | Port: {}",
                            grammar_config.lt_jar_path.display(),
                            grammar_config.port
                        ),
                        palette,
                    );
                }
                RibbonTab::Picture => {
                    ribbon_picture_group(ui, document, canvas, status_message, history, palette);
                }
                RibbonTab::Table => {
                    ribbon_font_group(ui, document, canvas, history, palette);
                    ribbon_color_group(ui, document, canvas, history, palette);
                    ribbon_insert_group(ui, document, canvas, status_message, history, palette);
                    table_format_group(ui, document, canvas, status_message, history, palette);
                }
            });
        });
    output
}

pub(super) fn paint_backstage(
    ui: &mut egui::Ui,
    state: &mut BackstageState,
    document: &DocumentState,
    current_path: &Option<PathBuf>,
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
                    let response = backstage_nav_row(
                        ui,
                        section.label(),
                        width,
                        active,
                        active_bg,
                        active_fg,
                        nav_fg,
                        palette,
                    );
                    if response.clicked() {
                        match section {
                            BackstageSection::Save => output.save_requested = true,
                            BackstageSection::Open => output.open_requested = true,
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

#[allow(clippy::too_many_arguments)]
fn paint_backstage_details(
    ui: &mut egui::Ui,
    state: &mut BackstageState,
    document: &DocumentState,
    current_path: &Option<PathBuf>,
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

fn backstage_nav_row(
    ui: &mut egui::Ui,
    label: &str,
    width: f32,
    selected: bool,
    selected_bg: egui::Color32,
    selected_fg: egui::Color32,
    normal_fg: egui::Color32,
    palette: ThemePalette,
) -> egui::Response {
    let (rect, response) = ui.allocate_exact_size(egui::vec2(width, 36.0), egui::Sense::click());
    let fill = if selected {
        selected_bg
    } else if response.hovered() {
        palette.accent.gamma_multiply(0.18)
    } else {
        egui::Color32::TRANSPARENT
    };
    ui.painter().rect_filled(rect, 0.0, fill);
    ui.painter().text(
        rect.left_center() + egui::vec2(24.0, 0.0),
        egui::Align2::LEFT_CENTER,
        label,
        egui::FontId::proportional(15.0),
        if selected { selected_fg } else { normal_fg },
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

#[cfg(not(target_arch = "wasm32"))]
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

pub(super) fn paint_status_bar(
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
        });
    });
}

fn ribbon_file_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    current_path: &mut Option<PathBuf>,
    history: &mut ChangeHistory,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Clipboard", palette, |ui| {
        if ui.button("📂 Open").clicked() {
            open_document(document, canvas, status_message, current_path, history);
        }
        if ui.button("💾 Save").clicked() {
            save_document(document, status_message, current_path);
        }
        if ui.button("Save As").clicked() {
            save_document_as(document, status_message, current_path);
        }
    });
}

fn ribbon_font_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Font", palette, |ui| {
        let active_font = FontChoice::from_style(canvas.active_style);
        egui::ComboBox::from_id_salt("font_choice")
            .selected_text(active_font.label())
            .width(160.0)
            .show_ui(ui, |ui| {
                for font in FontChoice::ALL {
                    if ui
                        .selectable_label(active_font == font, font.label())
                        .clicked()
                    {
                        set_font_choice(document, canvas, font, history);
                    }
                }
            });

        let mut font_size = canvas.active_style.font_size_points;
        let resp = ui.add(
            egui::DragValue::new(&mut font_size)
                .range(8.0..=72.0)
                .speed(0.25)
                .fixed_decimals(1),
        );
        if resp.changed() {
            let now = ui.input(|i| i.time);
            set_font_size(document, canvas, font_size.clamp(8.0, 72.0), history, now);
        }

        ui.separator();

        if format_button(ui, canvas.active_style.bold, "B", palette).clicked() {
            toggle_bold(document, canvas, history);
        }
        if format_button(ui, canvas.active_style.italic, "I", palette).clicked() {
            toggle_italic(document, canvas, history);
        }
        if format_button(ui, canvas.active_style.underline, "U", palette).clicked() {
            toggle_underline(document, canvas, history);
        }
        if format_button(ui, canvas.active_style.strikethrough, "S", palette).clicked() {
            toggle_strikethrough(document, canvas, history);
        }
    });
}

fn ribbon_insert_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    history: &mut ChangeHistory,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Insert", palette, |ui| {
        if ui.button("Image").clicked() {
            insert_image(document, canvas, status_message, history);
        }
        if ui.button("Page Break").clicked() {
            insert_page_break(document, canvas, status_message, history);
        }
        ui.separator();
        ui.menu_button("Table", |ui| {
            ui.label(egui::RichText::new("Insert Table").size(12.0).strong());
            ui.add_space(4.0);
            let grid_size = 8;
            let cell_size = 18.0;
            let mut hovered_rows = 0usize;
            let mut hovered_cols = 0usize;
            for row in 0..grid_size {
                ui.horizontal(|ui| {
                    ui.spacing_mut().item_spacing = egui::vec2(2.0, 2.0);
                    for col in 0..grid_size {
                        let is_selected = row < hovered_rows && col < hovered_cols;
                        let fill = if is_selected {
                            palette.accent.gamma_multiply(0.35)
                        } else {
                            palette.ribbon_group_bg
                        };
                        let stroke = egui::Stroke::new(
                            1.0,
                            if is_selected {
                                palette.accent
                            } else {
                                palette.border
                            },
                        );
                        let btn = egui::Button::new("")
                            .min_size(egui::vec2(cell_size, cell_size))
                            .fill(fill)
                            .stroke(stroke)
                            .corner_radius(2.0);
                        let resp = ui.add(btn);
                        if resp.hovered() {
                            hovered_rows = row + 1;
                            hovered_cols = col + 1;
                        }
                        if resp.clicked() {
                            insert_table(
                                document,
                                canvas,
                                row + 1,
                                col + 1,
                                status_message,
                                history,
                            );
                            ui.close();
                        }
                    }
                });
            }
            if hovered_rows > 0 && hovered_cols > 0 {
                ui.label(
                    egui::RichText::new(format!("{}×{}", hovered_rows, hovered_cols))
                        .size(11.0)
                        .color(palette.text_muted),
                );
            }
        });
    });
}

fn ribbon_paragraph_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Paragraph", palette, |ui| {
        for alignment in ParagraphAlignment::ALL {
            if alignment_button(
                ui,
                canvas.active_paragraph_style.alignment == alignment,
                alignment,
                palette,
            )
            .on_hover_text(alignment.label())
            .clicked()
            {
                set_paragraph_alignment(document, canvas, alignment, history);
            }
        }

        ui.separator();

        if format_button(
            ui,
            canvas.active_paragraph_style.list_kind == ListKind::Bullet,
            "•",
            palette,
        )
        .on_hover_text(ListKind::Bullet.label())
        .clicked()
        {
            toggle_bullet_list(document, canvas, history);
        }
        if format_button(
            ui,
            canvas.active_paragraph_style.list_kind == ListKind::Ordered,
            "1.",
            palette,
        )
        .on_hover_text(ListKind::Ordered.label())
        .clicked()
        {
            toggle_ordered_list(document, canvas, history);
        }
    });
}

fn ribbon_color_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Colors", palette, |ui| {
        let mut text_color = canvas.active_style.text_color;
        let resp = ui.color_edit_button_srgba(&mut text_color);
        if resp.changed() {
            let now = ui.input(|i| i.time);
            set_text_color(document, canvas, text_color, history, now);
        }
        ui.label(
            egui::RichText::new("Text")
                .size(11.0)
                .color(palette.text_muted),
        );

        let mut highlight = canvas.active_style.highlight_color;
        let resp = ui.color_edit_button_srgba(&mut highlight);
        if resp.changed() {
            let now = ui.input(|i| i.time);
            set_highlight_color(document, canvas, highlight, history, now);
        }
        ui.label(
            egui::RichText::new("Highlight")
                .size(11.0)
                .color(palette.text_muted),
        );
    });
}

fn ribbon_view_group(
    ui: &mut egui::Ui,
    canvas: &mut CanvasState,
    status_message: &mut String,
    theme_mode: &mut ThemeMode,
    palette: ThemePalette,
) {
    ribbon_group(ui, "View", palette, |ui| {
        ui.vertical(|ui| {
            let mut zoom_percent = canvas.zoom * 100.0;
            if ui
                .add(
                    egui::DragValue::new(&mut zoom_percent)
                        .range(50.0..=300.0)
                        .speed(1.0)
                        .fixed_decimals(0)
                        .suffix("%"),
                )
                .changed()
            {
                canvas.zoom_mode = ZoomMode::Manual;
                canvas.zoom = (zoom_percent / 100.0).clamp(0.5, 3.0);
            }
        });
        if ui.button("↺").clicked() {
            canvas.zoom_mode = if canvas.imported_docx_view {
                ZoomMode::FitPage
            } else {
                ZoomMode::Manual
            };
            canvas.zoom = 1.0;
            canvas.pan = egui::Vec2::ZERO;
            *status_message = "View reset".to_owned();
        }
        ui.separator();
        if theme_switch(ui, theme_mode, palette, false) {
            *status_message = format!("Theme switched to {}", theme_mode.label());
        }
    });
}

fn ribbon_grammer_actions_group(
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

fn ribbon_grammer_settings_group(
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

fn table_format_group(
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

fn ribbon_picture_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    history: &mut ChangeHistory,
    palette: ThemePalette,
) {
    let Some(image_id) = canvas.selected_image_id else {
        ribbon_info_group(
            ui,
            "Picture Format",
            "Click an image to select it.",
            palette,
        );
        return;
    };

    let image_opt = document
        .paragraph_images
        .iter()
        .flatten()
        .find(|img| img.id == image_id)
        .cloned();

    let Some(image) = image_opt else {
        return;
    };

    ribbon_group(ui, "Size", palette, |ui| {
        ui.label(
            egui::RichText::new("W:")
                .size(11.0)
                .color(palette.text_muted),
        );
        let mut width = image.width_points;
        let aspect = image.height_points / image.width_points.max(1.0);
        let resp = ui.add(
            egui::DragValue::new(&mut width)
                .speed(1.0)
                .range(24.0..=1200.0)
                .fixed_decimals(0)
                .suffix(" pt"),
        );
        if resp.changed() {
            let now = ui.input(|i| i.time);
            history.checkpoint_coalesced(document, now);
            let new_h = (width * aspect).max(24.0);
            document.resize_image_by_id(image_id, width, new_h);
            *status_message = format!("Image: {:.0} × {:.0} pt", width, new_h);
        }

        ui.label(
            egui::RichText::new("H:")
                .size(11.0)
                .color(palette.text_muted),
        );
        let mut height = image.height_points;
        let aspect_inv = image.width_points / image.height_points.max(1.0);
        let resp = ui.add(
            egui::DragValue::new(&mut height)
                .speed(1.0)
                .range(24.0..=1200.0)
                .fixed_decimals(0)
                .suffix(" pt"),
        );
        if resp.changed() {
            let now = ui.input(|i| i.time);
            history.checkpoint_coalesced(document, now);
            let new_w = (height * aspect_inv).max(24.0);
            document.resize_image_by_id(image_id, new_w, height);
            *status_message = format!("Image: {:.0} × {:.0} pt", new_w, height);
        }
    });

    ribbon_group(ui, "Adjust", palette, |ui| {
        if ui.button("Reset Size").clicked() {
            reset_image_size(document, canvas, image_id, status_message, history);
        }
        ui.separator();
        ui.label(
            egui::RichText::new(format!("Alt: {}", image.alt_text))
                .size(11.0)
                .color(palette.text_muted),
        );
    });

    ribbon_group(ui, "Transparency", palette, |ui| {
        let mut opacity_pct = image.opacity * 100.0;
        let resp = ui.add(
            egui::DragValue::new(&mut opacity_pct)
                .speed(1.0)
                .range(0.0..=100.0)
                .fixed_decimals(0)
                .suffix("%"),
        );
        if resp.changed() {
            let now = ui.input(|i| i.time);
            set_image_opacity(
                document,
                image_id,
                opacity_pct / 100.0,
                status_message,
                history,
                now,
            );
        }
        ui.vertical(|ui| {
            ui.spacing_mut().slider_width = 80.0;
            let mut opacity_val = image.opacity;
            let resp = ui.add(egui::Slider::new(&mut opacity_val, 0.0..=1.0).show_value(false));
            if resp.changed() {
                let now = ui.input(|i| i.time);
                set_image_opacity(
                    document,
                    image_id,
                    opacity_val,
                    status_message,
                    history,
                    now,
                );
            }
        });
    });

    ribbon_group(ui, "Text Wrap", palette, |ui| {
        for wrap in WrapMode::ALL {
            let selected = image.wrap_mode == wrap;
            if format_button(ui, selected, wrap.label(), palette)
                .on_hover_text(wrap.label())
                .clicked()
            {
                let now = ui.input(|i| i.time);
                history.checkpoint(document, now);
                set_image_wrap_mode(document, image_id, wrap, status_message, history);
                // Auto-switch layout mode based on wrap
                if wrap == WrapMode::Inline {
                    document.set_image_layout_mode(image_id, ImageLayoutMode::Inline);
                } else {
                    document.set_image_layout_mode(image_id, ImageLayoutMode::Floating);
                }
            }
        }
    });

    ribbon_group(ui, "Layout", palette, |ui| {
        let is_inline = image.layout_mode == ImageLayoutMode::Inline;
        if format_button(ui, is_inline, "Inline", palette)
            .on_hover_text("Inline with text")
            .clicked()
        {
            let now = ui.input(|i| i.time);
            history.checkpoint(document, now);
            document.set_image_layout_mode(image_id, ImageLayoutMode::Inline);
            *status_message = "Layout: Inline".to_owned();
        }
        if format_button(ui, !is_inline, "Float", palette)
            .on_hover_text("Floating (independent of text)")
            .clicked()
        {
            let now = ui.input(|i| i.time);
            history.checkpoint(document, now);
            document.set_image_layout_mode(image_id, ImageLayoutMode::Floating);
            *status_message = "Layout: Floating".to_owned();
        }

        ui.separator();

        let mut lock_ar = image.lock_aspect_ratio;
        if ui
            .checkbox(&mut lock_ar, "Lock Ratio")
            .on_hover_text("Lock aspect ratio when resizing")
            .changed()
        {
            let now = ui.input(|i| i.time);
            history.checkpoint(document, now);
            document.set_image_lock_aspect_ratio(image_id, lock_ar);
        }

        let mut move_text = image.move_with_text;
        if ui
            .checkbox(&mut move_text, "Move with text")
            .on_hover_text("Image moves when anchor paragraph moves")
            .changed()
        {
            let now = ui.input(|i| i.time);
            history.checkpoint(document, now);
            document.set_image_move_with_text(image_id, move_text);
        }
    });

    ribbon_group(ui, "Arrange", palette, |ui| {
        if ui
            .button("▲ Forward")
            .on_hover_text("Bring forward (increase z-order)")
            .clicked()
        {
            let now = ui.input(|i| i.time);
            history.checkpoint(document, now);
            document.set_image_z_index(image_id, image.z_index + 1);
            *status_message = format!("Z-order: {}", image.z_index + 1);
        }
        if ui
            .button("▼ Backward")
            .on_hover_text("Send backward (decrease z-order)")
            .clicked()
        {
            let now = ui.input(|i| i.time);
            history.checkpoint(document, now);
            document.set_image_z_index(image_id, image.z_index - 1);
            *status_message = format!("Z-order: {}", image.z_index - 1);
        }
    });

    ribbon_group(ui, "Quality", palette, |ui| {
        if format_button(
            ui,
            image.rendering == ImageRendering::Smooth,
            "Smooth",
            palette,
        )
        .on_hover_text("Bilinear filtering (smooth edges)")
        .clicked()
        {
            set_image_rendering(
                document,
                canvas,
                image_id,
                ImageRendering::Smooth,
                status_message,
                history,
            );
        }
        if format_button(
            ui,
            image.rendering == ImageRendering::Crisp,
            "Crisp",
            palette,
        )
        .on_hover_text("Nearest-neighbor (pixel-perfect / sharp)")
        .clicked()
        {
            set_image_rendering(
                document,
                canvas,
                image_id,
                ImageRendering::Crisp,
                status_message,
                history,
            );
        }
    });
}

fn ribbon_info_group(ui: &mut egui::Ui, title: &str, message: &str, palette: ThemePalette) {
    ribbon_group(ui, title, palette, |ui| {
        ui.label(
            egui::RichText::new(message)
                .size(11.0)
                .color(palette.text_muted),
        );
    });
}

fn ribbon_group(
    ui: &mut egui::Ui,
    title: &str,
    palette: ThemePalette,
    add_contents: impl FnOnce(&mut egui::Ui),
) {
    const RIBBON_GROUP_CONTENT_HEIGHT: f32 = 64.0;

    egui::Frame::new()
        .fill(palette.ribbon_group_bg)
        .inner_margin(egui::Margin::symmetric(8, 6))
        .stroke(egui::Stroke::new(1.0, palette.border))
        .corner_radius(4.0)
        .show(ui, |ui| {
            ui.set_min_height(RIBBON_GROUP_CONTENT_HEIGHT);
            ui.vertical(|ui| {
                ui.horizontal(|ui| {
                    add_contents(ui);
                });
                ui.add_space(4.0);
                ui.label(
                    egui::RichText::new(title)
                        .size(10.0)
                        .color(palette.text_muted),
                );
            });
        });
}

fn format_button(
    ui: &mut egui::Ui,
    active: bool,
    label: &str,
    palette: ThemePalette,
) -> egui::Response {
    let fill = if active {
        palette.accent.gamma_multiply(0.22)
    } else {
        palette.ribbon_group_bg
    };
    let stroke = if active {
        egui::Stroke::new(1.0, palette.accent)
    } else {
        egui::Stroke::new(1.0, palette.border)
    };
    ui.add(
        egui::Button::new(egui::RichText::new(label).strong().color(if active {
            palette.tab_active_fg
        } else {
            palette.text_primary
        }))
        .min_size(egui::vec2(24.0, 24.0))
        .fill(fill)
        .stroke(stroke)
        .corner_radius(3.0),
    )
}

fn alignment_button(
    ui: &mut egui::Ui,
    active: bool,
    alignment: ParagraphAlignment,
    palette: ThemePalette,
) -> egui::Response {
    let fill = if active {
        palette.accent.gamma_multiply(0.22)
    } else {
        palette.ribbon_group_bg
    };
    let stroke = if active {
        egui::Stroke::new(1.0, palette.accent)
    } else {
        egui::Stroke::new(1.0, palette.border)
    };
    let response = ui.add(
        egui::Button::new("")
            .min_size(egui::vec2(24.0, 24.0))
            .fill(fill)
            .stroke(stroke)
            .corner_radius(3.0),
    );

    let stroke = egui::Stroke::new(
        1.6,
        if active {
            palette.tab_active_fg
        } else {
            palette.text_primary
        },
    );
    let rect = response.rect.shrink2(egui::vec2(5.0, 5.0));
    let line_gap = rect.height() / 3.0;
    let line_y = [
        rect.top(),
        rect.top() + line_gap,
        rect.top() + line_gap * 2.0,
        rect.bottom(),
    ];

    for (index, y) in line_y.into_iter().enumerate() {
        let width_factor = match alignment {
            ParagraphAlignment::Left => [1.0, 0.78, 0.92, 0.64][index],
            ParagraphAlignment::Center => [0.72, 1.0, 0.84, 0.6][index],
            ParagraphAlignment::Right => [0.7, 1.0, 0.82, 0.62][index],
            ParagraphAlignment::Justify => 1.0,
        };
        let line_width = rect.width() * width_factor;
        let x = match alignment {
            ParagraphAlignment::Left | ParagraphAlignment::Justify => rect.left(),
            ParagraphAlignment::Center => rect.center().x - line_width * 0.5,
            ParagraphAlignment::Right => rect.right() - line_width,
        };
        ui.painter()
            .line_segment([egui::pos2(x, y), egui::pos2(x + line_width, y)], stroke);
    }

    response
}
