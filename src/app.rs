use std::path::PathBuf;

use eframe::{egui, App, CreationContext, Frame};
use rfd::FileDialog;

use crate::{
    canvas::paint_document_canvas,
    document::{CharacterStyle, DocumentState, FontChoice},
};

#[derive(Clone, Copy, PartialEq, Eq)]
pub enum ThemeMode {
    Light,
    Dark,
}

impl ThemeMode {
    pub const ALL: [Self; 2] = [Self::Light, Self::Dark];

    pub const fn label(self) -> &'static str {
        match self {
            Self::Light => "Light",
            Self::Dark => "Dark",
        }
    }
}

#[derive(Clone, Copy, PartialEq, Eq)]
enum RibbonTab {
    Home,
    Insert,
    Design,
    Layout,
    View,
}

impl RibbonTab {
    const ALL: [Self; 5] = [
        Self::Home,
        Self::Insert,
        Self::Design,
        Self::Layout,
        Self::View,
    ];

    const fn label(self) -> &'static str {
        match self {
            Self::Home => "Home",
            Self::Insert => "Insert",
            Self::Design => "Design",
            Self::Layout => "Layout",
            Self::View => "View",
        }
    }
}

#[derive(Clone, Copy)]
struct ThemePalette {
    title_bg: egui::Color32,
    title_fg: egui::Color32,
    title_muted: egui::Color32,
    title_button_bg: egui::Color32,
    tab_bg: egui::Color32,
    tab_fg: egui::Color32,
    tab_active_bg: egui::Color32,
    tab_active_fg: egui::Color32,
    ribbon_bg: egui::Color32,
    ribbon_group_bg: egui::Color32,
    border: egui::Color32,
    text_primary: egui::Color32,
    text_muted: egui::Color32,
    workspace_bg: egui::Color32,
    status_bg: egui::Color32,
    accent: egui::Color32,
}

fn theme_palette(mode: ThemeMode) -> ThemePalette {
    match mode {
        ThemeMode::Light => ThemePalette {
            title_bg: egui::Color32::from_rgb(43, 87, 154),
            title_fg: egui::Color32::from_rgb(247, 250, 255),
            title_muted: egui::Color32::from_rgb(214, 227, 247),
            title_button_bg: egui::Color32::from_rgba_premultiplied(255, 255, 255, 24),
            tab_bg: egui::Color32::from_rgb(43, 87, 154),
            tab_fg: egui::Color32::from_rgb(239, 246, 255),
            tab_active_bg: egui::Color32::from_rgb(245, 248, 252),
            tab_active_fg: egui::Color32::from_rgb(31, 64, 115),
            ribbon_bg: egui::Color32::from_rgb(244, 246, 249),
            ribbon_group_bg: egui::Color32::from_rgb(251, 252, 254),
            border: egui::Color32::from_rgb(202, 210, 224),
            text_primary: egui::Color32::from_rgb(30, 34, 40),
            text_muted: egui::Color32::from_rgb(94, 101, 114),
            workspace_bg: egui::Color32::from_rgb(215, 217, 220),
            status_bg: egui::Color32::from_rgb(235, 238, 243),
            accent: egui::Color32::from_rgb(54, 109, 193),
        },
        ThemeMode::Dark => ThemePalette {
            title_bg: egui::Color32::from_rgb(28, 34, 47),
            title_fg: egui::Color32::from_rgb(236, 241, 251),
            title_muted: egui::Color32::from_rgb(156, 170, 197),
            title_button_bg: egui::Color32::from_rgba_premultiplied(255, 255, 255, 20),
            tab_bg: egui::Color32::from_rgb(28, 34, 47),
            tab_fg: egui::Color32::from_rgb(213, 222, 240),
            tab_active_bg: egui::Color32::from_rgb(65, 79, 105),
            tab_active_fg: egui::Color32::from_rgb(241, 247, 255),
            ribbon_bg: egui::Color32::from_rgb(49, 55, 66),
            ribbon_group_bg: egui::Color32::from_rgb(57, 64, 77),
            border: egui::Color32::from_rgb(84, 94, 112),
            text_primary: egui::Color32::from_rgb(233, 238, 248),
            text_muted: egui::Color32::from_rgb(172, 181, 197),
            workspace_bg: egui::Color32::from_rgb(50, 52, 56),
            status_bg: egui::Color32::from_rgb(43, 49, 59),
            accent: egui::Color32::from_rgb(109, 157, 228),
        },
    }
}

pub struct CanvasState {
    pub zoom: f32,
    pub pan: egui::Vec2,
    pub selection: egui::text_selection::CCursorRange,
    pub active_style: CharacterStyle,
    pub last_interaction_time: f64,
}

impl Default for CanvasState {
    fn default() -> Self {
        Self {
            zoom: 1.0,
            pan: egui::Vec2::ZERO,
            selection: egui::text_selection::CCursorRange::default(),
            active_style: CharacterStyle::default(),
            last_interaction_time: 0.0,
        }
    }
}

pub struct WorsApp {
    document: DocumentState,
    canvas: CanvasState,
    active_tab: RibbonTab,
    theme_mode: ThemeMode,
    status_message: String,
    current_path: Option<PathBuf>,
}

impl WorsApp {
    pub fn new(cc: &CreationContext<'_>) -> Self {
        cc.egui_ctx
            .set_pixels_per_point(cc.egui_ctx.pixels_per_point());

        let theme_mode = ThemeMode::Light;
        configure_theme(&cc.egui_ctx, theme_mode, theme_palette(theme_mode));

        Self {
            document: DocumentState::bootstrap(),
            canvas: CanvasState::default(),
            active_tab: RibbonTab::Home,
            theme_mode,
            status_message: "Ready".to_owned(),
            current_path: None,
        }
    }
}

impl App for WorsApp {
    fn ui(&mut self, ui: &mut egui::Ui, _frame: &mut Frame) {
        handle_global_shortcuts(
            ui,
            &mut self.document,
            &mut self.current_path,
            &mut self.status_message,
        );

        let palette = theme_palette(self.theme_mode);
        let status_line = self.status_message.clone();
        configure_theme(ui.ctx(), self.theme_mode, palette);

        egui::Panel::top("title_bar")
            .frame(egui::Frame::new().fill(palette.title_bg))
            .show_inside(ui, |ui| {
                paint_title_bar(
                    ui,
                    &self.document,
                    &self.current_path,
                    &status_line,
                    &mut self.theme_mode,
                    &mut self.status_message,
                    palette,
                );
            });

        egui::Panel::top("tabs_bar")
            .frame(egui::Frame::new().fill(palette.tab_bg))
            .show_inside(ui, |ui| {
                paint_tab_row(ui, &mut self.active_tab, palette);
            });

        egui::Panel::top("ribbon")
            .frame(
                egui::Frame::new()
                    .fill(palette.ribbon_bg)
                    .stroke(egui::Stroke::new(1.0, palette.border)),
            )
            .show_inside(ui, |ui| {
                paint_ribbon(
                    ui,
                    &mut self.document,
                    &mut self.canvas,
                    &mut self.active_tab,
                    &mut self.status_message,
                    &mut self.current_path,
                    &mut self.theme_mode,
                    palette,
                );
            });

        egui::CentralPanel::default()
            .frame(egui::Frame::new().fill(palette.workspace_bg))
            .show_inside(ui, |ui| {
                paint_document_canvas(ui, &mut self.document, &mut self.canvas, self.theme_mode);
            });

        egui::Panel::bottom("status")
            .frame(
                egui::Frame::new()
                    .fill(palette.status_bg)
                    .stroke(egui::Stroke::new(1.0, palette.border))
                    .inner_margin(egui::Margin::symmetric(10, 6)),
            )
            .show_inside(ui, |ui| {
                paint_status_bar(
                    ui,
                    &self.document,
                    &self.canvas,
                    &self.status_message,
                    palette,
                );
            });
    }
}

fn configure_theme(ctx: &egui::Context, mode: ThemeMode, palette: ThemePalette) {
    let mut style = (*ctx.global_style()).clone();
    style.spacing.item_spacing = egui::vec2(8.0, 8.0);
    style.spacing.button_padding = egui::vec2(8.0, 5.0);
    style.spacing.combo_width = 130.0;
    style.visuals = match mode {
        ThemeMode::Light => egui::Visuals::light(),
        ThemeMode::Dark => egui::Visuals::dark(),
    };
    style.visuals.override_text_color = Some(palette.text_primary);
    style.visuals.widgets.inactive.bg_fill = palette.ribbon_group_bg;
    style.visuals.widgets.inactive.weak_bg_fill = palette.status_bg;
    style.visuals.widgets.inactive.bg_stroke = egui::Stroke::new(1.0, palette.border);
    style.visuals.widgets.inactive.fg_stroke = egui::Stroke::new(1.0, palette.text_primary);
    style.visuals.widgets.hovered.bg_fill = palette.ribbon_group_bg;
    style.visuals.widgets.hovered.weak_bg_fill = palette.tab_active_bg;
    style.visuals.widgets.hovered.bg_stroke = egui::Stroke::new(1.0, palette.accent);
    style.visuals.widgets.hovered.fg_stroke = egui::Stroke::new(1.0, palette.text_primary);
    style.visuals.widgets.active.bg_fill = palette.tab_active_bg;
    style.visuals.widgets.active.weak_bg_fill = palette.tab_active_bg;
    style.visuals.widgets.active.bg_stroke = egui::Stroke::new(1.0, palette.accent);
    style.visuals.widgets.active.fg_stroke = egui::Stroke::new(1.0, palette.text_primary);
    style.visuals.widgets.open = style.visuals.widgets.active;
    style.visuals.selection.bg_fill = palette.accent.gamma_multiply(0.35);
    style.visuals.selection.stroke = egui::Stroke::new(1.0, palette.accent);
    style.visuals.panel_fill = palette.ribbon_bg;
    style.visuals.window_fill = palette.ribbon_group_bg;
    ctx.set_global_style(style);
}

fn paint_title_bar(
    ui: &mut egui::Ui,
    document: &DocumentState,
    current_path: &Option<PathBuf>,
    status_message: &str,
    theme_mode: &mut ThemeMode,
    status_target: &mut String,
    palette: ThemePalette,
) {
    let path_label = current_path
        .as_ref()
        .map(|path| path.to_string_lossy().into_owned())
        .unwrap_or_else(|| "Unsaved document".to_owned());

    egui::Frame::new()
        .inner_margin(egui::Margin::symmetric(12, 8))
        .show(ui, |ui| {
            ui.horizontal(|ui| {
                ui.label(
                    egui::RichText::new("wors")
                        .size(15.0)
                        .strong()
                        .color(palette.title_fg),
                );
                ui.separator();
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

                ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
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
}

fn theme_switch(
    ui: &mut egui::Ui,
    theme_mode: &mut ThemeMode,
    palette: ThemePalette,
    dark_surface: bool,
) -> bool {
    let original = *theme_mode;
    let inactive_fill = if dark_surface {
        palette.title_button_bg
    } else {
        palette.ribbon_group_bg
    };
    let inactive_text = if dark_surface {
        palette.title_muted
    } else {
        palette.text_primary
    };
    let active_fill = if dark_surface {
        palette.tab_active_bg
    } else {
        palette.status_bg
    };
    let active_text = if dark_surface {
        palette.tab_active_fg
    } else {
        palette.text_primary
    };

    ui.horizontal(|ui| {
        ui.spacing_mut().item_spacing.x = 4.0;
        for mode in ThemeMode::ALL {
            let selected = *theme_mode == mode;
            let button = egui::Button::new(
                egui::RichText::new(mode.label())
                    .size(11.0)
                    .color(if selected { active_text } else { inactive_text }),
            )
            .min_size(egui::vec2(54.0, 22.0))
            .fill(if selected { active_fill } else { inactive_fill })
            .stroke(egui::Stroke::new(1.0, palette.border))
            .corner_radius(4.0);
            if ui.add(button).clicked() {
                *theme_mode = mode;
            }
        }
    });
    *theme_mode != original
}

fn paint_tab_row(ui: &mut egui::Ui, active_tab: &mut RibbonTab, palette: ThemePalette) {
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
                ui.add(file_button);

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
                        .corner_radius(4.0);
                    if ui.add(button).clicked() {
                        *active_tab = tab;
                    }
                }
            });
        });
}

fn paint_ribbon(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    active_tab: &mut RibbonTab,
    status_message: &mut String,
    current_path: &mut Option<PathBuf>,
    theme_mode: &mut ThemeMode,
    palette: ThemePalette,
) {
    sync_active_style(document, canvas);

    egui::Frame::new()
        .inner_margin(egui::Margin::symmetric(8, 8))
        .show(ui, |ui| {
            ui.horizontal_wrapped(|ui| match active_tab {
                RibbonTab::Home => {
                    ribbon_file_group(ui, document, canvas, status_message, current_path, palette);
                    ribbon_font_group(ui, document, canvas, palette);
                    ribbon_color_group(ui, document, canvas, palette);
                    ribbon_view_group(ui, canvas, status_message, theme_mode, palette);
                }
                RibbonTab::Insert => {
                    ribbon_file_group(ui, document, canvas, status_message, current_path, palette);
                    ribbon_info_group(
                        ui,
                        "Insert",
                        "Import supports .txt, .md, .markdown, and .docx.",
                        palette,
                    );
                }
                RibbonTab::Design => {
                    ribbon_font_group(ui, document, canvas, palette);
                    ribbon_color_group(ui, document, canvas, palette);
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
                        "Command+S Save, Command+B Bold, Command+I Italic, Command+U Underline",
                        palette,
                    );
                }
            });
        });
}

fn ribbon_file_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    current_path: &mut Option<PathBuf>,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Clipboard", palette, |ui| {
        if ui.button("Open").clicked() {
            open_document(document, canvas, status_message, current_path);
        }
        if ui.button("Save").clicked() {
            save_document(document, status_message, current_path);
        }
    });
}

fn ribbon_font_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Font", palette, |ui| {
        egui::ComboBox::from_id_salt("font_choice")
            .selected_text(canvas.active_style.font_choice.label())
            .width(130.0)
            .show_ui(ui, |ui| {
                for font in FontChoice::ALL {
                    if ui
                        .selectable_label(canvas.active_style.font_choice == font, font.label())
                        .clicked()
                    {
                        set_font_choice(document, canvas, font);
                    }
                }
            });

        let mut font_size = canvas.active_style.font_size_points;
        if ui
            .add(
                egui::DragValue::new(&mut font_size)
                    .range(8.0..=72.0)
                    .speed(0.25)
                    .fixed_decimals(1),
            )
            .changed()
        {
            set_font_size(document, canvas, font_size.clamp(8.0, 72.0));
        }

        ui.separator();

        if format_button(ui, canvas.active_style.bold, "B", palette).clicked() {
            toggle_bold(document, canvas);
        }
        if format_button(ui, canvas.active_style.italic, "I", palette).clicked() {
            toggle_italic(document, canvas);
        }
        if format_button(ui, canvas.active_style.underline, "U", palette).clicked() {
            toggle_underline(document, canvas);
        }
        if format_button(ui, canvas.active_style.strikethrough, "S", palette).clicked() {
            toggle_strikethrough(document, canvas);
        }
    });
}

fn ribbon_color_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Paragraph", palette, |ui| {
        let mut text_color = canvas.active_style.text_color;
        if ui.color_edit_button_srgba(&mut text_color).changed() {
            set_text_color(document, canvas, text_color);
        }
        ui.label(
            egui::RichText::new("Text")
                .size(11.0)
                .color(palette.text_muted),
        );

        let mut highlight = canvas.active_style.highlight_color;
        if ui.color_edit_button_srgba(&mut highlight).changed() {
            set_highlight_color(document, canvas, highlight);
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
        ui.add(
            egui::Slider::new(&mut canvas.zoom, 0.5..=3.0)
                .text("Zoom")
                .step_by(0.05),
        );
        if ui.button("Reset").clicked() {
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
    egui::Frame::new()
        .fill(palette.ribbon_group_bg)
        .inner_margin(egui::Margin::symmetric(8, 6))
        .stroke(egui::Stroke::new(1.0, palette.border))
        .corner_radius(4.0)
        .show(ui, |ui| {
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
    ui.add(
        egui::Button::new(egui::RichText::new(label).strong().color(if active {
            palette.tab_active_fg
        } else {
            palette.text_primary
        }))
        .min_size(egui::vec2(24.0, 24.0))
        .fill(if active {
            palette.tab_active_bg
        } else {
            palette.ribbon_group_bg
        })
        .stroke(egui::Stroke::new(1.0, palette.border))
        .corner_radius(3.0),
    )
}

fn paint_status_bar(
    ui: &mut egui::Ui,
    document: &DocumentState,
    canvas: &CanvasState,
    status_message: &str,
    palette: ThemePalette,
) {
    ui.horizontal(|ui| {
        let plain_text = document.plain_text();
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
        ui.with_layout(egui::Layout::right_to_left(egui::Align::Center), |ui| {
            ui.label(
                egui::RichText::new(format!("{:.0}%", canvas.zoom * 100.0))
                    .size(11.0)
                    .color(palette.text_muted),
            );
        });
    });
}

fn open_document(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    current_path: &mut Option<PathBuf>,
) {
    if let Some(path) = FileDialog::new()
        .add_filter("supported", &["txt", "md", "markdown", "docx"])
        .pick_file()
    {
        match DocumentState::load_from_path(&path) {
            Ok(new_document) => {
                *document = new_document;
                canvas.selection = egui::text_selection::CCursorRange::default();
                canvas.active_style = CharacterStyle::default();
                canvas.zoom = 1.0;
                canvas.pan = egui::Vec2::ZERO;
                *current_path = match path.extension().and_then(|ext| ext.to_str()) {
                    Some("docx") => None,
                    _ => Some(path.clone()),
                };
                *status_message = format!(
                    "Imported {}",
                    path.file_name()
                        .and_then(|name| name.to_str())
                        .unwrap_or("document")
                );
            }
            Err(error) => *status_message = error,
        }
    }
}

fn save_document(
    document: &DocumentState,
    status_message: &mut String,
    current_path: &mut Option<PathBuf>,
) {
    let path = match current_path.clone() {
        Some(path) => path,
        None => match FileDialog::new()
            .add_filter("text", &["txt"])
            .add_filter("markdown", &["md"])
            .set_file_name(&document.title)
            .save_file()
        {
            Some(path) => path,
            None => return,
        },
    };

    match document.save_to_path(&path) {
        Ok(()) => {
            *current_path = Some(path.clone());
            *status_message = format!(
                "Saved {}",
                path.file_name()
                    .and_then(|name| name.to_str())
                    .unwrap_or("document")
            );
        }
        Err(error) => *status_message = error,
    }
}

fn handle_global_shortcuts(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    current_path: &mut Option<PathBuf>,
    status_message: &mut String,
) {
    if ui.input_mut(|input| input.consume_key(egui::Modifiers::COMMAND, egui::Key::S)) {
        save_document(document, status_message, current_path);
    }
}

fn toggle_bold(document: &mut DocumentState, canvas: &mut CanvasState) {
    let next_value = !canvas.active_style.bold;
    apply_selection_or_active_style(document, canvas, move |style| style.bold = next_value);
}

fn toggle_italic(document: &mut DocumentState, canvas: &mut CanvasState) {
    let next_value = !canvas.active_style.italic;
    apply_selection_or_active_style(document, canvas, move |style| style.italic = next_value);
}

fn toggle_underline(document: &mut DocumentState, canvas: &mut CanvasState) {
    let next_value = !canvas.active_style.underline;
    apply_selection_or_active_style(document, canvas, move |style| style.underline = next_value);
}

fn toggle_strikethrough(document: &mut DocumentState, canvas: &mut CanvasState) {
    let next_value = !canvas.active_style.strikethrough;
    apply_selection_or_active_style(document, canvas, move |style| {
        style.strikethrough = next_value
    });
}

fn set_font_size(document: &mut DocumentState, canvas: &mut CanvasState, font_size: f32) {
    apply_selection_or_active_style(document, canvas, move |style| {
        style.font_size_points = font_size
    });
}

fn set_font_choice(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    font_choice: FontChoice,
) {
    apply_selection_or_active_style(document, canvas, move |style| {
        style.font_choice = font_choice
    });
}

fn set_text_color(document: &mut DocumentState, canvas: &mut CanvasState, color: egui::Color32) {
    apply_selection_or_active_style(document, canvas, move |style| style.text_color = color);
}

fn set_highlight_color(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    color: egui::Color32,
) {
    apply_selection_or_active_style(document, canvas, move |style| style.highlight_color = color);
}

fn apply_selection_or_active_style(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    mutate: impl Fn(&mut CharacterStyle) + Copy,
) {
    let range = canvas.selection.as_sorted_char_range();
    if range.start < range.end {
        document.apply_style_to_range(range, mutate);
    }
    mutate(&mut canvas.active_style);
}

fn sync_active_style(document: &DocumentState, canvas: &mut CanvasState) {
    let range = canvas.selection.as_sorted_char_range();
    let cursor_index = if range.start < range.end {
        range.end
    } else {
        canvas.selection.primary.index
    };
    canvas.active_style = document.typing_style_at(cursor_index);
}
