use std::path::PathBuf;

use eframe::egui;

use crate::document::{
    DocumentState, FontChoice, ImageLayoutMode, ImageRendering, ListKind, ParagraphAlignment,
    WrapMode, OBJECT_REPLACEMENT_CHAR,
};
use crate::grammar::{GrammarConfig, GrammarStatus, Language};

use super::{
    actions::{
        insert_image, insert_page_break, open_document, reset_image_size, save_document,
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
        }
    }
}

#[derive(Default)]
pub(super) struct GrammarRibbonOutput {
    pub manual_check_requested: bool,
    pub restart_requested: bool,
    pub download_requested: bool,
    pub settings_changed: bool,
}

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

    egui::Frame::new()
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
                {
                    if history.undo(document) {
                        canvas.image_textures.clear();
                        *status_target = "Undo".to_owned();
                    }
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
                {
                    if history.redo(document) {
                        canvas.image_textures.clear();
                        *status_target = "Redo".to_owned();
                    }
                }

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

pub(super) fn paint_tab_row(
    ui: &mut egui::Ui,
    active_tab: &mut RibbonTab,
    selected_image_id: Option<usize>,
    palette: ThemePalette,
) {
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
                    .corner_radius(4.0);
                    if ui.add(button).clicked() {
                        *active_tab = RibbonTab::Picture;
                    }
                }
            });
        });
}

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
            });
        });
    output
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
        egui::ComboBox::from_id_salt("font_choice")
            .selected_text(canvas.active_style.font_choice.label())
            .width(130.0)
            .show_ui(ui, |ui| {
                for font in FontChoice::ALL {
                    if ui
                        .selectable_label(canvas.active_style.font_choice == font, font.label())
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
