use std::path::PathBuf;

use eframe::egui;

use crate::document::{
    DocumentState, FontChoice, HeaderFooterKind, HeaderFooterVariant, ImageLayoutMode,
    ImageRendering, ListKind, PageMargins, PageSetup, PageSize, ParagraphAlignment, TextRun,
    WrapMode, OBJECT_REPLACEMENT_CHAR,
};
use crate::grammar::{GrammarConfig, GrammarStatus, Language};

use super::{
    actions::{
        delete_table_column, delete_table_row, insert_image, insert_page_break,
        insert_section_break, insert_table, insert_table_column, insert_table_row, open_document,
        reset_image_size, save_document, save_document_as, set_font_choice, set_font_size,
        set_highlight_color, set_image_opacity, set_image_rendering, set_image_wrap_mode,
        set_paragraph_alignment, set_text_color, sync_active_style, toggle_bold,
        toggle_bullet_list, toggle_italic, toggle_ordered_list, toggle_strikethrough,
        toggle_underline,
    },
    palette::{theme_switch, ThemeMode, ThemePalette},
    ActiveHeaderFooter, CanvasState, ChangeHistory, ZoomMode,
};

#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum RibbonTab {
    Home,
    Insert,
    Design,
    Layout,
    View,
    Grammar,
    HeaderFooter,
    Picture,
    Table,
}

fn current_section_id(
    document: &DocumentState,
    canvas: &CanvasState,
) -> crate::document::SectionId {
    if let Some(active) = canvas.active_header_footer {
        return active.section_id;
    }
    let paragraph_index = document
        .paragraphs()
        .iter()
        .position(|paragraph| {
            paragraph.range.contains(&canvas.selection.primary.index)
                || paragraph.range.start == canvas.selection.primary.index
        })
        .unwrap_or(0);
    document.section_at_paragraph(paragraph_index).id
}

impl RibbonTab {
    const ALL: [Self; 6] = [
        Self::Home,
        Self::Insert,
        Self::Design,
        Self::Layout,
        Self::View,
        Self::Grammar,
    ];

    const fn label(self) -> &'static str {
        match self {
            Self::Home => "Home",
            Self::Insert => "Insert",
            Self::Design => "Design",
            Self::Layout => "Layout",
            Self::View => "View",
            Self::Grammar => "Grammar",
            Self::HeaderFooter => "Header & Footer",
            Self::Picture => "Picture Format",
            Self::Table => "Table Format",
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
    active_header_footer: bool,
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

                if active_header_footer {
                    ui.separator();
                    let selected = *active_tab == RibbonTab::HeaderFooter;
                    let fg = if selected {
                        palette.tab_active_fg
                    } else {
                        palette.tab_fg
                    };
                    let bg = if selected {
                        palette.tab_active_bg
                    } else {
                        egui::Color32::TRANSPARENT
                    };
                    let button = egui::Button::new(
                        egui::RichText::new("Header & Footer")
                            .size(13.0)
                            .color(fg)
                            .strong(),
                    )
                    .min_size(egui::vec2(126.0, 28.0))
                    .fill(bg)
                    .stroke(if selected {
                        egui::Stroke::new(1.0, palette.accent)
                    } else {
                        egui::Stroke::NONE
                    })
                    .corner_radius(0.0);
                    if ui.add(button).clicked() {
                        *active_tab = RibbonTab::HeaderFooter;
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
        .inner_margin(egui::Margin::symmetric(8, 4))
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
                        "Import supports .txt, .md, .markdown, .docx, and .odt with images.",
                        palette,
                    );
                }
                RibbonTab::Design => {
                    ribbon_font_group(ui, document, canvas, history, palette);
                    ribbon_paragraph_group(ui, document, canvas, history, palette);
                    ribbon_color_group(ui, document, canvas, history, palette);
                }
                RibbonTab::Layout => {
                    ribbon_page_setup_group(ui, document, canvas, status_message, history, palette);
                    ribbon_flow_group(ui, document, canvas, status_message, history, palette);
                    ribbon_layout_header_footer_group(
                        ui,
                        document,
                        canvas,
                        status_message,
                        history,
                        palette,
                    );
                    ribbon_advanced_page_setup_group(
                        ui,
                        document,
                        canvas,
                        status_message,
                        history,
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
                RibbonTab::Grammar => {
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
                RibbonTab::HeaderFooter => {
                    ribbon_header_footer_insert_group(
                        ui,
                        document,
                        canvas,
                        status_message,
                        history,
                        palette,
                    );
                    ribbon_header_footer_options_group(
                        ui,
                        document,
                        canvas,
                        status_message,
                        history,
                        palette,
                    );
                    ribbon_header_footer_position_group(
                        ui,
                        document,
                        canvas,
                        status_message,
                        history,
                        palette,
                    );
                    ribbon_header_footer_actions_group(
                        ui,
                        document,
                        canvas,
                        status_message,
                        history,
                        palette,
                    );
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
            let _ = open_document(document, canvas, status_message, current_path, history);
        }
        if ui.button("💾 Save").clicked() {
            let _ = save_document(document, status_message, current_path);
        }
        if ui.button("Save As").clicked() {
            let _ = save_document_as(document, status_message, current_path);
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
        if ui.button("Section Break").clicked() {
            insert_section_break(document, canvas, status_message, history);
        }
        if ui.button("Header").clicked() {
            let section_id = current_section_id(document, canvas);
            let variant = HeaderFooterVariant::Default;
            canvas.active_header_footer = Some(ActiveHeaderFooter {
                kind: HeaderFooterKind::Header,
                section_id,
                variant,
                page_number: 1,
            });
            canvas.active_header_footer_cursor = document
                .resolve_header_footer_slot(section_id, HeaderFooterKind::Header, variant)
                .story
                .runs
                .iter()
                .map(|run| run.text.chars().count())
                .sum();
            canvas.active_header_footer_selection = egui::text_selection::CCursorRange::one(
                egui::epaint::text::cursor::CCursor::new(canvas.active_header_footer_cursor),
            );
            *status_message = "Editing header".to_owned();
        }
        if ui.button("Footer").clicked() {
            let section_id = current_section_id(document, canvas);
            let variant = HeaderFooterVariant::Default;
            canvas.active_header_footer = Some(ActiveHeaderFooter {
                kind: HeaderFooterKind::Footer,
                section_id,
                variant,
                page_number: 1,
            });
            canvas.active_header_footer_cursor = document
                .resolve_header_footer_slot(section_id, HeaderFooterKind::Footer, variant)
                .story
                .runs
                .iter()
                .map(|run| run.text.chars().count())
                .sum();
            canvas.active_header_footer_selection = egui::text_selection::CCursorRange::one(
                egui::epaint::text::cursor::CCursor::new(canvas.active_header_footer_cursor),
            );
            *status_message = "Editing footer".to_owned();
        }
        if ui.button("Page Number").clicked() {
            history.checkpoint(document, ui.input(|i| i.time));
            let (section_id, variant, kind) = canvas
                .active_header_footer
                .map(|active| (active.section_id, active.variant, active.kind))
                .unwrap_or_else(|| {
                    (
                        current_section_id(document, canvas),
                        HeaderFooterVariant::Default,
                        HeaderFooterKind::Footer,
                    )
                });
            let story = document
                .header_footer_story_mut_materialized(section_id, kind, variant)
                .expect("current section exists");
            let text = story.plain_text();
            if text.trim().is_empty() {
                story.runs = vec![TextRun {
                    text: "Page { PAGE } of { NUMPAGES }".to_owned(),
                    style: canvas.active_style,
                }];
            } else {
                story.runs.push(TextRun {
                    text: " { PAGE }".to_owned(),
                    style: canvas.active_style,
                });
            }
            document.sync_compat_from_first_section();
            canvas.active_header_footer = Some(ActiveHeaderFooter {
                kind,
                section_id,
                variant,
                page_number: 1,
            });
            canvas.active_header_footer_cursor = document
                .resolve_header_footer_slot(section_id, kind, variant)
                .story
                .plain_text()
                .chars()
                .count();
            canvas.active_header_footer_selection = egui::text_selection::CCursorRange::one(
                egui::epaint::text::cursor::CCursor::new(canvas.active_header_footer_cursor),
            );
            *status_message = "Page number inserted".to_owned();
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

fn ribbon_page_setup_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    history: &mut ChangeHistory,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Page Setup", palette, |ui| {
        ui.menu_button("Margins ▾", |ui| {
            if ui.button("Normal").clicked() {
                set_current_section_margins(
                    document,
                    canvas,
                    history,
                    status_message,
                    PageMargins {
                        top_points: 72.0,
                        right_points: 72.0,
                        bottom_points: 72.0,
                        left_points: 72.0,
                    },
                    "Normal margins",
                    ui.input(|i| i.time),
                );
                ui.close();
            }
            if ui.button("Narrow").clicked() {
                set_current_section_margins(
                    document,
                    canvas,
                    history,
                    status_message,
                    PageMargins {
                        top_points: 36.0,
                        right_points: 36.0,
                        bottom_points: 36.0,
                        left_points: 36.0,
                    },
                    "Narrow margins",
                    ui.input(|i| i.time),
                );
                ui.close();
            }
            if ui.button("Moderate").clicked() {
                set_current_section_margins(
                    document,
                    canvas,
                    history,
                    status_message,
                    PageMargins {
                        top_points: 72.0,
                        right_points: 54.0,
                        bottom_points: 72.0,
                        left_points: 54.0,
                    },
                    "Moderate margins",
                    ui.input(|i| i.time),
                );
                ui.close();
            }
            if ui.button("Wide").clicked() {
                set_current_section_margins(
                    document,
                    canvas,
                    history,
                    status_message,
                    PageMargins {
                        top_points: 72.0,
                        right_points: 144.0,
                        bottom_points: 72.0,
                        left_points: 144.0,
                    },
                    "Wide margins",
                    ui.input(|i| i.time),
                );
                ui.close();
            }
        });

        ui.menu_button("Size ▾", |ui| {
            for (label, size) in page_size_presets() {
                if ui.button(label).clicked() {
                    set_current_section_page_size(
                        document,
                        canvas,
                        history,
                        status_message,
                        size,
                        label,
                        ui.input(|i| i.time),
                    );
                    ui.close();
                }
            }
        });

        ui.menu_button("Orientation ▾", |ui| {
            if ui.button("Portrait").clicked() {
                set_current_section_orientation(
                    document,
                    canvas,
                    history,
                    status_message,
                    true,
                    ui.input(|i| i.time),
                );
                ui.close();
            }
            if ui.button("Landscape").clicked() {
                set_current_section_orientation(
                    document,
                    canvas,
                    history,
                    status_message,
                    false,
                    ui.input(|i| i.time),
                );
                ui.close();
            }
        });
    });
}

fn ribbon_flow_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    history: &mut ChangeHistory,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Flow", palette, |ui| {
        ui.menu_button("Columns ▾", |ui| {
            ui.add_enabled(false, egui::Button::new("One"));
            ui.add_enabled(false, egui::Button::new("Two"));
            ui.add_enabled(false, egui::Button::new("More Columns..."));
        });

        ui.menu_button("Breaks ▾", |ui| {
            if ui.button("Page").clicked() {
                insert_page_break(document, canvas, status_message, history);
                ui.close();
            }
            if ui.button("Section").clicked() {
                insert_section_break(document, canvas, status_message, history);
                ui.close();
            }
        });

        ui.menu_button("Line Numbers ▾", |ui| {
            ui.add_enabled(false, egui::Button::new("None"));
            ui.add_enabled(false, egui::Button::new("Continuous"));
            ui.add_enabled(false, egui::Button::new("Restart Each Page"));
        });
    });
}

fn ribbon_layout_header_footer_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    history: &mut ChangeHistory,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Header & Footer", palette, |ui| {
        ui.menu_button("Header ▾", |ui| {
            if ui.button("Edit Header").clicked() {
                enter_header_footer(document, canvas, HeaderFooterKind::Header, status_message);
                ui.close();
            }
            if ui.button("Blank Header").clicked() {
                set_blank_header_footer(
                    document,
                    canvas,
                    history,
                    HeaderFooterKind::Header,
                    status_message,
                    ui.input(|i| i.time),
                );
                ui.close();
            }
        });
        ui.menu_button("Footer ▾", |ui| {
            if ui.button("Edit Footer").clicked() {
                enter_header_footer(document, canvas, HeaderFooterKind::Footer, status_message);
                ui.close();
            }
            if ui.button("Blank Footer").clicked() {
                set_blank_header_footer(
                    document,
                    canvas,
                    history,
                    HeaderFooterKind::Footer,
                    status_message,
                    ui.input(|i| i.time),
                );
                ui.close();
            }
        });
        page_number_menu_button(ui, document, canvas, history, status_message, "Page # ▾");
    });
}

fn ribbon_advanced_page_setup_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    history: &mut ChangeHistory,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Advanced", palette, |ui| {
        ui.menu_button("Page Setup...", |ui| {
            let section_id = current_section_id(document, canvas);
            let mut setup = document
                .section_by_id(section_id)
                .map(|section| section.page_setup)
                .unwrap_or_else(PageSetup::standard);

            ui.label(
                egui::RichText::new("Margins")
                    .size(11.0)
                    .color(palette.text_muted),
            );
            ui.horizontal(|ui| {
                page_setup_drag(ui, "Top", &mut setup.margins.top_points);
                page_setup_drag(ui, "Bottom", &mut setup.margins.bottom_points);
            });
            ui.horizontal(|ui| {
                page_setup_drag(ui, "Left", &mut setup.margins.left_points);
                page_setup_drag(ui, "Right", &mut setup.margins.right_points);
            });
            ui.separator();
            ui.horizontal(|ui| {
                page_setup_drag(ui, "Width", &mut setup.page_size.width_points);
                page_setup_drag(ui, "Height", &mut setup.page_size.height_points);
            });
            ui.separator();
            let mut page_start = setup.page_number_start.unwrap_or(1) as i32;
            ui.horizontal(|ui| {
                ui.label("Page number start");
                if ui
                    .add(
                        egui::DragValue::new(&mut page_start)
                            .range(0..=9999)
                            .speed(1.0),
                    )
                    .changed()
                {
                    setup.page_number_start = Some(page_start.max(0) as usize);
                }
            });
            if ui.button("Apply").clicked() {
                history.checkpoint(document, ui.input(|i| i.time));
                setup.margins.top_points = setup.margins.top_points.max(0.0);
                setup.margins.right_points = setup.margins.right_points.max(0.0);
                setup.margins.bottom_points = setup.margins.bottom_points.max(0.0);
                setup.margins.left_points = setup.margins.left_points.max(0.0);
                setup.page_size.width_points = setup.page_size.width_points.max(72.0);
                setup.page_size.height_points = setup.page_size.height_points.max(72.0);
                if let Some(section) = document.section_by_id_mut(section_id) {
                    section.page_setup = setup;
                }
                document.sync_compat_from_first_section();
                *status_message = format!("Page setup updated for Section {section_id}");
                ui.close();
            }
        });
    });
}

fn ribbon_header_footer_insert_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    history: &mut ChangeHistory,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Insert", palette, |ui| {
        page_number_menu_button(ui, document, canvas, history, status_message, "Page # ▾");
        if ui.button("Date").clicked() {
            insert_header_footer_text(
                document,
                canvas,
                history,
                status_message,
                &today_label(),
                "Date inserted",
                ui.input(|i| i.time),
            );
        }
        ui.menu_button("Document Info ▾", |ui| {
            if ui.button("Title").clicked() {
                let title = document.title.clone();
                insert_header_footer_text(
                    document,
                    canvas,
                    history,
                    status_message,
                    &title,
                    "Document title inserted",
                    ui.input(|i| i.time),
                );
                ui.close();
            }
        });
    });
}

fn ribbon_header_footer_options_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    history: &mut ChangeHistory,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Options", palette, |ui| {
        let section_id = current_section_id(document, canvas);
        let active = canvas.active_header_footer.unwrap_or(ActiveHeaderFooter {
            kind: HeaderFooterKind::Header,
            section_id,
            variant: HeaderFooterVariant::Default,
            page_number: 1,
        });

        let mut different_first = document
            .section_by_id(section_id)
            .map(|section| section.different_first_page)
            .unwrap_or(false);
        if ui
            .checkbox(&mut different_first, "Different First Page")
            .changed()
        {
            history.checkpoint(document, ui.input(|i| i.time));
            if let Some(section) = document.section_by_id_mut(section_id) {
                section.different_first_page = different_first;
            }
            document.sync_compat_from_first_section();
            *status_message = format!("Different First Page updated for Section {section_id}");
        }

        let mut different_even = document.different_odd_even_pages;
        if ui.checkbox(&mut different_even, "Odd & Even").changed() {
            history.checkpoint(document, ui.input(|i| i.time));
            document.different_odd_even_pages = different_even;
            *status_message = "Odd/even header/footer setting updated".to_owned();
        }

        let mut linked =
            document.header_footer_linked(active.section_id, active.kind, active.variant);
        let link_enabled = document
            .sections
            .iter()
            .position(|section| section.id == active.section_id)
            .unwrap_or(0)
            > 0;
        if ui
            .add_enabled(
                link_enabled,
                egui::Checkbox::new(&mut linked, "Link to Previous"),
            )
            .changed()
        {
            history.checkpoint(document, ui.input(|i| i.time));
            document.set_header_footer_link(active.section_id, active.kind, active.variant, linked);
            *status_message = format!(
                "{} - Section {} {}",
                match active.kind {
                    HeaderFooterKind::Header => "Header",
                    HeaderFooterKind::Footer => "Footer",
                },
                active.section_id,
                if linked {
                    "linked to previous"
                } else {
                    "unlinked from previous"
                }
            );
        }
    });
}

fn ribbon_header_footer_position_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    history: &mut ChangeHistory,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Position", palette, |ui| {
        let section_id = current_section_id(document, canvas);
        ui.label("Header:");
        let mut header_from_top = document
            .section_by_id(section_id)
            .map(|section| section.page_setup.header_from_top_points)
            .unwrap_or(36.0);
        if ui
            .add(
                egui::DragValue::new(&mut header_from_top)
                    .range(0.0..=288.0)
                    .speed(1.0),
            )
            .changed()
        {
            history.checkpoint_coalesced(document, ui.input(|i| i.time));
            if let Some(section) = document.section_by_id_mut(section_id) {
                section.page_setup.header_from_top_points = header_from_top.max(0.0);
            }
            document.sync_compat_from_first_section();
            *status_message = "Header position updated".to_owned();
        }

        ui.label("Footer:");
        let mut footer_from_bottom = document
            .section_by_id(section_id)
            .map(|section| section.page_setup.footer_from_bottom_points)
            .unwrap_or(36.0);
        if ui
            .add(
                egui::DragValue::new(&mut footer_from_bottom)
                    .range(0.0..=288.0)
                    .speed(1.0),
            )
            .changed()
        {
            history.checkpoint_coalesced(document, ui.input(|i| i.time));
            if let Some(section) = document.section_by_id_mut(section_id) {
                section.page_setup.footer_from_bottom_points = footer_from_bottom.max(0.0);
            }
            document.sync_compat_from_first_section();
            *status_message = "Footer position updated".to_owned();
        }
    });
}

fn ribbon_header_footer_actions_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    history: &mut ChangeHistory,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Actions", palette, |ui| {
        let section_id = current_section_id(document, canvas);
        let active = canvas.active_header_footer.unwrap_or(ActiveHeaderFooter {
            kind: HeaderFooterKind::Header,
            section_id,
            variant: HeaderFooterVariant::Default,
            page_number: 1,
        });

        if ui.button("Remove Header").clicked() {
            history.checkpoint(document, ui.input(|i| i.time));
            document.clear_header_footer_slot(
                active.section_id,
                HeaderFooterKind::Header,
                active.variant,
            );
            document.sync_compat_from_first_section();
            *status_message = format!("Header - Section {} cleared", active.section_id);
        }
        if ui.button("Remove Footer").clicked() {
            history.checkpoint(document, ui.input(|i| i.time));
            document.clear_header_footer_slot(
                active.section_id,
                HeaderFooterKind::Footer,
                active.variant,
            );
            document.sync_compat_from_first_section();
            *status_message = format!("Footer - Section {} cleared", active.section_id);
        }
        if ui.button("Close").clicked() {
            canvas.active_header_footer = None;
            *status_message = "Closed header/footer".to_owned();
        }
    });
}

fn page_number_menu_button(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    status_message: &mut String,
    label: &str,
) {
    ui.menu_button(label, |ui| {
        if ui.button("Bottom of Page").clicked() {
            insert_page_number(
                document,
                canvas,
                history,
                status_message,
                HeaderFooterKind::Footer,
                ui.input(|i| i.time),
            );
            ui.close();
        }
        if ui.button("Top of Page").clicked() {
            insert_page_number(
                document,
                canvas,
                history,
                status_message,
                HeaderFooterKind::Header,
                ui.input(|i| i.time),
            );
            ui.close();
        }
        if ui.button("Current Position").clicked() {
            let kind = canvas
                .active_header_footer
                .map(|active| active.kind)
                .unwrap_or(HeaderFooterKind::Footer);
            insert_page_number(
                document,
                canvas,
                history,
                status_message,
                kind,
                ui.input(|i| i.time),
            );
            ui.close();
        }
    });
}

fn insert_page_number(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    status_message: &mut String,
    fallback_kind: HeaderFooterKind,
    now: f64,
) {
    history.checkpoint(document, now);
    let (section_id, variant, kind) = canvas
        .active_header_footer
        .map(|active| (active.section_id, active.variant, active.kind))
        .unwrap_or_else(|| {
            (
                current_section_id(document, canvas),
                HeaderFooterVariant::Default,
                fallback_kind,
            )
        });
    let story = document
        .header_footer_story_mut_materialized(section_id, kind, variant)
        .expect("current section exists");
    let text = story.plain_text();
    if text.trim().is_empty() {
        story.runs = vec![TextRun {
            text: "Page { PAGE } of { NUMPAGES }".to_owned(),
            style: canvas.active_style,
        }];
    } else {
        story.runs.push(TextRun {
            text: " { PAGE }".to_owned(),
            style: canvas.active_style,
        });
    }
    document.sync_compat_from_first_section();
    enter_header_footer_at_end(document, canvas, kind, section_id, variant);
    *status_message = "Page number inserted".to_owned();
}

fn enter_header_footer(
    document: &DocumentState,
    canvas: &mut CanvasState,
    kind: HeaderFooterKind,
    status_message: &mut String,
) {
    let section_id = current_section_id(document, canvas);
    enter_header_footer_at_end(
        document,
        canvas,
        kind,
        section_id,
        HeaderFooterVariant::Default,
    );
    *status_message = match kind {
        HeaderFooterKind::Header => "Editing header",
        HeaderFooterKind::Footer => "Editing footer",
    }
    .to_owned();
}

fn enter_header_footer_at_end(
    document: &DocumentState,
    canvas: &mut CanvasState,
    kind: HeaderFooterKind,
    section_id: crate::document::SectionId,
    variant: HeaderFooterVariant,
) {
    canvas.active_header_footer = Some(ActiveHeaderFooter {
        kind,
        section_id,
        variant,
        page_number: 1,
    });
    canvas.active_header_footer_cursor = document
        .resolve_header_footer_slot(section_id, kind, variant)
        .story
        .plain_text()
        .chars()
        .count();
    canvas.active_header_footer_selection = egui::text_selection::CCursorRange::one(
        egui::epaint::text::cursor::CCursor::new(canvas.active_header_footer_cursor),
    );
}

fn set_blank_header_footer(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    kind: HeaderFooterKind,
    status_message: &mut String,
    now: f64,
) {
    history.checkpoint(document, now);
    let section_id = current_section_id(document, canvas);
    let variant = HeaderFooterVariant::Default;
    let story = document
        .header_footer_story_mut_materialized(section_id, kind, variant)
        .expect("current section exists");
    story.runs = vec![TextRun {
        text: String::new(),
        style: canvas.active_style,
    }];
    document.sync_compat_from_first_section();
    enter_header_footer_at_end(document, canvas, kind, section_id, variant);
    *status_message = match kind {
        HeaderFooterKind::Header => "Blank header inserted",
        HeaderFooterKind::Footer => "Blank footer inserted",
    }
    .to_owned();
}

fn insert_header_footer_text(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    status_message: &mut String,
    text: &str,
    message: &str,
    now: f64,
) {
    let Some(active) = canvas.active_header_footer else {
        return;
    };
    history.checkpoint(document, now);
    let story = document
        .header_footer_story_mut_materialized(active.section_id, active.kind, active.variant)
        .expect("active section exists");
    let mut plain = story.plain_text();
    let cursor = canvas
        .active_header_footer_cursor
        .min(plain.chars().count());
    let byte_index = plain
        .char_indices()
        .nth(cursor)
        .map(|(index, _)| index)
        .unwrap_or(plain.len());
    plain.insert_str(byte_index, text);
    story.runs = vec![TextRun {
        text: plain,
        style: canvas.active_style,
    }];
    document.sync_compat_from_first_section();
    canvas.active_header_footer_cursor = cursor + text.chars().count();
    canvas.active_header_footer_selection = egui::text_selection::CCursorRange::one(
        egui::epaint::text::cursor::CCursor::new(canvas.active_header_footer_cursor),
    );
    *status_message = message.to_owned();
}

fn page_size_presets() -> [(&'static str, PageSize); 3] {
    [
        (
            "A4",
            PageSize {
                width_points: 595.0,
                height_points: 842.0,
            },
        ),
        (
            "Letter",
            PageSize {
                width_points: 612.0,
                height_points: 792.0,
            },
        ),
        (
            "Legal",
            PageSize {
                width_points: 612.0,
                height_points: 1008.0,
            },
        ),
    ]
}

fn set_current_section_margins(
    document: &mut DocumentState,
    canvas: &CanvasState,
    history: &mut ChangeHistory,
    status_message: &mut String,
    margins: PageMargins,
    label: &str,
    now: f64,
) {
    history.checkpoint(document, now);
    let section_id = current_section_id(document, canvas);
    if let Some(section) = document.section_by_id_mut(section_id) {
        section.page_setup.margins = margins;
    }
    document.sync_compat_from_first_section();
    *status_message = format!("{label} applied to Section {section_id}");
}

fn set_current_section_page_size(
    document: &mut DocumentState,
    canvas: &CanvasState,
    history: &mut ChangeHistory,
    status_message: &mut String,
    page_size: PageSize,
    label: &str,
    now: f64,
) {
    history.checkpoint(document, now);
    let section_id = current_section_id(document, canvas);
    if let Some(section) = document.section_by_id_mut(section_id) {
        section.page_setup.page_size = page_size;
    }
    document.sync_compat_from_first_section();
    *status_message = format!("Page size set to {label} for Section {section_id}");
}

fn set_current_section_orientation(
    document: &mut DocumentState,
    canvas: &CanvasState,
    history: &mut ChangeHistory,
    status_message: &mut String,
    portrait: bool,
    now: f64,
) {
    history.checkpoint(document, now);
    let section_id = current_section_id(document, canvas);
    if let Some(section) = document.section_by_id_mut(section_id) {
        let width = section.page_setup.page_size.width_points;
        let height = section.page_setup.page_size.height_points;
        section.page_setup.page_size = if portrait {
            PageSize {
                width_points: width.min(height),
                height_points: width.max(height),
            }
        } else {
            PageSize {
                width_points: width.max(height),
                height_points: width.min(height),
            }
        };
    }
    document.sync_compat_from_first_section();
    *status_message = format!(
        "{} orientation applied to Section {section_id}",
        if portrait { "Portrait" } else { "Landscape" }
    );
}

fn page_setup_drag(ui: &mut egui::Ui, label: &str, value: &mut f32) {
    ui.label(label);
    ui.add(
        egui::DragValue::new(value)
            .range(0.0..=2000.0)
            .speed(1.0)
            .fixed_decimals(0),
    );
}

fn today_label() -> String {
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
fn civil_from_days(days_since_unix_epoch: i64) -> (i32, u32, u32) {
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
        if ui.button("Page Width").clicked() {
            canvas.zoom_mode = ZoomMode::FitPage;
            *status_message = "Page width view".to_owned();
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
    const RIBBON_GROUP_CONTENT_HEIGHT: f32 = 44.0;

    egui::Frame::new()
        .fill(egui::Color32::TRANSPARENT)
        .inner_margin(egui::Margin::symmetric(6, 3))
        .stroke(egui::Stroke::NONE)
        .corner_radius(0.0)
        .show(ui, |ui| {
            ui.set_min_height(RIBBON_GROUP_CONTENT_HEIGHT);
            ui.vertical(|ui| {
                ui.horizontal(|ui| {
                    ui.spacing_mut().item_spacing = egui::vec2(5.0, 3.0);
                    add_contents(ui);
                });
                ui.add_space(2.0);
                ui.label(
                    egui::RichText::new(title)
                        .size(10.0)
                        .color(palette.text_muted),
                );
            });
        });
    ui.separator();
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

#[cfg(test)]
fn layout_tab_command_labels() -> &'static [&'static str] {
    &[
        "Margins",
        "Size",
        "Orientation",
        "Columns",
        "Breaks",
        "Line Numbers",
        "Header",
        "Footer",
        "Page #",
        "Page Setup",
    ]
}

#[cfg(test)]
fn layout_tab_removed_labels() -> &'static [&'static str] {
    &[
        "Zoom",
        "Dark",
        "Header from Top",
        "Footer from Bottom",
        "Remove Header",
        "Remove Footer",
        "Close Header and Footer",
    ]
}

#[cfg(test)]
fn header_footer_contextual_tab_visible(canvas: &CanvasState) -> bool {
    canvas.active_header_footer.is_some()
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn layout_tab_contract_lists_only_page_layout_commands() {
        assert_eq!(
            layout_tab_command_labels(),
            &[
                "Margins",
                "Size",
                "Orientation",
                "Columns",
                "Breaks",
                "Line Numbers",
                "Header",
                "Footer",
                "Page #",
                "Page Setup",
            ]
        );
        assert!(!layout_tab_command_labels()
            .iter()
            .any(|label| layout_tab_removed_labels().contains(label)));
    }

    #[test]
    fn header_footer_tab_visibility_tracks_editing_state() {
        let mut canvas = CanvasState::default();
        assert!(!header_footer_contextual_tab_visible(&canvas));

        canvas.active_header_footer = Some(ActiveHeaderFooter {
            kind: HeaderFooterKind::Header,
            section_id: 1,
            variant: HeaderFooterVariant::Default,
            page_number: 1,
        });
        assert!(header_footer_contextual_tab_visible(&canvas));
    }

    #[cfg(not(target_arch = "wasm32"))]
    #[test]
    fn unix_epoch_formats_as_civil_date() {
        assert_eq!(civil_from_days(0), (1970, 1, 1));
    }
}
