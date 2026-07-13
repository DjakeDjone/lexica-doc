use std::path::PathBuf;

use eframe::egui;

use super::header_layout::RIBBON_HEIGHT;
use super::RibbonTab;
use crate::app::{
    actions::sync_active_style,
    find_replace::FindReplaceState,
    palette::{ThemeMode, ThemePalette},
    CanvasState, ChangeHistory,
};
use crate::canvas::scrollbar_color;
use crate::document::DocumentState;
use crate::grammar::{GrammarConfig, GrammarStatus};

pub(crate) mod common;
pub(crate) mod grammar;
pub(crate) mod header_footer;
pub(crate) mod home;
pub(crate) mod insert;
pub(crate) mod layout;
pub(crate) mod picture;
pub(crate) mod table;

#[derive(Default)]
pub(crate) struct GrammarRibbonOutput {
    pub manual_check_requested: bool,
    pub restart_requested: bool,
    pub download_requested: bool,
    pub settings_changed: bool,
}

#[allow(clippy::too_many_arguments)]
pub(crate) fn paint_ribbon(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    active_tab: &mut RibbonTab,
    status_message: &mut String,
    current_path: &mut Option<PathBuf>,
    theme_mode: &mut ThemeMode,
    history: &mut ChangeHistory,
    find_replace: &mut FindReplaceState,
    grammar_config: &mut GrammarConfig,
    grammar_status: &GrammarStatus,
    grammar_auto_check: &mut bool,
    can_download_grammar: bool,
    #[cfg(not(target_arch = "wasm32"))] dialog_tx: &std::sync::mpsc::Sender<
        crate::app::DialogAction,
    >,
    palette: ThemePalette,
) -> GrammarRibbonOutput {
    sync_active_style(document, canvas);
    let mut output = GrammarRibbonOutput::default();

    egui::Frame::new()
        .inner_margin(egui::Margin::symmetric(8, 4))
        .show(ui, |ui| {
            let content_style = ui.style().clone();
            let color = scrollbar_color(*theme_mode);
            let visuals = ui.visuals_mut();
            visuals.extreme_bg_color = color;
            visuals.widgets.inactive.fg_stroke.color = color;
            visuals.widgets.hovered.fg_stroke.color = color;
            visuals.widgets.active.fg_stroke.color = color;
            ui.spacing_mut().scroll.active_background_opacity = 0.25;
            ui.spacing_mut().scroll.interact_background_opacity = 0.25;
            ui.spacing_mut().scroll.active_handle_opacity = 0.8;
            ui.spacing_mut().scroll.interact_handle_opacity = 0.8;

            egui::ScrollArea::horizontal()
                .id_salt("ribbon_horizontal_scroll")
                .max_height(RIBBON_HEIGHT)
                .auto_shrink([false, false])
                .show(ui, |ui| {
                    ui.set_style(content_style.clone());
                    ui.set_min_height(RIBBON_HEIGHT);
                    ui.horizontal(|ui| {
                        match active_tab {
                            RibbonTab::Home => {
                                #[cfg(not(target_arch = "wasm32"))]
                                common::ribbon_file_group(ui, document, canvas, status_message, current_path, history, dialog_tx, palette);
                                #[cfg(target_arch = "wasm32")]
                                common::ribbon_file_group(ui, document, canvas, status_message, current_path, history, palette);
                                home::ribbon_font_group(ui, document, canvas, history, palette);
                                home::ribbon_paragraph_group(ui, document, canvas, history, palette);
                                home::ribbon_color_group(ui, document, canvas, history, palette);
                                home::ribbon_editing_group(ui, find_replace, palette);
                                home::ribbon_view_group(ui, canvas, status_message, theme_mode, palette);
                            }
                            RibbonTab::Insert => {
                                #[cfg(not(target_arch = "wasm32"))]
                                {
                                    common::ribbon_file_group(ui, document, canvas, status_message, current_path, history, dialog_tx, palette);
                                    insert::ribbon_insert_group(ui, document, canvas, status_message, history, dialog_tx, palette);
                                }
                                #[cfg(target_arch = "wasm32")]
                                {
                                    common::ribbon_file_group(ui, document, canvas, status_message, current_path, history, palette);
                                    insert::ribbon_insert_group(ui, document, canvas, status_message, history, palette);
                                }
                                common::ribbon_info_group(
                                    ui,
                                    "Insert",
                                    "Import supports .txt, .md, .markdown, .docx, and .odt with images.",
                                    palette,
                                );
                            }
                            RibbonTab::Design => {
                                home::ribbon_font_group(ui, document, canvas, history, palette);
                                home::ribbon_paragraph_group(ui, document, canvas, history, palette);
                                home::ribbon_color_group(ui, document, canvas, history, palette);
                            }
                            RibbonTab::Layout => {
                                layout::ribbon_page_setup_group(ui, document, canvas, status_message, history, palette);
                                layout::ribbon_flow_group(ui, document, canvas, status_message, history, palette);
                                layout::ribbon_layout_header_footer_group(
                                    ui,
                                    document,
                                    canvas,
                                    status_message,
                                    history,
                                    palette,
                                );
                                layout::ribbon_advanced_page_setup_group(
                                    ui,
                                    document,
                                    canvas,
                                    status_message,
                                    history,
                                    palette,
                                );
                            }
                            RibbonTab::View => {
                                home::ribbon_view_group(ui, canvas, status_message, theme_mode, palette);
                                common::ribbon_info_group(
                                    ui,
                                    "Shortcuts",
                                    "Command+S Save, Command+Shift+S Save As, Ctrl+Z Undo, Ctrl+Shift+Z / Ctrl+Y Redo, Command+B Bold, Command+I Italic, Command+U Underline",
                                    palette,
                                );
                            }
                            RibbonTab::Grammar => {
                                grammar::ribbon_grammer_actions_group(
                                    ui,
                                    grammar_status,
                                    can_download_grammar,
                                    &mut output,
                                    palette,
                                );
                                grammar::ribbon_grammer_settings_group(
                                    ui,
                                    grammar_config,
                                    grammar_auto_check,
                                    &mut output,
                                    palette,
                                );
                                common::ribbon_info_group(
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
                                picture::ribbon_picture_group(ui, document, canvas, status_message, history, palette);
                            }
                            RibbonTab::Table => {
                                home::ribbon_font_group(ui, document, canvas, history, palette);
                                home::ribbon_color_group(ui, document, canvas, history, palette);
                                #[cfg(not(target_arch = "wasm32"))]
                                insert::ribbon_insert_group(ui, document, canvas, status_message, history, dialog_tx, palette);
                                #[cfg(target_arch = "wasm32")]
                                insert::ribbon_insert_group(ui, document, canvas, status_message, history, palette);
                                table::table_format_group(ui, document, canvas, status_message, history, palette);
                            }
                            RibbonTab::HeaderFooter => {
                                header_footer::ribbon_header_footer_insert_group(
                                    ui,
                                    document,
                                    canvas,
                                    status_message,
                                    history,
                                    palette,
                                );
                                header_footer::ribbon_header_footer_options_group(
                                    ui,
                                    document,
                                    canvas,
                                    status_message,
                                    history,
                                    palette,
                                );
                                header_footer::ribbon_header_footer_position_group(
                                    ui,
                                    document,
                                    canvas,
                                    status_message,
                                    history,
                                    palette,
                                );
                                header_footer::ribbon_header_footer_actions_group(
                                    ui,
                                    document,
                                    canvas,
                                    status_message,
                                    history,
                                    palette,
                                );
                            }
                        }
                    });
                });
            ui.set_style(content_style);
        });
    output
}
