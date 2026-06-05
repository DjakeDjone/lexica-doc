use std::path::PathBuf;

use eframe::egui;

use crate::app::{
    palette::{theme_switch, ThemeMode, ThemePalette},
    CanvasState, ChangeHistory,
};

#[allow(clippy::too_many_arguments)]
pub(crate) fn paint_title_bar(
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
