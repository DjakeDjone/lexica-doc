use std::path::PathBuf;

use eframe::egui;

use super::header_layout::{clipped_child_ui, title_action_width, TITLE_BAR_HEIGHT};
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
        .and_then(|path| path.file_name())
        .map(|name| name.to_string_lossy().into_owned())
        .unwrap_or_else(|| "Not saved yet".to_owned());
    let title_label = format!("{} — wors", document.title);

    // Render the title bar content first so buttons register their interactions
    // before the drag overlay.
    let _frame_response = egui::Frame::new()
        .inner_margin(egui::Margin::symmetric(12, 8))
        .show(ui, |ui| {
            let (row_rect, _) = ui.allocate_exact_size(
                egui::vec2(ui.available_width(), TITLE_BAR_HEIGHT),
                egui::Sense::hover(),
            );
            let action_width = title_action_width(row_rect.width());
            let action_rect = egui::Rect::from_min_max(
                egui::pos2(row_rect.right() - action_width, row_rect.top()),
                row_rect.right_bottom(),
            );
            let identity_rect = egui::Rect::from_min_max(
                row_rect.left_top(),
                egui::pos2(
                    (action_rect.left() - 12.0).max(row_rect.left()),
                    row_rect.bottom(),
                ),
            );

            let mut identity_ui = clipped_child_ui(
                ui,
                "title_identity",
                identity_rect,
                egui::Layout::left_to_right(egui::Align::Center),
            );
            paint_title_identity(
                &mut identity_ui,
                document,
                canvas,
                history,
                status_target,
                palette,
                logo,
                &title_label,
                &path_label,
            );

            if action_width > 0.0 {
                let mut actions_ui = clipped_child_ui(
                    ui,
                    "title_actions",
                    action_rect,
                    egui::Layout::right_to_left(egui::Align::Center),
                );
                paint_title_actions(
                    &mut actions_ui,
                    status_message,
                    theme_mode,
                    status_target,
                    palette,
                );
            }
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

#[allow(clippy::too_many_arguments)]
fn paint_title_identity(
    ui: &mut egui::Ui,
    document: &mut crate::document::DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    status_target: &mut String,
    palette: ThemePalette,
    logo: &egui::TextureHandle,
    title_label: &str,
    path_label: &str,
) {
    ui.spacing_mut().item_spacing = egui::vec2(6.0, 0.0);
    let full_width = ui.max_rect().width().max(0.0);
    if full_width >= 30.0 {
        ui.add(
            egui::Image::new(egui::load::SizedTexture::new(
                logo.id(),
                egui::vec2(24.0, 24.0),
            ))
            .sense(egui::Sense::hover()),
        );
    }

    let can_show_history = full_width >= 120.0;
    let history_width = if can_show_history { 56.0 } else { 0.0 };
    let text_width = (ui.available_width() - history_width).max(0.0);
    if text_width >= 54.0 {
        let title_width = compact_label_width(title_label, 8.0, 250.0).min(text_width);
        ui.add_sized(
            egui::vec2(title_width, TITLE_BAR_HEIGHT),
            egui::Label::new(
                egui::RichText::new(title_label)
                    .size(14.0)
                    .color(palette.title_fg),
            )
            .halign(egui::Align::Min)
            .truncate(),
        );

        let path_budget = (text_width - title_width - 6.0).max(0.0);
        let path_width = compact_label_width(path_label, 6.2, 260.0).min(path_budget);
        if path_width >= 42.0 {
            ui.add_sized(
                egui::vec2(path_width, TITLE_BAR_HEIGHT),
                egui::Label::new(
                    egui::RichText::new(path_label)
                        .size(11.0)
                        .color(palette.title_muted),
                )
                .halign(egui::Align::Min)
                .truncate(),
            );
        }
    }

    if can_show_history {
        // Undo / Redo buttons stay with the document identity block and are
        // dropped first on very narrow windows.
        let can_undo = history.can_undo();
        let can_redo = history.can_redo();
        let undo_btn = egui::Button::new(egui::RichText::new("↩").size(14.0).color(if can_undo {
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
        let redo_btn = egui::Button::new(egui::RichText::new("↪").size(14.0).color(if can_redo {
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
    }
}

fn compact_label_width(text: &str, px_per_char: f32, max_width: f32) -> f32 {
    ((text.chars().count() as f32 * px_per_char) + 12.0).clamp(42.0, max_width)
}

fn paint_title_actions(
    ui: &mut egui::Ui,
    status_message: &str,
    theme_mode: &mut ThemeMode,
    status_target: &mut String,
    palette: ThemePalette,
) {
    ui.spacing_mut().item_spacing = egui::vec2(6.0, 0.0);
    let full_width = ui.max_rect().width();
    #[cfg(not(target_arch = "wasm32"))]
    {
        paint_window_buttons(ui, palette);
        if full_width >= 128.0 {
            ui.separator();
        }
    }

    let has_room_for_theme = ui.available_width() >= 58.0;
    if has_room_for_theme && theme_switch(ui, theme_mode, palette, true) {
        *status_target = format!("Theme switched to {}", theme_mode.label());
    }

    let status_budget = ui.available_width().max(0.0);
    if status_budget >= 56.0 {
        if has_room_for_theme {
            ui.separator();
        }
        let status_width = compact_label_width(status_message, 6.2, 120.0).min(status_budget);
        ui.add_sized(
            egui::vec2(status_width, TITLE_BAR_HEIGHT),
            egui::Label::new(
                egui::RichText::new(status_message)
                    .size(11.0)
                    .color(palette.title_muted),
            )
            .halign(egui::Align::Max)
            .truncate(),
        );
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn paint_window_buttons(ui: &mut egui::Ui, palette: ThemePalette) {
    if ui.available_width() < 24.0 {
        return;
    }

    let close_btn = egui::Button::new(egui::RichText::new("🗙").size(14.0).color(palette.title_fg))
        .min_size(egui::vec2(24.0, 24.0))
        .fill(egui::Color32::TRANSPARENT)
        .stroke(egui::Stroke::NONE);
    if ui.add(close_btn).on_hover_text("Close").clicked() {
        ui.ctx().send_viewport_cmd(egui::ViewportCommand::Close);
    }
    if ui.available_width() < 24.0 {
        return;
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
    if ui.available_width() < 24.0 {
        return;
    }

    let min_btn = egui::Button::new(egui::RichText::new("🗕").size(14.0).color(palette.title_fg))
        .min_size(egui::vec2(24.0, 24.0))
        .fill(egui::Color32::TRANSPARENT)
        .stroke(egui::Stroke::NONE);
    if ui.add(min_btn).on_hover_text("Minimize").clicked() {
        ui.ctx()
            .send_viewport_cmd(egui::ViewportCommand::Minimized(true));
    }
}
