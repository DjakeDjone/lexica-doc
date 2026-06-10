use std::path::PathBuf;

use eframe::egui;

use crate::document::{DocumentState, ParagraphAlignment};
use crate::app::{
    actions::{open_document, save_document, save_document_as},
    CanvasState, ChangeHistory, palette::ThemePalette,
};

pub(crate) fn ribbon_file_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    current_path: &mut Option<PathBuf>,
    history: &mut ChangeHistory,
    #[cfg(not(target_arch = "wasm32"))]
    dialog_tx: &std::sync::mpsc::Sender<crate::app::DialogAction>,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Clipboard", palette, |ui| {
        if ui.button("📂 Open").clicked() {
            #[cfg(not(target_arch = "wasm32"))]
            let _ = open_document(document, canvas, status_message, current_path, history, dialog_tx);
            #[cfg(target_arch = "wasm32")]
            let _ = open_document(document, canvas, status_message, current_path, history);
        }
        if ui.button("💾 Save").clicked() {
            #[cfg(not(target_arch = "wasm32"))]
            let _ = save_document(document, status_message, current_path, dialog_tx);
            #[cfg(target_arch = "wasm32")]
            let _ = save_document(document, status_message, current_path);
        }
        if ui.button("Save As").clicked() {
            #[cfg(not(target_arch = "wasm32"))]
            let _ = save_document_as(document, status_message, current_path, dialog_tx);
            #[cfg(target_arch = "wasm32")]
            let _ = save_document_as(document, status_message, current_path);
        }
    });
}

pub(crate) fn ribbon_info_group(ui: &mut egui::Ui, title: &str, message: &str, palette: ThemePalette) {
    ribbon_group(ui, title, palette, |ui| {
        ui.label(
            egui::RichText::new(message)
                .size(11.0)
                .color(palette.text_muted),
        );
    });
}

pub(crate) fn ribbon_group(
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

pub(crate) fn format_button(
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

pub(crate) fn alignment_button(
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
