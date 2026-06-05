use eframe::egui;

use crate::document::{DocumentState, ImageLayoutMode, ImageRendering, WrapMode};
use crate::app::{
    actions::{reset_image_size, set_image_opacity, set_image_rendering, set_image_wrap_mode},
    CanvasState, ChangeHistory, palette::ThemePalette,
};
use super::common::{format_button, ribbon_group, ribbon_info_group};

pub(crate) fn ribbon_picture_group(
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
