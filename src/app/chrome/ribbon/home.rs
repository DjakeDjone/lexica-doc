use eframe::egui;

use super::common::{alignment_button, format_button, ribbon_group};
use crate::app::{
    actions::{
        set_font_choice, set_font_size, set_highlight_color, set_paragraph_alignment,
        set_paragraph_indents, set_text_color, toggle_bold, toggle_bullet_list, toggle_italic,
        toggle_ordered_list, toggle_strikethrough, toggle_subscript, toggle_superscript,
        toggle_underline,
    },
    find_replace::FindReplaceState,
    palette::{theme_switch, ThemeMode, ThemePalette},
    CanvasState, ChangeHistory,
};
use crate::document::{DocumentState, FontChoice, ListKind, ParagraphAlignment, VerticalAlign};

pub(crate) fn ribbon_font_group(
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
        if format_button(
            ui,
            canvas.active_style.vertical_align == VerticalAlign::Superscript,
            "X^2",
            palette,
        )
        .on_hover_text("Superscript")
        .clicked()
        {
            toggle_superscript(document, canvas, history);
        }
        if format_button(
            ui,
            canvas.active_style.vertical_align == VerticalAlign::Subscript,
            "X_2",
            palette,
        )
        .on_hover_text("Subscript")
        .clicked()
        {
            toggle_subscript(document, canvas, history);
        }
    });
}

pub(crate) fn ribbon_paragraph_group(
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

        ui.separator();
        ui.menu_button("Indent ▾", |ui| {
            let style = canvas.active_paragraph_style;
            let mut left = style.left_indent_points;
            let mut right = style.right_indent_points;

            ui.horizontal(|ui| {
                ui.label("Left");
                if ui
                    .add(
                        egui::DragValue::new(&mut left)
                            .range(-720.0..=720.0)
                            .speed(1.0)
                            .fixed_decimals(1)
                            .suffix(" pt"),
                    )
                    .changed()
                {
                    set_paragraph_indents(
                        document,
                        canvas,
                        left,
                        right,
                        style.first_line_indent_points,
                        history,
                        ui.input(|input| input.time),
                    );
                }
            });
            ui.horizontal(|ui| {
                ui.label("Right");
                if ui
                    .add(
                        egui::DragValue::new(&mut right)
                            .range(-720.0..=720.0)
                            .speed(1.0)
                            .fixed_decimals(1)
                            .suffix(" pt"),
                    )
                    .changed()
                {
                    set_paragraph_indents(
                        document,
                        canvas,
                        left,
                        right,
                        style.first_line_indent_points,
                        history,
                        ui.input(|input| input.time),
                    );
                }
            });

            let mut special = if style.first_line_indent_points < 0.0 {
                -1
            } else if style.first_line_indent_points > 0.0 {
                1
            } else {
                0
            };
            let previous_special = special;
            egui::ComboBox::from_id_salt("paragraph_indent_special")
                .selected_text(match special {
                    -1 => "Hanging",
                    1 => "First line",
                    _ => "None",
                })
                .show_ui(ui, |ui| {
                    ui.selectable_value(&mut special, 0, "None");
                    ui.selectable_value(&mut special, 1, "First line");
                    ui.selectable_value(&mut special, -1, "Hanging");
                });

            let mut magnitude = style.first_line_indent_points.abs();
            if special != previous_special {
                if special != 0 && magnitude == 0.0 {
                    magnitude = 36.0;
                }
                set_paragraph_indents(
                    document,
                    canvas,
                    left,
                    right,
                    magnitude * special as f32,
                    history,
                    ui.input(|input| input.time),
                );
            }

            ui.horizontal(|ui| {
                ui.label("By");
                if ui
                    .add_enabled(
                        special != 0,
                        egui::DragValue::new(&mut magnitude)
                            .range(0.0..=720.0)
                            .speed(1.0)
                            .fixed_decimals(1)
                            .suffix(" pt"),
                    )
                    .changed()
                {
                    set_paragraph_indents(
                        document,
                        canvas,
                        left,
                        right,
                        magnitude * special as f32,
                        history,
                        ui.input(|input| input.time),
                    );
                }
            });
        });
    });
}

pub(crate) fn ribbon_color_group(
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

pub(crate) fn ribbon_editing_group(
    ui: &mut egui::Ui,
    find_replace: &mut FindReplaceState,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Editing", palette, |ui| {
        if ui
            .button("Find")
            .on_hover_text("Find and replace document body text")
            .clicked()
        {
            find_replace.visible = true;
        }
    });
}

pub(crate) fn ribbon_view_group(
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
                canvas.zoom_mode = crate::app::ZoomMode::Manual;
                canvas.zoom = (zoom_percent / 100.0).clamp(0.5, 3.0);
            }
        });
        if ui.button("↺").clicked() {
            canvas.zoom_mode = if canvas.imported_docx_view {
                crate::app::ZoomMode::FitPage
            } else {
                crate::app::ZoomMode::Manual
            };
            canvas.zoom = 1.0;
            canvas.pan = egui::Vec2::ZERO;
            *status_message = "View reset".to_owned();
        }
        if ui.button("Page Width").clicked() {
            canvas.zoom_mode = crate::app::ZoomMode::FitPage;
            *status_message = "Page width view".to_owned();
        }
        ui.separator();
        if theme_switch(ui, theme_mode, palette, false) {
            *status_message = format!("Theme switched to {}", theme_mode.label());
        }
    });
}
