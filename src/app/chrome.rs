use std::path::PathBuf;

use eframe::egui;

use crate::document::{
    DocumentState, FontChoice, ListKind, ParagraphAlignment, OBJECT_REPLACEMENT_CHAR,
};

use super::{
    actions::{
        insert_image, insert_page_break, open_document, save_document, set_font_choice,
        set_font_size, set_highlight_color, set_paragraph_alignment, set_text_color,
        sync_active_style, toggle_bold, toggle_bullet_list, toggle_italic, toggle_ordered_list,
        toggle_strikethrough, toggle_underline,
    },
    palette::{theme_switch, ThemeMode, ThemePalette},
    CanvasState,
};

#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum RibbonTab {
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

pub(super) fn paint_title_bar(
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

pub(super) fn paint_tab_row(ui: &mut egui::Ui, active_tab: &mut RibbonTab, palette: ThemePalette) {
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

pub(super) fn paint_ribbon(
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
                    ribbon_paragraph_group(ui, document, canvas, palette);
                    ribbon_color_group(ui, document, canvas, palette);
                    ribbon_view_group(ui, canvas, status_message, theme_mode, palette);
                }
                RibbonTab::Insert => {
                    ribbon_file_group(ui, document, canvas, status_message, current_path, palette);
                    ribbon_insert_group(ui, document, canvas, status_message, palette);
                    ribbon_info_group(
                        ui,
                        "Insert",
                        "Import supports .txt, .md, .markdown, and .docx with images.",
                        palette,
                    );
                }
                RibbonTab::Design => {
                    ribbon_font_group(ui, document, canvas, palette);
                    ribbon_paragraph_group(ui, document, canvas, palette);
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

pub(super) fn paint_status_bar(
    ui: &mut egui::Ui,
    document: &DocumentState,
    canvas: &CanvasState,
    status_message: &str,
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
    palette: ThemePalette,
) {
    ribbon_group(ui, "Clipboard", palette, |ui| {
        if ui.button("📂 Open").clicked() {
            open_document(document, canvas, status_message, current_path);
        }
        if ui.button("💾 Save").clicked() {
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

fn ribbon_insert_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Insert", palette, |ui| {
        if ui.button("Image").clicked() {
            insert_image(document, canvas, status_message);
        }
        if ui.button("Page Break").clicked() {
            insert_page_break(document, canvas, status_message);
        }
    });
}

fn ribbon_paragraph_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
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
                set_paragraph_alignment(document, canvas, alignment);
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
            toggle_bullet_list(document, canvas);
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
            toggle_ordered_list(document, canvas);
        }
    });
}

fn ribbon_color_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Colors", palette, |ui| {
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
                canvas.zoom = (zoom_percent / 100.0).clamp(0.5, 3.0);
            }
        });
        if ui.button("↺").clicked() {
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
