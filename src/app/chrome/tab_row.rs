use eframe::egui;

use super::header_layout::TAB_ROW_HEIGHT;
use super::RibbonTab;
use crate::app::palette::ThemePalette;

pub(crate) fn paint_tab_row(
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
            egui::ScrollArea::horizontal()
                .id_salt("tab_row_horizontal_scroll")
                .max_height(TAB_ROW_HEIGHT)
                .auto_shrink([false, false])
                .show(ui, |ui| {
                    ui.set_min_height(TAB_ROW_HEIGHT);
                    ui.horizontal(|ui| {
                        ui.spacing_mut().item_spacing = egui::vec2(2.0, 0.0);
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
                            let button = egui::Button::new(
                                egui::RichText::new(tab.label())
                                    .size(13.0)
                                    .color(if selected {
                                        palette.tab_active_fg
                                    } else {
                                        palette.tab_fg
                                    }),
                            )
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
        });
    file_requested
}
