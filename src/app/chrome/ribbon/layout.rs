use eframe::egui;

use crate::document::{DocumentState, HeaderFooterKind, PageMargins, PageSetup, PageSize};
use crate::app::{CanvasState, ChangeHistory, palette::ThemePalette};
use super::common::ribbon_group;
use super::header_footer::{enter_header_footer, page_number_menu_button, set_blank_header_footer};

pub(crate) fn ribbon_page_setup_group(
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

pub(crate) fn ribbon_flow_group(
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
                crate::app::actions::insert_page_break(document, canvas, status_message, history);
                ui.close();
            }
            if ui.button("Section").clicked() {
                crate::app::actions::insert_section_break(document, canvas, status_message, history);
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

pub(crate) fn ribbon_layout_header_footer_group(
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

pub(crate) fn ribbon_advanced_page_setup_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    history: &mut ChangeHistory,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Advanced", palette, |ui| {
        ui.menu_button("Page Setup...", |ui| {
            let section_id = super::super::current_section_id(document, canvas);
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
    let section_id = super::super::current_section_id(document, canvas);
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
    let section_id = super::super::current_section_id(document, canvas);
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
    let section_id = super::super::current_section_id(document, canvas);
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
