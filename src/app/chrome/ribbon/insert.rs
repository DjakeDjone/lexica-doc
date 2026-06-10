use eframe::egui;

use super::common::ribbon_group;
use crate::app::{
    actions::{insert_image, insert_page_break, insert_section_break, insert_table},
    palette::ThemePalette,
    ActiveHeaderFooter, CanvasState, ChangeHistory,
};
use crate::document::{DocumentState, HeaderFooterKind, HeaderFooterVariant, TextRun};

pub(crate) fn ribbon_insert_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    history: &mut ChangeHistory,
    #[cfg(not(target_arch = "wasm32"))] dialog_tx: &std::sync::mpsc::Sender<
        crate::app::DialogAction,
    >,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Insert", palette, |ui| {
        if ui.button("Image").clicked() {
            #[cfg(not(target_arch = "wasm32"))]
            insert_image(document, canvas, status_message, history, dialog_tx);
            #[cfg(target_arch = "wasm32")]
            insert_image(document, canvas, status_message, history);
        }
        if ui.button("Page Break").clicked() {
            insert_page_break(document, canvas, status_message, history);
        }
        if ui.button("Section Break").clicked() {
            insert_section_break(document, canvas, status_message, history);
        }
        if ui.button("Header").clicked() {
            let section_id = super::super::current_section_id(document, canvas);
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
            let section_id = super::super::current_section_id(document, canvas);
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
                        super::super::current_section_id(document, canvas),
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
