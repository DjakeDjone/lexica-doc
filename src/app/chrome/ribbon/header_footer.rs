use eframe::egui;

use crate::document::{DocumentState, HeaderFooterKind, HeaderFooterVariant, TextRun};
use crate::app::{ActiveHeaderFooter, CanvasState, ChangeHistory, palette::ThemePalette};
use super::common::ribbon_group;

pub(crate) fn ribbon_header_footer_insert_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    history: &mut ChangeHistory,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Insert", palette, |ui| {
        page_number_menu_button(ui, document, canvas, history, status_message, "Page # ▾");
        if ui.button("Date").clicked() {
            insert_header_footer_text(
                document,
                canvas,
                history,
                status_message,
                &crate::app::chrome::status_bar::today_label(),
                "Date inserted",
                ui.input(|i| i.time),
            );
        }
        ui.menu_button("Document Info ▾", |ui| {
            if ui.button("Title").clicked() {
                let title = document.title.clone();
                insert_header_footer_text(
                    document,
                    canvas,
                    history,
                    status_message,
                    &title,
                    "Document title inserted",
                    ui.input(|i| i.time),
                );
                ui.close();
            }
        });
    });
}

pub(crate) fn ribbon_header_footer_options_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    history: &mut ChangeHistory,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Options", palette, |ui| {
        let section_id = super::super::current_section_id(document, canvas);
        let active = canvas.active_header_footer.unwrap_or(ActiveHeaderFooter {
            kind: HeaderFooterKind::Header,
            section_id,
            variant: HeaderFooterVariant::Default,
            page_number: 1,
        });

        let mut different_first = document
            .section_by_id(section_id)
            .map(|section| section.different_first_page)
            .unwrap_or(false);
        if ui
            .checkbox(&mut different_first, "Different First Page")
            .changed()
        {
            history.checkpoint(document, ui.input(|i| i.time));
            if let Some(section) = document.section_by_id_mut(section_id) {
                section.different_first_page = different_first;
            }
            document.sync_compat_from_first_section();
            *status_message = format!("Different First Page updated for Section {section_id}");
        }

        let mut different_even = document.different_odd_even_pages;
        if ui.checkbox(&mut different_even, "Odd & Even").changed() {
            history.checkpoint(document, ui.input(|i| i.time));
            document.different_odd_even_pages = different_even;
            *status_message = "Odd/even header/footer setting updated".to_owned();
        }

        let mut linked =
            document.header_footer_linked(active.section_id, active.kind, active.variant);
        let link_enabled = document
            .sections
            .iter()
            .position(|section| section.id == active.section_id)
            .unwrap_or(0)
            > 0;
        if ui
            .add_enabled(
                link_enabled,
                egui::Checkbox::new(&mut linked, "Link to Previous"),
            )
            .changed()
        {
            history.checkpoint(document, ui.input(|i| i.time));
            document.set_header_footer_link(active.section_id, active.kind, active.variant, linked);
            *status_message = format!(
                "{} - Section {} {}",
                match active.kind {
                    HeaderFooterKind::Header => "Header",
                    HeaderFooterKind::Footer => "Footer",
                },
                active.section_id,
                if linked {
                    "linked to previous"
                } else {
                    "unlinked from previous"
                }
            );
        }
    });
}

pub(crate) fn ribbon_header_footer_position_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    history: &mut ChangeHistory,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Position", palette, |ui| {
        let section_id = super::super::current_section_id(document, canvas);
        ui.label("Header:");
        let mut header_from_top = document
            .section_by_id(section_id)
            .map(|section| section.page_setup.header_from_top_points)
            .unwrap_or(36.0);
        if ui
            .add(
                egui::DragValue::new(&mut header_from_top)
                    .range(0.0..=288.0)
                    .speed(1.0),
            )
            .changed()
        {
            history.checkpoint_coalesced(document, ui.input(|i| i.time));
            if let Some(section) = document.section_by_id_mut(section_id) {
                section.page_setup.header_from_top_points = header_from_top.max(0.0);
            }
            document.sync_compat_from_first_section();
            *status_message = "Header position updated".to_owned();
        }

        ui.label("Footer:");
        let mut footer_from_bottom = document
            .section_by_id(section_id)
            .map(|section| section.page_setup.footer_from_bottom_points)
            .unwrap_or(36.0);
        if ui
            .add(
                egui::DragValue::new(&mut footer_from_bottom)
                    .range(0.0..=288.0)
                    .speed(1.0),
            )
            .changed()
        {
            history.checkpoint_coalesced(document, ui.input(|i| i.time));
            if let Some(section) = document.section_by_id_mut(section_id) {
                section.page_setup.footer_from_bottom_points = footer_from_bottom.max(0.0);
            }
            document.sync_compat_from_first_section();
            *status_message = "Footer position updated".to_owned();
        }
    });
}

pub(crate) fn ribbon_header_footer_actions_group(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    status_message: &mut String,
    history: &mut ChangeHistory,
    palette: ThemePalette,
) {
    ribbon_group(ui, "Actions", palette, |ui| {
        let section_id = super::super::current_section_id(document, canvas);
        let active = canvas.active_header_footer.unwrap_or(ActiveHeaderFooter {
            kind: HeaderFooterKind::Header,
            section_id,
            variant: HeaderFooterVariant::Default,
            page_number: 1,
        });

        if ui.button("Remove Header").clicked() {
            history.checkpoint(document, ui.input(|i| i.time));
            document.clear_header_footer_slot(
                active.section_id,
                HeaderFooterKind::Header,
                active.variant,
            );
            document.sync_compat_from_first_section();
            *status_message = format!("Header - Section {} cleared", active.section_id);
        }
        if ui.button("Remove Footer").clicked() {
            history.checkpoint(document, ui.input(|i| i.time));
            document.clear_header_footer_slot(
                active.section_id,
                HeaderFooterKind::Footer,
                active.variant,
            );
            document.sync_compat_from_first_section();
            *status_message = format!("Footer - Section {} cleared", active.section_id);
        }
        if ui.button("Close").clicked() {
            canvas.active_header_footer = None;
            *status_message = "Closed header/footer".to_owned();
        }
    });
}

pub(crate) fn page_number_menu_button(
    ui: &mut egui::Ui,
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    status_message: &mut String,
    label: &str,
) {
    ui.menu_button(label, |ui| {
        if ui.button("Bottom of Page").clicked() {
            insert_page_number(
                document,
                canvas,
                history,
                status_message,
                HeaderFooterKind::Footer,
                ui.input(|i| i.time),
            );
            ui.close();
        }
        if ui.button("Top of Page").clicked() {
            insert_page_number(
                document,
                canvas,
                history,
                status_message,
                HeaderFooterKind::Header,
                ui.input(|i| i.time),
            );
            ui.close();
        }
        if ui.button("Current Position").clicked() {
            let kind = canvas
                .active_header_footer
                .map(|active| active.kind)
                .unwrap_or(HeaderFooterKind::Footer);
            insert_page_number(
                document,
                canvas,
                history,
                status_message,
                kind,
                ui.input(|i| i.time),
            );
            ui.close();
        }
    });
}

fn insert_page_number(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    status_message: &mut String,
    fallback_kind: HeaderFooterKind,
    now: f64,
) {
    history.checkpoint(document, now);
    let (section_id, variant, kind) = canvas
        .active_header_footer
        .map(|active| (active.section_id, active.variant, active.kind))
        .unwrap_or_else(|| {
            (
                super::super::current_section_id(document, canvas),
                HeaderFooterVariant::Default,
                fallback_kind,
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
    enter_header_footer_at_end(document, canvas, kind, section_id, variant);
    *status_message = "Page number inserted".to_owned();
}

pub(crate) fn enter_header_footer(
    document: &DocumentState,
    canvas: &mut CanvasState,
    kind: HeaderFooterKind,
    status_message: &mut String,
) {
    let section_id = super::super::current_section_id(document, canvas);
    enter_header_footer_at_end(
        document,
        canvas,
        kind,
        section_id,
        HeaderFooterVariant::Default,
    );
    *status_message = match kind {
        HeaderFooterKind::Header => "Editing header",
        HeaderFooterKind::Footer => "Editing footer",
    }
    .to_owned();
}

fn enter_header_footer_at_end(
    document: &DocumentState,
    canvas: &mut CanvasState,
    kind: HeaderFooterKind,
    section_id: crate::document::SectionId,
    variant: HeaderFooterVariant,
) {
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
}

pub(crate) fn set_blank_header_footer(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    kind: HeaderFooterKind,
    status_message: &mut String,
    now: f64,
) {
    history.checkpoint(document, now);
    let section_id = super::super::current_section_id(document, canvas);
    let variant = HeaderFooterVariant::Default;
    let story = document
        .header_footer_story_mut_materialized(section_id, kind, variant)
        .expect("current section exists");
    story.runs = vec![TextRun {
        text: String::new(),
        style: canvas.active_style,
    }];
    document.sync_compat_from_first_section();
    enter_header_footer_at_end(document, canvas, kind, section_id, variant);
    *status_message = match kind {
        HeaderFooterKind::Header => "Blank header inserted",
        HeaderFooterKind::Footer => "Blank footer inserted",
    }
    .to_owned();
}

fn insert_header_footer_text(
    document: &mut DocumentState,
    canvas: &mut CanvasState,
    history: &mut ChangeHistory,
    status_message: &mut String,
    text: &str,
    message: &str,
    now: f64,
) {
    let Some(active) = canvas.active_header_footer else {
        return;
    };
    history.checkpoint(document, now);
    let story = document
        .header_footer_story_mut_materialized(active.section_id, active.kind, active.variant)
        .expect("active section exists");
    let mut plain = story.plain_text();
    let cursor = canvas
        .active_header_footer_cursor
        .min(plain.chars().count());
    let byte_index = plain
        .char_indices()
        .nth(cursor)
        .map(|(index, _)| index)
        .unwrap_or(plain.len());
    plain.insert_str(byte_index, text);
    story.runs = vec![TextRun {
        text: plain,
        style: canvas.active_style,
    }];
    document.sync_compat_from_first_section();
    canvas.active_header_footer_cursor = cursor + text.chars().count();
    canvas.active_header_footer_selection = egui::text_selection::CCursorRange::one(
        egui::epaint::text::cursor::CCursor::new(canvas.active_header_footer_cursor),
    );
    *status_message = message.to_owned();
}
