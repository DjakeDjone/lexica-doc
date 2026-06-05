use std::ops::Range;
use std::sync::Arc;
use eframe::egui::{
    self, Color32, CornerRadius, Rect,
};

use crate::app::{ActiveHeaderFooter, CanvasState};
use crate::document::{
    text_format, CharacterStyle, DocumentState, HeaderFooterKind, SectionId,
    TextRun,
};
use crate::layout::document_points_to_screen_points;

#[derive(Clone)]
pub(crate) struct HeaderRunPiece {
    pub(crate) text: String,
    pub(crate) style: CharacterStyle,
    pub(crate) range: Option<Range<usize>>,
}

impl HeaderRunPiece {
    pub(crate) fn runs_style(&self) -> CharacterStyle {
        self.style
    }
}

pub(crate) struct HeaderSegment {
    pub(crate) runs: Vec<HeaderRunPiece>,
    pub(crate) end: usize,
}

pub(crate) fn paint_page_header_footer(
    painter: &egui::Painter,
    document: &DocumentState,
    canvas: &CanvasState,
    page_rect: Rect,
    section_id: SectionId,
    page_index_within_section: usize,
    section_page_count: usize,
    page_number: usize,
    page_count: usize,
    active_header_footer: Option<ActiveHeaderFooter>,
    color: Color32,
) {
    let setup = document
        .section_by_id(section_id)
        .map(|section| section.page_setup)
        .unwrap_or_else(|| document.default_page_setup());
    let horizontal_margin =
        document_points_to_screen_points(setup.margins.left_points.max(18.0), canvas.zoom);
    let text_width = (page_rect.width() - horizontal_margin * 2.0).max(1.0);

    let header_variant = document.header_footer_variant_for_page(
        section_id,
        page_index_within_section,
        HeaderFooterKind::Header,
    );
    let header_story =
        document.resolve_header_footer_slot(section_id, HeaderFooterKind::Header, header_variant);
    let header_runs = rendered_header_footer_runs(
        document,
        &header_story.story.runs,
        section_id,
        page_index_within_section,
        page_number,
        page_count,
        section_page_count,
    );
    if !runs_plain_text(&header_runs).trim().is_empty()
        && active_header_footer
            != Some(ActiveHeaderFooter {
                kind: HeaderFooterKind::Header,
                section_id,
                variant: header_variant,
                page_number,
            })
    {
        let font_size = header_footer_base_font_size(&header_runs, canvas.zoom);
        let y = page_rect.top()
            + document_points_to_screen_points(setup.header_from_top_points, canvas.zoom);
        paint_tab_aligned_margin_runs(
            painter,
            &header_runs,
            canvas.zoom,
            color,
            Rect::from_min_size(
                egui::pos2(page_rect.left() + horizontal_margin, y),
                egui::vec2(text_width, font_size),
            ),
            None,
        );
    }

    let footer_variant = document.header_footer_variant_for_page(
        section_id,
        page_index_within_section,
        HeaderFooterKind::Footer,
    );
    let footer_story =
        document.resolve_header_footer_slot(section_id, HeaderFooterKind::Footer, footer_variant);
    let footer_runs = rendered_header_footer_runs(
        document,
        &footer_story.story.runs,
        section_id,
        page_index_within_section,
        page_number,
        page_count,
        section_page_count,
    );
    if !runs_plain_text(&footer_runs).trim().is_empty()
        && active_header_footer
            != Some(ActiveHeaderFooter {
                kind: HeaderFooterKind::Footer,
                section_id,
                variant: footer_variant,
                page_number,
            })
    {
        let font_size = header_footer_base_font_size(&footer_runs, canvas.zoom);
        let y = page_rect.bottom()
            - document_points_to_screen_points(setup.footer_from_bottom_points, canvas.zoom)
            - font_size;
        paint_tab_aligned_margin_runs(
            painter,
            &footer_runs,
            canvas.zoom,
            color,
            Rect::from_min_size(
                egui::pos2(page_rect.left() + horizontal_margin, y),
                egui::vec2(text_width, font_size),
            ),
            None,
        );
    }
}

pub(crate) fn rendered_header_footer_runs(
    document: &DocumentState,
    runs: &[TextRun],
    section_id: SectionId,
    page_index_within_section: usize,
    page_number: usize,
    page_count: usize,
    section_page_count: usize,
) -> Vec<TextRun> {
    runs.iter()
        .map(|run| TextRun {
            text: document.render_page_field_for_section_page(
                &run.text,
                section_id,
                page_index_within_section,
                page_number.saturating_sub(1),
                page_count,
                section_page_count,
            ),
            style: run.style,
        })
        .collect()
}

pub(crate) fn paint_tab_aligned_margin_runs(
    painter: &egui::Painter,
    runs: &[TextRun],
    zoom: f32,
    fallback_color: Color32,
    rect: Rect,
    selection: Option<Range<usize>>,
) {
    let segments = split_runs_for_header_tabs(runs);
    for slot in 0..3 {
        let Some(segment) = segments.get(slot) else {
            continue;
        };
        if segment.runs.is_empty() {
            continue;
        }
        let segment_width = measure_runs_width(painter, &segment.runs, zoom);
        let mut x = match slot {
            0 => rect.left(),
            1 => rect.center().x - segment_width * 0.5,
            _ => rect.right() - segment_width,
        };
        for piece in &segment.runs {
            let Some(piece_range) = piece.range.clone() else {
                continue;
            };
            if let Some(selection) = &selection {
                let start = piece_range.start.max(selection.start);
                let end = piece_range.end.min(selection.end);
                if start < end {
                    let before = slice_run_text_chars(&piece.text, 0..start - piece_range.start);
                    let selected = slice_run_text_chars(
                        &piece.text,
                        start - piece_range.start..end - piece_range.start,
                    );
                    let selected_x = x + measure_text_width(painter, &before, piece.style, zoom);
                    let selected_width = measure_text_width(painter, &selected, piece.style, zoom);
                    painter.rect_filled(
                        Rect::from_min_size(
                            egui::pos2(selected_x, rect.top()),
                            egui::vec2(
                                selected_width,
                                header_footer_line_height(piece.runs_style(), zoom),
                            ),
                        ),
                        CornerRadius::ZERO,
                        Color32::from_rgba_unmultiplied(80, 135, 230, 80),
                    );
                }
            }
            paint_run_text(
                painter,
                &piece.text,
                piece.style,
                zoom,
                egui::pos2(x, rect.top()),
                fallback_color,
            );
            x += measure_text_width(painter, &piece.text, piece.style, zoom);
        }
    }
}

pub(crate) fn split_runs_for_header_tabs(runs: &[TextRun]) -> Vec<HeaderSegment> {
    let mut segments = vec![HeaderSegment {
        runs: Vec::new(),
        end: 0,
    }];
    let mut slot = 0usize;
    let mut char_index = 0usize;

    for run in runs {
        let mut text = String::new();
        let mut piece_start = char_index;
        for ch in run.text.chars() {
            if ch == '\t' && slot < 2 {
                if !text.is_empty() {
                    segments[slot].runs.push(HeaderRunPiece {
                        text: std::mem::take(&mut text),
                        style: run.style,
                        range: Some(piece_start..char_index),
                    });
                }
                segments[slot].end = char_index;
                slot += 1;
                char_index += 1;
                piece_start = char_index;
                segments.push(HeaderSegment {
                    runs: Vec::new(),
                    end: char_index,
                });
            } else {
                text.push(if ch == '\t' { ' ' } else { ch });
                char_index += 1;
            }
        }
        if !text.is_empty() {
            segments[slot].runs.push(HeaderRunPiece {
                text,
                style: run.style,
                range: Some(piece_start..char_index),
            });
        }
    }
    if let Some(segment) = segments.get_mut(slot) {
        segment.end = char_index;
    }
    segments
}

pub(crate) fn measure_runs_width(painter: &egui::Painter, runs: &[HeaderRunPiece], zoom: f32) -> f32 {
    runs.iter()
        .map(|run| measure_text_width(painter, &run.text, run.style, zoom))
        .sum()
}

pub(crate) fn measure_text_width(
    painter: &egui::Painter,
    text: &str,
    style: CharacterStyle,
    zoom: f32,
) -> f32 {
    if text.is_empty() {
        return 0.0;
    }
    header_footer_text_galley(painter, text, style, zoom, Color32::BLACK)
        .size()
        .x
}

pub(crate) fn paint_run_text(
    painter: &egui::Painter,
    text: &str,
    style: CharacterStyle,
    zoom: f32,
    pos: egui::Pos2,
    fallback_color: Color32,
) {
    if text.is_empty() {
        return;
    }
    painter.galley(
        pos,
        header_footer_text_galley(painter, text, style, zoom, fallback_color),
        fallback_color,
    );
}

pub(crate) fn header_footer_text_galley(
    painter: &egui::Painter,
    text: &str,
    style: CharacterStyle,
    zoom: f32,
    fallback_color: Color32,
) -> Arc<egui::Galley> {
    let mut format = text_format(style, zoom);
    if format.color == Color32::TRANSPARENT {
        format.color = fallback_color;
    }
    let mut job = egui::epaint::text::LayoutJob::default();
    job.append(text, 0.0, format);
    painter.layout_job(job)
}

pub(crate) fn header_footer_line_height(style: CharacterStyle, zoom: f32) -> f32 {
    (style.font_size_points * zoom).max(1.0) * 1.2
}

pub(crate) fn header_footer_base_font_size(runs: &[TextRun], zoom: f32) -> f32 {
    runs.iter()
        .find(|run| !run.text.is_empty())
        .map(|run| run.style.font_size_points * zoom)
        .unwrap_or(9.0 * zoom)
        .clamp(7.0, 28.0)
}

pub(crate) fn header_footer_hit(
    page_layout: &crate::canvas::page_layout::PageLayout,
    document: &DocumentState,
    canvas: &CanvasState,
    pointer_pos: egui::Pos2,
) -> Option<ActiveHeaderFooter> {
    page_layout
        .pages
        .iter()
        .enumerate()
        .find_map(|(index, page)| {
            let page_number = index + 1;
            let header_variant = document.header_footer_variant_for_page(
                page.section_id,
                page.page_index_within_section,
                HeaderFooterKind::Header,
            );
            let footer_variant = document.header_footer_variant_for_page(
                page.section_id,
                page.page_index_within_section,
                HeaderFooterKind::Footer,
            );
            if page_header_rect(page.page_rect, document, canvas, page.section_id)
                .contains(pointer_pos)
            {
                Some(ActiveHeaderFooter {
                    kind: HeaderFooterKind::Header,
                    section_id: page.section_id,
                    variant: header_variant,
                    page_number,
                })
            } else if page_footer_rect(page.page_rect, document, canvas, page.section_id)
                .contains(pointer_pos)
            {
                Some(ActiveHeaderFooter {
                    kind: HeaderFooterKind::Footer,
                    section_id: page.section_id,
                    variant: footer_variant,
                    page_number,
                })
            } else {
                None
            }
        })
}

pub(crate) fn page_header_rect(
    page_rect: Rect,
    document: &DocumentState,
    canvas: &CanvasState,
    section_id: SectionId,
) -> Rect {
    let setup = document
        .section_by_id(section_id)
        .map(|section| section.page_setup)
        .unwrap_or_else(|| document.default_page_setup());
    let height = document_points_to_screen_points(setup.margins.top_points, canvas.zoom)
        .clamp(18.0, page_rect.height() * 0.25);
    Rect::from_min_size(page_rect.min, egui::vec2(page_rect.width(), height))
}

pub(crate) fn page_footer_rect(
    page_rect: Rect,
    document: &DocumentState,
    canvas: &CanvasState,
    section_id: SectionId,
) -> Rect {
    let setup = document
        .section_by_id(section_id)
        .map(|section| section.page_setup)
        .unwrap_or_else(|| document.default_page_setup());
    let height = document_points_to_screen_points(setup.margins.bottom_points, canvas.zoom)
        .clamp(18.0, page_rect.height() * 0.25);
    Rect::from_min_max(
        egui::pos2(page_rect.left(), page_rect.bottom() - height),
        page_rect.max,
    )
}

pub(crate) fn runs_plain_text(runs: &[TextRun]) -> String {
    runs.iter().map(|run| run.text.as_str()).collect()
}

pub(crate) fn runs_total_chars(runs: &[TextRun]) -> usize {
    runs.iter().map(|run| run.text.chars().count()).sum()
}

pub(crate) fn slice_run_text_chars(text: &str, range: Range<usize>) -> String {
    text.chars()
        .skip(range.start)
        .take(range.end.saturating_sub(range.start))
        .collect()
}
