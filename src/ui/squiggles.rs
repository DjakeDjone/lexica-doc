use eframe::egui::{self, Color32, Painter, Pos2, Rect, Sense, Stroke};

use crate::grammar::GrammarError;

const SQUIGGLE_AMPLITUDE: f32 = 1.5;
const SQUIGGLE_WAVELENGTH: f32 = 4.0;
const SQUIGGLE_STEP: f32 = 0.5;
const SQUIGGLE_Y_OFFSET: f32 = 1.0;

#[derive(Clone, Copy, Debug)]
pub struct SquigglePageSlice {
    pub content_rect: Rect,
    pub start_y: f32,
    pub end_y: f32,
}

#[derive(Clone, Debug)]
pub struct ReplacementAction {
    pub byte_start: usize,
    pub byte_end: usize,
    pub replacement: String,
}

pub fn paint_grammar_squiggles(
    ui: &mut egui::Ui,
    painter: &Painter,
    galley: &egui::Galley,
    pages: &[SquigglePageSlice],
    errors: &[GrammarError],
) -> Option<ReplacementAction> {
    let mut selected_replacement: Option<ReplacementAction> = None;
    let text = galley.text();

    for error in errors {
        let start_char = byte_to_char_index(text, error.byte_start);
        let mut end_char = byte_to_char_index(text, error.byte_end);
        if end_char <= start_char {
            end_char = (start_char + 1).min(text.chars().count());
        }

        let segments = range_segments(galley, pages, start_char, end_char);
        for segment_rect in &segments {
            draw_squiggle(painter, *segment_rect);
            let response =
                ui.allocate_rect(segment_rect.expand2(egui::vec2(0.0, 3.0)), Sense::hover());

            if selected_replacement.is_none() {
                egui::Tooltip::for_enabled(&response).show(|ui| {
                    ui.set_max_width(360.0);
                    ui.label(&error.message);
                    ui.separator();

                    for replacement in error.replacements.iter().take(5) {
                        if ui.button(replacement).clicked() {
                            selected_replacement = Some(ReplacementAction {
                                byte_start: error.byte_start,
                                byte_end: error.byte_end,
                                replacement: replacement.clone(),
                            });
                        }
                    }

                    if error.replacements.is_empty() {
                        ui.label("No replacements available.");
                    }
                });
            }
        }
    }

    selected_replacement
}

fn draw_squiggle(painter: &Painter, rect: Rect) {
    if rect.width() <= 0.0 {
        return;
    }

    let stroke = Stroke::new(1.0, Color32::from_rgb(210, 38, 38));
    let baseline = rect.bottom() + SQUIGGLE_Y_OFFSET;
    let mut x = rect.left();
    let mut last = Pos2::new(x, baseline);

    while x <= rect.right() {
        let phase = (x - rect.left()) * std::f32::consts::TAU / SQUIGGLE_WAVELENGTH;
        let current = Pos2::new(x, baseline + SQUIGGLE_AMPLITUDE * phase.sin());
        painter.line_segment([last, current], stroke);
        last = current;
        x += SQUIGGLE_STEP;
    }
}

fn range_segments(
    galley: &egui::Galley,
    pages: &[SquigglePageSlice],
    start_char: usize,
    end_char: usize,
) -> Vec<Rect> {
    let mut segments = Vec::new();
    let mut char_cursor = 0usize;

    for row in &galley.rows {
        let row_char_start = char_cursor;
        let row_char_end = row_char_start + row.char_count_excluding_newline();
        let row_char_with_newline = row.char_count_including_newline();

        let overlap_start = start_char.max(row_char_start);
        let overlap_end = end_char.min(row_char_end);
        if overlap_start < overlap_end {
            let start_column = overlap_start - row_char_start;
            let end_column = overlap_end - row_char_start;
            let start_x = row.pos.x + row.x_offset(start_column);
            let mut end_x = row.pos.x + row.x_offset(end_column);

            if end_x <= start_x {
                end_x = start_x + 4.0;
            }

            if let Some(page) = page_for_row(pages, row.pos.y) {
                let top = page.content_rect.top() + row.pos.y - page.start_y;
                let left = page.content_rect.left() + start_x;
                let right = page.content_rect.left() + end_x;
                let bottom = top + row.row.height();
                segments.push(Rect::from_min_max(
                    Pos2::new(left, top),
                    Pos2::new(right, bottom),
                ));
            }
        }

        char_cursor += row_char_with_newline;
    }

    segments
}

fn page_for_row(pages: &[SquigglePageSlice], row_y: f32) -> Option<&SquigglePageSlice> {
    pages
        .iter()
        .find(|page| row_y >= page.start_y && row_y <= page.end_y)
}

fn byte_to_char_index(text: &str, byte_offset: usize) -> usize {
    let clamped = byte_offset.min(text.len());
    let mut count = 0usize;
    for (idx, _) in text.char_indices() {
        if idx >= clamped {
            break;
        }
        count += 1;
    }
    if clamped == text.len() {
        text.chars().count()
    } else {
        count
    }
}
