use std::ops::Range;

use crate::document::types::{CharacterStyle, DocumentState, ParagraphStyle};

impl DocumentState {
    pub fn apply_style_to_range(
        &mut self,
        range: Range<usize>,
        mutate: impl Fn(&mut CharacterStyle),
    ) {
        if range.start >= range.end {
            return;
        }

        let start = range.start.min(self.total_chars());
        let end = range.end.min(self.total_chars());
        self.split_at_char(start);
        self.split_at_char(end);

        let mut offset = 0;
        for run in &mut self.runs {
            let run_chars = run.text.chars().count();
            if offset >= start && offset + run_chars <= end {
                mutate(&mut run.style);
            }
            offset += run_chars;
        }

        self.normalize_runs();
    }

    pub fn apply_paragraph_style_to_range(
        &mut self,
        range: Range<usize>,
        mutate: impl Fn(&mut ParagraphStyle),
    ) {
        let total_chars = self.total_chars();
        let start = range.start.min(total_chars);
        let end = range.end.min(total_chars);
        let start_paragraph = self.paragraph_index_at(start);
        let end_index = if start < end {
            end.saturating_sub(1)
        } else {
            start
        };
        let end_paragraph = self.paragraph_index_at(end_index);

        for paragraph_style in self
            .paragraph_styles
            .iter_mut()
            .skip(start_paragraph)
            .take(end_paragraph.saturating_sub(start_paragraph) + 1)
        {
            mutate(paragraph_style);
        }
    }

    pub fn apply_style_to_table_cell(
        &mut self,
        table_id: usize,
        row: usize,
        col: usize,
        mutate: impl Fn(&mut CharacterStyle) + Copy,
    ) {
        if let Some(table) = self.table_by_id_mut(table_id) {
            if let Some(cell) = table.rows.get_mut(row).and_then(|cells| cells.get_mut(col)) {
                cell.apply_style(mutate);
            }
        }
    }

    pub fn apply_style_to_table_cell_range(
        &mut self,
        table_id: usize,
        row: usize,
        col: usize,
        range: Range<usize>,
        mutate: impl Fn(&mut CharacterStyle) + Copy,
    ) {
        if let Some(table) = self.table_by_id_mut(table_id) {
            if let Some(cell) = table.rows.get_mut(row).and_then(|cells| cells.get_mut(col)) {
                cell.apply_style_to_range(range, mutate);
            }
        }
    }
}
