use std::ops::Range;
use serde::Serialize;
use eframe::egui::Color32;

use crate::document::text::char_to_byte_index;
use crate::document::types::{
    append_text_run, CharacterStyle, DocumentImage, TextRun, OBJECT_REPLACEMENT_CHAR,
    DocumentState, ListKind,
};

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct TableCell {
    pub runs: Vec<TextRun>,
    pub images: Vec<DocumentImage>,
    pub col_span: u16,
    pub row_span: u16,
}

impl TableCell {
    pub fn new(text: &str) -> Self {
        Self {
            runs: vec![TextRun {
                text: text.to_owned(),
                style: CharacterStyle::default(),
            }],
            images: Vec::new(),
            col_span: 1,
            row_span: 1,
        }
    }

    pub fn plain_text(&self) -> String {
        self.runs.iter().map(|run| run.text.as_str()).collect()
    }

    pub(crate) fn total_chars(&self) -> usize {
        self.runs.iter().map(|run| run.text.chars().count()).sum()
    }

    pub(crate) fn typing_style(&self) -> CharacterStyle {
        self.runs.last().map(|run| run.style).unwrap_or_default()
    }

    pub(crate) fn style_at(&self, char_index: usize) -> CharacterStyle {
        let target = char_index.min(self.total_chars());
        let mut offset = 0usize;
        for run in &self.runs {
            let run_chars = run.text.chars().count();
            if target < offset + run_chars {
                return run.style;
            }
            offset += run_chars;
        }
        self.typing_style()
    }

    pub(crate) fn selection_style_at(&self, range: Range<usize>) -> CharacterStyle {
        let total_chars = self.total_chars();
        let start = range.start.min(total_chars);
        let end = range.end.min(total_chars);
        if start < end {
            return self.style_at(end - 1);
        }

        self.style_at(start)
    }

    pub(crate) fn append_text(&mut self, text: &str, style: CharacterStyle) {
        self.insert_text(self.total_chars(), text, style);
    }

    pub(crate) fn apply_style(&mut self, mutate: impl Fn(&mut CharacterStyle) + Copy) {
        for run in &mut self.runs {
            mutate(&mut run.style);
        }
        self.normalize_runs();
    }

    pub(crate) fn apply_style_to_range(&mut self, range: Range<usize>, mutate: impl Fn(&mut CharacterStyle)) {
        if range.start >= range.end {
            return;
        }

        let start = range.start.min(self.total_chars());
        let end = range.end.min(self.total_chars());
        self.split_at_char(start);
        self.split_at_char(end);

        let mut offset = 0usize;
        for run in &mut self.runs {
            let run_chars = run.text.chars().count();
            if offset >= start && offset + run_chars <= end {
                mutate(&mut run.style);
            }
            offset += run_chars;
        }

        self.normalize_runs();
    }

    pub(crate) fn insert_text(&mut self, char_index: usize, text: &str, style: CharacterStyle) {
        if text.is_empty() {
            return;
        }

        let insertion_index = char_index.min(self.total_chars());
        self.split_at_char(insertion_index);

        let mut offset = 0usize;
        let mut target = self.runs.len();
        for (idx, run) in self.runs.iter().enumerate() {
            if offset == insertion_index {
                target = idx;
                break;
            }
            offset += run.text.chars().count();
        }

        self.runs.insert(
            target,
            TextRun {
                text: text.to_owned(),
                style,
            },
        );
        self.normalize_runs();
    }

    pub(crate) fn replace_range_with_text(
        &mut self,
        range: Range<usize>,
        text: &str,
        style: CharacterStyle,
    ) -> usize {
        let start = range.start.min(self.total_chars());
        let end = range.end.min(self.total_chars());
        self.delete_char_range(start..end);
        self.insert_text(start, text, style);
        start + text.chars().count()
    }

    pub(crate) fn delete_char_range(&mut self, range: Range<usize>) {
        if range.start >= range.end {
            return;
        }

        let start = range.start.min(self.total_chars());
        let end = range.end.min(self.total_chars());
        let mut next_runs = Vec::new();
        let mut char_index = 0usize;
        let mut image_index = 0usize;
        let mut removed_images = Vec::new();

        for run in &self.runs {
            let mut kept = String::new();
            for ch in run.text.chars() {
                let removing = char_index >= start && char_index < end;
                if ch == OBJECT_REPLACEMENT_CHAR {
                    if removing {
                        removed_images.push(image_index);
                    }
                    image_index += 1;
                }
                if !removing {
                    kept.push(ch);
                }
                char_index += 1;
            }
            append_text_run(&mut next_runs, &kept, run.style);
        }

        if !removed_images.is_empty() {
            self.images = self
                .images
                .drain(..)
                .enumerate()
                .filter_map(|(idx, image)| (!removed_images.contains(&idx)).then_some(image))
                .collect();
        }
        self.runs = next_runs;
        self.normalize_runs();
    }

    pub(crate) fn split_at_char(&mut self, char_index: usize) {
        if char_index == 0 || char_index >= self.total_chars() {
            return;
        }

        let mut offset = 0usize;
        for idx in 0..self.runs.len() {
            let run_chars = self.runs[idx].text.chars().count();
            if char_index > offset && char_index < offset + run_chars {
                let local = char_index - offset;
                let byte_index = char_to_byte_index(&self.runs[idx].text, local);
                let right = self.runs[idx].text.split_off(byte_index);
                let style = self.runs[idx].style;
                self.runs.insert(idx + 1, TextRun { text: right, style });
                break;
            }
            offset += run_chars;
        }
    }

    pub(crate) fn normalize_runs(&mut self) {
        let fallback_style = self.runs.last().map(|run| run.style).unwrap_or_default();
        self.runs.retain(|run| !run.text.is_empty());
        let mut normalized: Vec<TextRun> = Vec::with_capacity(self.runs.len().max(1));
        for run in self.runs.drain(..) {
            if let Some(last) = normalized.last_mut() {
                if last.style == run.style {
                    last.text.push_str(&run.text);
                    continue;
                }
            }
            normalized.push(run);
        }
        if normalized.is_empty() {
            normalized.push(TextRun {
                text: String::new(),
                style: fallback_style,
            });
        }
        self.runs = normalized;
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Serialize)]
pub struct TableBorders {
    #[serde(serialize_with = "crate::document::types::serialize_color32")]
    pub color: Color32,
    pub width_points: f32,
}

impl Default for TableBorders {
    fn default() -> Self {
        Self {
            color: Color32::from_rgb(180, 180, 180),
            width_points: 0.75,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct DocumentTable {
    pub id: usize,
    pub rows: Vec<Vec<TableCell>>,
    pub col_widths_points: Vec<f32>,
    pub row_heights_points: Vec<f32>,
    pub borders: TableBorders,
}

impl DocumentTable {
    pub fn new(id: usize, num_rows: usize, num_cols: usize, available_width: f32) -> Self {
        let col_width = (available_width / num_cols as f32).max(36.0);
        let rows = (0..num_rows)
            .map(|_| (0..num_cols).map(|_| TableCell::new("")).collect())
            .collect();
        Self {
            id,
            rows,
            col_widths_points: vec![col_width; num_cols],
            row_heights_points: vec![20.0; num_rows],
            borders: TableBorders::default(),
        }
    }

    pub fn num_rows(&self) -> usize {
        self.rows.len()
    }

    pub fn num_cols(&self) -> usize {
        self.col_widths_points.len()
    }

    pub fn total_width_points(&self) -> f32 {
        self.col_widths_points.iter().sum()
    }

    pub fn total_height_points(&self) -> f32 {
        self.row_heights_points.iter().sum()
    }
}

impl DocumentState {
    pub fn insert_table(&mut self, char_index: usize, num_rows: usize, num_cols: usize) -> usize {
        let available_width =
            self.page_size.width_points - self.margins.left_points - self.margins.right_points;
        let next_id = self.next_table_id();
        let table = DocumentTable::new(next_id, num_rows, num_cols, available_width);

        let insert_at = char_index.min(self.total_chars());
        let paragraph_index = self.paragraph_index_at(insert_at);
        let paragraph_range = self
            .paragraphs()
            .get(paragraph_index)
            .map(|p| p.range.clone())
            .unwrap_or(insert_at..insert_at);

        let placeholder = OBJECT_REPLACEMENT_CHAR.to_string();
        let insertion_text = if insert_at == paragraph_range.start {
            format!("{placeholder}\n")
        } else if insert_at == paragraph_range.end {
            format!("\n{placeholder}")
        } else {
            format!("\n{placeholder}\n")
        };

        self.insert_text(insert_at, &insertion_text, CharacterStyle::default());

        let table_paragraph = if insert_at == paragraph_range.start {
            paragraph_index
        } else {
            paragraph_index + 1
        };

        if let Some(slot) = self.paragraph_tables.get_mut(table_paragraph) {
            *slot = Some(table);
        }
        if let Some(style) = self.paragraph_styles.get_mut(table_paragraph) {
            style.list_kind = ListKind::None;
        }
        self.ensure_paragraph_style_count();

        self.paragraphs()
            .get(table_paragraph)
            .map(|p| p.range.end)
            .unwrap_or(insert_at)
    }

    pub fn table_by_id(&self, id: usize) -> Option<&DocumentTable> {
        self.paragraph_tables
            .iter()
            .flatten()
            .find(|table| table.id == id)
    }

    pub(crate) fn table_by_id_mut(&mut self, id: usize) -> Option<&mut DocumentTable> {
        self.paragraph_tables
            .iter_mut()
            .flatten()
            .find(|table| table.id == id)
    }

    pub fn insert_table_row(&mut self, table_id: usize, after_row: usize) {
        if let Some(table) = self.table_by_id_mut(table_id) {
            let num_cols = table.num_cols();
            let insert_at = if after_row == usize::MAX {
                0
            } else {
                (after_row + 1).min(table.rows.len())
            };
            let new_row: Vec<TableCell> = (0..num_cols).map(|_| TableCell::new("")).collect();
            table.rows.insert(insert_at, new_row);
            table.row_heights_points.insert(insert_at, 20.0);
        }
    }

    pub fn insert_table_column(&mut self, table_id: usize, after_col: usize) {
        if let Some(table) = self.table_by_id_mut(table_id) {
            let insert_at = if after_col == usize::MAX {
                0
            } else {
                (after_col + 1).min(table.num_cols())
            };
            // Reduce existing column widths to make room
            let total_width: f32 = table.col_widths_points.iter().sum();
            let new_col_count = table.num_cols() + 1;
            let new_col_width = total_width / new_col_count as f32;
            let scale = (total_width - new_col_width) / total_width.max(1.0);
            for w in table.col_widths_points.iter_mut() {
                *w *= scale;
            }
            table.col_widths_points.insert(insert_at, new_col_width);
            for row in &mut table.rows {
                row.insert(insert_at, TableCell::new(""));
            }
        }
    }

    pub fn delete_table_row(&mut self, table_id: usize, row_index: usize) {
        if let Some(table) = self.table_by_id_mut(table_id) {
            if table.rows.len() > 1 && row_index < table.rows.len() {
                table.rows.remove(row_index);
                table.row_heights_points.remove(row_index);
            }
        }
    }

    pub fn delete_table_column(&mut self, table_id: usize, col_index: usize) {
        if let Some(table) = self.table_by_id_mut(table_id) {
            if table.num_cols() > 1 && col_index < table.num_cols() {
                let removed_width = table.col_widths_points[col_index];
                table.col_widths_points.remove(col_index);
                // Redistribute removed width
                let remaining_cols = table.col_widths_points.len();
                let extra_each = removed_width / remaining_cols as f32;
                for w in table.col_widths_points.iter_mut() {
                    *w += extra_each;
                }
                for row in &mut table.rows {
                    if col_index < row.len() {
                        row.remove(col_index);
                    }
                }
            }
        }
    }

    pub fn set_table_cell_text(&mut self, table_id: usize, row: usize, col: usize, text: &str) {
        if let Some(table) = self.table_by_id_mut(table_id) {
            if let Some(cell) = table.rows.get_mut(row).and_then(|r| r.get_mut(col)) {
                cell.runs = vec![TextRun {
                    text: text.to_owned(),
                    style: CharacterStyle::default(),
                }];
                cell.images.clear();
            }
        }
    }

    pub fn table_cell_text(&self, table_id: usize, row: usize, col: usize) -> Option<String> {
        self.table_by_id(table_id)
            .and_then(|table| table.rows.get(row))
            .and_then(|cells| cells.get(col))
            .map(TableCell::plain_text)
    }

    pub fn table_cell_typing_style(
        &self,
        table_id: usize,
        row: usize,
        col: usize,
    ) -> Option<CharacterStyle> {
        self.table_by_id(table_id)
            .and_then(|table| table.rows.get(row))
            .and_then(|cells| cells.get(col))
            .map(TableCell::typing_style)
    }

    pub fn table_cell_style_at(
        &self,
        table_id: usize,
        row: usize,
        col: usize,
        char_index: usize,
    ) -> Option<CharacterStyle> {
        self.table_by_id(table_id)
            .and_then(|table| table.rows.get(row))
            .and_then(|cells| cells.get(col))
            .map(|cell| cell.style_at(char_index))
    }

    pub fn table_cell_selection_style_at(
        &self,
        table_id: usize,
        row: usize,
        col: usize,
        range: Range<usize>,
    ) -> Option<CharacterStyle> {
        self.table_by_id(table_id)
            .and_then(|table| table.rows.get(row))
            .and_then(|cells| cells.get(col))
            .map(|cell| cell.selection_style_at(range))
    }

    pub fn table_cell_len(&self, table_id: usize, row: usize, col: usize) -> Option<usize> {
        self.table_by_id(table_id)
            .and_then(|table| table.rows.get(row))
            .and_then(|cells| cells.get(col))
            .map(TableCell::total_chars)
    }

    pub fn append_table_cell_text(
        &mut self,
        table_id: usize,
        row: usize,
        col: usize,
        text: &str,
        style: CharacterStyle,
    ) {
        if let Some(table) = self.table_by_id_mut(table_id) {
            if let Some(cell) = table.rows.get_mut(row).and_then(|cells| cells.get_mut(col)) {
                cell.append_text(text, style);
            }
        }
    }

    pub fn replace_table_cell_range_with_text(
        &mut self,
        table_id: usize,
        row: usize,
        col: usize,
        range: Range<usize>,
        text: &str,
        style: CharacterStyle,
    ) -> Option<usize> {
        self.table_by_id_mut(table_id)
            .and_then(|table| table.rows.get_mut(row))
            .and_then(|cells| cells.get_mut(col))
            .map(|cell| cell.replace_range_with_text(range, text, style))
    }

    pub fn delete_table_cell_char_range(
        &mut self,
        table_id: usize,
        row: usize,
        col: usize,
        range: Range<usize>,
    ) {
        if let Some(table) = self.table_by_id_mut(table_id) {
            if let Some(cell) = table.rows.get_mut(row).and_then(|cells| cells.get_mut(col)) {
                cell.delete_char_range(range);
            }
        }
    }

    pub fn insert_table_cell_image(
        &mut self,
        table_id: usize,
        row: usize,
        col: usize,
        image: DocumentImage,
        style: CharacterStyle,
    ) {
        if let Some(table) = self.table_by_id_mut(table_id) {
            if let Some(cell) = table.rows.get_mut(row).and_then(|cells| cells.get_mut(col)) {
                cell.append_text(&OBJECT_REPLACEMENT_CHAR.to_string(), style);
                cell.images.push(image);
            }
        }
    }

    pub fn resize_table_column_pair(
        &mut self,
        table_id: usize,
        left_col: usize,
        left_width_points: f32,
        right_width_points: f32,
    ) {
        if let Some(table) = self.table_by_id_mut(table_id) {
            if left_col + 1 < table.col_widths_points.len() {
                table.col_widths_points[left_col] = left_width_points.max(18.0);
                table.col_widths_points[left_col + 1] = right_width_points.max(18.0);
            }
        }
    }

    pub fn resize_table_row_pair(
        &mut self,
        table_id: usize,
        top_row: usize,
        top_height_points: f32,
        bottom_height_points: f32,
    ) {
        if let Some(table) = self.table_by_id_mut(table_id) {
            if top_row + 1 < table.row_heights_points.len() {
                table.row_heights_points[top_row] = top_height_points.max(12.0);
                table.row_heights_points[top_row + 1] = bottom_height_points.max(12.0);
            }
        }
    }

    pub fn set_table_border_width(&mut self, table_id: usize, width_points: f32) {
        if let Some(table) = self.table_by_id_mut(table_id) {
            table.borders.width_points = width_points.clamp(0.0, 8.0);
        }
    }

    pub fn set_table_border_color(&mut self, table_id: usize, color: Color32) {
        if let Some(table) = self.table_by_id_mut(table_id) {
            table.borders.color = color;
        }
    }

    pub fn merge_table_cell_right(&mut self, table_id: usize, row: usize, col: usize) -> bool {
        let Some(table) = self.table_by_id_mut(table_id) else {
            return false;
        };
        let Some(row_cells) = table.rows.get_mut(row) else {
            return false;
        };
        if col + 1 >= row_cells.len() || row_cells[col].col_span == 0 {
            return false;
        }
        let next_span = row_cells[col + 1].col_span;
        if next_span == 0 {
            return false;
        }
        let mut merged_cell = row_cells[col + 1].clone();
        let merged_text = merged_cell.plain_text();
        if !merged_text.is_empty() {
            if !row_cells[col].plain_text().is_empty() {
                let style = row_cells[col].typing_style();
                row_cells[col].append_text(" ", style);
            }
            row_cells[col].runs.append(&mut merged_cell.runs);
            row_cells[col].normalize_runs();
        }
        row_cells[col].images.append(&mut merged_cell.images);
        row_cells[col].col_span = row_cells[col].col_span.saturating_add(next_span);
        row_cells[col + 1].col_span = 0;
        row_cells[col + 1].row_span = 0;
        row_cells[col + 1].runs.clear();
        row_cells[col + 1].images.clear();
        true
    }

    pub fn split_table_cell(&mut self, table_id: usize, row: usize, col: usize) -> bool {
        let Some(table) = self.table_by_id_mut(table_id) else {
            return false;
        };
        let Some(cell) = table.rows.get_mut(row).and_then(|cells| cells.get_mut(col)) else {
            return false;
        };
        let col_span = cell.col_span.max(1);
        let row_span = cell.row_span.max(1);
        if col_span == 1 && row_span == 1 {
            return false;
        }
        cell.col_span = 1;
        cell.row_span = 1;

        let max_row = (row + row_span as usize).min(table.rows.len());
        let max_col = (col + col_span as usize).min(table.num_cols());
        for row_idx in row..max_row {
            for col_idx in col..max_col {
                if row_idx == row && col_idx == col {
                    continue;
                }
                if let Some(covered) = table
                    .rows
                    .get_mut(row_idx)
                    .and_then(|cells| cells.get_mut(col_idx))
                {
                    if covered.col_span == 0 || covered.row_span == 0 {
                        *covered = TableCell::new("");
                    }
                }
            }
        }
        true
    }

    fn next_table_id(&self) -> usize {
        self.paragraph_tables
            .iter()
            .flatten()
            .map(|t| t.id)
            .max()
            .unwrap_or(0)
            + 1
    }
}
