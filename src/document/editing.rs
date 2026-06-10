use crate::document::text::char_to_byte_index;
use crate::document::types::{
    append_text_run, CharacterStyle, DistanceFromText, DocumentImage, DocumentState,
    HorizontalPosition, ImageLayoutMode, ImageRendering, ListKind, Paragraph, ParagraphStyle,
    TextRun, VerticalPosition, WrapMode, OBJECT_REPLACEMENT_CHAR,
};
use std::ops::Range;

impl DocumentState {
    pub fn insert_text(&mut self, char_index: usize, text: &str, style: CharacterStyle) {
        if text.is_empty() {
            return;
        }

        let insertion_index = char_index.min(self.total_chars());
        let inserted_paragraphs = text.chars().filter(|ch| *ch == '\n').count();
        if inserted_paragraphs > 0 {
            let paragraph_index = self.paragraph_index_at(insertion_index);
            let paragraph_style = self
                .paragraph_styles
                .get(paragraph_index)
                .copied()
                .unwrap_or_default();
            for offset in 0..inserted_paragraphs {
                let mut inserted_style = paragraph_style;
                inserted_style.page_break_before = false;
                self.paragraph_styles
                    .insert(paragraph_index + offset + 1, inserted_style);
                self.paragraph_images
                    .insert(paragraph_index + offset + 1, None);
                self.paragraph_tables
                    .insert(paragraph_index + offset + 1, None);
            }
        }

        self.split_at_char(insertion_index);

        let mut offset = 0;
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
        self.ensure_paragraph_style_count();
    }

    pub fn replace_range_with_runs(&mut self, range: Range<usize>, runs: Vec<TextRun>) {
        let start = range.start.min(self.total_chars());
        let end = range.end.min(self.total_chars());
        self.delete_range(start..end);

        let mut insert_at = start;
        for run in runs {
            if run.text.is_empty() {
                continue;
            }
            let run_len = run.text.chars().count();
            self.insert_text(insert_at, &run.text, run.style);
            insert_at += run_len;
        }
    }

    pub fn delete_range(&mut self, range: Range<usize>) {
        if range.start >= range.end {
            return;
        }

        let start = range.start.min(self.total_chars());
        let end = range.end.min(self.total_chars());
        let paragraph_index = self.paragraph_index_at(start);
        let removed_text = self.selected_text(start..end);
        let removed_paragraphs = removed_text.chars().filter(|ch| *ch == '\n').count();
        if removed_text.chars().any(|ch| ch == OBJECT_REPLACEMENT_CHAR) {
            let end_paragraph = self.paragraph_index_at(end.saturating_sub(1));
            for image in self
                .paragraph_images
                .iter_mut()
                .skip(paragraph_index)
                .take(end_paragraph.saturating_sub(paragraph_index) + 1)
            {
                *image = None;
            }
        }
        if removed_paragraphs > 0 {
            let drain_start = paragraph_index + 1;
            let drain_end = (drain_start + removed_paragraphs).min(self.paragraph_styles.len());
            self.paragraph_styles.drain(drain_start..drain_end);
            let image_drain_end =
                (drain_start + removed_paragraphs).min(self.paragraph_images.len());
            self.paragraph_images.drain(drain_start..image_drain_end);
            let table_drain_end =
                (drain_start + removed_paragraphs).min(self.paragraph_tables.len());
            self.paragraph_tables.drain(drain_start..table_drain_end);
        }

        self.split_at_char(start);
        self.split_at_char(end);

        let mut offset = 0;
        self.runs.retain(|run| {
            let run_chars = run.text.chars().count();
            let keep = offset + run_chars <= start || offset >= end;
            offset += run_chars;
            keep
        });

        self.normalize_runs();
        self.ensure_paragraph_style_count();
    }

    pub fn replace_with_runs(&mut self, title: String, runs: Vec<TextRun>) {
        self.title = title;
        self.runs = if runs.is_empty() {
            vec![TextRun {
                text: String::new(),
                style: CharacterStyle::default(),
            }]
        } else {
            runs
        };
        self.paragraph_styles = vec![ParagraphStyle::default(); self.paragraph_count()];
        self.paragraph_images = vec![None; self.paragraph_count()];
        self.paragraph_tables = vec![None; self.paragraph_count()];
        self.normalize_runs();
        self.ensure_paragraph_style_count();
    }

    pub fn insert_page_break(&mut self, char_index: usize) -> usize {
        let total_chars = self.total_chars();
        let insert_at = char_index.min(total_chars);
        let paragraph_count = self.paragraph_count();
        let paragraph_index = self.paragraph_index_at(insert_at);
        let paragraph_range = self
            .paragraphs()
            .get(paragraph_index)
            .map(|paragraph| paragraph.range.clone())
            .unwrap_or(insert_at..insert_at);

        let target_paragraph = if insert_at == paragraph_range.start {
            if paragraph_index == 0 {
                self.insert_text(0, "\n", CharacterStyle::default());
                1
            } else {
                paragraph_index
            }
        } else if insert_at == paragraph_range.end {
            if paragraph_index + 1 < paragraph_count {
                paragraph_index + 1
            } else {
                self.insert_text(insert_at, "\n", CharacterStyle::default());
                paragraph_index + 1
            }
        } else {
            self.insert_text(insert_at, "\n", CharacterStyle::default());
            paragraph_index + 1
        };

        if let Some(style) = self.paragraph_styles.get_mut(target_paragraph) {
            style.page_break_before = true;
        }
        self.ensure_paragraph_style_count();

        self.paragraphs()
            .get(target_paragraph)
            .map(|paragraph| paragraph.range.start)
            .unwrap_or(insert_at)
    }

    pub fn insert_image(&mut self, char_index: usize, image: DocumentImage) -> usize {
        let insert_at = char_index.min(self.total_chars());
        let paragraph_index = self.paragraph_index_at(insert_at);
        let paragraph_range = self
            .paragraphs()
            .get(paragraph_index)
            .map(|paragraph| paragraph.range.clone())
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

        let image_paragraph = if insert_at == paragraph_range.start {
            paragraph_index
        } else {
            paragraph_index + 1
        };

        if let Some(slot) = self.paragraph_images.get_mut(image_paragraph) {
            *slot = Some(image);
        }
        if let Some(style) = self.paragraph_styles.get_mut(image_paragraph) {
            style.list_kind = ListKind::None;
        }
        self.ensure_paragraph_style_count();

        self.paragraphs()
            .get(image_paragraph)
            .map(|paragraph| paragraph.range.end)
            .unwrap_or(insert_at)
    }

    pub fn resize_image_by_id(&mut self, id: usize, width_points: f32, height_points: f32) {
        if let Some(image) = self.image_by_id_mut(id) {
            image.width_points = width_points.max(24.0);
            image.height_points = height_points.max(24.0);
        }
    }

    pub fn set_image_offset_by_id(&mut self, id: usize, x_points: f32, y_points: f32) {
        if let Some(image) = self.image_by_id_mut(id) {
            image.set_manual_offset(x_points, y_points);
        }
    }

    pub fn adjust_image_offset_by_id(&mut self, id: usize, dx: f32, dy: f32) {
        if let Some(image) = self.image_by_id_mut(id) {
            image.adjust_manual_offset(dx, dy);
        }
    }

    pub fn set_image_layout_mode(&mut self, id: usize, mode: ImageLayoutMode) {
        if let Some(image) = self.image_by_id_mut(id) {
            image.layout_mode = mode;
            if mode == ImageLayoutMode::Inline {
                image.wrap_mode = WrapMode::Inline;
                image.set_manual_offset(0.0, 0.0);
            } else if image.wrap_mode == WrapMode::Inline {
                image.wrap_mode = WrapMode::Square;
            }
        }
    }

    pub fn set_image_horizontal_position(&mut self, id: usize, pos: HorizontalPosition) {
        if let Some(image) = self.image_by_id_mut(id) {
            image.horizontal_position = pos;
        }
    }

    pub fn set_image_vertical_position(&mut self, id: usize, pos: VerticalPosition) {
        if let Some(image) = self.image_by_id_mut(id) {
            image.vertical_position = pos;
        }
    }

    pub fn set_image_distance_from_text(&mut self, id: usize, dist: DistanceFromText) {
        if let Some(image) = self.image_by_id_mut(id) {
            image.distance_from_text = dist;
        }
    }

    pub fn set_image_z_index(&mut self, id: usize, z: i32) {
        if let Some(image) = self.image_by_id_mut(id) {
            image.z_index = z;
        }
    }

    pub fn set_image_move_with_text(&mut self, id: usize, flag: bool) {
        if let Some(image) = self.image_by_id_mut(id) {
            image.move_with_text = flag;
        }
    }

    pub fn set_image_lock_aspect_ratio(&mut self, id: usize, flag: bool) {
        if let Some(image) = self.image_by_id_mut(id) {
            image.lock_aspect_ratio = flag;
        }
    }

    pub fn set_image_opacity(&mut self, id: usize, opacity: f32) {
        if let Some(image) = self.image_by_id_mut(id) {
            image.opacity = opacity.clamp(0.0, 1.0);
        }
    }

    pub fn set_image_wrap_mode(&mut self, id: usize, wrap_mode: WrapMode) {
        if let Some(image) = self.image_by_id_mut(id) {
            image.wrap_mode = wrap_mode;
        }
    }

    pub fn set_image_rendering(&mut self, id: usize, rendering: ImageRendering) {
        if let Some(image) = self.image_by_id_mut(id) {
            image.rendering = rendering;
        }
    }

    pub fn image_by_id(&self, id: usize) -> Option<&DocumentImage> {
        self.paragraph_images
            .iter()
            .flatten()
            .find(|image| image.id == id)
    }

    pub fn move_image_paragraph_to_cursor(
        &mut self,
        id: usize,
        target_char_index: usize,
    ) -> Option<usize> {
        let mut paragraphs = self.paragraphs();
        let source_index = paragraphs
            .iter()
            .position(|paragraph| paragraph.image.as_ref().is_some_and(|image| image.id == id))?;

        let total_chars = self.total_chars();
        let mut target_index = if target_char_index >= total_chars {
            paragraphs.len()
        } else {
            self.paragraph_index_at(target_char_index)
                .min(paragraphs.len())
        };

        if source_index == target_index || source_index + 1 == target_index {
            return paragraphs
                .get(source_index)
                .map(|paragraph| paragraph.range.start);
        }

        let moved = paragraphs.remove(source_index);
        if source_index < target_index {
            target_index -= 1;
        }
        let target_index = target_index.min(paragraphs.len());
        paragraphs.insert(target_index, moved);
        self.replace_paragraphs(paragraphs);

        self.paragraphs()
            .into_iter()
            .find(|paragraph| paragraph.image.as_ref().is_some_and(|image| image.id == id))
            .map(|paragraph| paragraph.range.start)
    }

    pub fn paragraphs(&self) -> Vec<Paragraph> {
        let mut paragraphs = Vec::with_capacity(self.paragraph_count());
        let mut current_runs = Vec::new();
        let mut current_length = 0usize;
        let mut paragraph_start = 0usize;
        let mut paragraph_index = 0usize;
        let mut ordered_index = 0usize;
        let mut previous_was_ordered = false;

        let push_paragraph = |paragraphs: &mut Vec<Paragraph>,
                              current_runs: &mut Vec<TextRun>,
                              current_length: &mut usize,
                              paragraph_start: &mut usize,
                              paragraph_index: &mut usize,
                              ordered_index: &mut usize,
                              previous_was_ordered: &mut bool| {
            let style = self
                .paragraph_styles
                .get(*paragraph_index)
                .copied()
                .unwrap_or_default();
            let list_marker = match style.list_kind {
                ListKind::None => {
                    *ordered_index = 0;
                    *previous_was_ordered = false;
                    None
                }
                ListKind::Bullet => {
                    *ordered_index = 0;
                    *previous_was_ordered = false;
                    Some("•".to_owned())
                }
                ListKind::Ordered => {
                    if *previous_was_ordered {
                        *ordered_index += 1;
                    } else {
                        *ordered_index = 1;
                        *previous_was_ordered = true;
                    }
                    Some(format!("{}.", *ordered_index))
                }
            };

            paragraphs.push(Paragraph {
                index: *paragraph_index,
                range: *paragraph_start..(*paragraph_start + *current_length),
                style,
                runs: std::mem::take(current_runs),
                list_marker,
                image: self
                    .paragraph_images
                    .get(*paragraph_index)
                    .cloned()
                    .unwrap_or(None),
                table: self
                    .paragraph_tables
                    .get(*paragraph_index)
                    .cloned()
                    .unwrap_or(None),
            });

            *paragraph_start += *current_length + 1;
            *current_length = 0;
            *paragraph_index += 1;
        };

        for run in &self.runs {
            let mut segment = String::new();
            for ch in run.text.chars() {
                if ch == '\n' {
                    if !segment.is_empty() {
                        current_length += segment.chars().count();
                        append_text_run(&mut current_runs, &segment, run.style);
                        segment.clear();
                    }
                    push_paragraph(
                        &mut paragraphs,
                        &mut current_runs,
                        &mut current_length,
                        &mut paragraph_start,
                        &mut paragraph_index,
                        &mut ordered_index,
                        &mut previous_was_ordered,
                    );
                } else {
                    segment.push(ch);
                }
            }

            if !segment.is_empty() {
                current_length += segment.chars().count();
                append_text_run(&mut current_runs, &segment, run.style);
            }
        }

        push_paragraph(
            &mut paragraphs,
            &mut current_runs,
            &mut current_length,
            &mut paragraph_start,
            &mut paragraph_index,
            &mut ordered_index,
            &mut previous_was_ordered,
        );

        if paragraphs.is_empty() {
            paragraphs.push(Paragraph {
                index: 0,
                range: 0..0,
                style: ParagraphStyle::default(),
                runs: Vec::new(),
                list_marker: None,
                image: None,
                table: None,
            });
        }

        paragraphs
    }

    pub(crate) fn split_at_char(&mut self, char_index: usize) {
        if char_index == 0 || char_index >= self.total_chars() {
            return;
        }

        let mut offset = 0;
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

    pub(crate) fn image_by_id_mut(&mut self, id: usize) -> Option<&mut DocumentImage> {
        self.paragraph_images
            .iter_mut()
            .flatten()
            .find(|image| image.id == id)
    }

    pub fn next_image_id(&self) -> usize {
        let paragraph_max = self
            .paragraph_images
            .iter()
            .flatten()
            .map(|image| image.id)
            .max()
            .unwrap_or(0);
        let table_max = self
            .paragraph_tables
            .iter()
            .flatten()
            .flat_map(|table| table.rows.iter())
            .flat_map(|row| row.iter())
            .flat_map(|cell| cell.images.iter())
            .map(|image| image.id)
            .max()
            .unwrap_or(0);

        paragraph_max.max(table_max) + 1
    }

    fn replace_paragraphs(&mut self, paragraphs: Vec<Paragraph>) {
        let mut runs = Vec::new();
        let mut paragraph_styles = Vec::with_capacity(paragraphs.len());
        let mut paragraph_images = Vec::with_capacity(paragraphs.len());
        let mut paragraph_tables = Vec::with_capacity(paragraphs.len());
        let paragraph_count = paragraphs.len();

        for (index, paragraph) in paragraphs.into_iter().enumerate() {
            paragraph_styles.push(paragraph.style);
            paragraph_images.push(paragraph.image);
            paragraph_tables.push(paragraph.table);
            for run in paragraph.runs {
                append_text_run(&mut runs, &run.text, run.style);
            }
            if index + 1 < paragraph_count {
                append_text_run(&mut runs, "\n", CharacterStyle::default());
            }
        }

        if runs.is_empty() {
            runs.push(TextRun {
                text: String::new(),
                style: CharacterStyle::default(),
            });
        }

        self.runs = runs;
        self.paragraph_styles = paragraph_styles;
        self.paragraph_images = paragraph_images;
        self.paragraph_tables = paragraph_tables;
        self.normalize_runs();
        self.ensure_paragraph_style_count();
    }

    pub(crate) fn normalize_runs(&mut self) {
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
                style: CharacterStyle::default(),
            });
        }

        self.runs = normalized;
    }

    pub(crate) fn ensure_paragraph_style_count(&mut self) {
        let target = self.paragraph_count().max(1);
        self.paragraph_styles
            .resize(target, ParagraphStyle::default());
        self.paragraph_images.resize(target, None);
        self.paragraph_tables.resize(target, None);
    }
}
