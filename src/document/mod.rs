pub mod docx;
mod markdown;
mod text;

use std::{collections::BTreeMap, fmt::Write as _, fs, ops::Range, path::Path};

use base64::{engine::general_purpose::STANDARD as BASE64_STANDARD, Engine as _};
use eframe::egui::{
    epaint::text::{TextFormat, VariationCoords},
    Color32, FontFamily, FontId, Stroke,
};
use printpdf::{
    Base64OrRaw, BuiltinFont, GeneratePdfOptions, Op, PdfDocument, PdfFontHandle, PdfPage,
    PdfSaveOptions, Point, Pt, TextItem,
};
use serde::Serialize;

use docx::docx_to_document;
use markdown::markdown_to_runs;
use text::{char_to_byte_index, line_char_range, slice_char_range, word_char_range};

pub const OBJECT_REPLACEMENT_CHAR: char = '\u{fffc}';

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
pub enum FontChoice {
    Proportional,
    Monospace,
}

impl FontChoice {
    pub const ALL: [Self; 2] = [Self::Proportional, Self::Monospace];

    pub const fn label(self) -> &'static str {
        match self {
            Self::Proportional => "Body",
            Self::Monospace => "Monospace",
        }
    }

    pub const fn family(self) -> FontFamily {
        match self {
            Self::Proportional => FontFamily::Proportional,
            Self::Monospace => FontFamily::Monospace,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
pub enum ParagraphAlignment {
    Left,
    Center,
    Right,
    Justify,
}

impl ParagraphAlignment {
    pub const ALL: [Self; 4] = [Self::Left, Self::Center, Self::Right, Self::Justify];

    pub const fn label(self) -> &'static str {
        match self {
            Self::Left => "Left",
            Self::Center => "Center",
            Self::Right => "Right",
            Self::Justify => "Justify",
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
pub enum ListKind {
    None,
    Bullet,
    Ordered,
}

impl ListKind {
    pub const fn label(self) -> &'static str {
        match self {
            Self::None => "None",
            Self::Bullet => "Bullets",
            Self::Ordered => "Numbering",
        }
    }
}

fn serialize_color32<S: serde::Serializer>(color: &Color32, s: S) -> Result<S::Ok, S::Error> {
    s.serialize_str(&format!(
        "#{:02x}{:02x}{:02x}{:02x}",
        color.r(),
        color.g(),
        color.b(),
        color.a()
    ))
}

#[derive(Clone, Copy, Debug, PartialEq, Serialize)]
pub struct CharacterStyle {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub font_size_points: f32,
    pub font_choice: FontChoice,
    pub font_family_name: Option<&'static str>,
    #[serde(serialize_with = "serialize_color32")]
    pub text_color: Color32,
    #[serde(serialize_with = "serialize_color32")]
    pub highlight_color: Color32,
}

impl Default for CharacterStyle {
    fn default() -> Self {
        Self {
            bold: false,
            italic: false,
            underline: false,
            strikethrough: false,
            font_size_points: 12.0,
            font_choice: FontChoice::Proportional,
            font_family_name: None,
            text_color: Color32::from_rgb(36, 39, 46),
            highlight_color: Color32::TRANSPARENT,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Serialize)]
pub enum LineSpacingKind {
    AutoMultiplier,
    AtLeastPoints,
    ExactPoints,
}

#[derive(Clone, Copy, Debug, PartialEq, Serialize)]
pub struct LineSpacing {
    pub kind: LineSpacingKind,
    pub value: f32,
}

impl Default for LineSpacing {
    fn default() -> Self {
        Self {
            kind: LineSpacingKind::AutoMultiplier,
            value: 1.0,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Serialize)]
pub struct ParagraphStyle {
    pub alignment: ParagraphAlignment,
    pub list_kind: ListKind,
    pub page_break_before: bool,
    pub spacing_before_points: u16,
    pub spacing_after_points: u16,
    pub line_spacing: LineSpacing,
}

impl Default for ParagraphStyle {
    fn default() -> Self {
        Self {
            alignment: ParagraphAlignment::Left,
            list_kind: ListKind::None,
            page_break_before: false,
            spacing_before_points: 0,
            spacing_after_points: 0,
            line_spacing: LineSpacing::default(),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct TextRun {
    pub text: String,
    pub style: CharacterStyle,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
pub enum WrapMode {
    Inline,
    Square,
    Tight,
    Through,
    TopAndBottom,
    BehindText,
    InFrontOfText,
}

impl WrapMode {
    pub const ALL: [Self; 7] = [
        Self::Inline,
        Self::Square,
        Self::Tight,
        Self::Through,
        Self::TopAndBottom,
        Self::BehindText,
        Self::InFrontOfText,
    ];

    pub const fn label(self) -> &'static str {
        match self {
            Self::Inline => "Inline",
            Self::Square => "Square",
            Self::Tight => "Tight",
            Self::Through => "Through",
            Self::TopAndBottom => "Top & Bottom",
            Self::BehindText => "Behind Text",
            Self::InFrontOfText => "In Front",
        }
    }

    /// Returns true if this wrap mode is a floating mode (not inline).
    pub const fn is_floating(self) -> bool {
        !matches!(self, Self::Inline)
    }

    /// Returns true if text layout should not be affected by this image.
    pub const fn is_no_text_displacement(self) -> bool {
        matches!(self, Self::BehindText | Self::InFrontOfText)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
pub enum ImageRendering {
    Smooth,
    Crisp,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
pub enum ImageLayoutMode {
    Inline,
    Floating,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
pub enum HorizontalRelativeTo {
    Page,
    Margin,
    Column,
    Character,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
pub enum VerticalRelativeTo {
    Page,
    Margin,
    Paragraph,
    Line,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Serialize)]
pub enum PositionAlign {
    Start,
    Center,
    End,
}

#[derive(Clone, Copy, Debug, PartialEq, Serialize)]
pub struct HorizontalPosition {
    pub relative_to: HorizontalRelativeTo,
    pub align: Option<PositionAlign>,
    pub offset_points: f32,
}

impl Default for HorizontalPosition {
    fn default() -> Self {
        Self {
            relative_to: HorizontalRelativeTo::Column,
            align: Some(PositionAlign::Start),
            offset_points: 0.0,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Serialize)]
pub struct VerticalPosition {
    pub relative_to: VerticalRelativeTo,
    pub align: Option<PositionAlign>,
    pub offset_points: f32,
}

impl Default for VerticalPosition {
    fn default() -> Self {
        Self {
            relative_to: VerticalRelativeTo::Paragraph,
            align: Some(PositionAlign::Start),
            offset_points: 0.0,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Serialize)]
pub struct DistanceFromText {
    pub top_points: f32,
    pub right_points: f32,
    pub bottom_points: f32,
    pub left_points: f32,
}

impl Default for DistanceFromText {
    fn default() -> Self {
        Self {
            top_points: 0.0,
            right_points: 8.0,
            bottom_points: 0.0,
            left_points: 8.0,
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct DocumentImage {
    pub id: usize,
    #[serde(skip)]
    pub bytes: Vec<u8>,
    pub alt_text: String,
    pub width_points: f32,
    pub height_points: f32,
    pub lock_aspect_ratio: bool,
    pub opacity: f32,
    pub layout_mode: ImageLayoutMode,
    pub wrap_mode: WrapMode,
    pub rendering: ImageRendering,
    pub horizontal_position: HorizontalPosition,
    pub vertical_position: VerticalPosition,
    pub distance_from_text: DistanceFromText,
    pub z_index: i32,
    pub move_with_text: bool,
    pub allow_overlap: bool,
}

impl DocumentImage {
    pub fn offset_x_points(&self) -> f32 {
        self.horizontal_position.offset_points
    }

    pub fn offset_y_points(&self) -> f32 {
        self.vertical_position.offset_points
    }

    pub fn set_manual_offset(&mut self, x_points: f32, y_points: f32) {
        self.horizontal_position.align = None;
        self.vertical_position.align = None;
        self.horizontal_position.offset_points = x_points;
        self.vertical_position.offset_points = y_points;
    }

    pub fn adjust_manual_offset(&mut self, dx: f32, dy: f32) {
        self.set_manual_offset(
            self.horizontal_position.offset_points + dx,
            self.vertical_position.offset_points + dy,
        );
    }
}

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

    fn total_chars(&self) -> usize {
        self.runs.iter().map(|run| run.text.chars().count()).sum()
    }

    fn typing_style(&self) -> CharacterStyle {
        self.runs.last().map(|run| run.style).unwrap_or_default()
    }

    fn style_at(&self, char_index: usize) -> CharacterStyle {
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

    fn append_text(&mut self, text: &str, style: CharacterStyle) {
        self.insert_text(self.total_chars(), text, style);
    }

    fn apply_style(&mut self, mutate: impl Fn(&mut CharacterStyle) + Copy) {
        for run in &mut self.runs {
            mutate(&mut run.style);
        }
        self.normalize_runs();
    }

    fn apply_style_to_range(&mut self, range: Range<usize>, mutate: impl Fn(&mut CharacterStyle)) {
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

    fn insert_text(&mut self, char_index: usize, text: &str, style: CharacterStyle) {
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

    fn replace_range_with_text(
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

    fn delete_char_range(&mut self, range: Range<usize>) {
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

    fn split_at_char(&mut self, char_index: usize) {
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

    fn normalize_runs(&mut self) {
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
    #[serde(serialize_with = "serialize_color32")]
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

#[derive(Clone, Debug, PartialEq)]
pub struct Paragraph {
    pub index: usize,
    pub range: Range<usize>,
    pub style: ParagraphStyle,
    pub runs: Vec<TextRun>,
    pub list_marker: Option<String>,
    pub image: Option<DocumentImage>,
    pub table: Option<DocumentTable>,
}

#[derive(Clone, Copy, Debug, Serialize)]
pub struct PageSize {
    pub width_points: f32,
    pub height_points: f32,
}

#[derive(Clone, Copy, Debug, Serialize)]
pub struct PageMargins {
    pub top_points: f32,
    pub right_points: f32,
    pub bottom_points: f32,
    pub left_points: f32,
}

#[derive(Clone)]
pub struct DocumentState {
    pub title: String,
    pub runs: Vec<TextRun>,
    pub paragraph_styles: Vec<ParagraphStyle>,
    pub paragraph_images: Vec<Option<DocumentImage>>,
    pub paragraph_tables: Vec<Option<DocumentTable>>,
    pub page_size: PageSize,
    pub margins: PageMargins,
}

impl DocumentState {
    pub fn bootstrap() -> Self {
        Self {
            title: "Untitled".to_owned(),
            runs: vec![
                TextRun {
                    text: "wors".to_owned(),
                    style: CharacterStyle {
                        bold: true,
                        font_size_points: 22.0,
                        ..CharacterStyle::default()
                    },
                },
                TextRun {
                    text: " now edits text on a custom painter-backed page.\n\n".to_owned(),
                    style: CharacterStyle {
                        font_size_points: 13.0,
                        ..CharacterStyle::default()
                    },
                },
                TextRun {
                    text: "Use the ribbon above to change".to_owned(),
                    style: CharacterStyle {
                        bold: true,
                        ..CharacterStyle::default()
                    },
                },
                TextRun {
                    text: " bold, italic, underline, strike-through, text size, font family, text color, and highlight.".to_owned(),
                    style: CharacterStyle::default(),
                },
            ],
            paragraph_styles: vec![ParagraphStyle::default(); 3],
            paragraph_images: vec![None; 3],
            paragraph_tables: vec![None; 3],
            page_size: PageSize::a4(),
            margins: PageMargins::standard(),
        }
    }

    pub fn plain_text(&self) -> String {
        self.runs.iter().map(|run| run.text.as_str()).collect()
    }

    pub fn total_chars(&self) -> usize {
        self.runs.iter().map(|run| run.text.chars().count()).sum()
    }

    pub fn paragraph_count(&self) -> usize {
        self.plain_text().chars().filter(|ch| *ch == '\n').count() + 1
    }

    pub fn style_at(&self, char_index: usize) -> CharacterStyle {
        if self.runs.is_empty() {
            return CharacterStyle::default();
        }

        let mut offset = 0;
        for run in &self.runs {
            let run_chars = run.text.chars().count();
            if char_index < offset + run_chars {
                return run.style;
            }
            offset += run_chars;
        }

        self.runs.last().map(|run| run.style).unwrap_or_default()
    }

    pub fn line_range_at(&self, char_index: usize) -> Range<usize> {
        line_char_range(&self.plain_text(), char_index.min(self.total_chars()))
    }

    pub fn word_range_at(&self, char_index: usize) -> Option<Range<usize>> {
        word_char_range(&self.plain_text(), char_index.min(self.total_chars()))
    }

    pub fn typing_style_at(&self, char_index: usize) -> CharacterStyle {
        let cursor_index = char_index.min(self.total_chars());
        let line_range = self.line_range_at(cursor_index);

        if line_range.start == line_range.end {
            return CharacterStyle::default();
        }

        self.style_at(cursor_index)
    }

    pub fn paragraph_style_at(&self, char_index: usize) -> ParagraphStyle {
        let paragraph_index = self.paragraph_index_at(char_index);
        self.paragraph_styles
            .get(paragraph_index)
            .copied()
            .unwrap_or_default()
    }

    pub fn selected_text(&self, range: Range<usize>) -> String {
        let text = self.plain_text();
        slice_char_range(&text, range).to_owned()
    }

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

    pub fn load_from_path(path: &Path) -> Result<Self, String> {
        let title = path
            .file_stem()
            .and_then(|name| name.to_str())
            .unwrap_or("Imported")
            .to_owned();

        let extension = path
            .extension()
            .and_then(|ext| ext.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();

        let runs = match extension.as_str() {
            "docx" => {
                let imported = docx_to_document(
                    &fs::read(path)
                        .map_err(|error| format!("failed to read {}: {error}", path.display()))?,
                )?;

                let mut document = Self::bootstrap();
                document.title = title;
                document.runs = imported.runs;
                document.paragraph_styles = imported.paragraph_styles;
                document.paragraph_images = imported.paragraph_images;
                document.paragraph_tables = imported.paragraph_tables;
                if let Some(page_size) = imported.page_size {
                    document.page_size = page_size;
                }
                if let Some(margins) = imported.margins {
                    document.margins = margins;
                }
                document.normalize_runs();
                document.ensure_paragraph_style_count();
                return Ok(document);
            }
            "md" | "markdown" => {
                let source = fs::read_to_string(path)
                    .map_err(|error| format!("failed to read {}: {error}", path.display()))?;
                markdown_to_runs(&source)
            }
            _ => {
                let source = fs::read_to_string(path)
                    .map_err(|error| format!("failed to read {}: {error}", path.display()))?;
                vec![TextRun {
                    text: source,
                    style: CharacterStyle::default(),
                }]
            }
        };

        let mut document = Self::bootstrap();
        document.replace_with_runs(title, runs);
        Ok(document)
    }

    pub fn save_to_path(&self, path: &Path) -> Result<(), String> {
        let extension = path
            .extension()
            .and_then(|ext| ext.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();

        match extension.as_str() {
            "md" | "markdown" => fs::write(path, self.to_markdown())
                .map_err(|error| format!("failed to save {}: {error}", path.display())),
            "txt" | "" => fs::write(path, self.to_plain_text_export())
                .map_err(|error| format!("failed to save {}: {error}", path.display())),
            "html" | "htm" => fs::write(path, self.to_html())
                .map_err(|error| format!("failed to save {}: {error}", path.display())),
            "pdf" => fs::write(path, self.to_pdf_bytes()?)
                .map_err(|error| format!("failed to save {}: {error}", path.display())),
            other => Err(format!(
                "saving .{other} is not supported yet; use .txt, .md, .html, or .pdf"
            )),
        }
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

    fn split_at_char(&mut self, char_index: usize) {
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

    fn image_by_id_mut(&mut self, id: usize) -> Option<&mut DocumentImage> {
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

    pub fn table_by_id_mut(&mut self, id: usize) -> Option<&mut DocumentTable> {
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

    fn normalize_runs(&mut self) {
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

    fn paragraph_index_at(&self, char_index: usize) -> usize {
        let target = char_index.min(self.total_chars());
        let mut paragraph_index = 0;
        let mut offset = 0;
        for run in &self.runs {
            for ch in run.text.chars() {
                if offset >= target {
                    return paragraph_index;
                }
                if ch == '\n' {
                    paragraph_index += 1;
                }
                offset += 1;
            }
        }
        paragraph_index
    }

    fn ensure_paragraph_style_count(&mut self) {
        let target = self.paragraph_count().max(1);
        self.paragraph_styles
            .resize(target, ParagraphStyle::default());
        self.paragraph_images.resize(target, None);
        self.paragraph_tables.resize(target, None);
    }

    fn to_plain_text_export(&self) -> String {
        self.paragraphs()
            .into_iter()
            .map(|paragraph| {
                if let Some(table) = &paragraph.table {
                    return table_to_plain_text(table);
                }
                let mut text = plain_text_from_runs(&paragraph.runs);
                text.retain(|ch| ch != OBJECT_REPLACEMENT_CHAR);
                if paragraph.style.page_break_before {
                    if text.is_empty() {
                        text.push('\u{000C}');
                    } else {
                        text.insert(0, '\u{000C}');
                    }
                }
                if let Some(marker) = paragraph.list_marker {
                    if text.is_empty() {
                        marker
                    } else {
                        text.insert_str(0, &format!("{marker} "));
                        text
                    }
                } else {
                    text
                }
            })
            .collect::<Vec<_>>()
            .join("\n")
    }

    fn to_markdown(&self) -> String {
        self.paragraphs()
            .into_iter()
            .map(|paragraph| {
                if let Some(table) = &paragraph.table {
                    return table_to_markdown(table);
                }
                let mut text = markdown_text_from_runs(&paragraph.runs);
                if paragraph.style.page_break_before {
                    let break_marker = "<div style=\"page-break-before: always\"></div>";
                    text = if text.is_empty() {
                        break_marker.to_owned()
                    } else {
                        format!("{break_marker}\n\n{text}")
                    };
                }
                if paragraph.image.is_some() {
                    let alt = paragraph
                        .image
                        .as_ref()
                        .map(|image| image.alt_text.as_str())
                        .filter(|alt| !alt.is_empty())
                        .unwrap_or("Image");
                    if text.is_empty() {
                        text = format!("![{alt}](embedded-image)");
                    } else {
                        text = format!("{text}\n\n![{alt}](embedded-image)");
                    }
                }
                if let Some(marker) = paragraph.list_marker.as_deref() {
                    let prefix = match paragraph.style.list_kind {
                        ListKind::Bullet => "- ".to_owned(),
                        ListKind::Ordered => format!("{marker} "),
                        ListKind::None => String::new(),
                    };
                    text = format!("{prefix}{text}");
                }

                match paragraph.style.alignment {
                    ParagraphAlignment::Left => text,
                    ParagraphAlignment::Center => format!("<div align=\"center\">{text}</div>"),
                    ParagraphAlignment::Right => format!("<div align=\"right\">{text}</div>"),
                    ParagraphAlignment::Justify => format!("<div align=\"justify\">{text}</div>"),
                }
            })
            .collect::<Vec<_>>()
            .join("\n")
    }

    fn to_html(&self) -> String {
        let mut html = String::new();
        let _ = write!(
            html,
            "<!doctype html>\
<html lang=\"en\">\
<head>\
<meta charset=\"utf-8\">\
<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\
<title>{}</title>\
<style>\
body {{ margin: 0; padding: 18pt; background: #e7ebf0; }}\
.page {{ box-sizing: border-box; margin: 0 auto; width: {}pt; min-height: {}pt; padding: {}pt {}pt {}pt {}pt; background: #ffffff; color: #24272e; box-shadow: 0 1px 5px rgba(0, 0, 0, 0.18); }}\
.paragraph {{ margin: 0; white-space: pre-wrap; }}\
.page-break {{ break-before: page; page-break-before: always; height: 0; }}\
.image-block {{ display: block; max-width: 100%; }}\
@media print {{ body {{ background: transparent; padding: 0; }} .page {{ box-shadow: none; width: auto; min-height: auto; }} }}\
</style>\
</head>\
<body>\
<div class=\"page\">",
            html_escape(&self.title),
            self.page_size.width_points,
            self.page_size.height_points,
            self.margins.top_points,
            self.margins.right_points,
            self.margins.bottom_points,
            self.margins.left_points
        );

        for paragraph in self.paragraphs() {
            if paragraph.style.page_break_before {
                html.push_str("<div class=\"page-break\"></div>");
            }

            if let Some(table) = &paragraph.table {
                html.push_str(&table_to_html(table));
                continue;
            }

            let _ = write!(
                html,
                "<p class=\"paragraph\" style=\"text-align:{};margin-top:{}pt;margin-bottom:{}pt;{}\">",
                paragraph_alignment_css(paragraph.style.alignment),
                paragraph.style.spacing_before_points,
                paragraph.style.spacing_after_points,
                line_spacing_css(paragraph.style.line_spacing)
            );

            if let Some(marker) = paragraph.list_marker {
                let prefix = match paragraph.style.list_kind {
                    ListKind::Bullet | ListKind::Ordered => format!("{marker} "),
                    ListKind::None => String::new(),
                };
                html.push_str(&html_escape(&prefix));
            }

            for run in paragraph.runs {
                let text: String = run
                    .text
                    .chars()
                    .filter(|ch| *ch != OBJECT_REPLACEMENT_CHAR)
                    .collect();
                if text.is_empty() {
                    continue;
                }

                let _ = write!(
                    html,
                    "<span style=\"{}\">{}</span>",
                    run_style_css(run.style),
                    html_escape(&text)
                );
            }

            if let Some(image) = paragraph.image.as_ref() {
                if let Some(mime_type) = image_mime_type(&image.bytes) {
                    let _ = write!(
                        html,
                        "<img class=\"image-block\" alt=\"{}\" src=\"data:{};base64,{}\" style=\"width:{}pt;height:{}pt;opacity:{:.3};{}\" />",
                        html_escape(&image.alt_text),
                        mime_type,
                        BASE64_STANDARD.encode(&image.bytes),
                        image.width_points,
                        image.height_points,
                        image.opacity.clamp(0.0, 1.0),
                        image_position_css(image)
                    );
                }
            }

            html.push_str("</p>");
        }

        html.push_str("</div></body></html>");
        html
    }

    fn to_pdf_bytes(&self) -> Result<Vec<u8>, String> {
        let html = self.to_pdf_html();
        let options = GeneratePdfOptions {
            page_width: Some(points_to_mm(self.page_size.width_points)),
            page_height: Some(points_to_mm(self.page_size.height_points)),
            margin_top: Some(points_to_mm(self.margins.top_points)),
            margin_right: Some(points_to_mm(self.margins.right_points)),
            margin_bottom: Some(points_to_mm(self.margins.bottom_points)),
            margin_left: Some(points_to_mm(self.margins.left_points)),
            ..GeneratePdfOptions::default()
        };
        let images: BTreeMap<String, Base64OrRaw> = BTreeMap::new();
        let fonts: BTreeMap<String, Base64OrRaw> = BTreeMap::new();

        let mut warnings = Vec::new();
        let mut rendered = PdfDocument::from_html(&html, &images, &fonts, &options, &mut warnings)
            .map_err(|error| format!("failed to render PDF: {error}"))?;
        rendered.metadata.info.document_title = self.title.clone();
        rendered.metadata.info.conformance = Default::default();
        if rendered.pages.is_empty() {
            return Ok(self.to_plain_text_pdf_bytes());
        }

        Ok(rendered.save(&PdfSaveOptions::default(), &mut warnings))
    }

    fn to_pdf_html(&self) -> String {
        let mut html = String::new();
        let _ = write!(
            html,
            "<!doctype html>\
<html lang=\"en\">\
<head>\
<meta charset=\"utf-8\">\
<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\
<title>{}</title>\
<style>\
body {{ margin: 0; padding: 0; color: #24272e; font-family: Helvetica, Arial, sans-serif; }}\
p {{ margin: 0; white-space: pre-wrap; }}\
.page-break {{ break-before: page; page-break-before: always; height: 0; }}\
.image-block {{ display: block; max-width: 100%; }}\
</style>\
</head>\
<body>",
            html_escape(&self.title)
        );

        for paragraph in self.paragraphs() {
            if paragraph.style.page_break_before {
                html.push_str("<div class=\"page-break\"></div>");
            }

            let _ = write!(
                html,
                "<p style=\"text-align:{};margin-top:{:.2}px;margin-bottom:{:.2}px;{}\">",
                paragraph_alignment_css(paragraph.style.alignment),
                points_to_css_px(paragraph.style.spacing_before_points as f32),
                points_to_css_px(paragraph.style.spacing_after_points as f32),
                line_spacing_css_pdf(paragraph.style.line_spacing)
            );

            if let Some(marker) = paragraph.list_marker {
                let prefix = match paragraph.style.list_kind {
                    ListKind::Bullet | ListKind::Ordered => format!("{marker} "),
                    ListKind::None => String::new(),
                };
                html.push_str(&html_escape(&prefix));
            }

            for run in paragraph.runs {
                let text: String = run
                    .text
                    .chars()
                    .filter(|ch| *ch != OBJECT_REPLACEMENT_CHAR)
                    .collect();
                if text.is_empty() {
                    continue;
                }

                let escaped = html_escape(&text);
                let mut run_html = format!(
                    "<span style=\"{}\">{escaped}</span>",
                    run_style_css_pdf(run.style)
                );
                if run.style.bold {
                    run_html = format!("<strong>{run_html}</strong>");
                }
                if run.style.italic {
                    run_html = format!("<em>{run_html}</em>");
                }
                if run.style.underline {
                    run_html =
                        format!("<span style=\"text-decoration:underline;\">{run_html}</span>");
                }
                if run.style.strikethrough {
                    run_html =
                        format!("<span style=\"text-decoration:line-through;\">{run_html}</span>");
                }
                html.push_str(&run_html);
            }

            if let Some(image) = paragraph.image.as_ref() {
                if let Some(mime_type) = image_mime_type(&image.bytes) {
                    let _ = write!(
                        html,
                        "<img class=\"image-block\" alt=\"{}\" src=\"data:{};base64,{}\" style=\"width:{:.2}px;height:{:.2}px;opacity:{:.3};{}\" />",
                        html_escape(&image.alt_text),
                        mime_type,
                        BASE64_STANDARD.encode(&image.bytes),
                        points_to_css_px(image.width_points),
                        points_to_css_px(image.height_points),
                        image.opacity.clamp(0.0, 1.0),
                        image_position_css_pdf(image)
                    );
                }
            }

            html.push_str("</p>");
        }

        html.push_str("</body></html>");
        html
    }

    fn to_plain_text_pdf_bytes(&self) -> Vec<u8> {
        let page_width_mm = points_to_mm(self.page_size.width_points);
        let page_height_mm = points_to_mm(self.page_size.height_points);
        let left = self.margins.left_points.max(18.0);
        let top = self.margins.top_points.max(18.0);
        let bottom = self.margins.bottom_points.max(18.0);

        let font_size = 11.0_f32;
        let line_height = 14.0_f32;
        let max_lines =
            (((self.page_size.height_points - top - bottom) / line_height).floor() as usize).max(1);

        let mut logical_lines = Vec::new();
        for line in self
            .to_plain_text_export()
            .replace('\u{000C}', "\n\n----- Page Break -----\n\n")
            .lines()
        {
            let wrapped = wrap_text_for_pdf(line, 100);
            if wrapped.is_empty() {
                logical_lines.push(String::new());
            } else {
                logical_lines.extend(wrapped);
            }
        }
        if logical_lines.is_empty() {
            logical_lines.push(String::new());
        }

        let mut pages = Vec::new();
        for chunk in logical_lines.chunks(max_lines) {
            let mut y = self.page_size.height_points - top - font_size;
            let mut ops = vec![
                Op::StartTextSection,
                Op::SetFont {
                    font: PdfFontHandle::Builtin(BuiltinFont::Helvetica),
                    size: Pt(font_size),
                },
                Op::SetLineHeight {
                    lh: Pt(line_height),
                },
                Op::SetTextCursor {
                    pos: Point {
                        x: Pt(left),
                        y: Pt(y),
                    },
                },
            ];

            for (i, line) in chunk.iter().enumerate() {
                ops.push(Op::ShowText {
                    items: vec![TextItem::Text(line.clone())],
                });
                if i + 1 < chunk.len() {
                    ops.push(Op::AddLineBreak);
                    y -= line_height;
                    if y <= bottom {
                        break;
                    }
                }
            }

            ops.push(Op::EndTextSection);
            pages.push(PdfPage::new(
                printpdf::Mm(page_width_mm),
                printpdf::Mm(page_height_mm),
                ops,
            ));
        }

        let mut document = PdfDocument::new(&self.title);
        let document = document.with_pages(pages);
        let mut warnings = Vec::new();
        document.save(&PdfSaveOptions::default(), &mut warnings)
    }
}

impl PageSize {
    pub const fn a4() -> Self {
        Self {
            width_points: 595.0,
            height_points: 842.0,
        }
    }
}

impl PageMargins {
    pub const fn standard() -> Self {
        Self {
            top_points: 72.0,
            right_points: 72.0,
            bottom_points: 72.0,
            left_points: 72.0,
        }
    }
}

pub(crate) fn text_format(style: CharacterStyle, zoom: f32) -> TextFormat {
    let mut coords = VariationCoords::default();
    if style.bold {
        coords.push("wght", 700.0);
    }

    let line_color = style.text_color;
    let font_size = if style.bold {
        (style.font_size_points + 0.8) * zoom
    } else {
        style.font_size_points * zoom
    };

    let family = match style.font_family_name {
        Some(name) => FontFamily::Name(name.into()),
        None => style.font_choice.family(),
    };

    TextFormat {
        font_id: FontId::new(font_size, family),
        color: if style.bold {
            style.text_color.gamma_multiply(0.88)
        } else {
            style.text_color
        },
        background: style.highlight_color,
        italics: style.italic,
        underline: if style.underline {
            Stroke::new(1.0, line_color)
        } else {
            Stroke::NONE
        },
        strikethrough: if style.strikethrough {
            Stroke::new(1.0, line_color)
        } else {
            Stroke::NONE
        },
        coords,
        ..Default::default()
    }
}

fn append_text_run(runs: &mut Vec<TextRun>, text: &str, style: CharacterStyle) {
    if text.is_empty() {
        return;
    }

    if let Some(last) = runs.last_mut() {
        if last.style == style {
            last.text.push_str(text);
            return;
        }
    }

    runs.push(TextRun {
        text: text.to_owned(),
        style,
    });
}

fn plain_text_from_runs(runs: &[TextRun]) -> String {
    runs.iter().map(|run| run.text.as_str()).collect()
}

fn table_to_plain_text(table: &DocumentTable) -> String {
    let mut lines = Vec::new();
    for row in &table.rows {
        let cells: Vec<String> = row
            .iter()
            .map(|cell| {
                cell.plain_text()
                    .chars()
                    .filter(|ch| *ch != OBJECT_REPLACEMENT_CHAR)
                    .collect()
            })
            .collect();
        lines.push(cells.join("\t"));
    }
    lines.join("\n")
}

fn table_to_markdown(table: &DocumentTable) -> String {
    if table.rows.is_empty() {
        return String::new();
    }
    let num_cols = table.num_cols();
    let mut lines = Vec::new();

    // Header row (first row)
    let header: Vec<String> = table
        .rows
        .first()
        .map(|row| row.iter().map(markdown_text_from_table_cell).collect())
        .unwrap_or_default();
    lines.push(format!("| {} |", header.join(" | ")));

    // Separator
    let separator: Vec<&str> = (0..num_cols).map(|_| "---").collect();
    lines.push(format!("| {} |", separator.join(" | ")));

    // Data rows
    for row in table.rows.iter().skip(1) {
        let cells: Vec<String> = row.iter().map(markdown_text_from_table_cell).collect();
        // Pad if row has fewer cells
        let mut padded = cells;
        while padded.len() < num_cols {
            padded.push(String::new());
        }
        lines.push(format!("| {} |", padded.join(" | ")));
    }

    lines.join("\n")
}

fn table_to_html(table: &DocumentTable) -> String {
    let mut html = String::new();
    let border_color = css_color(table.borders.color);
    let border_width = table.borders.width_points;
    let _ = write!(
        html,
        "<table style=\"border-collapse:collapse;margin:4pt 0;width:100%;\">"
    );
    for (row_idx, row) in table.rows.iter().enumerate() {
        html.push_str("<tr>");
        for (col_idx, cell) in row.iter().enumerate() {
            let col_width = table
                .col_widths_points
                .get(col_idx)
                .copied()
                .unwrap_or(72.0);
            let tag = if row_idx == 0 { "th" } else { "td" };
            let _ = write!(
                html,
                "<{tag} style=\"border:{border_width:.1}pt solid {border_color};padding:4pt 6pt;width:{col_width:.1}pt;\">"
            );
            for run in &cell.runs {
                let text: String = run
                    .text
                    .chars()
                    .filter(|ch| *ch != OBJECT_REPLACEMENT_CHAR)
                    .collect();
                if !text.is_empty() {
                    let _ = write!(
                        html,
                        "<span style=\"{}\">{}</span>",
                        run_style_css(run.style),
                        html_escape(&text)
                    );
                }
            }
            for image in &cell.images {
                if let Some(mime_type) = image_mime_type(&image.bytes) {
                    let _ = write!(
                        html,
                        "<img class=\"image-block\" alt=\"{}\" src=\"data:{};base64,{}\" style=\"width:{}pt;height:{}pt;opacity:{:.3};\" />",
                        html_escape(&image.alt_text),
                        mime_type,
                        BASE64_STANDARD.encode(&image.bytes),
                        image.width_points,
                        image.height_points,
                        image.opacity.clamp(0.0, 1.0)
                    );
                }
            }
            let _ = write!(html, "</{tag}>");
        }
        html.push_str("</tr>");
    }
    html.push_str("</table>");
    html
}

fn markdown_text_from_table_cell(cell: &TableCell) -> String {
    let mut text = markdown_text_from_runs(&cell.runs);
    for image in &cell.images {
        let alt = if image.alt_text.is_empty() {
            "Image"
        } else {
            &image.alt_text
        };
        if !text.is_empty() {
            text.push(' ');
        }
        text.push_str(&format!("![{alt}](embedded-image)"));
    }
    text
}

fn markdown_text_from_runs(runs: &[TextRun]) -> String {
    let mut output = String::new();
    for run in runs {
        let mut text: String = run
            .text
            .chars()
            .filter(|ch| *ch != OBJECT_REPLACEMENT_CHAR)
            .collect();
        if text.is_empty() {
            continue;
        }
        if run.style.font_choice == FontChoice::Monospace {
            text = format!("`{text}`");
        }
        if run.style.bold {
            text = format!("**{text}**");
        }
        if run.style.italic {
            text = format!("*{text}*");
        }
        if run.style.strikethrough {
            text = format!("~~{text}~~");
        }
        if run.style.underline {
            text = format!("<u>{text}</u>");
        }
        output.push_str(&text);
    }
    output
}

fn points_to_mm(points: f32) -> f32 {
    points * (25.4 / 72.0)
}

fn points_to_css_px(points: f32) -> f32 {
    points * (96.0 / 72.0)
}

fn paragraph_alignment_css(alignment: ParagraphAlignment) -> &'static str {
    match alignment {
        ParagraphAlignment::Left => "left",
        ParagraphAlignment::Center => "center",
        ParagraphAlignment::Right => "right",
        ParagraphAlignment::Justify => "justify",
    }
}

fn line_spacing_css(line_spacing: LineSpacing) -> String {
    match line_spacing.kind {
        LineSpacingKind::AutoMultiplier => {
            format!("line-height:{:.3};", line_spacing.value.max(0.1))
        }
        LineSpacingKind::AtLeastPoints | LineSpacingKind::ExactPoints => {
            format!("line-height:{:.3}pt;", line_spacing.value.max(1.0))
        }
    }
}

fn line_spacing_css_pdf(line_spacing: LineSpacing) -> String {
    match line_spacing.kind {
        LineSpacingKind::AutoMultiplier => {
            format!("line-height:{:.3};", line_spacing.value.max(0.1))
        }
        LineSpacingKind::AtLeastPoints | LineSpacingKind::ExactPoints => {
            format!(
                "line-height:{:.2}px;",
                points_to_css_px(line_spacing.value.max(1.0))
            )
        }
    }
}

fn run_style_css(style: CharacterStyle) -> String {
    let mut css = format!(
        "font-family:{};font-size:{:.2}pt;color:{};",
        css_font_family(style),
        style.font_size_points.max(1.0),
        css_color(style.text_color)
    );
    if style.bold {
        css.push_str("font-weight:700;");
    }
    if style.italic {
        css.push_str("font-style:italic;");
    }
    if style.highlight_color != Color32::TRANSPARENT {
        let _ = write!(
            css,
            "background-color:{};",
            css_color(style.highlight_color)
        );
    }
    let decoration = text_decoration_css(style);
    if !decoration.is_empty() {
        let _ = write!(css, "text-decoration:{};", decoration);
    }
    css
}

fn run_style_css_pdf(style: CharacterStyle) -> String {
    let font_points = if style.bold {
        style.font_size_points + 0.8
    } else {
        style.font_size_points
    };

    let mut css = format!(
        "font-family:{};font-size:{:.2}px;color:{};",
        css_font_family(style),
        points_to_css_px(font_points.max(1.0)),
        css_color_rgb(style.text_color)
    );
    if style.italic {
        css.push_str("font-style:italic;");
    }
    if style.highlight_color != Color32::TRANSPARENT {
        let _ = write!(
            css,
            "background-color:{};",
            css_color_rgb(style.highlight_color)
        );
    }
    let decoration = text_decoration_css(style);
    if !decoration.is_empty() {
        let _ = write!(css, "text-decoration:{};", decoration);
    }
    css
}

fn text_decoration_css(style: CharacterStyle) -> &'static str {
    match (style.underline, style.strikethrough) {
        (true, true) => "underline line-through",
        (true, false) => "underline",
        (false, true) => "line-through",
        (false, false) => "",
    }
}

fn css_font_family(style: CharacterStyle) -> String {
    match style.font_family_name {
        Some("docx-carlito") => "Carlito, Calibri, sans-serif".to_owned(),
        Some("docx-caladea") => "Caladea, Cambria, serif".to_owned(),
        Some("docx-liberation-sans") => "\"Liberation Sans\", Arial, sans-serif".to_owned(),
        Some("docx-liberation-serif") => {
            "\"Liberation Serif\", \"Times New Roman\", serif".to_owned()
        }
        Some("docx-liberation-mono") => {
            "\"Liberation Mono\", \"Courier New\", Consolas, monospace".to_owned()
        }
        Some(name) => format!("\"{}\", sans-serif", name.replace('"', "\\\"")),
        None => match style.font_choice {
            FontChoice::Proportional => "sans-serif".to_owned(),
            FontChoice::Monospace => "monospace".to_owned(),
        },
    }
}

fn css_color(color: Color32) -> String {
    format!(
        "rgba({}, {}, {}, {:.3})",
        color.r(),
        color.g(),
        color.b(),
        (color.a() as f32 / 255.0).clamp(0.0, 1.0)
    )
}

fn css_color_rgb(color: Color32) -> String {
    format!("rgb({}, {}, {})", color.r(), color.g(), color.b())
}

fn image_mime_type(bytes: &[u8]) -> Option<&'static str> {
    match image::guess_format(bytes) {
        Ok(image::ImageFormat::Png) => Some("image/png"),
        Ok(image::ImageFormat::Jpeg) => Some("image/jpeg"),
        Ok(image::ImageFormat::Gif) => Some("image/gif"),
        Ok(image::ImageFormat::Bmp) => Some("image/bmp"),
        _ => None,
    }
}

fn wrap_text_for_pdf(text: &str, max_chars: usize) -> Vec<String> {
    if text.is_empty() {
        return Vec::new();
    }

    let mut out = Vec::new();
    let mut current = String::new();
    for word in text.split_whitespace() {
        let current_len = current.chars().count();
        let word_len = word.chars().count();
        let projected = if current.is_empty() {
            word_len
        } else {
            current_len + 1 + word_len
        };

        if projected > max_chars && !current.is_empty() {
            out.push(std::mem::take(&mut current));
        }

        if !current.is_empty() {
            current.push(' ');
        }
        current.push_str(word);
    }

    if !current.is_empty() {
        out.push(current);
    }
    out
}

fn image_position_css(image: &DocumentImage) -> String {
    if image.layout_mode == ImageLayoutMode::Floating {
        format!(
            "position:relative;left:{:.2}pt;top:{:.2}pt;z-index:{};",
            image.offset_x_points(),
            image.offset_y_points(),
            image.z_index
        )
    } else {
        String::new()
    }
}

fn image_position_css_pdf(image: &DocumentImage) -> String {
    if image.layout_mode == ImageLayoutMode::Floating {
        format!(
            "position:relative;left:{:.2}px;top:{:.2}px;z-index:{};",
            points_to_css_px(image.offset_x_points()),
            points_to_css_px(image.offset_y_points()),
            image.z_index
        )
    } else {
        String::new()
    }
}

fn html_escape(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&#39;"),
            _ => out.push(ch),
        }
    }
    out
}

#[cfg(test)]
mod tests {
    use std::{
        fs,
        time::{SystemTime, UNIX_EPOCH},
    };

    use super::{
        plain_text_from_runs, CharacterStyle, DocumentImage, DocumentState, ImageLayoutMode,
        ImageRendering, ListKind, WrapMode, OBJECT_REPLACEMENT_CHAR,
    };

    fn test_image(id: usize) -> DocumentImage {
        DocumentImage {
            id,
            bytes: vec![1, 2, 3],
            alt_text: format!("image-{id}"),
            width_points: 120.0,
            height_points: 60.0,
            lock_aspect_ratio: true,
            opacity: 1.0,
            layout_mode: ImageLayoutMode::Inline,
            wrap_mode: WrapMode::Inline,
            rendering: ImageRendering::Smooth,
            horizontal_position: Default::default(),
            vertical_position: Default::default(),
            distance_from_text: Default::default(),
            z_index: 7,
            move_with_text: true,
            allow_overlap: false,
        }
    }

    #[test]
    fn inserts_page_break_between_split_paragraphs() {
        let mut document = DocumentState::bootstrap();
        document.replace_with_runs(
            "Test".to_owned(),
            vec![super::TextRun {
                text: "alpha beta".to_owned(),
                style: CharacterStyle::default(),
            }],
        );

        let cursor = document.insert_page_break(6);
        let paragraphs = document.paragraphs();

        assert_eq!(cursor, paragraphs[1].range.start);
        assert_eq!(paragraphs.len(), 2);
        assert_eq!(plain_text_from_runs(&paragraphs[0].runs), "alpha ");
        assert_eq!(plain_text_from_runs(&paragraphs[1].runs), "beta");
        assert!(paragraphs[1].style.page_break_before);
    }

    #[test]
    fn inserts_block_image_as_its_own_paragraph() {
        let mut document = DocumentState::bootstrap();
        document.replace_with_runs(
            "Test".to_owned(),
            vec![super::TextRun {
                text: "alpha beta".to_owned(),
                style: CharacterStyle::default(),
            }],
        );

        let cursor = document.insert_image(
            6,
            DocumentImage {
                id: 1,
                bytes: vec![1, 2, 3],
                alt_text: "diagram".to_owned(),
                width_points: 120.0,
                height_points: 60.0,
                lock_aspect_ratio: true,
                opacity: 1.0,
                layout_mode: ImageLayoutMode::Inline,
                wrap_mode: WrapMode::Inline,
                rendering: ImageRendering::Smooth,
                horizontal_position: Default::default(),
                vertical_position: Default::default(),
                distance_from_text: Default::default(),
                z_index: 0,
                move_with_text: true,
                allow_overlap: false,
            },
        );
        let paragraphs = document.paragraphs();

        assert_eq!(cursor, paragraphs[1].range.end);
        assert_eq!(paragraphs.len(), 3);
        assert_eq!(
            paragraphs[1]
                .image
                .as_ref()
                .map(|image| image.alt_text.as_str()),
            Some("diagram")
        );
        assert_eq!(paragraphs[1].style.list_kind, ListKind::None);
        assert_eq!(
            document.plain_text(),
            format!("alpha \n{OBJECT_REPLACEMENT_CHAR}\nbeta")
        );
    }

    #[test]
    fn formats_empty_table_cell_and_uses_style_for_inserted_text() {
        let mut document = DocumentState::bootstrap();
        document.replace_with_runs("Test".to_owned(), Vec::new());
        document.insert_table(0, 1, 1);
        let table_id = document
            .paragraph_tables
            .iter()
            .flatten()
            .next()
            .unwrap()
            .id;

        document.apply_style_to_table_cell(table_id, 0, 0, |style| {
            style.bold = true;
            style.font_size_points = 18.0;
        });
        let active_style = document.table_cell_typing_style(table_id, 0, 0).unwrap();
        document.append_table_cell_text(table_id, 0, 0, "Styled", active_style);

        let cell = &document.table_by_id(table_id).unwrap().rows[0][0];
        assert_eq!(cell.plain_text(), "Styled");
        assert!(cell.runs[0].style.bold);
        assert_eq!(cell.runs[0].style.font_size_points, 18.0);
    }

    #[test]
    fn inserts_image_into_table_cell() {
        let mut document = DocumentState::bootstrap();
        document.replace_with_runs("Test".to_owned(), Vec::new());
        document.insert_table(0, 1, 1);
        let table_id = document
            .paragraph_tables
            .iter()
            .flatten()
            .next()
            .unwrap()
            .id;

        document.insert_table_cell_image(table_id, 0, 0, test_image(42), CharacterStyle::default());

        let cell = &document.table_by_id(table_id).unwrap().rows[0][0];
        assert_eq!(cell.images.len(), 1);
        assert_eq!(cell.plain_text(), OBJECT_REPLACEMENT_CHAR.to_string());
        assert!(document
            .to_markdown()
            .contains("![image-42](embedded-image)"));
    }

    #[test]
    fn moves_image_paragraph_later_without_extra_blank_lines() {
        let image = test_image(7);
        let mut document = DocumentState {
            title: "Test".to_owned(),
            runs: vec![super::TextRun {
                text: format!("alpha\n{OBJECT_REPLACEMENT_CHAR}\nbeta\ngamma"),
                style: CharacterStyle::default(),
            }],
            paragraph_styles: vec![
                Default::default(),
                super::ParagraphStyle {
                    page_break_before: true,
                    ..Default::default()
                },
                Default::default(),
                Default::default(),
            ],
            paragraph_images: vec![None, Some(image.clone()), None, None],
            paragraph_tables: vec![None; 4],
            page_size: super::PageSize::a4(),
            margins: super::PageMargins::standard(),
        };

        let cursor = document
            .move_image_paragraph_to_cursor(7, document.total_chars())
            .expect("image should move");
        let paragraphs = document.paragraphs();

        assert_eq!(
            document.plain_text(),
            format!("alpha\nbeta\ngamma\n{OBJECT_REPLACEMENT_CHAR}")
        );
        assert_eq!(cursor, paragraphs[3].range.start);
        assert_eq!(paragraphs[3].image.as_ref().map(|image| image.id), Some(7));
        assert_eq!(paragraphs[3].image.as_ref().unwrap().z_index, image.z_index);
        assert!(paragraphs[3].style.page_break_before);
        assert_eq!(
            document
                .plain_text()
                .chars()
                .filter(|ch| *ch == '\n')
                .count(),
            3
        );
    }

    #[test]
    fn moves_image_paragraph_earlier_without_losing_metadata() {
        let mut image = test_image(8);
        image.layout_mode = ImageLayoutMode::Floating;
        image.wrap_mode = WrapMode::Square;
        image.horizontal_position.offset_points = 42.0;
        let mut document = DocumentState {
            title: "Test".to_owned(),
            runs: vec![super::TextRun {
                text: format!("alpha\nbeta\n{OBJECT_REPLACEMENT_CHAR}\ngamma"),
                style: CharacterStyle::default(),
            }],
            paragraph_styles: vec![
                Default::default(),
                Default::default(),
                Default::default(),
                Default::default(),
            ],
            paragraph_images: vec![None, None, Some(image.clone()), None],
            paragraph_tables: vec![None; 4],
            page_size: super::PageSize::a4(),
            margins: super::PageMargins::standard(),
        };

        let cursor = document
            .move_image_paragraph_to_cursor(8, 0)
            .expect("image should move");
        let paragraphs = document.paragraphs();

        assert_eq!(
            document.plain_text(),
            format!("{OBJECT_REPLACEMENT_CHAR}\nalpha\nbeta\ngamma")
        );
        assert_eq!(cursor, 0);
        let moved = paragraphs[0].image.as_ref().expect("moved image");
        assert_eq!(moved.id, 8);
        assert_eq!(moved.layout_mode, ImageLayoutMode::Floating);
        assert_eq!(moved.wrap_mode, WrapMode::Square);
        assert_eq!(moved.horizontal_position.offset_points, 42.0);
    }

    #[test]
    fn exports_html_with_styled_runs() {
        let mut document = DocumentState::bootstrap();
        document.replace_with_runs(
            "Styled".to_owned(),
            vec![
                super::TextRun {
                    text: "Bold".to_owned(),
                    style: CharacterStyle {
                        bold: true,
                        ..CharacterStyle::default()
                    },
                },
                super::TextRun {
                    text: " + ".to_owned(),
                    style: CharacterStyle::default(),
                },
                super::TextRun {
                    text: "Mono".to_owned(),
                    style: CharacterStyle {
                        font_choice: super::FontChoice::Monospace,
                        ..CharacterStyle::default()
                    },
                },
            ],
        );

        let html = document.to_html();
        assert!(html.contains("<!doctype html>"));
        assert!(html.contains("font-weight:700;"));
        assert!(html.contains("Bold"));
        assert!(html.contains("Mono"));
    }

    #[test]
    fn exports_pdf_html_with_pdf_friendly_css() {
        let mut document = DocumentState::bootstrap();
        document.replace_with_runs(
            "Styled".to_owned(),
            vec![super::TextRun {
                text: "Bold".to_owned(),
                style: CharacterStyle {
                    bold: true,
                    ..CharacterStyle::default()
                },
            }],
        );

        let html = document.to_pdf_html();
        assert!(html.contains("font-family: Helvetica, Arial, sans-serif"));
        assert!(html.contains("font-size:"));
        assert!(html.contains("px"));
        assert!(html.contains("<strong>"));
        assert!(!html.contains("box-shadow"));
    }

    #[test]
    fn saves_pdf_extension() {
        let mut path = std::env::temp_dir();
        let stamp = SystemTime::now()
            .duration_since(UNIX_EPOCH)
            .expect("clock")
            .as_nanos();
        path.push(format!("wors-export-{stamp}.pdf"));

        let document = DocumentState::bootstrap();
        document
            .save_to_path(&path)
            .expect("pdf save should succeed");

        let bytes = fs::read(&path).expect("pdf should be readable");
        assert!(bytes.starts_with(b"%PDF"));

        let _ = fs::remove_file(path);
    }
}
