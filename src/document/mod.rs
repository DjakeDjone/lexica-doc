pub mod docx;
mod export;
mod markdown;
mod odt;
mod text;

use std::ops::Range;

use eframe::egui::{epaint::text::TextFormat, Color32, FontFamily, FontId, Stroke};
use serde::Serialize;

use text::{char_to_byte_index, line_char_range, slice_char_range, word_char_range};

pub const OBJECT_REPLACEMENT_CHAR: char = '\u{fffc}';
pub(crate) const DOCX_BODY_BOLD: &str = "docx-body-bold";
pub(crate) const DOCX_MONOSPACE_BOLD: &str = "docx-monospace-bold";
pub(crate) const DOCX_CARLITO_BOLD: &str = "docx-carlito-bold";
pub(crate) const DOCX_CALADEA_BOLD: &str = "docx-caladea-bold";
pub(crate) const DOCX_LIBERATION_SANS_BOLD: &str = "docx-liberation-sans-bold";
pub(crate) const DOCX_LIBERATION_SERIF_BOLD: &str = "docx-liberation-serif-bold";
pub(crate) const DOCX_LIBERATION_MONO_BOLD: &str = "docx-liberation-mono-bold";
pub(crate) const DOCX_COMIC_SANS_BOLD: &str = "docx-comic-sans-bold";

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize)]
pub enum FontChoice {
    Proportional,
    Monospace,
    Carlito,
    Caladea,
    LiberationSans,
    LiberationSerif,
    LiberationMono,
    ComicSans,
}

impl FontChoice {
    pub const ALL: [Self; 8] = [
        Self::Proportional,
        Self::Carlito,
        Self::Caladea,
        Self::LiberationSans,
        Self::LiberationSerif,
        Self::ComicSans,
        Self::Monospace,
        Self::LiberationMono,
    ];

    pub const fn label(self) -> &'static str {
        match self {
            Self::Proportional => "Body",
            Self::Monospace => "Monospace",
            Self::Carlito => "Carlito",
            Self::Caladea => "Caladea",
            Self::LiberationSans => "Liberation Sans",
            Self::LiberationSerif => "Liberation Serif",
            Self::LiberationMono => "Liberation Mono",
            Self::ComicSans => "Comic Sans",
        }
    }

    pub fn family(self) -> FontFamily {
        match self {
            Self::Proportional => FontFamily::Proportional,
            Self::Monospace => FontFamily::Monospace,
            Self::Carlito => FontFamily::Name("docx-carlito".into()),
            Self::Caladea => FontFamily::Name("docx-caladea".into()),
            Self::LiberationSans => FontFamily::Name("docx-liberation-sans".into()),
            Self::LiberationSerif => FontFamily::Name("docx-liberation-serif".into()),
            Self::LiberationMono => FontFamily::Name("docx-liberation-mono".into()),
            Self::ComicSans => FontFamily::Name("docx-comic-sans".into()),
        }
    }

    pub fn bold_family(self) -> FontFamily {
        match self {
            Self::Proportional => FontFamily::Name(DOCX_BODY_BOLD.into()),
            Self::Monospace => FontFamily::Name(DOCX_MONOSPACE_BOLD.into()),
            Self::Carlito => FontFamily::Name(DOCX_CARLITO_BOLD.into()),
            Self::Caladea => FontFamily::Name(DOCX_CALADEA_BOLD.into()),
            Self::LiberationSans => FontFamily::Name(DOCX_LIBERATION_SANS_BOLD.into()),
            Self::LiberationSerif => FontFamily::Name(DOCX_LIBERATION_SERIF_BOLD.into()),
            Self::LiberationMono => FontFamily::Name(DOCX_LIBERATION_MONO_BOLD.into()),
            Self::ComicSans => FontFamily::Name(DOCX_COMIC_SANS_BOLD.into()),
        }
    }

    pub const fn family_name(self) -> Option<&'static str> {
        match self {
            Self::Proportional | Self::Monospace => None,
            Self::Carlito => Some("docx-carlito"),
            Self::Caladea => Some("docx-caladea"),
            Self::LiberationSans => Some("docx-liberation-sans"),
            Self::LiberationSerif => Some("docx-liberation-serif"),
            Self::LiberationMono => Some("docx-liberation-mono"),
            Self::ComicSans => Some("docx-comic-sans"),
        }
    }

    pub const fn is_monospace(self) -> bool {
        matches!(self, Self::Monospace | Self::LiberationMono)
    }

    pub fn from_family_name(name: &'static str) -> Option<Self> {
        match name {
            "docx-carlito" => Some(Self::Carlito),
            "docx-caladea" => Some(Self::Caladea),
            "docx-liberation-sans" => Some(Self::LiberationSans),
            "docx-liberation-serif" => Some(Self::LiberationSerif),
            "docx-liberation-mono" => Some(Self::LiberationMono),
            "docx-comic-sans" => Some(Self::ComicSans),
            _ => None,
        }
    }

    pub fn from_style(style: CharacterStyle) -> Self {
        style
            .font_family_name
            .and_then(Self::from_family_name)
            .unwrap_or(style.font_choice)
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize)]
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

    fn selection_style_at(&self, range: Range<usize>) -> CharacterStyle {
        let total_chars = self.total_chars();
        let start = range.start.min(total_chars);
        let end = range.end.min(total_chars);
        if start < end {
            return self.style_at(end - 1);
        }

        self.style_at(start)
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

    pub fn selection_style_at(&self, range: Range<usize>) -> CharacterStyle {
        let total_chars = self.total_chars();
        let start = range.start.min(total_chars);
        let end = range.end.min(total_chars);
        if start < end {
            return self.style_at(end - 1);
        }

        self.typing_style_at(start)
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
    let line_color = style.text_color;
    let font_size = style.font_size_points * zoom;
    let family = text_font_family(style);

    TextFormat {
        font_id: FontId::new(font_size, family),
        color: style.text_color,
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
        ..Default::default()
    }
}

fn text_font_family(style: CharacterStyle) -> FontFamily {
    if style.bold {
        return match style.font_family_name {
            Some("docx-carlito") => FontFamily::Name(DOCX_CARLITO_BOLD.into()),
            Some("docx-caladea") => FontFamily::Name(DOCX_CALADEA_BOLD.into()),
            Some("docx-liberation-sans") => FontFamily::Name(DOCX_LIBERATION_SANS_BOLD.into()),
            Some("docx-liberation-serif") => FontFamily::Name(DOCX_LIBERATION_SERIF_BOLD.into()),
            Some("docx-liberation-mono") => FontFamily::Name(DOCX_LIBERATION_MONO_BOLD.into()),
            Some("docx-comic-sans") => FontFamily::Name(DOCX_COMIC_SANS_BOLD.into()),
            Some(name) => FontFamily::Name(name.into()),
            None => style.font_choice.bold_family(),
        };
    }

    match style.font_family_name {
        Some(name) => FontFamily::Name(name.into()),
        None => style.font_choice.family(),
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

#[cfg(test)]
mod tests {
    use std::{
        fs,
        time::{SystemTime, UNIX_EPOCH},
    };

    use super::{
        export::plain_text_from_runs, text_format, CharacterStyle, DocumentImage, DocumentState,
        FontChoice, ImageLayoutMode, ImageRendering, ListKind, WrapMode, DOCX_BODY_BOLD,
        DOCX_CARLITO_BOLD, DOCX_LIBERATION_MONO_BOLD, OBJECT_REPLACEMENT_CHAR,
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
    fn bold_text_uses_registered_bold_font_faces() {
        let body = text_format(
            CharacterStyle {
                bold: true,
                ..CharacterStyle::default()
            },
            1.0,
        );
        assert_eq!(
            body.font_id.family,
            eframe::egui::FontFamily::Name(DOCX_BODY_BOLD.into())
        );
        assert_eq!(
            body.font_id.size,
            CharacterStyle::default().font_size_points
        );

        let monospace = text_format(
            CharacterStyle {
                bold: true,
                font_choice: FontChoice::LiberationMono,
                ..CharacterStyle::default()
            },
            1.0,
        );
        assert_eq!(
            monospace.font_id.family,
            eframe::egui::FontFamily::Name(DOCX_LIBERATION_MONO_BOLD.into())
        );

        let imported_family = text_format(
            CharacterStyle {
                bold: true,
                font_family_name: Some("docx-carlito"),
                ..CharacterStyle::default()
            },
            1.0,
        );
        assert_eq!(
            imported_family.font_id.family,
            eframe::egui::FontFamily::Name(DOCX_CARLITO_BOLD.into())
        );
    }

    #[test]
    fn selected_style_uses_last_selected_character_at_run_boundary() {
        let mut document = DocumentState::bootstrap();
        document.replace_with_runs(
            "Test".to_owned(),
            vec![
                super::TextRun {
                    text: "Bold".to_owned(),
                    style: CharacterStyle {
                        bold: true,
                        ..CharacterStyle::default()
                    },
                },
                super::TextRun {
                    text: " plain".to_owned(),
                    style: CharacterStyle::default(),
                },
            ],
        );

        assert!(document.selection_style_at(0..4).bold);
        assert!(!document.selection_style_at(0..5).bold);
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
