pub mod docx;
mod markdown;
mod text;

use std::{fs, ops::Range, path::Path};

use eframe::egui::{
    epaint::text::{TextFormat, VariationCoords},
    Color32, FontFamily, FontId, Stroke,
};
use serde::Serialize;

use docx::docx_to_document;
use markdown::{
    markdown_cursor_index_in_line, markdown_heading_prefix, markdown_line_replacement,
    markdown_to_runs,
};
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
    /// Horizontal position offset in document points (convenience accessor).
    pub fn offset_x_points(&self) -> f32 {
        self.horizontal_position.offset_points
    }

    /// Vertical position offset in document points (convenience accessor).
    pub fn offset_y_points(&self) -> f32 {
        self.vertical_position.offset_points
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
        let line_text = self.selected_text(line_range.clone());

        if let Some((_, style)) = markdown_heading_prefix(&line_text) {
            return style;
        }

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

    pub fn apply_markdown_shortcuts_at(&mut self, char_index: usize) -> usize {
        let cursor_index = char_index.min(self.total_chars());
        let line_range = self.line_range_at(cursor_index);
        let line_text = self.selected_text(line_range.clone());
        let cursor_in_line = cursor_index.saturating_sub(line_range.start);
        let line_start = line_range.start;
        let Some(replacement_runs) = markdown_line_replacement(&line_text) else {
            return cursor_index;
        };

        self.replace_range_with_runs(line_range, replacement_runs);
        let new_cursor_in_line = markdown_cursor_index_in_line(&line_text, cursor_in_line);
        line_start + new_cursor_in_line
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
        for slot in &mut self.paragraph_images {
            if let Some(image) = slot {
                if image.id == id {
                    image.width_points = width_points.max(24.0);
                    image.height_points = height_points.max(24.0);
                    return;
                }
            }
        }
    }

    pub fn set_image_offset_by_id(&mut self, id: usize, x_points: f32, y_points: f32) {
        for slot in &mut self.paragraph_images {
            if let Some(image) = slot {
                if image.id == id {
                    image.horizontal_position.offset_points = x_points;
                    image.vertical_position.offset_points = y_points;
                    return;
                }
            }
        }
    }

    pub fn set_image_layout_mode(&mut self, id: usize, mode: ImageLayoutMode) {
        for slot in &mut self.paragraph_images {
            if let Some(image) = slot {
                if image.id == id {
                    image.layout_mode = mode;
                    // When switching to inline, reset wrap mode
                    if mode == ImageLayoutMode::Inline {
                        image.wrap_mode = WrapMode::Inline;
                    } else if image.wrap_mode == WrapMode::Inline {
                        image.wrap_mode = WrapMode::Square;
                    }
                    return;
                }
            }
        }
    }

    pub fn set_image_horizontal_position(&mut self, id: usize, pos: HorizontalPosition) {
        for slot in &mut self.paragraph_images {
            if let Some(image) = slot {
                if image.id == id {
                    image.horizontal_position = pos;
                    return;
                }
            }
        }
    }

    pub fn set_image_vertical_position(&mut self, id: usize, pos: VerticalPosition) {
        for slot in &mut self.paragraph_images {
            if let Some(image) = slot {
                if image.id == id {
                    image.vertical_position = pos;
                    return;
                }
            }
        }
    }

    pub fn set_image_distance_from_text(&mut self, id: usize, dist: DistanceFromText) {
        for slot in &mut self.paragraph_images {
            if let Some(image) = slot {
                if image.id == id {
                    image.distance_from_text = dist;
                    return;
                }
            }
        }
    }

    pub fn set_image_z_index(&mut self, id: usize, z: i32) {
        for slot in &mut self.paragraph_images {
            if let Some(image) = slot {
                if image.id == id {
                    image.z_index = z;
                    return;
                }
            }
        }
    }

    pub fn set_image_move_with_text(&mut self, id: usize, flag: bool) {
        for slot in &mut self.paragraph_images {
            if let Some(image) = slot {
                if image.id == id {
                    image.move_with_text = flag;
                    return;
                }
            }
        }
    }

    pub fn set_image_lock_aspect_ratio(&mut self, id: usize, flag: bool) {
        for slot in &mut self.paragraph_images {
            if let Some(image) = slot {
                if image.id == id {
                    image.lock_aspect_ratio = flag;
                    return;
                }
            }
        }
    }

    pub fn set_image_opacity(&mut self, id: usize, opacity: f32) {
        for slot in &mut self.paragraph_images {
            if let Some(image) = slot {
                if image.id == id {
                    image.opacity = opacity.clamp(0.0, 1.0);
                    return;
                }
            }
        }
    }

    pub fn set_image_wrap_mode(&mut self, id: usize, wrap_mode: WrapMode) {
        for slot in &mut self.paragraph_images {
            if let Some(image) = slot {
                if image.id == id {
                    image.wrap_mode = wrap_mode;
                    return;
                }
            }
        }
    }

    pub fn set_image_rendering(&mut self, id: usize, rendering: ImageRendering) {
        for slot in &mut self.paragraph_images {
            if let Some(image) = slot {
                if image.id == id {
                    image.rendering = rendering;
                    return;
                }
            }
        }
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

        let serialized = match extension.as_str() {
            "md" | "markdown" => self.to_markdown(),
            "txt" | "" => self.to_plain_text_export(),
            other => {
                return Err(format!(
                    "saving .{other} is not supported yet; use .txt or .md"
                ));
            }
        };

        fs::write(path, serialized)
            .map_err(|error| format!("failed to save {}: {error}", path.display()))
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
    }

    fn to_plain_text_export(&self) -> String {
        self.paragraphs()
            .into_iter()
            .map(|paragraph| {
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

#[cfg(test)]
mod tests {
    use super::{
        plain_text_from_runs, CharacterStyle, DocumentImage, DocumentState, ImageLayoutMode,
        ImageRendering, ListKind, WrapMode, OBJECT_REPLACEMENT_CHAR,
    };

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
}
