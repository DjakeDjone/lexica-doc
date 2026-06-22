use eframe::egui::{epaint::text::TextFormat, Align, Color32, FontFamily, FontId, Stroke};
use serde::Serialize;
use std::ops::Range;

use crate::document::text::{line_char_range, slice_char_range, word_char_range};

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

pub(crate) fn serialize_color32<S: serde::Serializer>(
    color: &Color32,
    s: S,
) -> Result<S::Ok, S::Error> {
    s.serialize_str(&format!(
        "#{:02x}{:02x}{:02x}{:02x}",
        color.r(),
        color.g(),
        color.b(),
        color.a()
    ))
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize)]
pub enum VerticalAlign {
    Baseline,
    Superscript,
    Subscript,
}

#[derive(Clone, Copy, Debug, PartialEq, Serialize)]
pub struct CharacterStyle {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub vertical_align: VerticalAlign,
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
            vertical_align: VerticalAlign::Baseline,
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
use super::table::DocumentTable;

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

#[derive(Clone, Copy, Debug, PartialEq, Serialize)]
pub struct PageSize {
    pub width_points: f32,
    pub height_points: f32,
}

impl PageSize {
    pub const fn a4() -> Self {
        Self {
            width_points: 595.0,
            height_points: 842.0,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Serialize)]
pub struct PageMargins {
    pub top_points: f32,
    pub right_points: f32,
    pub bottom_points: f32,
    pub left_points: f32,
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

pub type SectionId = usize;

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize)]
pub enum HeaderFooterKind {
    Header,
    Footer,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq, Hash, Serialize)]
pub enum HeaderFooterVariant {
    Default,
    First,
    Even,
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct HeaderFooterStory {
    pub runs: Vec<TextRun>,
}

impl HeaderFooterStory {
    pub fn empty() -> Self {
        Self {
            runs: empty_header_footer_runs(),
        }
    }

    pub fn from_runs(runs: Vec<TextRun>) -> Self {
        Self { runs }
    }

    pub fn plain_text(&self) -> String {
        plain_text_from_runs(&self.runs)
    }
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct HeaderFooterSlot {
    pub story: HeaderFooterStory,
    pub linked_to_previous: bool,
}

impl HeaderFooterSlot {
    pub fn empty(linked_to_previous: bool) -> Self {
        Self {
            story: HeaderFooterStory::empty(),
            linked_to_previous,
        }
    }

    pub(crate) fn story_ref(&self) -> &HeaderFooterStory {
        &self.story
    }
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct SectionHeaderFooter {
    pub header_default: HeaderFooterSlot,
    pub header_first: HeaderFooterSlot,
    pub header_even: HeaderFooterSlot,
    pub footer_default: HeaderFooterSlot,
    pub footer_first: HeaderFooterSlot,
    pub footer_even: HeaderFooterSlot,
}

impl SectionHeaderFooter {
    pub fn empty(linked_to_previous: bool) -> Self {
        Self {
            header_default: HeaderFooterSlot::empty(linked_to_previous),
            header_first: HeaderFooterSlot::empty(linked_to_previous),
            header_even: HeaderFooterSlot::empty(linked_to_previous),
            footer_default: HeaderFooterSlot::empty(linked_to_previous),
            footer_first: HeaderFooterSlot::empty(linked_to_previous),
            footer_even: HeaderFooterSlot::empty(linked_to_previous),
        }
    }

    pub fn slot(&self, kind: HeaderFooterKind, variant: HeaderFooterVariant) -> &HeaderFooterSlot {
        match (kind, variant) {
            (HeaderFooterKind::Header, HeaderFooterVariant::Default) => &self.header_default,
            (HeaderFooterKind::Header, HeaderFooterVariant::First) => &self.header_first,
            (HeaderFooterKind::Header, HeaderFooterVariant::Even) => &self.header_even,
            (HeaderFooterKind::Footer, HeaderFooterVariant::Default) => &self.footer_default,
            (HeaderFooterKind::Footer, HeaderFooterVariant::First) => &self.footer_first,
            (HeaderFooterKind::Footer, HeaderFooterVariant::Even) => &self.footer_even,
        }
    }

    pub fn slot_mut(
        &mut self,
        kind: HeaderFooterKind,
        variant: HeaderFooterVariant,
    ) -> &mut HeaderFooterSlot {
        match (kind, variant) {
            (HeaderFooterKind::Header, HeaderFooterVariant::Default) => &mut self.header_default,
            (HeaderFooterKind::Header, HeaderFooterVariant::First) => &mut self.header_first,
            (HeaderFooterKind::Header, HeaderFooterVariant::Even) => &mut self.header_even,
            (HeaderFooterKind::Footer, HeaderFooterVariant::Default) => &mut self.footer_default,
            (HeaderFooterKind::Footer, HeaderFooterVariant::First) => &mut self.footer_first,
            (HeaderFooterKind::Footer, HeaderFooterVariant::Even) => &mut self.footer_even,
        }
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Serialize)]
pub struct PageSetup {
    pub page_size: PageSize,
    pub margins: PageMargins,
    pub header_from_top_points: f32,
    pub footer_from_bottom_points: f32,
    pub page_number_start: Option<usize>,
}

impl PageSetup {
    pub const fn standard() -> Self {
        Self {
            page_size: PageSize::a4(),
            margins: PageMargins::standard(),
            header_from_top_points: 36.0,
            footer_from_bottom_points: 36.0,
            page_number_start: Some(1),
        }
    }
}

#[derive(Clone, Debug, PartialEq, Serialize)]
pub struct Section {
    pub id: SectionId,
    pub starts_at_paragraph: usize,
    pub page_setup: PageSetup,
    pub different_first_page: bool,
    pub header_footer: SectionHeaderFooter,
}

impl Section {
    pub fn first(page_setup: PageSetup) -> Self {
        Self {
            id: 1,
            starts_at_paragraph: 0,
            page_setup,
            different_first_page: false,
            header_footer: SectionHeaderFooter::empty(false),
        }
    }

    pub fn linked_from(id: SectionId, starts_at_paragraph: usize, previous: &Section) -> Self {
        let mut page_setup = previous.page_setup;
        page_setup.page_number_start = None;
        Self {
            id,
            starts_at_paragraph,
            page_setup,
            different_first_page: previous.different_first_page,
            header_footer: SectionHeaderFooter::empty(true),
        }
    }
}

pub struct ResolvedHeaderFooter<'a> {
    pub section_id: SectionId,
    pub source_section_id: SectionId,
    pub variant: HeaderFooterVariant,
    pub story: &'a HeaderFooterStory,
    pub inherited: bool,
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
    pub header_text: String,
    pub footer_text: String,
    pub first_page_header_text: String,
    pub first_page_footer_text: String,
    pub even_page_header_text: String,
    pub even_page_footer_text: String,
    pub header_runs: Vec<TextRun>,
    pub footer_runs: Vec<TextRun>,
    pub first_page_header_runs: Vec<TextRun>,
    pub first_page_footer_runs: Vec<TextRun>,
    pub even_page_header_runs: Vec<TextRun>,
    pub even_page_footer_runs: Vec<TextRun>,
    pub different_first_page: bool,
    pub different_odd_even_pages: bool,
    pub page_number_start: usize,
    pub sections: Vec<Section>,
}

impl DocumentState {
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

    pub(crate) fn paragraph_index_at(&self, char_index: usize) -> usize {
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
}

pub(crate) fn empty_header_footer_runs() -> Vec<TextRun> {
    vec![TextRun {
        text: String::new(),
        style: CharacterStyle::default(),
    }]
}

pub(crate) fn plain_text_from_runs(runs: &[TextRun]) -> String {
    runs.iter().map(|run| run.text.as_str()).collect()
}

pub(crate) fn text_format(style: CharacterStyle, zoom: f32) -> TextFormat {
    let line_color = style.text_color;
    let font_size = match style.vertical_align {
        VerticalAlign::Baseline => style.font_size_points,
        VerticalAlign::Superscript | VerticalAlign::Subscript => style.font_size_points * 0.65,
    } * zoom;
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
        valign: match style.vertical_align {
            VerticalAlign::Baseline | VerticalAlign::Subscript => Align::BOTTOM,
            VerticalAlign::Superscript => Align::TOP,
        },
        ..Default::default()
    }
}

pub(crate) fn text_font_family(style: CharacterStyle) -> FontFamily {
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

pub(crate) fn append_text_run(runs: &mut Vec<TextRun>, text: &str, style: CharacterStyle) {
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
