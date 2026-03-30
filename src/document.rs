use std::{
    fs,
    io::{Cursor, Read},
    ops::Range,
    path::Path,
};

use eframe::egui::{
    epaint::text::{LayoutJob, TextFormat, VariationCoords},
    Color32, FontFamily, FontId, Stroke,
};
use pulldown_cmark::{Event, HeadingLevel, Options, Parser, Tag, TagEnd};
use quick_xml::{events::Event as XmlEvent, Reader};
use zip::ZipArchive;

#[derive(Clone, Copy, Debug, PartialEq)]
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

#[derive(Clone, Copy, Debug, PartialEq)]
pub struct CharacterStyle {
    pub bold: bool,
    pub italic: bool,
    pub underline: bool,
    pub strikethrough: bool,
    pub font_size_points: f32,
    pub font_choice: FontChoice,
    pub text_color: Color32,
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
            text_color: Color32::from_rgb(36, 39, 46),
            highlight_color: Color32::TRANSPARENT,
        }
    }
}

#[derive(Clone, Debug, PartialEq)]
pub struct TextRun {
    pub text: String,
    pub style: CharacterStyle,
}

#[derive(Clone, Copy, Debug)]
pub struct PageSize {
    pub width_points: f32,
    pub height_points: f32,
}

#[derive(Clone, Copy, Debug)]
pub struct PageMargins {
    pub top_points: f32,
    pub right_points: f32,
    pub bottom_points: f32,
    pub left_points: f32,
}

pub struct DocumentState {
    pub title: String,
    pub runs: Vec<TextRun>,
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
        self.normalize_runs();
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
            "docx" => docx_to_runs(
                &fs::read(path)
                    .map_err(|error| format!("failed to read {}: {error}", path.display()))?,
            )?,
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
            "txt" | "" => self.plain_text(),
            other => {
                return Err(format!(
                    "saving .{other} is not supported yet; use .txt or .md"
                ));
            }
        };

        fs::write(path, serialized)
            .map_err(|error| format!("failed to save {}: {error}", path.display()))
    }

    pub fn layout_job(&self, zoom: f32, wrap_width: f32) -> LayoutJob {
        let mut job = LayoutJob::default();
        job.wrap.max_width = wrap_width;
        job.break_on_newline = true;

        for run in &self.runs {
            if run.text.is_empty() {
                continue;
            }

            let mut coords = VariationCoords::default();
            if run.style.bold {
                coords.push("wght", 700.0);
            }

            let line_color = run.style.text_color;
            let font_size = if run.style.bold {
                (run.style.font_size_points + 0.8) * zoom
            } else {
                run.style.font_size_points * zoom
            };
            let format = TextFormat {
                font_id: FontId::new(font_size, run.style.font_choice.family()),
                color: if run.style.bold {
                    run.style.text_color.gamma_multiply(0.88)
                } else {
                    run.style.text_color
                },
                background: run.style.highlight_color,
                italics: run.style.italic,
                underline: if run.style.underline {
                    Stroke::new(1.0, line_color)
                } else {
                    Stroke::NONE
                },
                strikethrough: if run.style.strikethrough {
                    Stroke::new(1.0, line_color)
                } else {
                    Stroke::NONE
                },
                coords,
                ..Default::default()
            };

            job.append(&run.text, 0.0, format);
        }

        if job.sections.is_empty() {
            job.append(
                "",
                0.0,
                TextFormat::simple(
                    FontId::new(12.0 * zoom, FontFamily::Proportional),
                    CharacterStyle::default().text_color,
                ),
            );
        }

        job
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

    fn to_markdown(&self) -> String {
        let mut output = String::new();
        for run in &self.runs {
            let mut text = run.text.clone();
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

fn markdown_to_runs(source: &str) -> Vec<TextRun> {
    let parser = Parser::new_ext(
        source,
        Options::ENABLE_STRIKETHROUGH | Options::ENABLE_TASKLISTS | Options::ENABLE_TABLES,
    );

    let mut runs = Vec::new();
    let mut stack = vec![CharacterStyle::default()];
    let mut pending_prefix = String::new();
    let mut heading_level = None;
    let mut list_depth = 0usize;

    for event in parser {
        match event {
            Event::Start(tag) => {
                let mut next = *stack.last().unwrap_or(&CharacterStyle::default());
                match tag {
                    Tag::Strong => next.bold = true,
                    Tag::Emphasis => next.italic = true,
                    Tag::Strikethrough => next.strikethrough = true,
                    Tag::CodeBlock(_) => {
                        next.font_choice = FontChoice::Monospace;
                        next.highlight_color = Color32::from_rgb(243, 243, 243);
                    }
                    Tag::Heading { level, .. } => {
                        next.bold = true;
                        next.font_size_points = heading_font_size(level);
                        heading_level = Some(level);
                    }
                    Tag::BlockQuote(_) => {
                        next.italic = true;
                        next.text_color = Color32::from_rgb(86, 90, 100);
                    }
                    Tag::Item => {
                        pending_prefix.push_str(&"  ".repeat(list_depth.saturating_sub(1)));
                        pending_prefix.push_str("• ");
                    }
                    Tag::List(_) => {
                        list_depth += 1;
                    }
                    _ => {}
                }
                stack.push(next);
            }
            Event::End(tag) => {
                match tag {
                    TagEnd::Paragraph | TagEnd::Heading(_) => append_plain(
                        &mut runs,
                        "\n\n",
                        *stack.last().unwrap_or(&CharacterStyle::default()),
                    ),
                    TagEnd::CodeBlock => append_plain(
                        &mut runs,
                        "\n\n",
                        *stack.last().unwrap_or(&CharacterStyle::default()),
                    ),
                    TagEnd::Item => append_plain(
                        &mut runs,
                        "\n",
                        *stack.last().unwrap_or(&CharacterStyle::default()),
                    ),
                    TagEnd::List(_) => {
                        list_depth = list_depth.saturating_sub(1);
                        append_plain(
                            &mut runs,
                            "\n",
                            *stack.last().unwrap_or(&CharacterStyle::default()),
                        );
                    }
                    _ => {}
                }
                stack.pop();
                if matches!(tag, TagEnd::Heading(_)) {
                    heading_level = None;
                }
            }
            Event::Text(text) => {
                if !pending_prefix.is_empty() {
                    append_plain(
                        &mut runs,
                        &pending_prefix,
                        *stack.last().unwrap_or(&CharacterStyle::default()),
                    );
                    pending_prefix.clear();
                }
                append_plain(
                    &mut runs,
                    &text,
                    *stack.last().unwrap_or(&CharacterStyle::default()),
                );
            }
            Event::Code(text) => {
                let mut style = *stack.last().unwrap_or(&CharacterStyle::default());
                style.font_choice = FontChoice::Monospace;
                style.highlight_color = Color32::from_rgb(243, 243, 243);
                append_plain(&mut runs, &text, style);
            }
            Event::SoftBreak | Event::HardBreak => {
                append_plain(
                    &mut runs,
                    "\n",
                    *stack.last().unwrap_or(&CharacterStyle::default()),
                );
            }
            Event::Rule => {
                append_plain(
                    &mut runs,
                    "\n--------------------\n",
                    CharacterStyle {
                        text_color: Color32::from_gray(90),
                        ..CharacterStyle::default()
                    },
                );
            }
            _ => {}
        }
    }

    if runs.is_empty() && heading_level.is_none() {
        runs.push(TextRun {
            text: String::new(),
            style: CharacterStyle::default(),
        });
    }

    runs
}

fn append_plain(runs: &mut Vec<TextRun>, text: &str, style: CharacterStyle) {
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

fn heading_font_size(level: HeadingLevel) -> f32 {
    match level {
        HeadingLevel::H1 => 28.0,
        HeadingLevel::H2 => 24.0,
        HeadingLevel::H3 => 20.0,
        HeadingLevel::H4 => 18.0,
        HeadingLevel::H5 => 16.0,
        HeadingLevel::H6 => 14.0,
    }
}

fn heading_style(level: HeadingLevel) -> CharacterStyle {
    CharacterStyle {
        bold: true,
        font_size_points: heading_font_size(level),
        ..CharacterStyle::default()
    }
}

fn markdown_heading_prefix(line: &str) -> Option<(usize, CharacterStyle)> {
    let hashes = line.chars().take_while(|ch| *ch == '#').count();
    if !(1..=6).contains(&hashes) {
        return None;
    }

    if line.chars().nth(hashes) != Some(' ') {
        return None;
    }

    let level = match hashes {
        1 => HeadingLevel::H1,
        2 => HeadingLevel::H2,
        3 => HeadingLevel::H3,
        4 => HeadingLevel::H4,
        5 => HeadingLevel::H5,
        _ => HeadingLevel::H6,
    };

    Some((hashes + 1, heading_style(level)))
}

fn markdown_line_replacement(line: &str) -> Option<Vec<TextRun>> {
    if line.trim().is_empty() {
        return None;
    }

    let runs = render_markdown_line(line);
    if runs.is_empty() {
        return None;
    }

    let rendered_text = plain_text_from_runs(&runs);
    let has_non_default_style = runs
        .iter()
        .any(|run| run.style != CharacterStyle::default());
    if rendered_text == line && !has_non_default_style {
        return None;
    }

    Some(runs)
}

fn render_markdown_line(line: &str) -> Vec<TextRun> {
    let trailing_whitespace = line
        .chars()
        .rev()
        .take_while(|ch| ch.is_whitespace() && *ch != '\n')
        .count();
    let content_len = line.chars().count().saturating_sub(trailing_whitespace);
    let content = slice_char_range(line, 0..content_len);
    let suffix = slice_char_range(line, content_len..line.chars().count());

    let mut runs = markdown_to_runs(content);
    trim_trailing_newlines(&mut runs);

    if !suffix.is_empty() {
        let suffix_style = runs
            .last()
            .map(|run| run.style)
            .or_else(|| markdown_heading_prefix(content).map(|(_, style)| style))
            .unwrap_or_default();
        append_plain(&mut runs, suffix, suffix_style);
    }

    runs
}

fn trim_trailing_newlines(runs: &mut Vec<TextRun>) {
    while let Some(last) = runs.last_mut() {
        while last.text.ends_with('\n') {
            last.text.pop();
        }

        if last.text.is_empty() {
            runs.pop();
        } else {
            break;
        }
    }
}

fn plain_text_from_runs(runs: &[TextRun]) -> String {
    runs.iter().map(|run| run.text.as_str()).collect()
}

fn markdown_cursor_index_in_line(line: &str, cursor_in_line: usize) -> usize {
    let prefix = slice_char_range(line, 0..cursor_in_line.min(line.chars().count()));
    plain_text_from_runs(&render_markdown_line(prefix))
        .chars()
        .count()
}

fn line_char_range(text: &str, char_index: usize) -> Range<usize> {
    let total_chars = text.chars().count();
    let target = char_index.min(total_chars);
    let mut start = 0;
    let mut end = total_chars;

    for (index, ch) in text.chars().enumerate() {
        if index < target && ch == '\n' {
            start = index + 1;
        }
        if index >= target && ch == '\n' {
            end = index;
            break;
        }
    }

    start..end
}

fn docx_to_runs(bytes: &[u8]) -> Result<Vec<TextRun>, String> {
    let cursor = Cursor::new(bytes);
    let mut archive =
        ZipArchive::new(cursor).map_err(|error| format!("invalid .docx archive: {error}"))?;
    let mut document_xml = String::new();
    archive
        .by_name("word/document.xml")
        .map_err(|error| format!("missing word/document.xml: {error}"))?
        .read_to_string(&mut document_xml)
        .map_err(|error| format!("failed to read word/document.xml: {error}"))?;

    let mut reader = Reader::from_str(&document_xml);
    reader.config_mut().trim_text(false);

    let mut runs = Vec::new();
    let mut run_style = CharacterStyle::default();
    let mut in_text = false;
    let mut paragraph_count = 0usize;

    loop {
        match reader.read_event() {
            Ok(XmlEvent::Start(event)) => match local_name(event.name().as_ref()) {
                b"p" => {
                    if paragraph_count > 0 {
                        append_plain(&mut runs, "\n\n", CharacterStyle::default());
                    }
                    paragraph_count += 1;
                }
                b"r" => {
                    run_style = CharacterStyle::default();
                }
                b"t" => in_text = true,
                b"br" => append_plain(&mut runs, "\n", run_style),
                b"tab" => append_plain(&mut runs, "\t", run_style),
                b"b" => run_style.bold = docx_flag(&event, true),
                b"i" => run_style.italic = docx_flag(&event, true),
                b"u" => {
                    run_style.underline =
                        !matches!(attr_value(&event, b"val").as_deref(), Some("none"))
                }
                b"strike" | b"dstrike" => run_style.strikethrough = docx_flag(&event, true),
                b"sz" => {
                    if let Some(value) = attr_value(&event, b"val") {
                        if let Ok(half_points) = value.parse::<f32>() {
                            run_style.font_size_points = (half_points / 2.0).clamp(8.0, 72.0);
                        }
                    }
                }
                b"color" => {
                    if let Some(value) = attr_value(&event, b"val") {
                        if let Some(color) = parse_hex_color(&value) {
                            run_style.text_color = color;
                        }
                    }
                }
                b"highlight" => {
                    if let Some(value) = attr_value(&event, b"val") {
                        run_style.highlight_color = highlight_color(&value);
                    }
                }
                _ => {}
            },
            Ok(XmlEvent::Empty(event)) => match local_name(event.name().as_ref()) {
                b"br" => append_plain(&mut runs, "\n", run_style),
                b"tab" => append_plain(&mut runs, "\t", run_style),
                b"b" => run_style.bold = docx_flag(&event, true),
                b"i" => run_style.italic = docx_flag(&event, true),
                b"u" => {
                    run_style.underline =
                        !matches!(attr_value(&event, b"val").as_deref(), Some("none"))
                }
                b"strike" | b"dstrike" => run_style.strikethrough = docx_flag(&event, true),
                b"sz" => {
                    if let Some(value) = attr_value(&event, b"val") {
                        if let Ok(half_points) = value.parse::<f32>() {
                            run_style.font_size_points = (half_points / 2.0).clamp(8.0, 72.0);
                        }
                    }
                }
                b"color" => {
                    if let Some(value) = attr_value(&event, b"val") {
                        if let Some(color) = parse_hex_color(&value) {
                            run_style.text_color = color;
                        }
                    }
                }
                b"highlight" => {
                    if let Some(value) = attr_value(&event, b"val") {
                        run_style.highlight_color = highlight_color(&value);
                    }
                }
                _ => {}
            },
            Ok(XmlEvent::Text(text)) => {
                if in_text {
                    let decoded = text
                        .xml_content()
                        .map_err(|error| format!("failed to decode document text: {error}"))?;
                    append_plain(&mut runs, decoded.as_ref(), run_style);
                }
            }
            Ok(XmlEvent::End(event)) => {
                if local_name(event.name().as_ref()) == b"t" {
                    in_text = false;
                }
            }
            Ok(XmlEvent::Eof) => break,
            Err(error) => return Err(format!("failed to parse word/document.xml: {error}")),
            _ => {}
        }
    }

    if runs.is_empty() {
        runs.push(TextRun {
            text: String::new(),
            style: CharacterStyle::default(),
        });
    }

    Ok(runs)
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn attr_value(event: &quick_xml::events::BytesStart<'_>, key: &[u8]) -> Option<String> {
    event
        .attributes()
        .flatten()
        .find(|attr| local_name(attr.key.as_ref()) == key)
        .and_then(|attr| String::from_utf8(attr.value.into_owned()).ok())
}

fn docx_flag(event: &quick_xml::events::BytesStart<'_>, default: bool) -> bool {
    match attr_value(event, b"val").as_deref() {
        Some("0" | "false") => false,
        Some("1" | "true") => true,
        Some(_) => default,
        None => default,
    }
}

fn parse_hex_color(value: &str) -> Option<Color32> {
    if value.len() != 6 {
        return None;
    }

    let red = u8::from_str_radix(&value[0..2], 16).ok()?;
    let green = u8::from_str_radix(&value[2..4], 16).ok()?;
    let blue = u8::from_str_radix(&value[4..6], 16).ok()?;
    Some(Color32::from_rgb(red, green, blue))
}

fn highlight_color(value: &str) -> Color32 {
    match value {
        "yellow" => Color32::from_rgb(255, 242, 129),
        "green" => Color32::from_rgb(187, 232, 172),
        "cyan" => Color32::from_rgb(163, 231, 240),
        "magenta" => Color32::from_rgb(244, 188, 231),
        "blue" => Color32::from_rgb(177, 205, 252),
        "red" => Color32::from_rgb(248, 188, 188),
        "darkYellow" => Color32::from_rgb(215, 185, 90),
        "darkGreen" => Color32::from_rgb(104, 170, 112),
        "darkBlue" => Color32::from_rgb(99, 129, 207),
        _ => Color32::TRANSPARENT,
    }
}

fn char_to_byte_index(text: &str, char_index: usize) -> usize {
    text.char_indices()
        .nth(char_index)
        .map(|(idx, _)| idx)
        .unwrap_or(text.len())
}

fn slice_char_range(text: &str, range: Range<usize>) -> &str {
    let start = char_to_byte_index(text, range.start);
    let end = char_to_byte_index(text, range.end);
    &text[start..end]
}
