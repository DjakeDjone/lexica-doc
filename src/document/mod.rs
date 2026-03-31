mod docx;
mod markdown;
mod text;

use std::{fs, ops::Range, path::Path};

use eframe::egui::{
    epaint::text::{LayoutJob, TextFormat, VariationCoords},
    Color32, FontFamily, FontId, Stroke,
};

use docx::docx_to_runs;
use markdown::{
    markdown_cursor_index_in_line, markdown_heading_prefix, markdown_line_replacement,
    markdown_to_runs,
};
use text::{char_to_byte_index, line_char_range, slice_char_range};

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
