use eframe::egui;
use pulldown_cmark::{Event, HeadingLevel, Options, Parser, Tag, TagEnd};

use crate::document::{CharacterStyle, FontChoice, TextRun};

use super::text::slice_char_range;

pub(super) fn markdown_to_runs(source: &str) -> Vec<TextRun> {
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
                        next.highlight_color = egui::Color32::from_rgb(243, 243, 243);
                    }
                    Tag::Heading { level, .. } => {
                        next.bold = true;
                        next.font_size_points = heading_font_size(level);
                        heading_level = Some(level);
                    }
                    Tag::BlockQuote(_) => {
                        next.italic = true;
                        next.text_color = egui::Color32::from_rgb(86, 90, 100);
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
                style.highlight_color = egui::Color32::from_rgb(243, 243, 243);
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
                        text_color: egui::Color32::from_gray(90),
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

pub(super) fn markdown_heading_prefix(line: &str) -> Option<(usize, CharacterStyle)> {
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

pub(super) fn markdown_line_replacement(line: &str) -> Option<Vec<TextRun>> {
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

pub(super) fn markdown_cursor_index_in_line(line: &str, cursor_in_line: usize) -> usize {
    let prefix = slice_char_range(line, 0..cursor_in_line.min(line.chars().count()));
    plain_text_from_runs(&render_markdown_line(prefix))
        .chars()
        .count()
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
