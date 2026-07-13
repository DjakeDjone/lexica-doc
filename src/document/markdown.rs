use eframe::egui;
use pulldown_cmark::{Event, HeadingLevel, Options, Parser, Tag, TagEnd};

use crate::document::{
    CharacterStyle, DocumentTable, FontChoice, PageMargins, PageSize, TableCell, TextRun,
    OBJECT_REPLACEMENT_CHAR,
};

pub(super) struct MarkdownImport {
    pub runs: Vec<TextRun>,
    pub paragraph_tables: Vec<Option<DocumentTable>>,
}

pub(super) fn import_markdown(source: &str) -> MarkdownImport {
    let parser = Parser::new_ext(
        source,
        Options::ENABLE_STRIKETHROUGH | Options::ENABLE_TASKLISTS | Options::ENABLE_TABLES,
    );

    let mut runs = Vec::new();
    let mut stack = vec![CharacterStyle::default()];
    let mut pending_prefix = String::new();
    let mut heading_level = None;
    let mut list_depth = 0usize;
    let mut tables = Vec::new();
    let mut table_columns = 0usize;
    let mut table_rows = Vec::new();
    let mut table_row = None;
    let mut table_cell = None;
    let placeholder = OBJECT_REPLACEMENT_CHAR.to_string();

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
                    Tag::Table(alignments) => {
                        table_columns = alignments.len();
                        table_rows.clear();
                        append_plain(
                            &mut runs,
                            &placeholder,
                            CharacterStyle::default(),
                        );
                    }
                    Tag::TableHead | Tag::TableRow => table_row = Some(Vec::new()),
                    Tag::TableCell => table_cell = Some(Vec::new()),
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
                    TagEnd::TableCell => {
                        let runs = table_cell.take().unwrap_or_default();
                        table_row.get_or_insert_default().push(TableCell {
                            runs: if runs.is_empty() {
                                TableCell::new("").runs
                            } else {
                                runs
                            },
                            images: Vec::new(),
                            col_span: 1,
                            row_span: 1,
                        });
                    }
                    TagEnd::TableHead | TagEnd::TableRow => {
                        table_rows.push(table_row.take().unwrap_or_default());
                    }
                    TagEnd::Table => {
                        for row in &mut table_rows {
                            row.resize_with(table_columns, || TableCell::new(""));
                        }
                        let page_size = PageSize::a4();
                        let margins = PageMargins::standard();
                        let mut table = DocumentTable::new(
                            tables.len() + 1,
                            table_rows.len(),
                            table_columns,
                            page_size.width_points
                                - margins.left_points
                                - margins.right_points,
                        );
                        table.rows = std::mem::take(&mut table_rows);
                        tables.push(table);
                        append_plain(&mut runs, "\n\n", CharacterStyle::default());
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
                    let target = table_cell.as_mut().unwrap_or(&mut runs);
                    append_plain(
                        target,
                        &pending_prefix,
                        *stack.last().unwrap_or(&CharacterStyle::default()),
                    );
                    pending_prefix.clear();
                }
                let target = table_cell.as_mut().unwrap_or(&mut runs);
                append_plain(
                    target,
                    &text,
                    *stack.last().unwrap_or(&CharacterStyle::default()),
                );
            }
            Event::Code(text) => {
                let mut style = *stack.last().unwrap_or(&CharacterStyle::default());
                style.font_choice = FontChoice::Monospace;
                style.highlight_color = egui::Color32::from_rgb(243, 243, 243);
                append_plain(table_cell.as_mut().unwrap_or(&mut runs), &text, style);
            }
            Event::SoftBreak | Event::HardBreak => {
                append_plain(
                    table_cell.as_mut().unwrap_or(&mut runs),
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

    let text: String = runs.iter().map(|run| run.text.as_str()).collect();
    let mut tables = tables.into_iter();
    let paragraph_tables = text
        .split('\n')
        .map(|paragraph| {
            (paragraph == placeholder)
                .then(|| tables.next())
                .flatten()
        })
        .collect();

    MarkdownImport {
        runs,
        paragraph_tables,
    }
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
