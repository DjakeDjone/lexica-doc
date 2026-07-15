use eframe::egui;
use pulldown_cmark::{Event, HeadingLevel, Options, Parser, Tag, TagEnd};

use crate::document::{
    CharacterStyle, DocumentTable, FontChoice, ListKind, PageMargins, PageSize, ParagraphStyle,
    TableCell, TextRun, OBJECT_REPLACEMENT_CHAR,
};

pub(super) struct MarkdownImport {
    pub runs: Vec<TextRun>,
    pub paragraph_styles: Vec<ParagraphStyle>,
    pub paragraph_tables: Vec<Option<DocumentTable>>,
}

pub(super) fn import_markdown(source: &str) -> MarkdownImport {
    let parser = Parser::new_ext(
        source,
        Options::ENABLE_STRIKETHROUGH | Options::ENABLE_TASKLISTS | Options::ENABLE_TABLES,
    );

    let mut runs = Vec::new();
    let mut stack = vec![CharacterStyle::default()];
    let mut pending_list_kind = None;
    let mut list_kinds = Vec::new();
    let mut list_paragraphs = Vec::new();
    let mut heading_level = None;
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
                        pending_list_kind = list_kinds.last().copied();
                    }
                    Tag::List(start) => {
                        list_kinds.push(if start.is_some() {
                            ListKind::Ordered
                        } else {
                            ListKind::Bullet
                        });
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
                    TagEnd::Item => {
                        pending_list_kind = None;
                        append_plain(
                            &mut runs,
                            "\n",
                            *stack.last().unwrap_or(&CharacterStyle::default()),
                        );
                    }
                    TagEnd::List(_) => {
                        list_kinds.pop();
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
                mark_list_paragraph(
                    &runs,
                    table_cell.is_none(),
                    &mut pending_list_kind,
                    &mut list_paragraphs,
                );
                let target = table_cell.as_mut().unwrap_or(&mut runs);
                append_plain(
                    target,
                    &text,
                    *stack.last().unwrap_or(&CharacterStyle::default()),
                );
            }
            Event::Code(text) => {
                mark_list_paragraph(
                    &runs,
                    table_cell.is_none(),
                    &mut pending_list_kind,
                    &mut list_paragraphs,
                );
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
    let mut paragraph_styles = vec![ParagraphStyle::default(); text.matches('\n').count() + 1];
    for (index, kind) in list_paragraphs {
        paragraph_styles[index].list_kind = kind;
    }
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
        paragraph_styles,
        paragraph_tables,
    }
}

fn mark_list_paragraph(
    runs: &[TextRun],
    is_document_text: bool,
    pending: &mut Option<ListKind>,
    paragraphs: &mut Vec<(usize, ListKind)>,
) {
    if let Some(kind) = pending.take() {
        if is_document_text {
            paragraphs.push((
                runs.iter().map(|run| run.text.matches('\n').count()).sum(),
                kind,
            ));
        }
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
