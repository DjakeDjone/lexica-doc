use crate::document::{
    DocumentState, DocumentTable, FontChoice, ListKind, ParagraphAlignment, TableCell, TextRun,
    VerticalAlign, OBJECT_REPLACEMENT_CHAR,
};

impl DocumentState {
    pub(crate) fn to_markdown(&self) -> String {
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

pub(super) fn markdown_text_from_runs(runs: &[TextRun]) -> String {
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
        if FontChoice::from_style(run.style).is_monospace() {
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
        match run.style.vertical_align {
            VerticalAlign::Baseline => {}
            VerticalAlign::Superscript => text = format!("<sup>{text}</sup>"),
            VerticalAlign::Subscript => text = format!("<sub>{text}</sub>"),
        }
        output.push_str(&text);
    }
    output
}
