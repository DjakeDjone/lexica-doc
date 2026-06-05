use crate::document::{DocumentState, DocumentTable, TextRun, OBJECT_REPLACEMENT_CHAR};

impl DocumentState {
    pub(crate) fn to_plain_text_export(&self) -> String {
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
}

pub(crate) fn plain_text_from_runs(runs: &[TextRun]) -> String {
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
