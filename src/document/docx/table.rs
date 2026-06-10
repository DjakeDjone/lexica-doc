use quick_xml::{events::Event as XmlEvent, Reader};

use super::styles::{apply_resolved_font, resolve_font_from_event_without_theme};
use super::{append_plain, attr_value, docx_flag, local_name, parse_hex_color, twips_to_points};
use crate::document::{CharacterStyle, DocumentTable, TableCell, TextRun};

pub(crate) fn parse_docx_table(
    reader: &mut Reader<&[u8]>,
    table_id: usize,
    available_width: f32,
) -> Result<DocumentTable, String> {
    let mut rows: Vec<Vec<TableCell>> = Vec::new();
    let mut col_widths_points: Vec<f32> = Vec::new();
    let mut row_heights_points: Vec<f32> = Vec::new();
    let mut current_row: Option<Vec<TableCell>> = None;
    let mut current_cell_runs: Option<Vec<TextRun>> = None;
    let mut current_text = false;
    let mut current_run_style = CharacterStyle::default();
    let mut current_row_height = None::<f32>;

    loop {
        match reader.read_event() {
            Ok(XmlEvent::Start(event)) => match local_name(event.name().as_ref()) {
                b"tr" => {
                    current_row = Some(Vec::new());
                    current_row_height = None;
                }
                b"tc" => current_cell_runs = Some(Vec::new()),
                b"r" => current_run_style = CharacterStyle::default(),
                b"t" => current_text = true,
                b"tab" if current_cell_runs.is_some() => {
                    if let Some(runs) = current_cell_runs.as_mut() {
                        append_plain(runs, "\t", current_run_style);
                    }
                }
                b"br" | b"cr" if current_cell_runs.is_some() => {
                    if let Some(runs) = current_cell_runs.as_mut() {
                        append_plain(runs, "\n", current_run_style);
                    }
                }
                b"gridCol" => {
                    if let Some(width) = attr_value(&event, b"w")
                        .and_then(|value| value.parse::<f32>().ok())
                        .map(twips_to_points)
                    {
                        col_widths_points.push(width.max(18.0));
                    }
                }
                b"trHeight" => {
                    current_row_height = attr_value(&event, b"val")
                        .and_then(|value| value.parse::<f32>().ok())
                        .map(twips_to_points);
                }
                b"rFonts" => {
                    if let Some(font) = resolve_font_from_event_without_theme(&event) {
                        apply_resolved_font(&mut current_run_style, font);
                    }
                }
                b"b" => current_run_style.bold = docx_flag(&event, true),
                b"i" => current_run_style.italic = docx_flag(&event, true),
                b"u" => {
                    current_run_style.underline =
                        !matches!(attr_value(&event, b"val").as_deref(), Some("none"))
                }
                b"sz" => {
                    if let Some(value) = attr_value(&event, b"val") {
                        if let Ok(half_points) = value.parse::<f32>() {
                            current_run_style.font_size_points =
                                (half_points / 2.0).clamp(8.0, 72.0);
                        }
                    }
                }
                b"color" => {
                    if let Some(value) = attr_value(&event, b"val") {
                        if let Some(color) = parse_hex_color(&value) {
                            current_run_style.text_color = color;
                        }
                    }
                }
                _ => {}
            },
            Ok(XmlEvent::Empty(event)) => match local_name(event.name().as_ref()) {
                b"tab" if current_cell_runs.is_some() => {
                    if let Some(runs) = current_cell_runs.as_mut() {
                        append_plain(runs, "\t", current_run_style);
                    }
                }
                b"br" | b"cr" if current_cell_runs.is_some() => {
                    if let Some(runs) = current_cell_runs.as_mut() {
                        append_plain(runs, "\n", current_run_style);
                    }
                }
                b"gridCol" => {
                    if let Some(width) = attr_value(&event, b"w")
                        .and_then(|value| value.parse::<f32>().ok())
                        .map(twips_to_points)
                    {
                        col_widths_points.push(width.max(18.0));
                    }
                }
                b"trHeight" => {
                    current_row_height = attr_value(&event, b"val")
                        .and_then(|value| value.parse::<f32>().ok())
                        .map(twips_to_points);
                }
                b"b" => current_run_style.bold = docx_flag(&event, true),
                b"i" => current_run_style.italic = docx_flag(&event, true),
                b"u" => {
                    current_run_style.underline =
                        !matches!(attr_value(&event, b"val").as_deref(), Some("none"))
                }
                b"sz" => {
                    if let Some(value) = attr_value(&event, b"val") {
                        if let Ok(half_points) = value.parse::<f32>() {
                            current_run_style.font_size_points =
                                (half_points / 2.0).clamp(8.0, 72.0);
                        }
                    }
                }
                b"color" => {
                    if let Some(value) = attr_value(&event, b"val") {
                        if let Some(color) = parse_hex_color(&value) {
                            current_run_style.text_color = color;
                        }
                    }
                }
                b"rFonts" => {
                    if let Some(font) = resolve_font_from_event_without_theme(&event) {
                        apply_resolved_font(&mut current_run_style, font);
                    }
                }
                _ => {}
            },
            Ok(XmlEvent::Text(text)) => {
                if current_text {
                    let decoded = text
                        .xml_content()
                        .map_err(|error| format!("failed to decode table text: {error}"))?;
                    if let Some(runs) = current_cell_runs.as_mut() {
                        append_plain(runs, decoded.as_ref(), current_run_style);
                    }
                }
            }
            Ok(XmlEvent::End(event)) => match local_name(event.name().as_ref()) {
                b"t" => current_text = false,
                b"tc" => {
                    let runs = current_cell_runs.take().unwrap_or_default();
                    let cell = TableCell {
                        runs: if runs.is_empty() {
                            vec![TextRun {
                                text: String::new(),
                                style: CharacterStyle::default(),
                            }]
                        } else {
                            runs
                        },
                        images: Vec::new(),
                        col_span: 1,
                        row_span: 1,
                    };
                    if let Some(row) = current_row.as_mut() {
                        row.push(cell);
                    }
                }
                b"tr" => {
                    if let Some(row) = current_row.take() {
                        rows.push(row);
                        row_heights_points.push(current_row_height.unwrap_or(20.0).max(12.0));
                    }
                }
                b"tbl" => break,
                _ => {}
            },
            Ok(XmlEvent::Eof) => break,
            Err(error) => return Err(format!("failed to parse table: {error}")),
            _ => {}
        }
    }

    let num_cols = rows.iter().map(Vec::len).max().unwrap_or(1).max(1);
    if rows.is_empty() {
        rows.push((0..num_cols).map(|_| TableCell::new("")).collect());
        row_heights_points.push(20.0);
    }
    for row in &mut rows {
        while row.len() < num_cols {
            row.push(TableCell::new(""));
        }
    }
    if col_widths_points.len() < num_cols {
        let known_width: f32 = col_widths_points.iter().sum();
        let remaining = (available_width - known_width).max(36.0);
        let fill = remaining / (num_cols - col_widths_points.len()).max(1) as f32;
        col_widths_points.resize(num_cols, fill.max(36.0));
    } else {
        col_widths_points.truncate(num_cols);
    }
    row_heights_points.resize(rows.len(), 20.0);

    Ok(DocumentTable {
        id: table_id,
        rows,
        col_widths_points,
        row_heights_points,
        borders: Default::default(),
    })
}
