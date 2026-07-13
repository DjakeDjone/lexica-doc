use std::{
    collections::HashMap,
    io::{Cursor, Read},
};

use eframe::egui::Color32;
use quick_xml::{events::Event as XmlEvent, Reader};
use serde::Serialize;

use crate::document::{
    CharacterStyle, DocumentImage, DocumentTable, FontChoice, ImageLayoutMode, ImageRendering,
    LineSpacing, LineSpacingKind, ListKind, PageMargins, PageSize, ParagraphAlignment,
    ParagraphStyle, TableCell, TextRun, VerticalAlign, WrapMode, OBJECT_REPLACEMENT_CHAR,
};

#[derive(Debug, Serialize)]
pub struct ImportedOdt {
    pub runs: Vec<TextRun>,
    pub paragraph_styles: Vec<ParagraphStyle>,
    pub paragraph_images: Vec<Option<DocumentImage>>,
    pub paragraph_tables: Vec<Option<DocumentTable>>,
    pub page_size: Option<PageSize>,
    pub margins: Option<PageMargins>,
}

#[derive(Clone, Debug, Default)]
struct OdtStyles {
    text: HashMap<String, CharacterStyle>,
    paragraph: HashMap<String, ParagraphStyle>,
    list: HashMap<String, ListKind>,
    page_size: Option<PageSize>,
    margins: Option<PageMargins>,
}

pub fn odt_to_document(bytes: &[u8]) -> Result<ImportedOdt, String> {
    let cursor = Cursor::new(bytes);
    let mut archive =
        zip::ZipArchive::new(cursor).map_err(|error| format!("invalid .odt archive: {error}"))?;

    let mut content_xml = String::new();
    archive
        .by_name("content.xml")
        .map_err(|error| format!("missing content.xml: {error}"))?
        .read_to_string(&mut content_xml)
        .map_err(|error| format!("failed to read content.xml: {error}"))?;

    let mut styles = OdtStyles::default();
    if let Ok(mut styles_file) = archive.by_name("styles.xml") {
        let mut styles_xml = String::new();
        styles_file
            .read_to_string(&mut styles_xml)
            .map_err(|error| format!("failed to read styles.xml: {error}"))?;
        styles = parse_styles_xml(&styles_xml, styles)?;
    }
    styles = parse_styles_xml(&content_xml, styles)?;

    let pictures = load_pictures(&mut archive)?;
    parse_content_xml(&content_xml, &styles, &pictures)
}

fn load_pictures(
    archive: &mut zip::ZipArchive<Cursor<&[u8]>>,
) -> Result<HashMap<String, Vec<u8>>, String> {
    let mut pictures = HashMap::new();
    for index in 0..archive.len() {
        let mut file = archive
            .by_index(index)
            .map_err(|error| format!("failed to inspect .odt entry: {error}"))?;
        let name = file.name().to_owned();
        if !name.starts_with("Pictures/") || name.ends_with('/') {
            continue;
        }
        let mut bytes = Vec::new();
        file.read_to_end(&mut bytes)
            .map_err(|error| format!("failed to read {name}: {error}"))?;
        pictures.insert(name, bytes);
    }
    Ok(pictures)
}

fn parse_content_xml(
    content_xml: &str,
    styles: &OdtStyles,
    pictures: &HashMap<String, Vec<u8>>,
) -> Result<ImportedOdt, String> {
    let mut reader = Reader::from_str(content_xml);
    reader.config_mut().trim_text(false);

    let mut runs = Vec::new();
    let mut paragraph_styles = Vec::new();
    let mut paragraph_images = Vec::new();
    let mut paragraph_tables = Vec::new();
    let mut current_runs = Vec::<TextRun>::new();
    let mut style_stack = vec![CharacterStyle::default()];
    let mut paragraph_style = ParagraphStyle::default();
    let mut current_image = None;
    let mut in_paragraph = false;
    let mut in_list = Vec::<ListKind>::new();
    let mut next_image_id = 1usize;
    let mut next_table_id = 1usize;

    loop {
        match reader.read_event() {
            Ok(XmlEvent::Start(event)) => match local_name(event.name().as_ref()) {
                b"p" | b"h" => {
                    in_paragraph = true;
                    current_runs.clear();
                    current_image = None;
                    paragraph_style = attr_value(&event, b"style-name")
                        .and_then(|name| styles.paragraph.get(&name).copied())
                        .unwrap_or_default();
                    if let Some(kind) = in_list.last().copied() {
                        paragraph_style.list_kind = kind;
                    }
                    if local_name(event.name().as_ref()) == b"h" {
                        let heading_style = CharacterStyle {
                            bold: true,
                            font_size_points: 20.0,
                            ..Default::default()
                        };
                        style_stack = vec![heading_style];
                    } else {
                        style_stack = vec![CharacterStyle::default()];
                    }
                }
                b"span" if in_paragraph => {
                    let mut style = *style_stack.last().unwrap_or(&CharacterStyle::default());
                    if let Some(style_name) = attr_value(&event, b"style-name") {
                        if let Some(saved) = styles.text.get(&style_name) {
                            style = *saved;
                        }
                    }
                    style_stack.push(style);
                }
                b"line-break" if in_paragraph => append_plain(
                    &mut current_runs,
                    "\n",
                    *style_stack.last().unwrap_or(&CharacterStyle::default()),
                ),
                b"tab" if in_paragraph => append_plain(
                    &mut current_runs,
                    "\t",
                    *style_stack.last().unwrap_or(&CharacterStyle::default()),
                ),
                b"s" if in_paragraph => {
                    let count = attr_value(&event, b"c")
                        .and_then(|value| value.parse::<usize>().ok())
                        .unwrap_or(1);
                    append_plain(
                        &mut current_runs,
                        &" ".repeat(count),
                        *style_stack.last().unwrap_or(&CharacterStyle::default()),
                    );
                }
                b"list" => {
                    let kind = attr_value(&event, b"style-name")
                        .and_then(|name| styles.list.get(&name).copied())
                        .unwrap_or(ListKind::Bullet);
                    in_list.push(kind);
                }
                b"table" => {
                    let table = parse_odt_table(&mut reader, styles, next_table_id)?;
                    next_table_id += 1;
                    push_paragraph(
                        &mut runs,
                        &mut paragraph_styles,
                        &mut paragraph_images,
                        &mut paragraph_tables,
                        vec![TextRun {
                            text: OBJECT_REPLACEMENT_CHAR.to_string(),
                            style: CharacterStyle::default(),
                        }],
                        ParagraphStyle::default(),
                        None,
                        Some(table),
                    );
                }
                b"frame" if in_paragraph => {
                    current_image =
                        parse_frame_image(&event, &mut reader, pictures, next_image_id)?;
                    if current_image.is_some() {
                        next_image_id += 1;
                        append_plain(
                            &mut current_runs,
                            &OBJECT_REPLACEMENT_CHAR.to_string(),
                            *style_stack.last().unwrap_or(&CharacterStyle::default()),
                        );
                    }
                }
                _ => {}
            },
            Ok(XmlEvent::End(event)) => match local_name(event.name().as_ref()) {
                b"p" | b"h" if in_paragraph => {
                    push_paragraph(
                        &mut runs,
                        &mut paragraph_styles,
                        &mut paragraph_images,
                        &mut paragraph_tables,
                        current_runs.clone(),
                        paragraph_style,
                        current_image.take(),
                        None,
                    );
                    in_paragraph = false;
                    style_stack = vec![CharacterStyle::default()];
                }
                b"span" if in_paragraph => {
                    style_stack.pop();
                    if style_stack.is_empty() {
                        style_stack.push(CharacterStyle::default());
                    }
                }
                b"list" => {
                    in_list.pop();
                }
                _ => {}
            },
            Ok(XmlEvent::Text(text)) if in_paragraph => {
                let text = text
                    .decode()
                    .map_err(|error| format!("failed to decode ODT text: {error}"))?;
                append_plain(
                    &mut current_runs,
                    &text,
                    *style_stack.last().unwrap_or(&CharacterStyle::default()),
                );
            }
            Ok(XmlEvent::Empty(event)) if in_paragraph => match local_name(event.name().as_ref()) {
                b"line-break" => append_plain(
                    &mut current_runs,
                    "\n",
                    *style_stack.last().unwrap_or(&CharacterStyle::default()),
                ),
                b"tab" => append_plain(
                    &mut current_runs,
                    "\t",
                    *style_stack.last().unwrap_or(&CharacterStyle::default()),
                ),
                b"s" => {
                    let count = attr_value(&event, b"c")
                        .and_then(|value| value.parse::<usize>().ok())
                        .unwrap_or(1);
                    append_plain(
                        &mut current_runs,
                        &" ".repeat(count),
                        *style_stack.last().unwrap_or(&CharacterStyle::default()),
                    );
                }
                _ => {}
            },
            Ok(XmlEvent::Eof) => break,
            Err(error) => return Err(format!("failed to parse content.xml: {error}")),
            _ => {}
        }
    }

    if runs.is_empty() {
        runs.push(TextRun {
            text: String::new(),
            style: CharacterStyle::default(),
        });
        paragraph_styles.push(ParagraphStyle::default());
        paragraph_images.push(None);
        paragraph_tables.push(None);
    }

    Ok(ImportedOdt {
        runs,
        paragraph_styles,
        paragraph_images,
        paragraph_tables,
        page_size: styles.page_size,
        margins: styles.margins,
    })
}

#[allow(clippy::too_many_arguments)]
fn push_paragraph(
    runs: &mut Vec<TextRun>,
    paragraph_styles: &mut Vec<ParagraphStyle>,
    paragraph_images: &mut Vec<Option<DocumentImage>>,
    paragraph_tables: &mut Vec<Option<DocumentTable>>,
    mut next_runs: Vec<TextRun>,
    paragraph_style: ParagraphStyle,
    image: Option<DocumentImage>,
    table: Option<DocumentTable>,
) {
    if !paragraph_styles.is_empty() {
        append_plain(runs, "\n", CharacterStyle::default());
    }
    if next_runs.is_empty() {
        next_runs.push(TextRun {
            text: String::new(),
            style: CharacterStyle::default(),
        });
    }
    for run in next_runs {
        append_plain(runs, &run.text, run.style);
    }
    paragraph_styles.push(paragraph_style);
    paragraph_images.push(image);
    paragraph_tables.push(table);
}

fn parse_frame_image(
    frame: &quick_xml::events::BytesStart<'_>,
    reader: &mut Reader<&[u8]>,
    pictures: &HashMap<String, Vec<u8>>,
    id: usize,
) -> Result<Option<DocumentImage>, String> {
    let width_points = attr_value(frame, b"width")
        .and_then(|value| parse_length_points(&value))
        .unwrap_or(144.0);
    let height_points = attr_value(frame, b"height")
        .and_then(|value| parse_length_points(&value))
        .unwrap_or(96.0);
    let alt_text = attr_value(frame, b"name").unwrap_or_default();
    let mut href = None::<String>;
    let mut depth = 1usize;

    loop {
        match reader.read_event() {
            Ok(XmlEvent::Start(event)) => {
                if local_name(event.name().as_ref()) == b"image" {
                    href = attr_value(&event, b"href");
                } else {
                    depth += 1;
                }
            }
            Ok(XmlEvent::Empty(event)) => {
                if local_name(event.name().as_ref()) == b"image" {
                    href = attr_value(&event, b"href");
                }
            }
            Ok(XmlEvent::End(_)) => {
                depth -= 1;
                if depth == 0 {
                    break;
                }
            }
            Ok(XmlEvent::Eof) => break,
            Err(error) => return Err(format!("failed to parse ODT image: {error}")),
            _ => {}
        }
    }

    let Some(href) = href else {
        return Ok(None);
    };
    let normalized = href.trim_start_matches("./").to_owned();
    let Some(bytes) = pictures.get(&normalized).cloned() else {
        return Ok(None);
    };

    Ok(Some(DocumentImage {
        id,
        bytes: bytes.into(),
        alt_text,
        width_points,
        height_points,
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
    }))
}

fn parse_odt_table(
    reader: &mut Reader<&[u8]>,
    styles: &OdtStyles,
    id: usize,
) -> Result<DocumentTable, String> {
    let mut rows = Vec::<Vec<TableCell>>::new();
    let mut current_row = Vec::<TableCell>::new();
    let mut current_cell_runs = Vec::<TextRun>::new();
    let mut style_stack = vec![CharacterStyle::default()];
    let mut in_cell = false;
    let mut in_cell_paragraph = false;

    loop {
        match reader.read_event() {
            Ok(XmlEvent::Start(event)) => match local_name(event.name().as_ref()) {
                b"table-row" => current_row.clear(),
                b"table-cell" | b"covered-table-cell" => {
                    in_cell = true;
                    current_cell_runs.clear();
                }
                b"p" | b"h" if in_cell => in_cell_paragraph = true,
                b"span" if in_cell_paragraph => {
                    let mut style = *style_stack.last().unwrap_or(&CharacterStyle::default());
                    if let Some(style_name) = attr_value(&event, b"style-name") {
                        if let Some(saved) = styles.text.get(&style_name) {
                            style = *saved;
                        }
                    }
                    style_stack.push(style);
                }
                b"line-break" if in_cell_paragraph => {
                    append_plain(&mut current_cell_runs, "\n", *style_stack.last().unwrap())
                }
                b"tab" if in_cell_paragraph => {
                    append_plain(&mut current_cell_runs, "\t", *style_stack.last().unwrap())
                }
                b"s" if in_cell_paragraph => append_plain(
                    &mut current_cell_runs,
                    " ",
                    *style_stack.last().unwrap_or(&CharacterStyle::default()),
                ),
                _ => {}
            },
            Ok(XmlEvent::Empty(event)) => match local_name(event.name().as_ref()) {
                b"line-break" if in_cell_paragraph => {
                    append_plain(&mut current_cell_runs, "\n", *style_stack.last().unwrap())
                }
                b"tab" if in_cell_paragraph => {
                    append_plain(&mut current_cell_runs, "\t", *style_stack.last().unwrap())
                }
                b"s" if in_cell_paragraph => append_plain(
                    &mut current_cell_runs,
                    " ",
                    *style_stack.last().unwrap_or(&CharacterStyle::default()),
                ),
                _ => {}
            },
            Ok(XmlEvent::End(event)) => match local_name(event.name().as_ref()) {
                b"table-row" => rows.push(current_row.clone()),
                b"table-cell" | b"covered-table-cell" => {
                    in_cell = false;
                    if current_cell_runs.is_empty() {
                        current_cell_runs.push(TextRun {
                            text: String::new(),
                            style: CharacterStyle::default(),
                        });
                    }
                    current_row.push(TableCell {
                        runs: current_cell_runs.clone(),
                        images: Vec::new(),
                        col_span: 1,
                        row_span: 1,
                    });
                }
                b"p" | b"h" if in_cell_paragraph => {
                    in_cell_paragraph = false;
                    append_plain(&mut current_cell_runs, "\n", CharacterStyle::default());
                }
                b"span" if in_cell_paragraph => {
                    style_stack.pop();
                    if style_stack.is_empty() {
                        style_stack.push(CharacterStyle::default());
                    }
                }
                b"table" => break,
                _ => {}
            },
            Ok(XmlEvent::Text(text)) if in_cell_paragraph => {
                let text = text
                    .decode()
                    .map_err(|error| format!("failed to decode ODT table text: {error}"))?;
                append_plain(&mut current_cell_runs, &text, *style_stack.last().unwrap());
            }
            Ok(XmlEvent::Eof) => break,
            Err(error) => return Err(format!("failed to parse ODT table: {error}")),
            _ => {}
        }
    }

    let cols = rows.iter().map(Vec::len).max().unwrap_or(1);
    for row in &mut rows {
        while row.len() < cols {
            row.push(TableCell::new(""));
        }
    }
    if rows.is_empty() {
        rows.push(vec![TableCell::new("")]);
    }
    Ok(DocumentTable {
        id,
        col_widths_points: vec![72.0; cols],
        row_heights_points: vec![20.0; rows.len()],
        rows,
        borders: Default::default(),
    })
}

fn parse_styles_xml(xml: &str, mut styles: OdtStyles) -> Result<OdtStyles, String> {
    let mut reader = Reader::from_str(xml);
    reader.config_mut().trim_text(false);
    let mut current_name = None::<String>;
    let mut current_family = None::<String>;

    loop {
        match reader.read_event() {
            Ok(XmlEvent::Start(event)) => match local_name(event.name().as_ref()) {
                b"style" | b"default-style" => {
                    current_name = attr_value(&event, b"name");
                    current_family = attr_value(&event, b"family");
                }
                b"text-properties" => {
                    if current_family.as_deref() == Some("text") {
                        if let Some(name) = current_name.as_ref() {
                            let mut style = styles.text.get(name).copied().unwrap_or_default();
                            apply_text_properties(&event, &mut style);
                            styles.text.insert(name.clone(), style);
                        }
                    }
                }
                b"paragraph-properties" => {
                    if current_family.as_deref() == Some("paragraph") {
                        if let Some(name) = current_name.as_ref() {
                            let mut style = styles.paragraph.get(name).copied().unwrap_or_default();
                            apply_paragraph_properties(&event, &mut style);
                            styles.paragraph.insert(name.clone(), style);
                        }
                    }
                }
                b"page-layout-properties" => {
                    if let Some(size) = parse_page_size(&event) {
                        styles.page_size = Some(size);
                    }
                    if let Some(margins) = parse_page_margins(&event) {
                        styles.margins = Some(margins);
                    }
                }
                b"list-style" => {
                    current_name = attr_value(&event, b"name");
                    current_family = Some("list".to_owned());
                }
                b"list-level-style-number" => {
                    if let Some(name) = current_name.as_ref() {
                        styles.list.insert(name.clone(), ListKind::Ordered);
                    }
                }
                b"list-level-style-bullet" => {
                    if let Some(name) = current_name.as_ref() {
                        styles.list.entry(name.clone()).or_insert(ListKind::Bullet);
                    }
                }
                _ => {}
            },
            Ok(XmlEvent::Empty(event)) => match local_name(event.name().as_ref()) {
                b"text-properties" => {
                    if current_family.as_deref() == Some("text") {
                        if let Some(name) = current_name.as_ref() {
                            let mut style = styles.text.get(name).copied().unwrap_or_default();
                            apply_text_properties(&event, &mut style);
                            styles.text.insert(name.clone(), style);
                        }
                    }
                }
                b"paragraph-properties" => {
                    if current_family.as_deref() == Some("paragraph") {
                        if let Some(name) = current_name.as_ref() {
                            let mut style = styles.paragraph.get(name).copied().unwrap_or_default();
                            apply_paragraph_properties(&event, &mut style);
                            styles.paragraph.insert(name.clone(), style);
                        }
                    }
                }
                b"page-layout-properties" => {
                    if let Some(size) = parse_page_size(&event) {
                        styles.page_size = Some(size);
                    }
                    if let Some(margins) = parse_page_margins(&event) {
                        styles.margins = Some(margins);
                    }
                }
                b"list-level-style-number" => {
                    if let Some(name) = current_name.as_ref() {
                        styles.list.insert(name.clone(), ListKind::Ordered);
                    }
                }
                b"list-level-style-bullet" => {
                    if let Some(name) = current_name.as_ref() {
                        styles.list.entry(name.clone()).or_insert(ListKind::Bullet);
                    }
                }
                _ => {}
            },
            Ok(XmlEvent::End(event)) => match local_name(event.name().as_ref()) {
                b"style" | b"default-style" | b"list-style" => {
                    current_name = None;
                    current_family = None;
                }
                _ => {}
            },
            Ok(XmlEvent::Eof) => break,
            Err(error) => return Err(format!("failed to parse ODT styles: {error}")),
            _ => {}
        }
    }
    Ok(styles)
}

fn apply_text_properties(event: &quick_xml::events::BytesStart<'_>, style: &mut CharacterStyle) {
    if attr_value(event, b"font-weight").as_deref() == Some("bold") {
        style.bold = true;
    }
    if attr_value(event, b"font-style").as_deref() == Some("italic") {
        style.italic = true;
    }
    if attr_value(event, b"text-underline-style").is_some_and(|value| value != "none") {
        style.underline = true;
    }
    if attr_value(event, b"text-line-through-style").is_some_and(|value| value != "none") {
        style.strikethrough = true;
    }
    if let Some(size) =
        attr_value(event, b"font-size").and_then(|value| parse_length_points(&value))
    {
        style.font_size_points = size.clamp(8.0, 72.0);
    }
    if let Some(color) = attr_value(event, b"color").and_then(|value| parse_color(&value)) {
        style.text_color = color;
    }
    if let Some(color) =
        attr_value(event, b"background-color").and_then(|value| parse_color(&value))
    {
        style.highlight_color = color;
    }
    if let Some(font) = attr_value(event, b"font-name") {
        apply_font_name(style, &font);
    }
    if let Some(position) = attr_value(event, b"text-position") {
        style.vertical_align = parse_text_position(&position);
    }
}

fn apply_paragraph_properties(
    event: &quick_xml::events::BytesStart<'_>,
    style: &mut ParagraphStyle,
) {
    if let Some(align) = attr_value(event, b"text-align") {
        style.alignment = match align.as_str() {
            "center" => ParagraphAlignment::Center,
            "end" | "right" => ParagraphAlignment::Right,
            "justify" => ParagraphAlignment::Justify,
            _ => ParagraphAlignment::Left,
        };
    }
    if attr_value(event, b"break-before").is_some_and(|value| value == "page") {
        style.page_break_before = true;
    }
    if let Some(value) = attr_value(event, b"margin-left").and_then(|v| parse_length_points(&v)) {
        style.left_indent_points = value;
    }
    if let Some(value) = attr_value(event, b"margin-right").and_then(|v| parse_length_points(&v)) {
        style.right_indent_points = value;
    }
    if let Some(value) = attr_value(event, b"text-indent").and_then(|v| parse_length_points(&v)) {
        style.first_line_indent_points = value;
    }
    if let Some(value) = attr_value(event, b"margin-top").and_then(|v| parse_length_points(&v)) {
        style.spacing_before_points = value.max(0.0).round() as u16;
    }
    if let Some(value) = attr_value(event, b"margin-bottom").and_then(|v| parse_length_points(&v)) {
        style.spacing_after_points = value.max(0.0).round() as u16;
    }
    if let Some(value) = attr_value(event, b"line-height") {
        style.line_spacing = parse_line_height(&value);
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

fn attr_value(event: &quick_xml::events::BytesStart<'_>, name: &[u8]) -> Option<String> {
    event
        .attributes()
        .with_checks(false)
        .filter_map(Result::ok)
        .find(|attr| local_name(attr.key.as_ref()) == name)
        .and_then(|attr| {
            String::from_utf8(attr.value.as_ref().to_vec())
                .ok()
                .map(|value| value.replace("&quot;", "\""))
        })
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn parse_length_points(value: &str) -> Option<f32> {
    let value = value.trim();
    if value.is_empty() {
        return None;
    }
    let split = value
        .find(|ch: char| !(ch.is_ascii_digit() || ch == '.' || ch == '-'))
        .unwrap_or(value.len());
    let number = value[..split].parse::<f32>().ok()?;
    let unit = value[split..].trim();
    Some(match unit {
        "in" => number * 72.0,
        "cm" => number * 72.0 / 2.54,
        "mm" => number * 72.0 / 25.4,
        "pt" | "" => number,
        "pc" => number * 12.0,
        "px" => number * 72.0 / 96.0,
        _ => number,
    })
}

fn parse_line_height(value: &str) -> LineSpacing {
    if let Some(percent) = value.strip_suffix('%') {
        return LineSpacing {
            kind: LineSpacingKind::AutoMultiplier,
            value: percent.parse::<f32>().unwrap_or(100.0).max(10.0) / 100.0,
        };
    }
    LineSpacing {
        kind: LineSpacingKind::ExactPoints,
        value: parse_length_points(value).unwrap_or(12.0).max(1.0),
    }
}

fn parse_page_size(event: &quick_xml::events::BytesStart<'_>) -> Option<PageSize> {
    Some(PageSize {
        width_points: attr_value(event, b"page-width").and_then(|v| parse_length_points(&v))?,
        height_points: attr_value(event, b"page-height").and_then(|v| parse_length_points(&v))?,
    })
}

fn parse_page_margins(event: &quick_xml::events::BytesStart<'_>) -> Option<PageMargins> {
    let margin = attr_value(event, b"margin").and_then(|v| parse_length_points(&v));
    Some(PageMargins {
        top_points: attr_value(event, b"margin-top")
            .and_then(|v| parse_length_points(&v))
            .or(margin)?,
        right_points: attr_value(event, b"margin-right")
            .and_then(|v| parse_length_points(&v))
            .or(margin)?,
        bottom_points: attr_value(event, b"margin-bottom")
            .and_then(|v| parse_length_points(&v))
            .or(margin)?,
        left_points: attr_value(event, b"margin-left")
            .and_then(|v| parse_length_points(&v))
            .or(margin)?,
    })
}

fn parse_color(value: &str) -> Option<Color32> {
    let hex = value.trim().trim_start_matches('#');
    if hex.len() != 6 {
        return None;
    }
    let r = u8::from_str_radix(&hex[0..2], 16).ok()?;
    let g = u8::from_str_radix(&hex[2..4], 16).ok()?;
    let b = u8::from_str_radix(&hex[4..6], 16).ok()?;
    Some(Color32::from_rgb(r, g, b))
}

fn parse_text_position(value: &str) -> VerticalAlign {
    let normalized = value.trim().to_ascii_lowercase();
    if normalized.starts_with("super") {
        VerticalAlign::Superscript
    } else if normalized.starts_with("sub") {
        VerticalAlign::Subscript
    } else {
        VerticalAlign::Baseline
    }
}

fn apply_font_name(style: &mut CharacterStyle, font: &str) {
    let normalized = font.to_ascii_lowercase();
    if normalized.contains("liberation mono") || normalized.contains("courier") {
        style.font_choice = FontChoice::LiberationMono;
        style.font_family_name = Some("docx-liberation-mono");
    } else if normalized.contains("liberation serif") || normalized.contains("times") {
        style.font_choice = FontChoice::LiberationSerif;
        style.font_family_name = Some("docx-liberation-serif");
    } else if normalized.contains("liberation sans") || normalized.contains("arial") {
        style.font_choice = FontChoice::LiberationSans;
        style.font_family_name = Some("docx-liberation-sans");
    } else if normalized.contains("carlito") || normalized.contains("calibri") {
        style.font_choice = FontChoice::Carlito;
        style.font_family_name = Some("docx-carlito");
    } else if normalized.contains("caladea") || normalized.contains("cambria") {
        style.font_choice = FontChoice::Caladea;
        style.font_family_name = Some("docx-caladea");
    } else if normalized.contains("comic") {
        style.font_choice = FontChoice::ComicSans;
        style.font_family_name = Some("docx-comic-sans");
    }
}
