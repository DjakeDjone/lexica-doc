use std::{
    collections::{BTreeMap, HashMap, HashSet},
    fmt::Write as _,
    io::{Cursor, Read, Write},
};

use eframe::egui::Color32;
use quick_xml::{events::Event as XmlEvent, Reader};
use serde::Serialize;
use zip::{write::SimpleFileOptions, CompressionMethod, ZipArchive, ZipWriter};

use crate::document::{
    CharacterStyle, DocumentImage, DocumentState, DocumentTable, FontChoice, ImageLayoutMode,
    ImageRendering, LineSpacing, LineSpacingKind, ListKind, PageMargins, PageSize,
    ParagraphAlignment, ParagraphStyle, TableCell, TextRun, WrapMode, OBJECT_REPLACEMENT_CHAR,
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
        ZipArchive::new(cursor).map_err(|error| format!("invalid .odt archive: {error}"))?;

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
    archive: &mut ZipArchive<Cursor<&[u8]>>,
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
                        let mut heading_style = CharacterStyle::default();
                        heading_style.bold = true;
                        heading_style.font_size_points = 20.0;
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
                    let table = parse_odt_table(&mut reader, next_table_id)?;
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
        bytes,
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

fn parse_odt_table(reader: &mut Reader<&[u8]>, id: usize) -> Result<DocumentTable, String> {
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
                b"span" if in_cell_paragraph => style_stack.push(*style_stack.last().unwrap()),
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

pub fn document_to_odt(document: &DocumentState) -> Result<Vec<u8>, String> {
    let mut buffer = Cursor::new(Vec::new());
    let mut zip = ZipWriter::new(&mut buffer);
    let stored = SimpleFileOptions::default().compression_method(CompressionMethod::Stored);
    let deflated = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);

    zip.start_file("mimetype", stored)
        .map_err(|error| format!("failed to write ODT mimetype: {error}"))?;
    zip.write_all(b"application/vnd.oasis.opendocument.text")
        .map_err(|error| format!("failed to write ODT mimetype: {error}"))?;

    let package = OdtPackage::from_document(document);
    write_zip_text(&mut zip, "content.xml", &package.content_xml, deflated)?;
    write_zip_text(&mut zip, "styles.xml", &package.styles_xml, deflated)?;
    write_zip_text(&mut zip, "meta.xml", &package.meta_xml, deflated)?;
    write_zip_text(
        &mut zip,
        "META-INF/manifest.xml",
        &package.manifest_xml,
        deflated,
    )?;

    for image in package.images {
        zip.start_file(&image.path, deflated)
            .map_err(|error| format!("failed to write {}: {error}", image.path))?;
        zip.write_all(&image.bytes)
            .map_err(|error| format!("failed to write {}: {error}", image.path))?;
    }

    zip.finish()
        .map_err(|error| format!("failed to finish ODT package: {error}"))?;
    Ok(buffer.into_inner())
}

struct OdtPackage {
    content_xml: String,
    styles_xml: String,
    meta_xml: String,
    manifest_xml: String,
    images: Vec<OdtImage>,
}

struct OdtImage {
    path: String,
    media_type: &'static str,
    bytes: Vec<u8>,
}

impl OdtPackage {
    fn from_document(document: &DocumentState) -> Self {
        let mut style_builder = StyleBuilder::default();
        let mut body = String::new();
        let mut images = Vec::new();
        let mut active_list = ListKind::None;
        let mut image_paths_by_id = HashMap::<usize, String>::new();

        for paragraph in document.paragraphs() {
            if active_list != paragraph.style.list_kind {
                if active_list != ListKind::None {
                    body.push_str("</text:list>");
                }
                active_list = paragraph.style.list_kind;
                if active_list != ListKind::None {
                    let list_style = match active_list {
                        ListKind::Bullet => "L-bullet",
                        ListKind::Ordered => "L-number",
                        ListKind::None => "",
                    };
                    let _ = write!(body, "<text:list text:style-name=\"{list_style}\">");
                }
            }

            if active_list != ListKind::None {
                body.push_str("<text:list-item>");
            }

            if let Some(table) = paragraph.table.as_ref() {
                write_table(
                    &mut body,
                    table,
                    &mut style_builder,
                    &mut images,
                    &mut image_paths_by_id,
                );
            } else {
                let paragraph_style = style_builder.paragraph_style(paragraph.style);
                let tag = if paragraph
                    .runs
                    .first()
                    .is_some_and(|run| run.style.bold && run.style.font_size_points >= 18.0)
                {
                    "text:h"
                } else {
                    "text:p"
                };
                let _ = write!(body, "<{tag} text:style-name=\"{paragraph_style}\">");
                for run in paragraph.runs {
                    write_run(&mut body, run, &mut style_builder);
                }
                if let Some(image) = paragraph.image.as_ref() {
                    write_image(&mut body, image, &mut images, &mut image_paths_by_id);
                }
                let _ = write!(body, "</{tag}>");
            }

            if active_list != ListKind::None {
                body.push_str("</text:list-item>");
            }
        }
        if active_list != ListKind::None {
            body.push_str("</text:list>");
        }

        let automatic_styles = style_builder.automatic_styles();
        let content_xml = format!(
            "{}<office:document-content {} office:version=\"1.3\"><office:scripts/><office:automatic-styles>{automatic_styles}</office:automatic-styles><office:body><office:text>{body}</office:text></office:body></office:document-content>",
            xml_decl(),
            namespaces()
        );
        let styles_xml = styles_xml(document);
        let meta_xml = format!(
            "{}<office:document-meta {} office:version=\"1.3\"><office:meta><dc:title>{}</dc:title><meta:generator>wors</meta:generator></office:meta></office:document-meta>",
            xml_decl(),
            namespaces(),
            xml_escape(&document.title)
        );
        let manifest_xml = manifest_xml(&images);
        Self {
            content_xml,
            styles_xml,
            meta_xml,
            manifest_xml,
            images,
        }
    }
}

#[derive(Default)]
struct StyleBuilder {
    text_styles: BTreeMap<String, CharacterStyle>,
    paragraph_styles: BTreeMap<String, ParagraphStyle>,
    text_ids: HashMap<StyleKey, String>,
    paragraph_ids: HashMap<ParagraphKey, String>,
}

impl StyleBuilder {
    fn text_style(&mut self, style: CharacterStyle) -> String {
        let key = StyleKey::from(style);
        if let Some(name) = self.text_ids.get(&key) {
            return name.clone();
        }
        let name = format!("T{}", self.text_ids.len() + 1);
        self.text_ids.insert(key, name.clone());
        self.text_styles.insert(name.clone(), style);
        name
    }

    fn paragraph_style(&mut self, style: ParagraphStyle) -> String {
        let key = ParagraphKey::from(style);
        if let Some(name) = self.paragraph_ids.get(&key) {
            return name.clone();
        }
        let name = format!("P{}", self.paragraph_ids.len() + 1);
        self.paragraph_ids.insert(key, name.clone());
        self.paragraph_styles.insert(name.clone(), style);
        name
    }

    fn automatic_styles(&self) -> String {
        let mut xml = String::new();
        xml.push_str("<text:list-style style:name=\"L-bullet\"><text:list-level-style-bullet text:level=\"1\" text:bullet-char=\"&#8226;\"><style:list-level-properties text:min-label-width=\"0.25in\"/></text:list-level-style-bullet></text:list-style>");
        xml.push_str("<text:list-style style:name=\"L-number\"><text:list-level-style-number text:level=\"1\" style:num-format=\"1\"><style:list-level-properties text:min-label-width=\"0.25in\"/></text:list-level-style-number></text:list-style>");
        for (name, style) in &self.paragraph_styles {
            let _ = write!(
                xml,
                "<style:style style:name=\"{}\" style:family=\"paragraph\"><style:paragraph-properties fo:text-align=\"{}\" fo:margin-top=\"{}pt\" fo:margin-bottom=\"{}pt\" {} {}/></style:style>",
                name,
                odt_alignment(style.alignment),
                style.spacing_before_points,
                style.spacing_after_points,
                odt_line_height(style.line_spacing),
                if style.page_break_before { "fo:break-before=\"page\"" } else { "" }
            );
        }
        for (name, style) in &self.text_styles {
            let _ = write!(
                xml,
                "<style:style style:name=\"{}\" style:family=\"text\"><style:text-properties fo:font-size=\"{:.1}pt\" fo:color=\"{}\" {} {} {} {} {} {}/></style:style>",
                name,
                style.font_size_points.max(1.0),
                xml_color(style.text_color),
                if style.bold { "fo:font-weight=\"bold\"" } else { "" },
                if style.italic { "fo:font-style=\"italic\"" } else { "" },
                if style.underline { "style:text-underline-style=\"solid\"" } else { "" },
                if style.strikethrough { "style:text-line-through-style=\"solid\"" } else { "" },
                if style.highlight_color != Color32::TRANSPARENT {
                    format!("fo:background-color=\"{}\"", xml_color(style.highlight_color))
                } else {
                    String::new()
                },
                odt_font_name(*style)
            );
        }
        xml
    }
}

#[derive(Clone, Debug, Hash, PartialEq, Eq)]
struct StyleKey {
    bold: bool,
    italic: bool,
    underline: bool,
    strikethrough: bool,
    font_size_tenths: u16,
    font_choice: FontChoice,
    font_family_name: Option<&'static str>,
    text_color: [u8; 4],
    highlight_color: [u8; 4],
}

impl From<CharacterStyle> for StyleKey {
    fn from(style: CharacterStyle) -> Self {
        Self {
            bold: style.bold,
            italic: style.italic,
            underline: style.underline,
            strikethrough: style.strikethrough,
            font_size_tenths: (style.font_size_points * 10.0).round() as u16,
            font_choice: style.font_choice,
            font_family_name: style.font_family_name,
            text_color: style.text_color.to_array(),
            highlight_color: style.highlight_color.to_array(),
        }
    }
}

#[derive(Clone, Debug, Hash, PartialEq, Eq)]
struct ParagraphKey {
    alignment: ParagraphAlignment,
    page_break_before: bool,
    spacing_before_points: u16,
    spacing_after_points: u16,
    line_spacing_kind: u8,
    line_spacing_tenths: u16,
}

impl From<ParagraphStyle> for ParagraphKey {
    fn from(style: ParagraphStyle) -> Self {
        let kind = match style.line_spacing.kind {
            LineSpacingKind::AutoMultiplier => 0,
            LineSpacingKind::AtLeastPoints => 1,
            LineSpacingKind::ExactPoints => 2,
        };
        Self {
            alignment: style.alignment,
            page_break_before: style.page_break_before,
            spacing_before_points: style.spacing_before_points,
            spacing_after_points: style.spacing_after_points,
            line_spacing_kind: kind,
            line_spacing_tenths: (style.line_spacing.value * 10.0).round() as u16,
        }
    }
}

fn write_run(body: &mut String, run: TextRun, style_builder: &mut StyleBuilder) {
    let text: String = run
        .text
        .chars()
        .filter(|ch| *ch != OBJECT_REPLACEMENT_CHAR)
        .collect();
    if text.is_empty() {
        return;
    }
    let style_name = style_builder.text_style(run.style);
    let _ = write!(body, "<text:span text:style-name=\"{style_name}\">");
    write_text(body, &text);
    body.push_str("</text:span>");
}

fn write_text(body: &mut String, text: &str) {
    for ch in text.chars() {
        match ch {
            '\n' => body.push_str("<text:line-break/>"),
            '\t' => body.push_str("<text:tab/>"),
            ' ' => body.push_str("<text:s/>"),
            _ => body.push_str(&xml_escape_char(ch)),
        }
    }
}

fn write_table(
    body: &mut String,
    table: &DocumentTable,
    style_builder: &mut StyleBuilder,
    images: &mut Vec<OdtImage>,
    image_paths_by_id: &mut HashMap<usize, String>,
) {
    let _ = write!(body, "<table:table table:name=\"Table{}\">", table.id);
    for row in &table.rows {
        body.push_str("<table:table-row>");
        for cell in row {
            body.push_str("<table:table-cell office:value-type=\"string\"><text:p>");
            for run in &cell.runs {
                write_run(body, run.clone(), style_builder);
            }
            for image in &cell.images {
                write_image(body, image, images, image_paths_by_id);
            }
            body.push_str("</text:p></table:table-cell>");
        }
        body.push_str("</table:table-row>");
    }
    body.push_str("</table:table>");
}

fn write_image(
    body: &mut String,
    image: &DocumentImage,
    images: &mut Vec<OdtImage>,
    image_paths_by_id: &mut HashMap<usize, String>,
) {
    let path = image_paths_by_id.entry(image.id).or_insert_with(|| {
        let ext = image_extension(&image.bytes);
        let path = format!("Pictures/image-{}.{}", image.id, ext);
        images.push(OdtImage {
            path: path.clone(),
            media_type: image_mime_type_for_extension(ext),
            bytes: image.bytes.clone(),
        });
        path
    });
    let _ = write!(
        body,
        "<draw:frame draw:name=\"{}\" text:anchor-type=\"as-char\" svg:width=\"{:.2}pt\" svg:height=\"{:.2}pt\"><draw:image xlink:href=\"{}\" xlink:type=\"simple\" xlink:show=\"embed\" xlink:actuate=\"onLoad\"/></draw:frame>",
        xml_escape(&image.alt_text),
        image.width_points.max(1.0),
        image.height_points.max(1.0),
        xml_escape(path)
    );
}

fn styles_xml(document: &DocumentState) -> String {
    format!(
        "{}<office:document-styles {} office:version=\"1.3\"><office:styles><style:style style:name=\"Standard\" style:family=\"paragraph\"/></office:styles><office:automatic-styles><style:page-layout style:name=\"pm1\"><style:page-layout-properties fo:page-width=\"{:.2}pt\" fo:page-height=\"{:.2}pt\" fo:margin-top=\"{:.2}pt\" fo:margin-right=\"{:.2}pt\" fo:margin-bottom=\"{:.2}pt\" fo:margin-left=\"{:.2}pt\"/></style:page-layout></office:automatic-styles><office:master-styles><style:master-page style:name=\"Standard\" style:page-layout-name=\"pm1\"/></office:master-styles></office:document-styles>",
        xml_decl(),
        namespaces(),
        document.page_size.width_points,
        document.page_size.height_points,
        document.margins.top_points,
        document.margins.right_points,
        document.margins.bottom_points,
        document.margins.left_points
    )
}

fn manifest_xml(images: &[OdtImage]) -> String {
    let mut xml = format!(
        "{}<manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\" manifest:version=\"1.3\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.text\"/><manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/><manifest:file-entry manifest:full-path=\"styles.xml\" manifest:media-type=\"text/xml\"/><manifest:file-entry manifest:full-path=\"meta.xml\" manifest:media-type=\"text/xml\"/>",
        xml_decl()
    );
    let mut seen = HashSet::new();
    for image in images {
        if seen.insert(&image.path) {
            let _ = write!(
                xml,
                "<manifest:file-entry manifest:full-path=\"{}\" manifest:media-type=\"{}\"/>",
                xml_escape(&image.path),
                image.media_type
            );
        }
    }
    xml.push_str("</manifest:manifest>");
    xml
}

fn write_zip_text(
    zip: &mut ZipWriter<&mut Cursor<Vec<u8>>>,
    name: &str,
    content: &str,
    options: SimpleFileOptions,
) -> Result<(), String> {
    zip.start_file(name, options)
        .map_err(|error| format!("failed to write {name}: {error}"))?;
    zip.write_all(content.as_bytes())
        .map_err(|error| format!("failed to write {name}: {error}"))
}

fn xml_decl() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8"?>"#
}

fn namespaces() -> &'static str {
    r#"xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0""#
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

fn odt_alignment(alignment: ParagraphAlignment) -> &'static str {
    match alignment {
        ParagraphAlignment::Left => "start",
        ParagraphAlignment::Center => "center",
        ParagraphAlignment::Right => "end",
        ParagraphAlignment::Justify => "justify",
    }
}

fn odt_line_height(line_spacing: LineSpacing) -> String {
    match line_spacing.kind {
        LineSpacingKind::AutoMultiplier => {
            format!(
                "fo:line-height=\"{:.0}%\"",
                line_spacing.value.max(0.1) * 100.0
            )
        }
        LineSpacingKind::AtLeastPoints | LineSpacingKind::ExactPoints => {
            format!("fo:line-height=\"{:.1}pt\"", line_spacing.value.max(1.0))
        }
    }
}

fn odt_font_name(style: CharacterStyle) -> String {
    let name = match style.font_family_name {
        Some("docx-carlito") => "Carlito",
        Some("docx-caladea") => "Caladea",
        Some("docx-liberation-sans") => "Liberation Sans",
        Some("docx-liberation-serif") => "Liberation Serif",
        Some("docx-liberation-mono") => "Liberation Mono",
        Some("docx-comic-sans") => "Comic Neue",
        Some(name) => name,
        None => match style.font_choice {
            FontChoice::Proportional | FontChoice::LiberationSans => "Liberation Sans",
            FontChoice::Monospace | FontChoice::LiberationMono => "Liberation Mono",
            FontChoice::Carlito => "Carlito",
            FontChoice::Caladea => "Caladea",
            FontChoice::LiberationSerif => "Liberation Serif",
            FontChoice::ComicSans => "Comic Neue",
        },
    };
    format!("style:font-name=\"{}\"", xml_escape(name))
}

fn image_extension(bytes: &[u8]) -> &'static str {
    if bytes.starts_with(b"\x89PNG\r\n\x1a\n") {
        "png"
    } else if bytes.starts_with(b"\xff\xd8\xff") {
        "jpg"
    } else if bytes.starts_with(b"GIF87a") || bytes.starts_with(b"GIF89a") {
        "gif"
    } else if bytes.starts_with(b"BM") {
        "bmp"
    } else {
        "bin"
    }
}

fn image_mime_type_for_extension(extension: &str) -> &'static str {
    match extension {
        "png" => "image/png",
        "jpg" | "jpeg" => "image/jpeg",
        "gif" => "image/gif",
        "bmp" => "image/bmp",
        _ => "application/octet-stream",
    }
}

fn xml_color(color: Color32) -> String {
    format!("#{:02x}{:02x}{:02x}", color.r(), color.g(), color.b())
}

fn xml_escape(value: &str) -> String {
    value.chars().map(xml_escape_char).collect()
}

fn xml_escape_char(ch: char) -> String {
    match ch {
        '&' => "&amp;".to_owned(),
        '<' => "&lt;".to_owned(),
        '>' => "&gt;".to_owned(),
        '"' => "&quot;".to_owned(),
        '\'' => "&apos;".to_owned(),
        _ => ch.to_string(),
    }
}

#[cfg(test)]
mod tests {
    use super::{document_to_odt, odt_to_document};
    use crate::document::{
        CharacterStyle, DocumentImage, DocumentState, DocumentTable, ImageLayoutMode,
        ImageRendering, ParagraphAlignment, ParagraphStyle, TableCell, TextRun, WrapMode,
        OBJECT_REPLACEMENT_CHAR,
    };

    const PNG_1X1: &[u8] = &[
        137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6,
        0, 0, 0, 31, 21, 196, 137, 0, 0, 0, 10, 73, 68, 65, 84, 120, 156, 99, 0, 1, 0, 0, 5, 0, 1,
        13, 10, 45, 180, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
    ];

    #[test]
    fn exports_and_imports_basic_odt_package() {
        let mut document = DocumentState::bootstrap();
        document.title = "Roundtrip".to_owned();
        document.replace_with_runs(
            "Roundtrip".to_owned(),
            vec![TextRun {
                text: "Hello ODT".to_owned(),
                style: CharacterStyle {
                    bold: true,
                    italic: true,
                    ..CharacterStyle::default()
                },
            }],
        );
        document.paragraph_styles[0].alignment = ParagraphAlignment::Center;

        let bytes = document_to_odt(&document).expect("odt should export");
        assert!(bytes.starts_with(b"PK"));

        let imported = odt_to_document(&bytes).expect("odt should import");
        assert_eq!(imported.runs[0].text, "Hello ODT");
        assert!(imported.runs[0].style.bold);
        assert!(imported.runs[0].style.italic);
        assert_eq!(
            imported.paragraph_styles[0].alignment,
            ParagraphAlignment::Center
        );
    }

    #[test]
    fn exports_and_imports_odt_images_and_tables() {
        let mut document = DocumentState::bootstrap();
        let table = DocumentTable {
            id: 1,
            rows: vec![vec![TableCell::new("Cell")]],
            col_widths_points: vec![144.0],
            row_heights_points: vec![24.0],
            borders: Default::default(),
        };
        document.runs = vec![
            TextRun {
                text: OBJECT_REPLACEMENT_CHAR.to_string(),
                style: CharacterStyle::default(),
            },
            TextRun {
                text: "\n".to_owned(),
                style: CharacterStyle::default(),
            },
            TextRun {
                text: OBJECT_REPLACEMENT_CHAR.to_string(),
                style: CharacterStyle::default(),
            },
        ];
        document.paragraph_styles = vec![ParagraphStyle::default(), ParagraphStyle::default()];
        document.paragraph_images = vec![Some(test_image()), None];
        document.paragraph_tables = vec![None, Some(table)];

        let bytes = document_to_odt(&document).expect("odt should export");
        let imported = odt_to_document(&bytes).expect("odt should import");

        let image = imported.paragraph_images[0]
            .as_ref()
            .expect("image imports");
        assert_eq!(image.bytes, PNG_1X1);
        let table = imported.paragraph_tables[1]
            .as_ref()
            .expect("table imports");
        assert_eq!(table.rows[0][0].plain_text().trim(), "Cell");
    }

    fn test_image() -> DocumentImage {
        DocumentImage {
            id: 1,
            bytes: PNG_1X1.to_vec(),
            alt_text: "pixel".to_owned(),
            width_points: 12.0,
            height_points: 12.0,
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
        }
    }
}
