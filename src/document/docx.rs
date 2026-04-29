use std::{
    collections::{HashMap, HashSet},
    io::{Cursor, Read},
    path::Path,
};

use eframe::egui::Color32;
use quick_xml::{events::Event as XmlEvent, Reader};
use zip::ZipArchive;

use crate::document::{
    CharacterStyle, DocumentImage, DocumentTable, FontChoice, LineSpacing, LineSpacingKind,
    ListKind, PageMargins, PageSize, ParagraphAlignment, ParagraphStyle, TableCell, TextRun,
    OBJECT_REPLACEMENT_CHAR,
};
use serde::Serialize;

const DOCX_CARLITO: &str = "docx-carlito";
const DOCX_CALADEA: &str = "docx-caladea";
const DOCX_LIBERATION_SANS: &str = "docx-liberation-sans";
const DOCX_LIBERATION_SERIF: &str = "docx-liberation-serif";
const DOCX_LIBERATION_MONO: &str = "docx-liberation-mono";

#[derive(Debug, Serialize)]
pub struct ImportedDocx {
    pub runs: Vec<TextRun>,
    pub paragraph_styles: Vec<ParagraphStyle>,
    pub paragraph_images: Vec<Option<DocumentImage>>,
    pub paragraph_tables: Vec<Option<DocumentTable>>,
    pub page_size: Option<PageSize>,
    pub margins: Option<PageMargins>,
}

pub fn docx_to_document(bytes: &[u8]) -> Result<ImportedDocx, String> {
    let cursor = Cursor::new(bytes);
    let mut archive =
        ZipArchive::new(cursor).map_err(|error| format!("invalid .docx archive: {error}"))?;
    let mut document_xml = String::new();
    archive
        .by_name("word/document.xml")
        .map_err(|error| format!("missing word/document.xml: {error}"))?
        .read_to_string(&mut document_xml)
        .map_err(|error| format!("failed to read word/document.xml: {error}"))?;

    let numbering = load_numbering_definitions(&mut archive)?;
    let theme_fonts = load_theme_fonts(&mut archive)?;
    let styles = load_styles(&mut archive, &theme_fonts)?;
    let relationships = load_document_relationships(&mut archive)?;
    let media = load_media_store(&mut archive, &relationships)?;
    parse_document_xml(
        &document_xml,
        &numbering,
        &styles,
        &theme_fonts,
        &relationships,
        &media,
    )
}

fn load_numbering_definitions(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
) -> Result<NumberingDefinitions, String> {
    let Ok(mut numbering_file) = archive.by_name("word/numbering.xml") else {
        return Ok(NumberingDefinitions::default());
    };

    let mut numbering_xml = String::new();
    numbering_file
        .read_to_string(&mut numbering_xml)
        .map_err(|error| format!("failed to read word/numbering.xml: {error}"))?;
    parse_numbering_xml(&numbering_xml)
}

fn load_document_relationships(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
) -> Result<DocumentRelationships, String> {
    let Ok(mut relationships_file) = archive.by_name("word/_rels/document.xml.rels") else {
        return Ok(DocumentRelationships::default());
    };

    let mut relationships_xml = String::new();
    relationships_file
        .read_to_string(&mut relationships_xml)
        .map_err(|error| format!("failed to read word/_rels/document.xml.rels: {error}"))?;
    parse_document_relationships(&relationships_xml)
}

fn load_styles(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
    theme_fonts: &ThemeFonts,
) -> Result<DocxStyles, String> {
    let Ok(mut styles_file) = archive.by_name("word/styles.xml") else {
        return Ok(DocxStyles::default());
    };

    let mut styles_xml = String::new();
    styles_file
        .read_to_string(&mut styles_xml)
        .map_err(|error| format!("failed to read word/styles.xml: {error}"))?;
    parse_styles_xml(&styles_xml, theme_fonts)
}

fn load_theme_fonts(archive: &mut ZipArchive<Cursor<&[u8]>>) -> Result<ThemeFonts, String> {
    let Ok(mut theme_file) = archive.by_name("word/theme/theme1.xml") else {
        return Ok(ThemeFonts::default());
    };

    let mut theme_xml = String::new();
    theme_file
        .read_to_string(&mut theme_xml)
        .map_err(|error| format!("failed to read word/theme/theme1.xml: {error}"))?;
    parse_theme_xml(&theme_xml)
}

fn load_media_store(
    archive: &mut ZipArchive<Cursor<&[u8]>>,
    relationships: &DocumentRelationships,
) -> Result<HashMap<String, Vec<u8>>, String> {
    let mut media = HashMap::new();

    for target in HashSet::<String>::from_iter(relationships.image_targets.values().cloned()) {
        let Ok(mut file) = archive.by_name(&target) else {
            continue;
        };
        let mut bytes = Vec::new();
        file.read_to_end(&mut bytes)
            .map_err(|error| format!("failed to read {target}: {error}"))?;
        media.insert(target, bytes);
    }

    Ok(media)
}

fn parse_document_xml(
    document_xml: &str,
    numbering: &NumberingDefinitions,
    styles: &DocxStyles,
    theme_fonts: &ThemeFonts,
    relationships: &DocumentRelationships,
    media: &HashMap<String, Vec<u8>>,
) -> Result<ImportedDocx, String> {
    let mut reader = Reader::from_str(document_xml);
    reader.config_mut().trim_text(false);

    let mut runs = Vec::new();
    let mut paragraph_styles = Vec::new();
    let mut paragraph_images = Vec::new();
    let mut paragraph_tables = Vec::new();
    let mut paragraph_run_style = styles.default_run_style();
    let mut run_style = paragraph_run_style;
    let mut paragraph_style = styles.default_paragraph_style();
    let mut current_paragraph_image = None;
    let mut in_text = false;
    let mut current_num_id = None;
    let mut current_ilvl = None;
    let mut current_drawing = None::<DrawingState>;
    let mut page_size = None;
    let mut margins = None;
    let mut next_image_id = 1usize;
    let mut next_table_id = 1usize;

    loop {
        match reader.read_event() {
            Ok(XmlEvent::Start(event)) => match local_name(event.name().as_ref()) {
                b"p" => {
                    if !paragraph_styles.is_empty() {
                        append_plain(&mut runs, "\n", CharacterStyle::default());
                    }
                    paragraph_style = styles.default_paragraph_style();
                    paragraph_run_style = styles.default_run_style();
                    current_paragraph_image = None;
                    current_num_id = None;
                    current_ilvl = None;
                }
                b"tbl" => {
                    let available_width = page_size.unwrap_or_else(PageSize::a4).width_points
                        - margins.unwrap_or_else(PageMargins::standard).left_points
                        - margins.unwrap_or_else(PageMargins::standard).right_points;
                    let table = parse_docx_table(&mut reader, next_table_id, available_width)?;
                    next_table_id += 1;
                    if !paragraph_styles.is_empty() {
                        append_plain(&mut runs, "\n", CharacterStyle::default());
                    }
                    append_plain(
                        &mut runs,
                        &OBJECT_REPLACEMENT_CHAR.to_string(),
                        CharacterStyle::default(),
                    );
                    paragraph_styles.push(ParagraphStyle::default());
                    paragraph_images.push(None);
                    paragraph_tables.push(Some(table));
                }
                b"r" => {
                    run_style = paragraph_run_style;
                }
                b"t" => in_text = true,
                b"br" | b"cr" => append_plain(&mut runs, "\n", run_style),
                b"tab" => append_plain(&mut runs, "\t", run_style),
                b"pStyle" => {
                    if let Some(style_id) = attr_value(&event, b"val") {
                        styles.apply_paragraph_style(&style_id, &mut paragraph_style);
                        paragraph_run_style = styles.run_style_for_paragraph(&style_id);
                    }
                }
                b"rStyle" => {
                    if let Some(style_id) = attr_value(&event, b"val") {
                        run_style = styles.apply_run_style(&style_id, run_style);
                    }
                }
                b"rFonts" => {
                    if let Some(font) = resolve_rfonts(&event, theme_fonts) {
                        apply_resolved_font(&mut run_style, font);
                    }
                }
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
                b"jc" => {
                    paragraph_style.alignment = paragraph_alignment_for(
                        attr_value(&event, b"val").as_deref().unwrap_or_default(),
                    );
                }
                b"spacing" => apply_spacing(&event, &mut paragraph_style),
                b"pageBreakBefore" => paragraph_style.page_break_before = docx_flag(&event, true),
                b"numId" => current_num_id = attr_value(&event, b"val"),
                b"ilvl" => current_ilvl = attr_value(&event, b"val"),
                b"pgSz" => page_size = parse_page_size(&event),
                b"pgMar" => margins = parse_page_margins(&event),
                b"drawing" | b"pict" => current_drawing = Some(DrawingState::default()),
                b"anchor" => {
                    if let Some(drawing) = current_drawing.as_mut() {
                        drawing.is_anchor = true;
                        // Parse distance-from-text attributes on the anchor element
                        drawing.distance_from_text = parse_anchor_distance(&event);
                        // behindDoc attribute determines behind-text vs normal
                        if attr_value(&event, b"behindDoc").as_deref() == Some("1") {
                            drawing.wrap_mode = Some(crate::document::WrapMode::BehindText);
                        }
                    }
                }
                b"inline" => {
                    if let Some(drawing) = current_drawing.as_mut() {
                        drawing.is_anchor = false;
                    }
                }
                b"wrapSquare" => {
                    if let Some(drawing) = current_drawing.as_mut() {
                        drawing.wrap_mode = Some(crate::document::WrapMode::Square);
                    }
                }
                b"wrapTight" => {
                    if let Some(drawing) = current_drawing.as_mut() {
                        drawing.wrap_mode = Some(crate::document::WrapMode::Tight);
                    }
                }
                b"wrapThrough" => {
                    if let Some(drawing) = current_drawing.as_mut() {
                        drawing.wrap_mode = Some(crate::document::WrapMode::Through);
                    }
                }
                b"wrapTopAndBottom" => {
                    if let Some(drawing) = current_drawing.as_mut() {
                        drawing.wrap_mode = Some(crate::document::WrapMode::TopAndBottom);
                    }
                }
                b"wrapNone" => {
                    if let Some(drawing) = current_drawing.as_mut() {
                        // wrapNone means no text wrapping — could be behind or in-front
                        if drawing.wrap_mode.is_none() {
                            drawing.wrap_mode = Some(crate::document::WrapMode::InFrontOfText);
                        }
                    }
                }
                b"docPr" => {
                    if let Some(drawing) = current_drawing.as_mut() {
                        drawing.alt_text = attr_value(&event, b"descr")
                            .or_else(|| attr_value(&event, b"name"))
                            .or_else(|| attr_value(&event, b"title"));
                    }
                }
                b"extent" => {
                    if let Some(drawing) = current_drawing.as_mut() {
                        drawing.size_points = parse_emu_extent(&event);
                    }
                }
                b"blip" => {
                    if let Some(drawing) = current_drawing.as_mut() {
                        drawing.rel_id = attr_value(&event, b"embed");
                    }
                }
                b"imagedata" => {
                    if let Some(drawing) = current_drawing.as_mut() {
                        drawing.rel_id = attr_value(&event, b"id");
                        drawing.alt_text = drawing
                            .alt_text
                            .clone()
                            .or_else(|| attr_value(&event, b"title"));
                    }
                }
                _ => {}
            },
            Ok(XmlEvent::Empty(event)) => match local_name(event.name().as_ref()) {
                b"br" | b"cr" => append_plain(&mut runs, "\n", run_style),
                b"tab" => append_plain(&mut runs, "\t", run_style),
                b"pStyle" => {
                    if let Some(style_id) = attr_value(&event, b"val") {
                        styles.apply_paragraph_style(&style_id, &mut paragraph_style);
                        paragraph_run_style = styles.run_style_for_paragraph(&style_id);
                    }
                }
                b"rStyle" => {
                    if let Some(style_id) = attr_value(&event, b"val") {
                        run_style = styles.apply_run_style(&style_id, run_style);
                    }
                }
                b"rFonts" => {
                    if let Some(font) = resolve_rfonts(&event, theme_fonts) {
                        apply_resolved_font(&mut run_style, font);
                    }
                }
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
                b"jc" => {
                    paragraph_style.alignment = paragraph_alignment_for(
                        attr_value(&event, b"val").as_deref().unwrap_or_default(),
                    );
                }
                b"spacing" => apply_spacing(&event, &mut paragraph_style),
                b"pageBreakBefore" => paragraph_style.page_break_before = docx_flag(&event, true),
                b"numId" => current_num_id = attr_value(&event, b"val"),
                b"ilvl" => current_ilvl = attr_value(&event, b"val"),
                b"pgSz" => page_size = parse_page_size(&event),
                b"pgMar" => margins = parse_page_margins(&event),
                b"wrapSquare" => {
                    if let Some(drawing) = current_drawing.as_mut() {
                        drawing.wrap_mode = Some(crate::document::WrapMode::Square);
                    }
                }
                b"wrapTight" => {
                    if let Some(drawing) = current_drawing.as_mut() {
                        drawing.wrap_mode = Some(crate::document::WrapMode::Tight);
                    }
                }
                b"wrapThrough" => {
                    if let Some(drawing) = current_drawing.as_mut() {
                        drawing.wrap_mode = Some(crate::document::WrapMode::Through);
                    }
                }
                b"wrapTopAndBottom" => {
                    if let Some(drawing) = current_drawing.as_mut() {
                        drawing.wrap_mode = Some(crate::document::WrapMode::TopAndBottom);
                    }
                }
                b"wrapNone" => {
                    if let Some(drawing) = current_drawing.as_mut() {
                        if drawing.wrap_mode.is_none() {
                            drawing.wrap_mode = Some(crate::document::WrapMode::InFrontOfText);
                        }
                    }
                }
                b"docPr" => {
                    if let Some(drawing) = current_drawing.as_mut() {
                        drawing.alt_text = attr_value(&event, b"descr")
                            .or_else(|| attr_value(&event, b"name"))
                            .or_else(|| attr_value(&event, b"title"));
                    }
                }
                b"extent" => {
                    if let Some(drawing) = current_drawing.as_mut() {
                        drawing.size_points = parse_emu_extent(&event);
                    }
                }
                b"blip" => {
                    if let Some(drawing) = current_drawing.as_mut() {
                        drawing.rel_id = attr_value(&event, b"embed");
                    }
                }
                b"imagedata" => {
                    if let Some(drawing) = current_drawing.as_mut() {
                        drawing.rel_id = attr_value(&event, b"id");
                        drawing.alt_text = drawing
                            .alt_text
                            .clone()
                            .or_else(|| attr_value(&event, b"title"));
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
            Ok(XmlEvent::End(event)) => match local_name(event.name().as_ref()) {
                b"t" => in_text = false,
                b"drawing" | b"pict" => {
                    if current_paragraph_image.is_none() {
                        if let Some(image) = resolve_drawing(
                            current_drawing.take(),
                            relationships,
                            media,
                            &mut next_image_id,
                        ) {
                            current_paragraph_image = Some(image);
                            append_plain(
                                &mut runs,
                                &OBJECT_REPLACEMENT_CHAR.to_string(),
                                CharacterStyle::default(),
                            );
                        }
                    } else {
                        current_drawing = None;
                    }
                }
                b"p" => {
                    paragraph_style.list_kind =
                        numbering.lookup(current_num_id.as_deref(), current_ilvl.as_deref());
                    paragraph_styles.push(paragraph_style);
                    paragraph_images.push(current_paragraph_image.clone());
                    paragraph_tables.push(None);
                }
                _ => {}
            },
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

    if paragraph_styles.is_empty() {
        paragraph_styles.push(ParagraphStyle::default());
    }
    if paragraph_images.is_empty() {
        paragraph_images.push(None);
    }
    if paragraph_tables.is_empty() {
        paragraph_tables.push(None);
    }

    Ok(ImportedDocx {
        runs,
        paragraph_styles,
        paragraph_images,
        paragraph_tables,
        page_size,
        margins,
    })
}

fn parse_docx_table(
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

fn resolve_font_from_event_without_theme(
    event: &quick_xml::events::BytesStart<'_>,
) -> Option<ResolvedFont> {
    for key in [b"ascii".as_slice(), b"hAnsi", b"cs", b"eastAsia"] {
        if let Some(value) = attr_value(event, key).filter(|value| !value.trim().is_empty()) {
            return Some(resolve_font_name(&value));
        }
    }
    None
}

#[derive(Default)]
struct NumberingDefinitions {
    num_to_abstract: HashMap<String, String>,
    level_kinds: HashMap<(String, String), ListKind>,
}

impl NumberingDefinitions {
    fn lookup(&self, num_id: Option<&str>, ilvl: Option<&str>) -> ListKind {
        let Some(num_id) = num_id else {
            return ListKind::None;
        };
        if num_id == "0" {
            return ListKind::None;
        }

        let Some(abstract_id) = self.num_to_abstract.get(num_id) else {
            return ListKind::None;
        };
        let level = ilvl.unwrap_or("0");
        self.level_kinds
            .get(&(abstract_id.clone(), level.to_owned()))
            .copied()
            .or_else(|| {
                self.level_kinds
                    .get(&(abstract_id.clone(), "0".to_owned()))
                    .copied()
            })
            .unwrap_or(ListKind::None)
    }
}

#[derive(Default)]
struct DocumentRelationships {
    image_targets: HashMap<String, String>,
}

#[derive(Default)]
struct DrawingState {
    rel_id: Option<String>,
    alt_text: Option<String>,
    size_points: Option<(f32, f32)>,
    is_anchor: bool,
    wrap_mode: Option<crate::document::WrapMode>,
    distance_from_text: Option<crate::document::DistanceFromText>,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
struct ResolvedFont {
    font_family_name: Option<&'static str>,
    font_choice: FontChoice,
}

#[derive(Default)]
struct ThemeFonts {
    major_latin: Option<String>,
    minor_latin: Option<String>,
}

#[derive(Clone, Copy, Default)]
struct CharacterStylePatch {
    bold: Option<bool>,
    italic: Option<bool>,
    underline: Option<bool>,
    strikethrough: Option<bool>,
    font_size_points: Option<f32>,
    font: Option<ResolvedFont>,
    text_color: Option<Color32>,
    highlight_color: Option<Color32>,
}

impl CharacterStylePatch {
    fn apply(self, mut style: CharacterStyle) -> CharacterStyle {
        if let Some(value) = self.bold {
            style.bold = value;
        }
        if let Some(value) = self.italic {
            style.italic = value;
        }
        if let Some(value) = self.underline {
            style.underline = value;
        }
        if let Some(value) = self.strikethrough {
            style.strikethrough = value;
        }
        if let Some(value) = self.font_size_points {
            style.font_size_points = value;
        }
        if let Some(value) = self.font {
            style.font_family_name = value.font_family_name;
            style.font_choice = value.font_choice;
        }
        if let Some(value) = self.text_color {
            style.text_color = value;
        }
        if let Some(value) = self.highlight_color {
            style.highlight_color = value;
        }
        style
    }

    fn overlay(&mut self, other: Self) {
        if other.bold.is_some() {
            self.bold = other.bold;
        }
        if other.italic.is_some() {
            self.italic = other.italic;
        }
        if other.underline.is_some() {
            self.underline = other.underline;
        }
        if other.strikethrough.is_some() {
            self.strikethrough = other.strikethrough;
        }
        if other.font_size_points.is_some() {
            self.font_size_points = other.font_size_points;
        }
        if other.font.is_some() {
            self.font = other.font;
        }
        if other.text_color.is_some() {
            self.text_color = other.text_color;
        }
        if other.highlight_color.is_some() {
            self.highlight_color = other.highlight_color;
        }
    }
}

#[derive(Clone, Copy, Default)]
struct ParagraphStylePatch {
    alignment: Option<ParagraphAlignment>,
    page_break_before: Option<bool>,
    spacing_before_points: Option<u16>,
    spacing_after_points: Option<u16>,
    line_spacing: Option<LineSpacing>,
}

impl ParagraphStylePatch {
    fn apply(self, mut style: ParagraphStyle) -> ParagraphStyle {
        if let Some(value) = self.alignment {
            style.alignment = value;
        }
        if let Some(value) = self.page_break_before {
            style.page_break_before = value;
        }
        if let Some(value) = self.spacing_before_points {
            style.spacing_before_points = value;
        }
        if let Some(value) = self.spacing_after_points {
            style.spacing_after_points = value;
        }
        if let Some(value) = self.line_spacing {
            style.line_spacing = value;
        }
        style
    }

    fn overlay(&mut self, other: Self) {
        if other.alignment.is_some() {
            self.alignment = other.alignment;
        }
        if other.page_break_before.is_some() {
            self.page_break_before = other.page_break_before;
        }
        if other.spacing_before_points.is_some() {
            self.spacing_before_points = other.spacing_before_points;
        }
        if other.spacing_after_points.is_some() {
            self.spacing_after_points = other.spacing_after_points;
        }
        if other.line_spacing.is_some() {
            self.line_spacing = other.line_spacing;
        }
    }
}

#[derive(Clone, Default)]
struct RawParagraphStyleDefinition {
    based_on: Option<String>,
    paragraph: ParagraphStylePatch,
    run: CharacterStylePatch,
}

#[derive(Clone, Default)]
struct RawCharacterStyleDefinition {
    based_on: Option<String>,
    run: CharacterStylePatch,
}

#[derive(Clone, Copy, Default)]
struct ResolvedParagraphStyle {
    paragraph: ParagraphStylePatch,
    run: CharacterStylePatch,
}

#[derive(Default)]
struct RawDocxStyles {
    default_paragraph: ParagraphStylePatch,
    default_run: CharacterStylePatch,
    paragraph_styles: HashMap<String, RawParagraphStyleDefinition>,
    character_styles: HashMap<String, RawCharacterStyleDefinition>,
}

#[derive(Default)]
struct DocxStyles {
    default_paragraph: ParagraphStyle,
    default_run: CharacterStyle,
    paragraph_styles: HashMap<String, ResolvedParagraphStyle>,
    character_styles: HashMap<String, CharacterStylePatch>,
}

impl DocxStyles {
    fn default_paragraph_style(&self) -> ParagraphStyle {
        self.default_paragraph
    }

    fn default_run_style(&self) -> CharacterStyle {
        self.default_run
    }

    fn apply_paragraph_style(&self, style_id: &str, style: &mut ParagraphStyle) {
        if let Some(resolved) = self.paragraph_styles.get(style_id) {
            *style = resolved.paragraph.apply(*style);
        }
    }

    fn run_style_for_paragraph(&self, style_id: &str) -> CharacterStyle {
        self.paragraph_styles
            .get(style_id)
            .map(|resolved| resolved.run.apply(self.default_run_style()))
            .unwrap_or_else(|| self.default_run_style())
    }

    fn apply_run_style(&self, style_id: &str, base: CharacterStyle) -> CharacterStyle {
        self.character_styles
            .get(style_id)
            .copied()
            .map(|style| style.apply(base))
            .unwrap_or(base)
    }
}

fn parse_theme_xml(theme_xml: &str) -> Result<ThemeFonts, String> {
    let mut reader = Reader::from_str(theme_xml);
    reader.config_mut().trim_text(false);

    let mut theme_fonts = ThemeFonts::default();
    let mut in_major_font = false;
    let mut in_minor_font = false;

    loop {
        match reader.read_event() {
            Ok(XmlEvent::Start(event)) | Ok(XmlEvent::Empty(event)) => {
                match local_name(event.name().as_ref()) {
                    b"majorFont" => in_major_font = true,
                    b"minorFont" => in_minor_font = true,
                    b"latin" => {
                        let typeface = attr_value(&event, b"typeface")
                            .filter(|value| !value.trim().is_empty());
                        if in_major_font {
                            theme_fonts.major_latin = typeface;
                        } else if in_minor_font {
                            theme_fonts.minor_latin = typeface;
                        }
                    }
                    _ => {}
                }
            }
            Ok(XmlEvent::End(event)) => match local_name(event.name().as_ref()) {
                b"majorFont" => in_major_font = false,
                b"minorFont" => in_minor_font = false,
                _ => {}
            },
            Ok(XmlEvent::Eof) => break,
            Err(error) => return Err(format!("failed to parse word/theme/theme1.xml: {error}")),
            _ => {}
        }
    }

    Ok(theme_fonts)
}

fn apply_resolved_font(style: &mut CharacterStyle, font: ResolvedFont) {
    style.font_family_name = font.font_family_name;
    style.font_choice = font.font_choice;
}

fn resolve_rfonts(
    event: &quick_xml::events::BytesStart<'_>,
    theme_fonts: &ThemeFonts,
) -> Option<ResolvedFont> {
    for key in [b"ascii".as_slice(), b"hAnsi", b"cs", b"eastAsia"] {
        if let Some(value) = attr_value(event, key).filter(|value| !value.trim().is_empty()) {
            return Some(resolve_font_name(&value));
        }
    }

    for key in [
        b"asciiTheme".as_slice(),
        b"hAnsiTheme",
        b"csTheme",
        b"eastAsiaTheme",
    ] {
        if let Some(value) = attr_value(event, key).filter(|value| !value.trim().is_empty()) {
            return resolve_theme_font(&value, theme_fonts);
        }
    }

    None
}

fn resolve_theme_font(slot: &str, theme_fonts: &ThemeFonts) -> Option<ResolvedFont> {
    let font_name = match slot {
        "majorAscii" | "majorHAnsi" | "majorBidi" | "majorEastAsia" => {
            theme_fonts.major_latin.as_deref()
        }
        "minorAscii" | "minorHAnsi" | "minorBidi" | "minorEastAsia" => {
            theme_fonts.minor_latin.as_deref()
        }
        _ => None,
    }?;
    Some(resolve_font_name(font_name))
}

fn resolve_font_name(name: &str) -> ResolvedFont {
    let normalized = name.trim().to_ascii_lowercase();
    let family_name = match normalized.as_str() {
        "calibri" | "calibri light" | "aptos" | "aptos display" => Some(DOCX_CARLITO),
        "cambria" => Some(DOCX_CALADEA),
        "arial" => Some(DOCX_LIBERATION_SANS),
        "times new roman" => Some(DOCX_LIBERATION_SERIF),
        "courier new" | "consolas" => Some(DOCX_LIBERATION_MONO),
        _ => None,
    };

    let monospace = matches!(
        normalized.as_str(),
        "courier new" | "consolas" | "menlo" | "monaco" | "source code pro"
    ) || normalized.contains("mono");

    ResolvedFont {
        font_family_name: family_name,
        font_choice: if monospace {
            FontChoice::Monospace
        } else {
            FontChoice::Proportional
        },
    }
}

fn parse_styles_xml(styles_xml: &str, theme_fonts: &ThemeFonts) -> Result<DocxStyles, String> {
    let mut reader = Reader::from_str(styles_xml);
    reader.config_mut().trim_text(false);

    let mut raw = RawDocxStyles::default();
    let mut current_style_id = None::<String>;
    let mut current_style_type = None::<String>;
    let mut current_paragraph_style = RawParagraphStyleDefinition::default();
    let mut current_character_style = RawCharacterStyleDefinition::default();
    let mut in_doc_defaults = false;
    let mut in_doc_defaults_paragraph = false;
    let mut in_doc_defaults_run = false;
    let mut in_style_paragraph = false;
    let mut in_style_run = false;

    loop {
        match reader.read_event() {
            Ok(XmlEvent::Start(event)) => match local_name(event.name().as_ref()) {
                b"docDefaults" => in_doc_defaults = true,
                b"style" => {
                    current_style_id = attr_value(&event, b"styleId");
                    current_style_type = attr_value(&event, b"type");
                    current_paragraph_style = RawParagraphStyleDefinition::default();
                    current_character_style = RawCharacterStyleDefinition::default();
                }
                b"basedOn" => {
                    let based_on = attr_value(&event, b"val");
                    match current_style_type.as_deref() {
                        Some("paragraph") => current_paragraph_style.based_on = based_on,
                        Some("character") => current_character_style.based_on = based_on,
                        _ => {}
                    }
                }
                b"pPr" => {
                    if in_doc_defaults {
                        in_doc_defaults_paragraph = true;
                    } else if current_style_type.as_deref() == Some("paragraph") {
                        in_style_paragraph = true;
                    }
                }
                b"rPr" => {
                    if in_doc_defaults {
                        in_doc_defaults_run = true;
                    } else if current_style_id.is_some() {
                        in_style_run = true;
                    }
                }
                name => {
                    if in_doc_defaults_paragraph {
                        apply_paragraph_style_patch_event(name, &event, &mut raw.default_paragraph);
                    }
                    if in_doc_defaults_run {
                        apply_run_style_patch_event(
                            name,
                            &event,
                            &mut raw.default_run,
                            theme_fonts,
                        );
                    }
                    if in_style_run {
                        match current_style_type.as_deref() {
                            Some("paragraph") => apply_run_style_patch_event(
                                name,
                                &event,
                                &mut current_paragraph_style.run,
                                theme_fonts,
                            ),
                            Some("character") => apply_run_style_patch_event(
                                name,
                                &event,
                                &mut current_character_style.run,
                                theme_fonts,
                            ),
                            _ => {}
                        }
                    }
                    if in_style_paragraph {
                        apply_paragraph_style_patch_event(
                            name,
                            &event,
                            &mut current_paragraph_style.paragraph,
                        );
                    }
                }
            },
            Ok(XmlEvent::Empty(event)) => match local_name(event.name().as_ref()) {
                b"basedOn" => {
                    let based_on = attr_value(&event, b"val");
                    match current_style_type.as_deref() {
                        Some("paragraph") => current_paragraph_style.based_on = based_on,
                        Some("character") => current_character_style.based_on = based_on,
                        _ => {}
                    }
                }
                name => {
                    if in_doc_defaults && name == b"rPr" {
                        continue;
                    }
                    if in_doc_defaults_paragraph {
                        apply_paragraph_style_patch_event(name, &event, &mut raw.default_paragraph);
                    }
                    if in_doc_defaults_run {
                        apply_run_style_patch_event(
                            name,
                            &event,
                            &mut raw.default_run,
                            theme_fonts,
                        );
                    }
                    if in_style_run {
                        match current_style_type.as_deref() {
                            Some("paragraph") => apply_run_style_patch_event(
                                name,
                                &event,
                                &mut current_paragraph_style.run,
                                theme_fonts,
                            ),
                            Some("character") => apply_run_style_patch_event(
                                name,
                                &event,
                                &mut current_character_style.run,
                                theme_fonts,
                            ),
                            _ => {}
                        }
                    }
                    if in_style_paragraph {
                        apply_paragraph_style_patch_event(
                            name,
                            &event,
                            &mut current_paragraph_style.paragraph,
                        );
                    }
                }
            },
            Ok(XmlEvent::End(event)) => match local_name(event.name().as_ref()) {
                b"docDefaults" => in_doc_defaults = false,
                b"rPr" => {
                    in_doc_defaults_run = false;
                    in_style_run = false;
                }
                b"pPr" => {
                    in_doc_defaults_paragraph = false;
                    in_style_paragraph = false;
                }
                b"style" => {
                    if let Some(style_id) = current_style_id.take() {
                        match current_style_type.as_deref() {
                            Some("paragraph") => {
                                raw.paragraph_styles
                                    .insert(style_id, current_paragraph_style.clone());
                            }
                            Some("character") => {
                                raw.character_styles
                                    .insert(style_id, current_character_style.clone());
                            }
                            _ => {}
                        }
                    }
                    current_style_type = None;
                }
                _ => {}
            },
            Ok(XmlEvent::Eof) => break,
            Err(error) => return Err(format!("failed to parse word/styles.xml: {error}")),
            _ => {}
        }
    }

    Ok(resolve_styles(raw))
}

fn resolve_styles(raw: RawDocxStyles) -> DocxStyles {
    let mut paragraph_styles = HashMap::new();
    let mut character_styles = HashMap::new();

    for style_id in raw.paragraph_styles.keys() {
        let mut active = HashSet::new();
        let resolved = resolve_paragraph_style(style_id, &raw, &mut active);
        paragraph_styles.insert(style_id.clone(), resolved);
    }

    for style_id in raw.character_styles.keys() {
        let mut active = HashSet::new();
        let resolved = resolve_character_style(style_id, &raw, &mut active);
        character_styles.insert(style_id.clone(), resolved);
    }

    DocxStyles {
        default_paragraph: raw.default_paragraph.apply(ParagraphStyle::default()),
        default_run: raw.default_run.apply(CharacterStyle::default()),
        paragraph_styles,
        character_styles,
    }
}

fn resolve_paragraph_style(
    style_id: &str,
    raw: &RawDocxStyles,
    active: &mut HashSet<String>,
) -> ResolvedParagraphStyle {
    if !active.insert(style_id.to_owned()) {
        return ResolvedParagraphStyle::default();
    }

    let Some(style) = raw.paragraph_styles.get(style_id) else {
        active.remove(style_id);
        return ResolvedParagraphStyle::default();
    };

    let mut resolved = if let Some(parent) = style.based_on.as_deref() {
        resolve_paragraph_style(parent, raw, active)
    } else {
        ResolvedParagraphStyle::default()
    };
    resolved.paragraph.overlay(style.paragraph);
    resolved.run.overlay(style.run);
    active.remove(style_id);
    resolved
}

fn resolve_character_style(
    style_id: &str,
    raw: &RawDocxStyles,
    active: &mut HashSet<String>,
) -> CharacterStylePatch {
    if !active.insert(style_id.to_owned()) {
        return CharacterStylePatch::default();
    }

    let Some(style) = raw.character_styles.get(style_id) else {
        active.remove(style_id);
        return CharacterStylePatch::default();
    };

    let mut resolved = if let Some(parent) = style.based_on.as_deref() {
        resolve_character_style(parent, raw, active)
    } else {
        CharacterStylePatch::default()
    };
    resolved.overlay(style.run);
    active.remove(style_id);
    resolved
}

fn apply_run_style_patch_event(
    name: &[u8],
    event: &quick_xml::events::BytesStart<'_>,
    patch: &mut CharacterStylePatch,
    theme_fonts: &ThemeFonts,
) {
    match name {
        b"rFonts" => patch.font = resolve_rfonts(event, theme_fonts),
        b"b" => patch.bold = Some(docx_flag(event, true)),
        b"i" => patch.italic = Some(docx_flag(event, true)),
        b"u" => {
            patch.underline = Some(!matches!(
                attr_value(event, b"val").as_deref(),
                Some("none")
            ))
        }
        b"strike" | b"dstrike" => patch.strikethrough = Some(docx_flag(event, true)),
        b"sz" => {
            if let Some(value) = attr_value(event, b"val") {
                if let Ok(half_points) = value.parse::<f32>() {
                    patch.font_size_points = Some((half_points / 2.0).clamp(8.0, 72.0));
                }
            }
        }
        b"color" => {
            if let Some(value) = attr_value(event, b"val") {
                patch.text_color = parse_hex_color(&value);
            }
        }
        b"highlight" => {
            if let Some(value) = attr_value(event, b"val") {
                patch.highlight_color = Some(highlight_color(&value));
            }
        }
        _ => {}
    }
}

fn apply_paragraph_style_patch_event(
    name: &[u8],
    event: &quick_xml::events::BytesStart<'_>,
    patch: &mut ParagraphStylePatch,
) {
    match name {
        b"jc" => {
            patch.alignment = Some(paragraph_alignment_for(
                attr_value(event, b"val").as_deref().unwrap_or_default(),
            ));
        }
        b"spacing" => apply_spacing_patch(event, patch),
        b"pageBreakBefore" => patch.page_break_before = Some(docx_flag(event, true)),
        _ => {}
    }
}

fn parse_numbering_xml(numbering_xml: &str) -> Result<NumberingDefinitions, String> {
    let mut reader = Reader::from_str(numbering_xml);
    reader.config_mut().trim_text(false);

    let mut numbering = NumberingDefinitions::default();
    let mut current_abstract = None::<String>;
    let mut current_level = None::<String>;
    let mut current_num = None::<String>;

    loop {
        match reader.read_event() {
            Ok(XmlEvent::Start(event)) => match local_name(event.name().as_ref()) {
                b"abstractNum" => current_abstract = attr_value(&event, b"abstractNumId"),
                b"lvl" => current_level = attr_value(&event, b"ilvl"),
                b"num" => current_num = attr_value(&event, b"numId"),
                b"numFmt" => {
                    if let (Some(abstract_id), Some(level), Some(value)) = (
                        current_abstract.as_ref(),
                        current_level.as_ref(),
                        attr_value(&event, b"val"),
                    ) {
                        numbering.level_kinds.insert(
                            (abstract_id.clone(), level.clone()),
                            list_kind_for_numbering(&value),
                        );
                    }
                }
                b"abstractNumId" => {
                    if let (Some(num_id), Some(abstract_id)) =
                        (current_num.as_ref(), attr_value(&event, b"val"))
                    {
                        numbering
                            .num_to_abstract
                            .insert(num_id.clone(), abstract_id);
                    }
                }
                _ => {}
            },
            Ok(XmlEvent::Empty(event)) => match local_name(event.name().as_ref()) {
                b"abstractNum" => current_abstract = attr_value(&event, b"abstractNumId"),
                b"lvl" => current_level = attr_value(&event, b"ilvl"),
                b"num" => current_num = attr_value(&event, b"numId"),
                b"numFmt" => {
                    if let (Some(abstract_id), Some(level), Some(value)) = (
                        current_abstract.as_ref(),
                        current_level.as_ref(),
                        attr_value(&event, b"val"),
                    ) {
                        numbering.level_kinds.insert(
                            (abstract_id.clone(), level.clone()),
                            list_kind_for_numbering(&value),
                        );
                    }
                }
                b"abstractNumId" => {
                    if let (Some(num_id), Some(abstract_id)) =
                        (current_num.as_ref(), attr_value(&event, b"val"))
                    {
                        numbering
                            .num_to_abstract
                            .insert(num_id.clone(), abstract_id);
                    }
                }
                _ => {}
            },
            Ok(XmlEvent::End(event)) => match local_name(event.name().as_ref()) {
                b"abstractNum" => current_abstract = None,
                b"lvl" => current_level = None,
                b"num" => current_num = None,
                _ => {}
            },
            Ok(XmlEvent::Eof) => break,
            Err(error) => return Err(format!("failed to parse word/numbering.xml: {error}")),
            _ => {}
        }
    }

    Ok(numbering)
}

fn parse_document_relationships(relationships_xml: &str) -> Result<DocumentRelationships, String> {
    let mut reader = Reader::from_str(relationships_xml);
    reader.config_mut().trim_text(false);

    let mut relationships = DocumentRelationships::default();

    loop {
        match reader.read_event() {
            Ok(XmlEvent::Start(event)) | Ok(XmlEvent::Empty(event)) => {
                if local_name(event.name().as_ref()) != b"Relationship" {
                    continue;
                }

                let Some(rel_type) = attr_value(&event, b"Type") else {
                    continue;
                };
                if !rel_type.contains("/image") {
                    continue;
                }

                let (Some(id), Some(target)) =
                    (attr_value(&event, b"Id"), attr_value(&event, b"Target"))
                else {
                    continue;
                };
                relationships
                    .image_targets
                    .insert(id, normalize_relationship_target(&target));
            }
            Ok(XmlEvent::Eof) => break,
            Err(error) => {
                return Err(format!(
                    "failed to parse word/_rels/document.xml.rels: {error}"
                ));
            }
            _ => {}
        }
    }

    Ok(relationships)
}

fn normalize_relationship_target(target: &str) -> String {
    if target.starts_with('/') {
        target.trim_start_matches('/').to_owned()
    } else if target.starts_with("word/") {
        target.to_owned()
    } else {
        format!("word/{target}")
    }
}

fn resolve_drawing(
    drawing: Option<DrawingState>,
    relationships: &DocumentRelationships,
    media: &HashMap<String, Vec<u8>>,
    next_image_id: &mut usize,
) -> Option<DocumentImage> {
    let drawing = drawing?;
    let rel_id = drawing.rel_id?;
    let target = relationships.image_targets.get(&rel_id)?;
    let bytes = media.get(target)?.clone();

    let (width_points, height_points) = drawing.size_points.unwrap_or((240.0, 180.0));
    let alt_text = drawing.alt_text.unwrap_or_else(|| {
        Path::new(target)
            .file_stem()
            .and_then(|value| value.to_str())
            .unwrap_or("Image")
            .to_owned()
    });

    let layout_mode = if drawing.is_anchor {
        crate::document::ImageLayoutMode::Floating
    } else {
        crate::document::ImageLayoutMode::Inline
    };
    let wrap_mode = drawing.wrap_mode.unwrap_or(if drawing.is_anchor {
        crate::document::WrapMode::Square
    } else {
        crate::document::WrapMode::Inline
    });
    let distance_from_text = drawing.distance_from_text.unwrap_or_default();

    let image = DocumentImage {
        id: *next_image_id,
        bytes,
        alt_text,
        width_points,
        height_points,
        lock_aspect_ratio: true,
        opacity: 1.0,
        layout_mode,
        wrap_mode,
        rendering: crate::document::ImageRendering::Smooth,
        horizontal_position: Default::default(),
        vertical_position: Default::default(),
        distance_from_text,
        z_index: 0,
        move_with_text: true,
        allow_overlap: false,
    };
    *next_image_id += 1;
    Some(image)
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

fn paragraph_alignment_for(value: &str) -> ParagraphAlignment {
    match value {
        "center" => ParagraphAlignment::Center,
        "right" => ParagraphAlignment::Right,
        "both" | "distribute" => ParagraphAlignment::Justify,
        _ => ParagraphAlignment::Left,
    }
}

fn list_kind_for_numbering(value: &str) -> ListKind {
    match value {
        "bullet" => ListKind::Bullet,
        "none" => ListKind::None,
        _ => ListKind::Ordered,
    }
}

fn apply_spacing(event: &quick_xml::events::BytesStart<'_>, paragraph_style: &mut ParagraphStyle) {
    if let Some(value) = attr_value(event, b"before")
        .and_then(|value| value.parse::<f32>().ok())
        .map(twips_to_points)
    {
        paragraph_style.spacing_before_points = value.round().clamp(0.0, u16::MAX as f32) as u16;
    }
    if let Some(value) = attr_value(event, b"after")
        .and_then(|value| value.parse::<f32>().ok())
        .map(twips_to_points)
    {
        paragraph_style.spacing_after_points = value.round().clamp(0.0, u16::MAX as f32) as u16;
    }
    if let Some(line_spacing) = parse_line_spacing(event) {
        paragraph_style.line_spacing = line_spacing;
    }
}

fn apply_spacing_patch(event: &quick_xml::events::BytesStart<'_>, patch: &mut ParagraphStylePatch) {
    if let Some(value) = attr_value(event, b"before")
        .and_then(|value| value.parse::<f32>().ok())
        .map(twips_to_points)
    {
        patch.spacing_before_points = Some(value.round().clamp(0.0, u16::MAX as f32) as u16);
    }
    if let Some(value) = attr_value(event, b"after")
        .and_then(|value| value.parse::<f32>().ok())
        .map(twips_to_points)
    {
        patch.spacing_after_points = Some(value.round().clamp(0.0, u16::MAX as f32) as u16);
    }
    if let Some(line_spacing) = parse_line_spacing(event) {
        patch.line_spacing = Some(line_spacing);
    }
}

fn parse_line_spacing(event: &quick_xml::events::BytesStart<'_>) -> Option<LineSpacing> {
    let line = attr_value(event, b"line")?.parse::<f32>().ok()?;
    let line_rule = attr_value(event, b"lineRule").unwrap_or_else(|| "auto".to_owned());
    Some(match line_rule.as_str() {
        "atLeast" => LineSpacing {
            kind: LineSpacingKind::AtLeastPoints,
            value: twips_to_points(line),
        },
        "exact" => LineSpacing {
            kind: LineSpacingKind::ExactPoints,
            value: twips_to_points(line),
        },
        _ => LineSpacing {
            kind: LineSpacingKind::AutoMultiplier,
            value: line / 240.0,
        },
    })
}

fn parse_page_size(event: &quick_xml::events::BytesStart<'_>) -> Option<PageSize> {
    let width_twips = attr_value(event, b"w")?.parse::<f32>().ok()?;
    let height_twips = attr_value(event, b"h")?.parse::<f32>().ok()?;
    Some(PageSize {
        width_points: twips_to_points(width_twips),
        height_points: twips_to_points(height_twips),
    })
}

fn parse_page_margins(event: &quick_xml::events::BytesStart<'_>) -> Option<PageMargins> {
    Some(PageMargins {
        top_points: twips_to_points(attr_value(event, b"top")?.parse::<f32>().ok()?),
        right_points: twips_to_points(attr_value(event, b"right")?.parse::<f32>().ok()?),
        bottom_points: twips_to_points(attr_value(event, b"bottom")?.parse::<f32>().ok()?),
        left_points: twips_to_points(attr_value(event, b"left")?.parse::<f32>().ok()?),
    })
}

fn parse_emu_extent(event: &quick_xml::events::BytesStart<'_>) -> Option<(f32, f32)> {
    let width = attr_value(event, b"cx")?.parse::<f32>().ok()?;
    let height = attr_value(event, b"cy")?.parse::<f32>().ok()?;
    Some((emu_to_points(width), emu_to_points(height)))
}

fn twips_to_points(value: f32) -> f32 {
    value / 20.0
}

fn emu_to_points(value: f32) -> f32 {
    value / 12_700.0
}

fn parse_anchor_distance(
    event: &quick_xml::events::BytesStart<'_>,
) -> Option<crate::document::DistanceFromText> {
    let top = attr_value(event, b"distT")
        .and_then(|v| v.parse::<f32>().ok())
        .map(emu_to_points)
        .unwrap_or(0.0);
    let bottom = attr_value(event, b"distB")
        .and_then(|v| v.parse::<f32>().ok())
        .map(emu_to_points)
        .unwrap_or(0.0);
    let left = attr_value(event, b"distL")
        .and_then(|v| v.parse::<f32>().ok())
        .map(emu_to_points)
        .unwrap_or(8.0);
    let right = attr_value(event, b"distR")
        .and_then(|v| v.parse::<f32>().ok())
        .map(emu_to_points)
        .unwrap_or(8.0);
    Some(crate::document::DistanceFromText {
        top_points: top,
        right_points: right,
        bottom_points: bottom,
        left_points: left,
    })
}

#[cfg(test)]
mod tests {
    use std::collections::HashMap;

    use super::{
        parse_document_relationships, parse_document_xml, parse_numbering_xml, parse_styles_xml,
        parse_theme_xml,
    };
    use crate::document::{LineSpacingKind, ListKind, ParagraphAlignment, OBJECT_REPLACEMENT_CHAR};

    #[test]
    fn parses_lists_alignment_and_page_settings_from_docx_xml() {
        let numbering = parse_numbering_xml(
            r#"
            <w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:abstractNum w:abstractNumId="10">
                <w:lvl w:ilvl="0">
                  <w:numFmt w:val="decimal"/>
                </w:lvl>
              </w:abstractNum>
              <w:abstractNum w:abstractNumId="11">
                <w:lvl w:ilvl="0">
                  <w:numFmt w:val="bullet"/>
                </w:lvl>
              </w:abstractNum>
              <w:num w:numId="7">
                <w:abstractNumId w:val="10"/>
              </w:num>
              <w:num w:numId="8">
                <w:abstractNumId w:val="11"/>
              </w:num>
            </w:numbering>
            "#,
        )
        .unwrap();

        let imported = parse_document_xml(
            r#"
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:pPr>
                    <w:jc w:val="center"/>
                    <w:numPr>
                      <w:ilvl w:val="0"/>
                      <w:numId w:val="7"/>
                    </w:numPr>
                  </w:pPr>
                  <w:r><w:t>First</w:t></w:r>
                </w:p>
                <w:p>
                  <w:pPr>
                    <w:numPr>
                      <w:ilvl w:val="0"/>
                      <w:numId w:val="8"/>
                    </w:numPr>
                  </w:pPr>
                  <w:r><w:t>Second</w:t></w:r>
                </w:p>
                <w:sectPr>
                  <w:pgSz w:w="12240" w:h="15840"/>
                  <w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/>
                </w:sectPr>
              </w:body>
            </w:document>
            "#,
            &numbering,
            &Default::default(),
            &Default::default(),
            &Default::default(),
            &HashMap::new(),
        )
        .unwrap();

        assert_eq!(imported.runs.len(), 1);
        assert_eq!(imported.runs[0].text, "First\nSecond");
        assert_eq!(imported.paragraph_styles.len(), 2);
        assert_eq!(imported.paragraph_images, vec![None, None]);
        assert_eq!(
            imported.paragraph_styles[0],
            crate::document::ParagraphStyle {
                alignment: ParagraphAlignment::Center,
                list_kind: ListKind::Ordered,
                page_break_before: false,
                spacing_before_points: 0,
                spacing_after_points: 0,
                line_spacing: crate::document::LineSpacing::default(),
            }
        );
        assert_eq!(imported.paragraph_styles[1].list_kind, ListKind::Bullet);
        assert_eq!(imported.page_size.unwrap().width_points, 612.0);
        assert_eq!(imported.margins.unwrap().left_points, 90.0);
    }

    #[test]
    fn imports_image_paragraphs_from_docx_xml() {
        let numbering = parse_numbering_xml(
            r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#,
        )
        .unwrap();
        let relationships = parse_document_relationships(
            r#"
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship
                Id="rId5"
                Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
                Target="media/image1.png"
              />
            </Relationships>
            "#,
        )
        .unwrap();

        let imported = parse_document_xml(
            r#"
            <w:document
              xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
              xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"
              xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
              xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <w:body>
                <w:p>
                  <w:r>
                    <w:drawing>
                      <wp:inline>
                        <wp:extent cx="914400" cy="457200"/>
                        <wp:docPr name="Logo" descr="Logo"/>
                        <a:graphic>
                          <a:graphicData>
                            <a:blip r:embed="rId5"/>
                          </a:graphicData>
                        </a:graphic>
                      </wp:inline>
                    </w:drawing>
                  </w:r>
                </w:p>
              </w:body>
            </w:document>
            "#,
            &numbering,
            &Default::default(),
            &Default::default(),
            &relationships,
            &HashMap::from([(String::from("word/media/image1.png"), vec![1, 2, 3, 4])]),
        )
        .unwrap();

        assert_eq!(imported.runs[0].text, OBJECT_REPLACEMENT_CHAR.to_string());
        assert_eq!(imported.paragraph_images.len(), 1);
        let image = imported.paragraph_images[0].as_ref().unwrap();
        assert_eq!(image.alt_text, "Logo");
        assert_eq!(image.width_points, 72.0);
        assert_eq!(image.height_points, 36.0);
        assert_eq!(image.bytes, vec![1, 2, 3, 4]);
    }

    #[test]
    fn imports_tables_from_docx_xml() {
        let numbering = parse_numbering_xml(
            r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#,
        )
        .unwrap();

        let imported = parse_document_xml(
            r#"
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:tbl>
                  <w:tblGrid>
                    <w:gridCol w:w="1440"/>
                    <w:gridCol w:w="2880"/>
                  </w:tblGrid>
                  <w:tr>
                    <w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
                    <w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
                  </w:tr>
                  <w:tr>
                    <w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>
                    <w:tc><w:p><w:r><w:t>B2</w:t></w:r></w:p></w:tc>
                  </w:tr>
                </w:tbl>
              </w:body>
            </w:document>
            "#,
            &numbering,
            &Default::default(),
            &Default::default(),
            &Default::default(),
            &HashMap::new(),
        )
        .unwrap();

        assert_eq!(imported.runs[0].text, OBJECT_REPLACEMENT_CHAR.to_string());
        assert_eq!(imported.paragraph_images, vec![None]);
        let table = imported.paragraph_tables[0].as_ref().unwrap();
        assert_eq!(table.num_rows(), 2);
        assert_eq!(table.num_cols(), 2);
        assert_eq!(table.col_widths_points, vec![72.0, 144.0]);
        assert_eq!(table.rows[0][0].plain_text(), "A1");
        assert_eq!(table.rows[1][1].plain_text(), "B2");
    }

    #[test]
    fn falls_back_to_default_paragraph_style_without_numbering() {
        let numbering = parse_numbering_xml(
            r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#,
        )
        .unwrap();

        let imported = parse_document_xml(
            r#"
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:r><w:t xml:space="preserve"> plain </w:t></w:r>
                </w:p>
              </w:body>
            </w:document>
            "#,
            &numbering,
            &Default::default(),
            &Default::default(),
            &Default::default(),
            &HashMap::new(),
        )
        .unwrap();

        assert_eq!(imported.runs[0].text, " plain ");
        assert_eq!(imported.paragraph_styles.len(), 1);
        assert_eq!(imported.paragraph_styles[0].list_kind, ListKind::None);
        assert_eq!(
            imported.paragraph_styles[0].alignment,
            ParagraphAlignment::Left
        );
    }

    #[test]
    fn resolves_word_styles_for_paragraph_spacing_and_run_formatting() {
        let numbering = parse_numbering_xml(
            r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#,
        )
        .unwrap();
        let styles = parse_styles_xml(
            r#"
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:docDefaults>
                <w:pPrDefault>
                  <w:pPr>
                    <w:spacing w:line="360"/>
                  </w:pPr>
                </w:pPrDefault>
                <w:rPrDefault>
                  <w:rPr>
                    <w:sz w:val="22"/>
                  </w:rPr>
                </w:rPrDefault>
              </w:docDefaults>
              <w:style w:type="paragraph" w:styleId="Normal">
                <w:pPr>
                  <w:spacing w:after="160"/>
                </w:pPr>
              </w:style>
              <w:style w:type="paragraph" w:styleId="Title">
                <w:basedOn w:val="Normal"/>
                <w:pPr>
                  <w:spacing w:after="240"/>
                </w:pPr>
                <w:rPr>
                  <w:rFonts w:ascii="Calibri"/>
                  <w:b/>
                  <w:sz w:val="56"/>
                </w:rPr>
              </w:style>
              <w:style w:type="character" w:styleId="Accent">
                <w:rPr>
                  <w:rFonts w:ascii="Consolas"/>
                  <w:i/>
                </w:rPr>
              </w:style>
            </w:styles>
            "#,
            &Default::default(),
        )
        .unwrap();

        let imported = parse_document_xml(
            r#"
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:pPr><w:pStyle w:val="Title"/></w:pPr>
                  <w:r><w:t>Heading</w:t></w:r>
                </w:p>
                <w:p>
                  <w:pPr><w:pStyle w:val="Normal"/></w:pPr>
                  <w:r>
                    <w:rPr><w:rStyle w:val="Accent"/></w:rPr>
                    <w:t>Body</w:t>
                  </w:r>
                </w:p>
              </w:body>
            </w:document>
            "#,
            &numbering,
            &styles,
            &Default::default(),
            &Default::default(),
            &HashMap::new(),
        )
        .unwrap();

        assert_eq!(imported.paragraph_styles[0].spacing_after_points, 12);
        assert_eq!(imported.paragraph_styles[1].spacing_after_points, 8);
        assert_eq!(
            imported.paragraph_styles[1].line_spacing.kind,
            LineSpacingKind::AutoMultiplier
        );
        assert_eq!(imported.paragraph_styles[1].line_spacing.value, 1.5);
        assert_eq!(imported.runs.len(), 3);
        assert_eq!(imported.runs[0].text, "Heading");
        assert!(imported.runs[0].style.bold);
        assert_eq!(imported.runs[0].style.font_size_points, 28.0);
        assert_eq!(
            imported.runs[0].style.font_family_name,
            Some("docx-carlito")
        );
        assert_eq!(imported.runs[2].text, "Body");
        assert!(imported.runs[2].style.italic);
        assert_eq!(imported.runs[2].style.font_size_points, 11.0);
        assert_eq!(
            imported.runs[2].style.font_family_name,
            Some("docx-liberation-mono")
        );
    }

    #[test]
    fn resolves_theme_fonts_and_direct_run_font_override() {
        let numbering = parse_numbering_xml(
            r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#,
        )
        .unwrap();
        let theme_fonts = parse_theme_xml(
            r#"
            <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
              <a:themeElements>
                <a:fontScheme>
                  <a:majorFont><a:latin typeface="Cambria"/></a:majorFont>
                  <a:minorFont><a:latin typeface="Aptos"/></a:minorFont>
                </a:fontScheme>
              </a:themeElements>
            </a:theme>
            "#,
        )
        .unwrap();
        let styles = parse_styles_xml(
            r#"
            <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:style w:type="paragraph" w:styleId="Body">
                <w:rPr>
                  <w:rFonts w:asciiTheme="minorHAnsi"/>
                </w:rPr>
              </w:style>
            </w:styles>
            "#,
            &theme_fonts,
        )
        .unwrap();

        let imported = parse_document_xml(
            r#"
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:pPr><w:pStyle w:val="Body"/></w:pPr>
                  <w:r><w:t>Body</w:t></w:r>
                </w:p>
                <w:p>
                  <w:pPr><w:pStyle w:val="Body"/></w:pPr>
                  <w:r>
                    <w:rPr><w:rFonts w:ascii="Cambria"/></w:rPr>
                    <w:t>Override</w:t>
                  </w:r>
                </w:p>
              </w:body>
            </w:document>
            "#,
            &numbering,
            &styles,
            &theme_fonts,
            &Default::default(),
            &HashMap::new(),
        )
        .unwrap();

        assert_eq!(
            imported.runs[0].style.font_family_name,
            Some("docx-carlito")
        );
        assert_eq!(
            imported.runs[2].style.font_family_name,
            Some("docx-caladea")
        );
    }

    #[test]
    fn parses_exact_and_at_least_line_spacing() {
        let numbering = parse_numbering_xml(
            r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>"#,
        )
        .unwrap();

        let imported = parse_document_xml(
            r#"
            <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
              <w:body>
                <w:p>
                  <w:pPr><w:spacing w:line="480" w:lineRule="exact"/></w:pPr>
                  <w:r><w:t>Exact</w:t></w:r>
                </w:p>
                <w:p>
                  <w:pPr><w:spacing w:line="360" w:lineRule="atLeast"/></w:pPr>
                  <w:r><w:t>AtLeast</w:t></w:r>
                </w:p>
              </w:body>
            </w:document>
            "#,
            &numbering,
            &Default::default(),
            &Default::default(),
            &Default::default(),
            &HashMap::new(),
        )
        .unwrap();

        assert_eq!(
            imported.paragraph_styles[0].line_spacing.kind,
            LineSpacingKind::ExactPoints
        );
        assert_eq!(imported.paragraph_styles[0].line_spacing.value, 24.0);
        assert_eq!(
            imported.paragraph_styles[1].line_spacing.kind,
            LineSpacingKind::AtLeastPoints
        );
        assert_eq!(imported.paragraph_styles[1].line_spacing.value, 18.0);
    }
}
