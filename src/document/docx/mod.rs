pub mod export;
pub mod helpers;
pub mod numbering;
pub mod styles;
pub mod table;

pub use export::document_to_docx;

#[cfg(test)]
mod tests;

use std::{
    collections::{HashMap, HashSet},
    io::{Cursor, Read},
};

use quick_xml::{events::Event as XmlEvent, Reader};
use zip::ZipArchive;

use crate::document::{
    CharacterStyle, DocumentImage, DocumentTable, PageMargins, PageSetup, PageSize, ParagraphStyle,
    Section, TextRun, OBJECT_REPLACEMENT_CHAR,
};
use serde::Serialize;

use helpers::*;
use numbering::{parse_numbering_xml, NumberingDefinitions};
use styles::{
    apply_resolved_font, parse_styles_xml, parse_theme_xml, resolve_rfonts, DocxStyles, ThemeFonts,
};
use table::parse_docx_table;

#[derive(Debug, Serialize)]
pub struct ImportedDocx {
    pub runs: Vec<TextRun>,
    pub paragraph_styles: Vec<ParagraphStyle>,
    pub paragraph_images: Vec<Option<DocumentImage>>,
    pub paragraph_tables: Vec<Option<DocumentTable>>,
    pub page_size: Option<PageSize>,
    pub margins: Option<PageMargins>,
    pub different_odd_even_pages: bool,
    pub sections: Vec<Section>,
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
    let different_odd_even_pages = load_even_and_odd_setting(&mut archive)?;
    let theme_fonts = load_theme_fonts(&mut archive)?;
    let styles = load_styles(&mut archive, &theme_fonts)?;
    let relationships = load_document_relationships(&mut archive)?;
    let media = load_media_store(&mut archive, &relationships)?;
    let mut imported = parse_document_xml(
        &document_xml,
        &numbering,
        &styles,
        &theme_fonts,
        &relationships,
        &media,
    )?;
    imported.different_odd_even_pages = different_odd_even_pages;
    if imported.sections.is_empty() {
        let mut setup = PageSetup::standard();
        if let Some(page_size) = imported.page_size {
            setup.page_size = page_size;
        }
        if let Some(margins) = imported.margins {
            setup.margins = margins;
        }
        imported.sections.push(Section::first(setup));
    }
    Ok(imported)
}

fn load_even_and_odd_setting(archive: &mut ZipArchive<Cursor<&[u8]>>) -> Result<bool, String> {
    let Ok(mut settings_file) = archive.by_name("word/settings.xml") else {
        return Ok(false);
    };
    let mut settings_xml = String::new();
    settings_file
        .read_to_string(&mut settings_xml)
        .map_err(|error| format!("failed to read word/settings.xml: {error}"))?;
    Ok(
        settings_xml.contains("<w:evenAndOddHeaders")
            || settings_xml.contains("<evenAndOddHeaders"),
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

#[derive(Default)]
pub(crate) struct DrawingState {
    pub(crate) rel_id: Option<String>,
    pub(crate) alt_text: Option<String>,
    pub(crate) size_points: Option<(f32, f32)>,
    pub(crate) is_anchor: bool,
    pub(crate) wrap_mode: Option<crate::document::WrapMode>,
    pub(crate) distance_from_text: Option<crate::document::DistanceFromText>,
}

#[derive(Default)]
pub(crate) struct DocumentRelationships {
    pub(crate) image_targets: HashMap<String, String>,
}

pub(crate) fn parse_document_xml(
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
                        drawing.distance_from_text = parse_anchor_distance(&event);
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
        different_odd_even_pages: false,
        sections: Vec::new(),
    })
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
