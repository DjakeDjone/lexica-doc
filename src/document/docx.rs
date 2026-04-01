use std::{
    collections::HashMap,
    io::{Cursor, Read},
};

use eframe::egui::Color32;
use quick_xml::{events::Event as XmlEvent, Reader};
use zip::ZipArchive;

use crate::document::{
    CharacterStyle, ListKind, PageMargins, PageSize, ParagraphAlignment, ParagraphStyle, TextRun,
};

pub(super) struct ImportedDocx {
    pub(super) runs: Vec<TextRun>,
    pub(super) paragraph_styles: Vec<ParagraphStyle>,
    pub(super) page_size: Option<PageSize>,
    pub(super) margins: Option<PageMargins>,
}

pub(super) fn docx_to_document(bytes: &[u8]) -> Result<ImportedDocx, String> {
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
    parse_document_xml(&document_xml, &numbering)
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

fn parse_document_xml(
    document_xml: &str,
    numbering: &NumberingDefinitions,
) -> Result<ImportedDocx, String> {
    let mut reader = Reader::from_str(document_xml);
    reader.config_mut().trim_text(false);

    let mut runs = Vec::new();
    let mut paragraph_styles = Vec::new();
    let mut run_style = CharacterStyle::default();
    let mut paragraph_style = ParagraphStyle::default();
    let mut in_text = false;
    let mut current_num_id = None;
    let mut current_ilvl = None;
    let mut page_size = None;
    let mut margins = None;

    loop {
        match reader.read_event() {
            Ok(XmlEvent::Start(event)) => match local_name(event.name().as_ref()) {
                b"p" => {
                    if !paragraph_styles.is_empty() {
                        append_plain(&mut runs, "\n", CharacterStyle::default());
                    }
                    paragraph_style = ParagraphStyle::default();
                    current_num_id = None;
                    current_ilvl = None;
                }
                b"r" => {
                    run_style = CharacterStyle::default();
                }
                b"t" => in_text = true,
                b"br" => append_plain(&mut runs, "\n", run_style),
                b"tab" => append_plain(&mut runs, "\t", run_style),
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
                b"numId" => current_num_id = attr_value(&event, b"val"),
                b"ilvl" => current_ilvl = attr_value(&event, b"val"),
                b"pgSz" => page_size = parse_page_size(&event),
                b"pgMar" => margins = parse_page_margins(&event),
                _ => {}
            },
            Ok(XmlEvent::Empty(event)) => match local_name(event.name().as_ref()) {
                b"br" => append_plain(&mut runs, "\n", run_style),
                b"tab" => append_plain(&mut runs, "\t", run_style),
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
                b"numId" => current_num_id = attr_value(&event, b"val"),
                b"ilvl" => current_ilvl = attr_value(&event, b"val"),
                b"pgSz" => page_size = parse_page_size(&event),
                b"pgMar" => margins = parse_page_margins(&event),
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
                b"p" => {
                    paragraph_style.list_kind =
                        numbering.lookup(current_num_id.as_deref(), current_ilvl.as_deref());
                    paragraph_styles.push(paragraph_style);
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

    Ok(ImportedDocx {
        runs,
        paragraph_styles,
        page_size,
        margins,
    })
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

fn twips_to_points(value: f32) -> f32 {
    value / 20.0
}

#[cfg(test)]
mod tests {
    use super::{parse_document_xml, parse_numbering_xml};
    use crate::document::{ListKind, ParagraphAlignment};

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
        )
        .unwrap();

        assert_eq!(imported.runs.len(), 1);
        assert_eq!(imported.runs[0].text, "First\nSecond");
        assert_eq!(imported.paragraph_styles.len(), 2);
        assert_eq!(
            imported.paragraph_styles[0],
            crate::document::ParagraphStyle {
                alignment: ParagraphAlignment::Center,
                list_kind: ListKind::Ordered,
            }
        );
        assert_eq!(imported.paragraph_styles[1].list_kind, ListKind::Bullet);
        assert_eq!(imported.page_size.unwrap().width_points, 612.0);
        assert_eq!(imported.margins.unwrap().left_points, 90.0);
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
}
