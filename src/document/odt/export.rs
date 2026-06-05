use std::{
    collections::{BTreeMap, HashMap, HashSet},
    fmt::Write as _,
    io::{Cursor, Write},
};

use eframe::egui::Color32;
use zip::{write::SimpleFileOptions, CompressionMethod, ZipWriter};

use crate::document::{
    CharacterStyle, DocumentImage, DocumentState, DocumentTable, FontChoice, LineSpacing,
    LineSpacingKind, ListKind, ParagraphAlignment, ParagraphStyle, TextRun,
    OBJECT_REPLACEMENT_CHAR,
};

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
    let mut master_page =
        "<style:master-page style:name=\"Standard\" style:page-layout-name=\"pm1\">".to_owned();
    if !document.header_text.trim().is_empty() {
        let _ = write!(
            master_page,
            "<style:header><text:p>{}</text:p></style:header>",
            odt_page_field_xml(&document.header_text)
        );
    }
    if !document.footer_text.trim().is_empty() {
        let _ = write!(
            master_page,
            "<style:footer><text:p>{}</text:p></style:footer>",
            odt_page_field_xml(&document.footer_text)
        );
    }
    master_page.push_str("</style:master-page>");

    format!(
        "{}<office:document-styles {} office:version=\"1.3\"><office:styles><style:style style:name=\"Standard\" style:family=\"paragraph\"/></office:styles><office:automatic-styles><style:page-layout style:name=\"pm1\"><style:page-layout-properties fo:page-width=\"{:.2}pt\" fo:page-height=\"{:.2}pt\" fo:margin-top=\"{:.2}pt\" fo:margin-right=\"{:.2}pt\" fo:margin-bottom=\"{:.2}pt\" fo:margin-left=\"{:.2}pt\"/></style:page-layout></office:automatic-styles><office:master-styles>{}</office:master-styles></office:document-styles>",
        xml_decl(),
        namespaces(),
        document.page_size.width_points,
        document.page_size.height_points,
        document.margins.top_points,
        document.margins.right_points,
        document.margins.bottom_points,
        document.margins.left_points,
        master_page
    )
}

fn odt_page_field_xml(template: &str) -> String {
    let mut xml = String::new();
    let mut remaining = template;
    while !remaining.is_empty() {
        let Some(start) = remaining.find('{') else {
            xml.push_str(&xml_escape(remaining));
            break;
        };
        xml.push_str(&xml_escape(&remaining[..start]));
        let tail = &remaining[start..];
        if let Some(next) = tail
            .strip_prefix("{ NUMPAGES }")
            .or_else(|| tail.strip_prefix("{NUMPAGES}"))
            .or_else(|| tail.strip_prefix("{ SECTIONPAGES }"))
            .or_else(|| tail.strip_prefix("{SECTIONPAGES}"))
            .or_else(|| tail.strip_prefix("{pagecount}"))
            .or_else(|| tail.strip_prefix("{sectionpages}"))
            .or_else(|| tail.strip_prefix("{numpages}"))
        {
            xml.push_str("<text:page-count/>");
            remaining = next;
        } else if let Some(next) = tail.strip_prefix("{pages}") {
            xml.push_str("<text:page-count/>");
            remaining = next;
        } else if let Some(next) = tail
            .strip_prefix("{ PAGE }")
            .or_else(|| tail.strip_prefix("{PAGE}"))
            .or_else(|| tail.strip_prefix("{page}"))
        {
            xml.push_str("<text:page-number/>");
            remaining = next;
        } else {
            xml.push('{');
            remaining = &remaining[start + 1..];
        }
    }
    xml
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
