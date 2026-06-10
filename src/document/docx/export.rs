use std::{
    collections::{BTreeMap, HashMap},
    fmt::Write as _,
    io::{Cursor, Write},
};

use eframe::egui::Color32;
use zip::{write::SimpleFileOptions, CompressionMethod, ZipWriter};

use crate::document::{
    CharacterStyle, DocumentImage, DocumentState, DocumentTable, FontChoice, HeaderFooterKind,
    HeaderFooterVariant, ImageLayoutMode, LineSpacing, LineSpacingKind, ListKind, PageSetup,
    Paragraph, ParagraphAlignment, SectionId, TextRun, WrapMode, OBJECT_REPLACEMENT_CHAR,
};

pub fn document_to_docx(document: &DocumentState) -> Result<Vec<u8>, String> {
    DocxPackage::from_document(document)?.into_bytes()
}

struct DocxPackage {
    parts: BTreeMap<String, Vec<u8>>,
}

impl DocxPackage {
    fn from_document(document: &DocumentState) -> Result<Self, String> {
        let mut builder = PackageBuilder::default();
        let section_refs = builder.build_header_footer_parts(document);
        let document_xml = builder.document_xml(document, &section_refs);

        builder.add_text_part("word/document.xml", document_xml);
        builder.add_text_part("word/styles.xml", styles_xml());
        builder.add_text_part("word/numbering.xml", numbering_xml());
        builder.add_text_part("word/settings.xml", settings_xml(document));
        builder.add_text_part("docProps/core.xml", core_xml(document));
        builder.add_text_part("docProps/app.xml", app_xml());
        builder.add_text_part("word/_rels/document.xml.rels", builder.document_rels_xml());
        builder.add_text_part("_rels/.rels", root_rels_xml());
        builder.add_text_part("[Content_Types].xml", builder.content_types_xml());

        Ok(Self {
            parts: builder.parts,
        })
    }

    fn into_bytes(self) -> Result<Vec<u8>, String> {
        let mut buffer = Cursor::new(Vec::new());
        let mut zip = ZipWriter::new(&mut buffer);
        let options = SimpleFileOptions::default().compression_method(CompressionMethod::Deflated);
        for (name, bytes) in self.parts {
            zip.start_file(&name, options)
                .map_err(|error| format!("failed to write DOCX part {name}: {error}"))?;
            zip.write_all(&bytes)
                .map_err(|error| format!("failed to write DOCX part {name}: {error}"))?;
        }
        zip.finish()
            .map_err(|error| format!("failed to finish DOCX package: {error}"))?;
        Ok(buffer.into_inner())
    }
}

#[derive(Default)]
struct PackageBuilder {
    parts: BTreeMap<String, Vec<u8>>,
    document_rels: Vec<Relationship>,
    next_rel_id: usize,
    next_header_id: usize,
    next_footer_id: usize,
    next_doc_pr_id: usize,
    image_paths_by_id: HashMap<usize, String>,
}

#[derive(Clone)]
struct Relationship {
    id: String,
    rel_type: &'static str,
    target: String,
}

#[derive(Clone)]
struct HeaderFooterReference {
    kind: HeaderFooterKind,
    variant: HeaderFooterVariant,
    rel_id: String,
}

impl PackageBuilder {
    fn add_text_part(&mut self, path: &str, xml: String) {
        self.parts.insert(path.to_owned(), xml.into_bytes());
    }

    fn add_binary_part(&mut self, path: &str, bytes: Vec<u8>) {
        self.parts.insert(path.to_owned(), bytes);
    }

    fn add_document_relationship(&mut self, rel_type: &'static str, target: String) -> String {
        self.next_rel_id += 1;
        let id = format!("rId{}", self.next_rel_id);
        self.document_rels.push(Relationship {
            id: id.clone(),
            rel_type,
            target,
        });
        id
    }

    fn add_image_relationship(&mut self, image: &DocumentImage) -> (String, String) {
        let path = if let Some(path) = self.image_paths_by_id.get(&image.id) {
            path.clone()
        } else {
            let ext = image_extension(&image.bytes);
            let path = format!("word/media/image-{}.{}", image.id, ext);
            self.add_binary_part(&path, image.bytes.clone());
            self.image_paths_by_id.insert(image.id, path.clone());
            path
        };
        let target = path
            .strip_prefix("word/")
            .map(str::to_owned)
            .unwrap_or_else(|| path.clone());
        let rel_id = self.add_document_relationship(REL_IMAGE, target);
        (rel_id, path)
    }

    fn build_header_footer_parts(
        &mut self,
        document: &DocumentState,
    ) -> HashMap<SectionId, Vec<HeaderFooterReference>> {
        let mut refs = HashMap::new();
        let variants = [
            HeaderFooterVariant::Default,
            HeaderFooterVariant::First,
            HeaderFooterVariant::Even,
        ];
        let kinds = [HeaderFooterKind::Header, HeaderFooterKind::Footer];

        for section in &document.sections {
            let mut section_refs = Vec::new();
            for kind in kinds {
                for variant in variants {
                    let resolved = document.resolve_header_footer_slot(section.id, kind, variant);
                    if resolved.story.plain_text().trim().is_empty() {
                        continue;
                    }
                    let (path, target, rel_type) = match kind {
                        HeaderFooterKind::Header => {
                            self.next_header_id += 1;
                            let file = format!("header{}.xml", self.next_header_id);
                            (format!("word/{file}"), file, REL_HEADER)
                        }
                        HeaderFooterKind::Footer => {
                            self.next_footer_id += 1;
                            let file = format!("footer{}.xml", self.next_footer_id);
                            (format!("word/{file}"), file, REL_FOOTER)
                        }
                    };
                    let xml = header_footer_xml(kind, &resolved.story.runs);
                    self.add_text_part(&path, xml);
                    let rel_id = self.add_document_relationship(rel_type, target);
                    section_refs.push(HeaderFooterReference {
                        kind,
                        variant,
                        rel_id,
                    });
                }
            }
            refs.insert(section.id, section_refs);
        }

        refs
    }

    fn document_xml(
        &mut self,
        document: &DocumentState,
        section_refs: &HashMap<SectionId, Vec<HeaderFooterReference>>,
    ) -> String {
        let paragraphs = document.paragraphs();
        let mut xml = String::new();
        xml.push_str(xml_decl());
        xml.push_str(
            r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"><w:body>"#,
        );

        for (idx, paragraph) in paragraphs.iter().enumerate() {
            let ending_section = document
                .sections
                .iter()
                .find(|section| section.starts_at_paragraph == idx + 1)
                .and_then(|_| {
                    document
                        .sections
                        .iter()
                        .rev()
                        .find(|s| s.starts_at_paragraph <= idx)
                });
            if let Some(table) = paragraph.table.as_ref() {
                if let Some(section) = ending_section {
                    write_paragraph(
                        &mut xml,
                        self,
                        paragraph,
                        Some(section_xml(
                            document,
                            section.id,
                            section.page_setup,
                            section_refs,
                        )),
                    );
                }
                write_table(&mut xml, self, table);
                continue;
            }

            let sect_pr = ending_section
                .map(|section| section_xml(document, section.id, section.page_setup, section_refs));
            write_paragraph(&mut xml, self, paragraph, sect_pr);
        }

        let final_section = document
            .sections
            .last()
            .map(|section| (section.id, section.page_setup))
            .unwrap_or_else(|| {
                let setup = document.default_page_setup();
                (document.first_section_id(), setup)
            });
        xml.push_str(&section_xml(
            document,
            final_section.0,
            final_section.1,
            section_refs,
        ));
        xml.push_str("</w:body></w:document>");
        xml
    }

    fn document_rels_xml(&self) -> String {
        let mut xml = String::new();
        xml.push_str(xml_decl());
        xml.push_str(r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">"#);
        for rel in &self.document_rels {
            let _ = write!(
                xml,
                r#"<Relationship Id="{}" Type="{}" Target="{}"/>"#,
                xml_escape(&rel.id),
                rel.rel_type,
                xml_escape(&rel.target)
            );
        }
        xml.push_str("</Relationships>");
        xml
    }

    fn content_types_xml(&self) -> String {
        let mut xml = String::new();
        xml.push_str(xml_decl());
        xml.push_str(
            r#"<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">"#,
        );
        xml.push_str(r#"<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>"#);
        xml.push_str(r#"<Default Extension="xml" ContentType="application/xml"/>"#);

        let mut media_defaults = BTreeMap::<&str, &str>::new();
        for path in self.parts.keys() {
            if let Some(rest) = path.strip_prefix("word/media/") {
                if let Some(ext) = rest.rsplit_once('.').map(|(_, ext)| ext) {
                    media_defaults.insert(ext, image_mime_type_for_extension(ext));
                }
            }
        }
        for (ext, content_type) in media_defaults {
            let _ = write!(
                xml,
                r#"<Default Extension="{}" ContentType="{}"/>"#,
                xml_escape(ext),
                content_type
            );
        }

        let overrides = [
            (
                "/word/document.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
            ),
            (
                "/word/styles.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
            ),
            (
                "/word/numbering.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
            ),
            (
                "/word/settings.xml",
                "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml",
            ),
            (
                "/docProps/core.xml",
                "application/vnd.openxmlformats-package.core-properties+xml",
            ),
            (
                "/docProps/app.xml",
                "application/vnd.openxmlformats-officedocument.extended-properties+xml",
            ),
        ];
        for (part_name, content_type) in overrides {
            let _ = write!(
                xml,
                r#"<Override PartName="{part_name}" ContentType="{content_type}"/>"#
            );
        }
        for path in self.parts.keys() {
            if path.starts_with("word/header") {
                let _ = write!(
                    xml,
                    r#"<Override PartName="/{}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>"#,
                    xml_escape(path)
                );
            } else if path.starts_with("word/footer") {
                let _ = write!(
                    xml,
                    r#"<Override PartName="/{}" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>"#,
                    xml_escape(path)
                );
            }
        }
        xml.push_str("</Types>");
        xml
    }
}

fn write_paragraph(
    xml: &mut String,
    builder: &mut PackageBuilder,
    paragraph: &Paragraph,
    sect_pr: Option<String>,
) {
    xml.push_str("<w:p>");
    write_paragraph_properties(xml, paragraph.style, sect_pr);
    for run in &paragraph.runs {
        write_run(xml, run);
    }
    if let Some(image) = paragraph.image.as_ref() {
        write_image_run(xml, builder, image);
    }
    xml.push_str("</w:p>");
}

fn write_paragraph_properties(
    xml: &mut String,
    style: crate::document::ParagraphStyle,
    sect_pr: Option<String>,
) {
    if style == Default::default() && sect_pr.is_none() {
        return;
    }
    xml.push_str("<w:pPr>");
    if style.alignment != ParagraphAlignment::Left {
        let _ = write!(
            xml,
            r#"<w:jc w:val="{}"/>"#,
            docx_alignment(style.alignment)
        );
    }
    if style.list_kind != ListKind::None {
        let num_id = match style.list_kind {
            ListKind::Bullet => 1,
            ListKind::Ordered => 2,
            ListKind::None => 0,
        };
        let _ = write!(
            xml,
            r#"<w:numPr><w:ilvl w:val="0"/><w:numId w:val="{num_id}"/></w:numPr>"#
        );
    }
    if style.spacing_before_points > 0
        || style.spacing_after_points > 0
        || style.line_spacing != LineSpacing::default()
    {
        let (line, line_rule) = docx_line_spacing(style.line_spacing);
        let _ = write!(
            xml,
            r#"<w:spacing w:before="{}" w:after="{}" w:line="{}" w:lineRule="{}"/>"#,
            points_to_twips(style.spacing_before_points as f32),
            points_to_twips(style.spacing_after_points as f32),
            line,
            line_rule
        );
    }
    if style.page_break_before {
        xml.push_str("<w:pageBreakBefore/>");
    }
    if let Some(sect_pr) = sect_pr {
        xml.push_str(&sect_pr);
    }
    xml.push_str("</w:pPr>");
}

fn write_run(xml: &mut String, run: &TextRun) {
    let mut text = String::new();
    for ch in run.text.chars() {
        if ch != OBJECT_REPLACEMENT_CHAR {
            text.push(ch);
        }
    }
    if text.is_empty() {
        return;
    }
    for segment in split_text_controls(&text) {
        match segment {
            TextSegment::Text(value) => write_text_run(xml, &value, run.style),
            TextSegment::Tab => write_simple_run(xml, run.style, "<w:tab/>"),
            TextSegment::Break => write_simple_run(xml, run.style, "<w:br/>"),
        }
    }
}

fn write_text_run(xml: &mut String, text: &str, style: CharacterStyle) {
    if text.is_empty() {
        return;
    }
    xml.push_str("<w:r>");
    write_run_properties(xml, style);
    if preserve_space(text) {
        let _ = write!(
            xml,
            r#"<w:t xml:space="preserve">{}</w:t>"#,
            xml_escape(text)
        );
    } else {
        let _ = write!(xml, "<w:t>{}</w:t>", xml_escape(text));
    }
    xml.push_str("</w:r>");
}

fn write_simple_run(xml: &mut String, style: CharacterStyle, content: &str) {
    xml.push_str("<w:r>");
    write_run_properties(xml, style);
    xml.push_str(content);
    xml.push_str("</w:r>");
}

fn write_run_properties(xml: &mut String, style: CharacterStyle) {
    xml.push_str("<w:rPr>");
    let font = word_font_name(style);
    let _ = write!(
        xml,
        r#"<w:rFonts w:ascii="{}" w:hAnsi="{}" w:cs="{}"/>"#,
        xml_escape(font),
        xml_escape(font),
        xml_escape(font)
    );
    if style.bold {
        xml.push_str("<w:b/>");
    }
    if style.italic {
        xml.push_str("<w:i/>");
    }
    if style.underline {
        xml.push_str(r#"<w:u w:val="single"/>"#);
    }
    if style.strikethrough {
        xml.push_str("<w:strike/>");
    }
    let _ = write!(
        xml,
        r#"<w:sz w:val="{}"/>"#,
        (style.font_size_points.max(1.0) * 2.0).round() as u16
    );
    let _ = write!(xml, r#"<w:color w:val="{}"/>"#, hex_color(style.text_color));
    if style.highlight_color != Color32::TRANSPARENT {
        let _ = write!(
            xml,
            r#"<w:highlight w:val="{}"/>"#,
            highlight_name(style.highlight_color)
        );
    }
    xml.push_str("</w:rPr>");
}

fn write_table(xml: &mut String, builder: &mut PackageBuilder, table: &DocumentTable) {
    xml.push_str("<w:tbl><w:tblPr>");
    let border = table_border_xml(table);
    let _ = write!(
        xml,
        "<w:tblBorders>{border}{border}{border}{border}{border}{border}</w:tblBorders>"
    );
    xml.push_str("</w:tblPr><w:tblGrid>");
    for width in &table.col_widths_points {
        let _ = write!(xml, r#"<w:gridCol w:w="{}"/>"#, points_to_twips(*width));
    }
    xml.push_str("</w:tblGrid>");
    for (row_idx, row) in table.rows.iter().enumerate() {
        xml.push_str("<w:tr>");
        if let Some(height) = table.row_heights_points.get(row_idx) {
            let _ = write!(
                xml,
                r#"<w:trPr><w:trHeight w:val="{}"/></w:trPr>"#,
                points_to_twips(*height)
            );
        }
        for cell in row {
            xml.push_str("<w:tc><w:tcPr>");
            if cell.col_span > 1 {
                let _ = write!(xml, r#"<w:gridSpan w:val="{}"/>"#, cell.col_span);
            }
            xml.push_str("</w:tcPr><w:p>");
            for run in &cell.runs {
                write_run(xml, run);
            }
            for image in &cell.images {
                write_image_run(xml, builder, image);
            }
            xml.push_str("</w:p></w:tc>");
        }
        xml.push_str("</w:tr>");
    }
    xml.push_str("</w:tbl>");
}

fn write_image_run(xml: &mut String, builder: &mut PackageBuilder, image: &DocumentImage) {
    let (rel_id, _) = builder.add_image_relationship(image);
    builder.next_doc_pr_id += 1;
    let doc_pr_id = builder.next_doc_pr_id;
    let width = points_to_emu(image.width_points.max(1.0));
    let height = points_to_emu(image.height_points.max(1.0));
    xml.push_str("<w:r><w:drawing>");
    if image.layout_mode == ImageLayoutMode::Inline || image.wrap_mode == WrapMode::Inline {
        let _ = write!(
            xml,
            r#"<wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="{width}" cy="{height}"/><wp:docPr id="{doc_pr_id}" name="Picture {doc_pr_id}" descr="{}"/>{}</wp:inline>"#,
            xml_escape(&image.alt_text),
            picture_xml(&rel_id, doc_pr_id, width, height)
        );
    } else {
        let behind_doc = matches!(image.wrap_mode, WrapMode::BehindText);
        let relative_height = image.z_index.max(0) as u32;
        let _ = write!(
            xml,
            r#"<wp:anchor distT="{}" distB="{}" distL="{}" distR="{}" simplePos="0" relativeHeight="{relative_height}" behindDoc="{}" locked="0" layoutInCell="1" allowOverlap="{}"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>{}</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>{}</wp:posOffset></wp:positionV><wp:extent cx="{width}" cy="{height}"/>{}<wp:docPr id="{doc_pr_id}" name="Picture {doc_pr_id}" descr="{}"/>{}</wp:anchor>"#,
            points_to_emu(image.distance_from_text.top_points),
            points_to_emu(image.distance_from_text.bottom_points),
            points_to_emu(image.distance_from_text.left_points),
            points_to_emu(image.distance_from_text.right_points),
            if behind_doc { "1" } else { "0" },
            if image.allow_overlap { "1" } else { "0" },
            points_to_emu(image.horizontal_position.offset_points),
            points_to_emu(image.vertical_position.offset_points),
            wrap_xml(image.wrap_mode),
            xml_escape(&image.alt_text),
            picture_xml(&rel_id, doc_pr_id, width, height)
        );
    }
    xml.push_str("</w:drawing></w:r>");
}

fn picture_xml(rel_id: &str, doc_pr_id: usize, width: i64, height: i64) -> String {
    format!(
        r#"<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="{doc_pr_id}" name="Picture {doc_pr_id}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="{}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{width}" cy="{height}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic>"#,
        xml_escape(rel_id)
    )
}

fn section_xml(
    document: &DocumentState,
    section_id: SectionId,
    setup: PageSetup,
    section_refs: &HashMap<SectionId, Vec<HeaderFooterReference>>,
) -> String {
    let mut xml = String::new();
    xml.push_str("<w:sectPr>");
    if let Some(refs) = section_refs.get(&section_id) {
        for reference in refs {
            let tag = match reference.kind {
                HeaderFooterKind::Header => "w:headerReference",
                HeaderFooterKind::Footer => "w:footerReference",
            };
            let _ = write!(
                xml,
                r#"<{tag} w:type="{}" r:id="{}"/>"#,
                variant_name(reference.variant),
                xml_escape(&reference.rel_id)
            );
        }
    }
    let _ = write!(
        xml,
        r#"<w:pgSz w:w="{}" w:h="{}"/><w:pgMar w:top="{}" w:right="{}" w:bottom="{}" w:left="{}" w:header="{}" w:footer="{}" w:gutter="0"/>"#,
        points_to_twips(setup.page_size.width_points),
        points_to_twips(setup.page_size.height_points),
        points_to_twips(setup.margins.top_points),
        points_to_twips(setup.margins.right_points),
        points_to_twips(setup.margins.bottom_points),
        points_to_twips(setup.margins.left_points),
        points_to_twips(setup.header_from_top_points),
        points_to_twips(setup.footer_from_bottom_points)
    );
    if let Some(section) = document.section_by_id(section_id) {
        if section.different_first_page {
            xml.push_str("<w:titlePg/>");
        }
    }
    if let Some(start) = setup.page_number_start {
        let _ = write!(xml, r#"<w:pgNumType w:start="{start}"/>"#);
    }
    xml.push_str("</w:sectPr>");
    xml
}

fn header_footer_xml(kind: HeaderFooterKind, runs: &[TextRun]) -> String {
    let tag = match kind {
        HeaderFooterKind::Header => "hdr",
        HeaderFooterKind::Footer => "ftr",
    };
    let mut xml = String::new();
    xml.push_str(xml_decl());
    let _ = write!(
        xml,
        r#"<w:{tag} xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:p>"#
    );
    for run in runs {
        write_header_footer_run(&mut xml, run);
    }
    let _ = write!(xml, "</w:p></w:{tag}>");
    xml
}

fn write_header_footer_run(xml: &mut String, run: &TextRun) {
    let mut remaining = run.text.as_str();
    while !remaining.is_empty() {
        let Some((start, token, field)) = find_next_field(remaining) else {
            write_text_run(xml, remaining, run.style);
            break;
        };
        if start > 0 {
            write_text_run(xml, &remaining[..start], run.style);
        }
        write_field(xml, field, run.style);
        remaining = &remaining[start + token.len()..];
    }
}

fn write_field(xml: &mut String, instruction: &str, style: CharacterStyle) {
    write_simple_run(xml, style, r#"<w:fldChar w:fldCharType="begin"/>"#);
    xml.push_str("<w:r>");
    write_run_properties(xml, style);
    let _ = write!(
        xml,
        r#"<w:instrText xml:space="preserve"> {} </w:instrText>"#,
        instruction
    );
    xml.push_str("</w:r>");
    write_simple_run(xml, style, r#"<w:fldChar w:fldCharType="separate"/>"#);
    write_text_run(xml, "1", style);
    write_simple_run(xml, style, r#"<w:fldChar w:fldCharType="end"/>"#);
}

fn find_next_field(value: &str) -> Option<(usize, &'static str, &'static str)> {
    const FIELDS: [(&str, &str); 12] = [
        ("{ SECTIONPAGES }", "SECTIONPAGES"),
        ("{SECTIONPAGES}", "SECTIONPAGES"),
        ("{sectionpages}", "SECTIONPAGES"),
        ("{ NUMPAGES }", "NUMPAGES"),
        ("{NUMPAGES}", "NUMPAGES"),
        ("{pagecount}", "NUMPAGES"),
        ("{pages}", "NUMPAGES"),
        ("{numpages}", "NUMPAGES"),
        ("{ PAGE }", "PAGE"),
        ("{PAGE}", "PAGE"),
        ("{page}", "PAGE"),
        ("{ PAGE}", "PAGE"),
    ];
    FIELDS
        .iter()
        .filter_map(|(token, field)| value.find(token).map(|idx| (idx, *token, *field)))
        .min_by_key(|(idx, _, _)| *idx)
}

fn styles_xml() -> String {
    format!(
        "{}{}",
        xml_decl(),
        r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:cs="Calibri"/><w:sz w:val="24"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults><w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:qFormat/></w:style><w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont"><w:name w:val="Default Paragraph Font"/></w:style></w:styles>"#
    )
}

fn numbering_xml() -> String {
    format!(
        "{}{}",
        xml_decl(),
        r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:abstractNum w:abstractNumId="1"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="•"/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum><w:abstractNum w:abstractNumId="2"><w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:lvlJc w:val="left"/></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="1"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="2"/></w:num></w:numbering>"#
    )
}

fn settings_xml(document: &DocumentState) -> String {
    let mut xml = String::new();
    xml.push_str(xml_decl());
    xml.push_str(
        r#"<w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">"#,
    );
    if document.different_odd_even_pages {
        xml.push_str("<w:evenAndOddHeaders/>");
    }
    xml.push_str(
        r#"<w:compat><w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>"#,
    );
    xml.push_str("</w:settings>");
    xml
}

fn root_rels_xml() -> String {
    format!(
        "{}{}",
        xml_decl(),
        r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>"#
    )
}

fn core_xml(document: &DocumentState) -> String {
    format!(
        r#"{}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>{}</dc:title><dc:creator>wors</dc:creator><cp:lastModifiedBy>wors</cp:lastModifiedBy></cp:coreProperties>"#,
        xml_decl(),
        xml_escape(&document.title)
    )
}

fn app_xml() -> String {
    format!(
        "{}{}",
        xml_decl(),
        r#"<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>wors</Application></Properties>"#
    )
}

fn split_text_controls(text: &str) -> Vec<TextSegment> {
    let mut out = Vec::new();
    let mut current = String::new();
    for ch in text.chars() {
        match ch {
            '\t' => {
                if !current.is_empty() {
                    out.push(TextSegment::Text(std::mem::take(&mut current)));
                }
                out.push(TextSegment::Tab);
            }
            '\n' => {
                if !current.is_empty() {
                    out.push(TextSegment::Text(std::mem::take(&mut current)));
                }
                out.push(TextSegment::Break);
            }
            _ => current.push(ch),
        }
    }
    if !current.is_empty() {
        out.push(TextSegment::Text(current));
    }
    out
}

enum TextSegment {
    Text(String),
    Tab,
    Break,
}

fn table_border_xml(table: &DocumentTable) -> String {
    let size = (table.borders.width_points.max(0.25) * 8.0).round() as u16;
    let color = hex_color(table.borders.color);
    ["top", "left", "bottom", "right", "insideH", "insideV"]
        .into_iter()
        .map(|name| {
            format!(r#"<w:{name} w:val="single" w:sz="{size}" w:space="0" w:color="{color}"/>"#)
        })
        .collect()
}

fn wrap_xml(wrap_mode: WrapMode) -> &'static str {
    match wrap_mode {
        WrapMode::Square => r#"<wp:wrapSquare wrapText="bothSides"/>"#,
        WrapMode::Tight => r#"<wp:wrapTight wrapText="bothSides"/>"#,
        WrapMode::Through => r#"<wp:wrapThrough wrapText="bothSides"/>"#,
        WrapMode::TopAndBottom => "<wp:wrapTopAndBottom/>",
        WrapMode::BehindText | WrapMode::InFrontOfText => "<wp:wrapNone/>",
        WrapMode::Inline => r#"<wp:wrapSquare wrapText="bothSides"/>"#,
    }
}

fn docx_alignment(alignment: ParagraphAlignment) -> &'static str {
    match alignment {
        ParagraphAlignment::Left => "left",
        ParagraphAlignment::Center => "center",
        ParagraphAlignment::Right => "right",
        ParagraphAlignment::Justify => "both",
    }
}

fn docx_line_spacing(line_spacing: LineSpacing) -> (i32, &'static str) {
    match line_spacing.kind {
        LineSpacingKind::AutoMultiplier => {
            ((line_spacing.value.max(0.1) * 240.0).round() as i32, "auto")
        }
        LineSpacingKind::AtLeastPoints => (points_to_twips(line_spacing.value.max(1.0)), "atLeast"),
        LineSpacingKind::ExactPoints => (points_to_twips(line_spacing.value.max(1.0)), "exact"),
    }
}

fn variant_name(variant: HeaderFooterVariant) -> &'static str {
    match variant {
        HeaderFooterVariant::Default => "default",
        HeaderFooterVariant::First => "first",
        HeaderFooterVariant::Even => "even",
    }
}

fn word_font_name(style: CharacterStyle) -> &'static str {
    match style.font_family_name {
        Some("docx-carlito") => "Carlito",
        Some("docx-caladea") => "Caladea",
        Some("docx-liberation-sans") => "Liberation Sans",
        Some("docx-liberation-serif") => "Liberation Serif",
        Some("docx-liberation-mono") => "Liberation Mono",
        Some("docx-comic-sans") => "Comic Neue",
        Some(name) => name,
        None => match style.font_choice {
            FontChoice::Proportional | FontChoice::Carlito => "Carlito",
            FontChoice::Monospace | FontChoice::LiberationMono => "Liberation Mono",
            FontChoice::Caladea => "Caladea",
            FontChoice::LiberationSans => "Liberation Sans",
            FontChoice::LiberationSerif => "Liberation Serif",
            FontChoice::ComicSans => "Comic Neue",
        },
    }
}

fn preserve_space(text: &str) -> bool {
    text.starts_with(' ') || text.ends_with(' ') || text.contains("  ")
}

fn highlight_name(color: Color32) -> &'static str {
    let candidates = [
        ("yellow", Color32::from_rgb(255, 242, 129)),
        ("green", Color32::from_rgb(187, 232, 172)),
        ("cyan", Color32::from_rgb(163, 231, 240)),
        ("magenta", Color32::from_rgb(244, 188, 231)),
        ("blue", Color32::from_rgb(177, 205, 252)),
        ("red", Color32::from_rgb(248, 188, 188)),
        ("darkYellow", Color32::from_rgb(215, 185, 90)),
        ("darkGreen", Color32::from_rgb(104, 170, 112)),
        ("darkBlue", Color32::from_rgb(99, 129, 207)),
    ];
    candidates
        .iter()
        .min_by_key(|(_, candidate)| color_distance(color, *candidate))
        .map(|(name, _)| *name)
        .unwrap_or("yellow")
}

fn color_distance(a: Color32, b: Color32) -> u32 {
    let dr = a.r() as i32 - b.r() as i32;
    let dg = a.g() as i32 - b.g() as i32;
    let db = a.b() as i32 - b.b() as i32;
    (dr * dr + dg * dg + db * db) as u32
}

fn hex_color(color: Color32) -> String {
    format!("{:02X}{:02X}{:02X}", color.r(), color.g(), color.b())
}

fn points_to_twips(points: f32) -> i32 {
    (points * 20.0).round() as i32
}

fn points_to_emu(points: f32) -> i64 {
    (points * 12_700.0).round() as i64
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

fn xml_decl() -> &'static str {
    r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?>"#
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

const REL_IMAGE: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image";
const REL_HEADER: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header";
const REL_FOOTER: &str =
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer";
