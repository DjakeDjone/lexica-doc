use std::{
    collections::HashMap,
    io::{Cursor, Read},
};

use super::numbering::parse_numbering_xml;
use super::styles::parse_styles_xml;
use super::styles::parse_theme_xml;
use super::{document_to_docx, parse_document_relationships, parse_document_xml};
use crate::document::{
    CharacterStyle, DistanceFromText, DocumentImage, DocumentState, DocumentTable,
    HeaderFooterStory, ImageLayoutMode, ImageRendering, LineSpacing, LineSpacingKind, ListKind,
    PageMargins, PageSetup, PageSize, ParagraphAlignment, ParagraphStyle, PositionAlign, Section,
    TableCell, TextRun, VerticalAlign, VerticalPosition, VerticalRelativeTo, WrapMode,
    OBJECT_REPLACEMENT_CHAR,
};
use eframe::egui::Color32;
use zip::ZipArchive;

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

#[test]
fn exports_docx_package_with_rich_word_parts() {
    let bytes = document_to_docx(&rich_export_document()).expect("docx export");
    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("docx zip");

    let content_types = zip_text(&mut archive, "[Content_Types].xml");
    assert!(content_types.contains("wordprocessingml.document.main+xml"));
    assert!(content_types.contains("wordprocessingml.header+xml"));
    assert!(content_types.contains("wordprocessingml.footer+xml"));
    assert!(content_types.contains("image/png"));

    let rels = zip_text(&mut archive, "word/_rels/document.xml.rels");
    assert!(rels.contains("/relationships/image"));
    assert!(rels.contains("/relationships/header"));
    assert!(rels.contains("/relationships/footer"));

    assert!(archive.by_name("word/media/image-1.png").is_ok());
    assert!(archive.by_name("word/header1.xml").is_ok());
    assert!(archive.by_name("word/footer1.xml").is_ok());
    assert!(archive.by_name("word/styles.xml").is_ok());
    assert!(archive.by_name("word/numbering.xml").is_ok());
    assert!(archive.by_name("word/settings.xml").is_ok());
}

#[test]
fn parses_docx_vertical_align_runs() {
    let imported = parse_document_xml(
        r#"
        <w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
          <w:body>
            <w:p>
              <w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:t>2</w:t></w:r>
              <w:r><w:rPr><w:vertAlign w:val="subscript"/></w:rPr><w:t>n</w:t></w:r>
            </w:p>
          </w:body>
        </w:document>
        "#,
        &Default::default(),
        &Default::default(),
        &Default::default(),
        &Default::default(),
        &HashMap::new(),
    )
    .expect("document xml");

    assert_eq!(
        imported.runs[0].style.vertical_align,
        VerticalAlign::Superscript
    );
    assert_eq!(
        imported.runs[1].style.vertical_align,
        VerticalAlign::Subscript
    );
}

#[test]
fn exports_docx_vertical_align_runs() {
    let mut document = DocumentState::bootstrap();
    document.replace_with_runs(
        "Vertical".to_owned(),
        vec![
            TextRun {
                text: "2".to_owned(),
                style: CharacterStyle {
                    vertical_align: VerticalAlign::Superscript,
                    ..CharacterStyle::default()
                },
            },
            TextRun {
                text: "n".to_owned(),
                style: CharacterStyle {
                    vertical_align: VerticalAlign::Subscript,
                    ..CharacterStyle::default()
                },
            },
        ],
    );

    let bytes = document_to_docx(&document).expect("docx export");
    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("docx zip");
    let document_xml = zip_text(&mut archive, "word/document.xml");

    assert!(document_xml.contains(r#"<w:vertAlign w:val="superscript"/>"#));
    assert!(document_xml.contains(r#"<w:vertAlign w:val="subscript"/>"#));
}

#[test]
fn exports_docx_document_xml_for_formatting_tables_images_and_sections() {
    let bytes = document_to_docx(&rich_export_document()).expect("docx export");
    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("docx zip");
    let document_xml = zip_text(&mut archive, "word/document.xml");

    assert!(document_xml.contains("<w:b/>"));
    assert!(document_xml.contains("<w:i/>"));
    assert!(document_xml.contains(r#"<w:u w:val="single"/>"#));
    assert!(document_xml.contains("<w:strike/>"));
    assert!(document_xml.contains(r#"<w:sz w:val="32"/>"#));
    assert!(document_xml.contains(r#"<w:color w:val="102030"/>"#));
    assert!(document_xml.contains(r#"<w:highlight w:val="yellow"/>"#));
    assert!(document_xml.contains(r#"<w:rFonts w:ascii="Liberation Serif""#));
    assert!(document_xml.contains(r#"<w:jc w:val="center"/>"#));
    assert!(document_xml
        .contains(r#"<w:spacing w:before="120" w:after="160" w:line="360" w:lineRule="auto"/>"#));
    assert!(document_xml.contains("<w:pageBreakBefore/>"));
    assert!(document_xml.contains(r#"<w:numId w:val="1"/>"#));
    assert!(document_xml.contains(r#"<w:numId w:val="2"/>"#));
    assert!(document_xml.contains("<w:tbl>"));
    assert!(document_xml.contains(r#"<w:gridCol w:w="1440"/>"#));
    assert!(document_xml.contains(r#"<w:gridSpan w:val="2"/>"#));
    assert!(document_xml.contains("<wp:inline"));
    assert!(document_xml.contains("<wp:anchor"));
    assert!(document_xml.contains("<wp:wrapSquare"));
    assert!(document_xml.contains(r#"<w:pgSz w:w="12240" w:h="15840"/>"#));
    assert!(document_xml
        .contains(r#"<w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800""#));
    assert!(document_xml.contains("<w:titlePg/>"));
    assert!(document_xml.contains(r#"<w:pgNumType w:start="3"/>"#));
}

#[test]
fn exports_docx_numbering_headers_footers_and_save_bytes() {
    let document = rich_export_document();
    let bytes = document
        .export_bytes_for_extension("docx")
        .expect("docx bytes");
    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("docx zip");

    let numbering_xml = zip_text(&mut archive, "word/numbering.xml");
    assert!(numbering_xml.contains(r#"<w:num w:numId="1">"#));
    assert!(numbering_xml.contains(r#"<w:num w:numId="2">"#));

    let settings_xml = zip_text(&mut archive, "word/settings.xml");
    assert!(settings_xml.contains("<w:evenAndOddHeaders/>"));

    let header_xml = zip_text(&mut archive, "word/header1.xml");
    assert!(header_xml.contains("<w:instrText"));
    assert!(header_xml.contains(" PAGE "));
    let footer_xml = zip_text(&mut archive, "word/footer1.xml");
    assert!(footer_xml.contains(" NUMPAGES "));
}

#[test]
fn saves_docx_extension_to_readable_zip() {
    let mut path = std::env::temp_dir();
    path.push(format!(
        "wors-docx-export-{}.docx",
        std::time::SystemTime::now()
            .duration_since(std::time::UNIX_EPOCH)
            .expect("clock")
            .as_nanos()
    ));

    rich_export_document()
        .save_to_path(&path)
        .expect("docx save");
    let bytes = std::fs::read(&path).expect("saved docx");
    let mut archive = ZipArchive::new(Cursor::new(bytes)).expect("saved docx zip");
    assert!(archive.by_name("word/document.xml").is_ok());

    let _ = std::fs::remove_file(path);
}

fn rich_export_document() -> DocumentState {
    let inline_image = test_docx_image(1, ImageLayoutMode::Inline, WrapMode::Inline);
    let mut floating_image = test_docx_image(2, ImageLayoutMode::Floating, WrapMode::Square);
    floating_image.distance_from_text = DistanceFromText {
        top_points: 2.0,
        right_points: 4.0,
        bottom_points: 3.0,
        left_points: 5.0,
    };
    floating_image.vertical_position = VerticalPosition {
        relative_to: VerticalRelativeTo::Paragraph,
        align: Some(PositionAlign::Start),
        offset_points: 12.0,
    };
    floating_image.set_manual_offset(18.0, 12.0);

    let mut table = DocumentTable::new(1, 1, 2, 216.0);
    table.col_widths_points = vec![72.0, 144.0];
    table.rows[0][0] = TableCell {
        runs: vec![TextRun {
            text: "Table".to_owned(),
            style: CharacterStyle {
                bold: true,
                ..CharacterStyle::default()
            },
        }],
        images: Vec::new(),
        col_span: 2,
        row_span: 1,
    };

    let mut document = DocumentState::bootstrap();
    document.title = "Word Export".to_owned();
    document.runs = vec![
        TextRun {
            text: "Styled".to_owned(),
            style: CharacterStyle {
                bold: true,
                italic: true,
                underline: true,
                strikethrough: true,
                font_size_points: 16.0,
                font_family_name: Some("docx-liberation-serif"),
                text_color: Color32::from_rgb(0x10, 0x20, 0x30),
                highlight_color: Color32::from_rgb(255, 242, 129),
                ..CharacterStyle::default()
            },
        },
        TextRun {
            text: format!("\nBullet\nNumber\n{OBJECT_REPLACEMENT_CHAR}\n{OBJECT_REPLACEMENT_CHAR}\n{OBJECT_REPLACEMENT_CHAR}"),
            style: CharacterStyle::default(),
        },
    ];
    document.paragraph_styles = vec![
        ParagraphStyle {
            alignment: ParagraphAlignment::Center,
            spacing_before_points: 6,
            spacing_after_points: 8,
            line_spacing: LineSpacing {
                kind: LineSpacingKind::AutoMultiplier,
                value: 1.5,
            },
            ..ParagraphStyle::default()
        },
        ParagraphStyle {
            list_kind: ListKind::Bullet,
            page_break_before: true,
            ..ParagraphStyle::default()
        },
        ParagraphStyle {
            list_kind: ListKind::Ordered,
            ..ParagraphStyle::default()
        },
        ParagraphStyle::default(),
        ParagraphStyle::default(),
        ParagraphStyle::default(),
    ];
    document.paragraph_images = vec![
        None,
        None,
        None,
        Some(inline_image),
        Some(floating_image),
        None,
    ];
    document.paragraph_tables = vec![None, None, None, None, None, Some(table)];
    document.page_size = PageSize {
        width_points: 612.0,
        height_points: 792.0,
    };
    document.margins = PageMargins {
        top_points: 72.0,
        right_points: 90.0,
        bottom_points: 72.0,
        left_points: 90.0,
    };
    document.different_odd_even_pages = true;
    document.sections = vec![Section::first(PageSetup {
        page_size: document.page_size,
        margins: document.margins,
        header_from_top_points: 36.0,
        footer_from_bottom_points: 36.0,
        page_number_start: Some(3),
    })];
    document.sections[0].different_first_page = true;
    document.sections[0].header_footer.header_default.story =
        HeaderFooterStory::from_runs(vec![TextRun {
            text: "Page { PAGE }".to_owned(),
            style: CharacterStyle::default(),
        }]);
    document.sections[0].header_footer.footer_default.story =
        HeaderFooterStory::from_runs(vec![TextRun {
            text: "Total { NUMPAGES }".to_owned(),
            style: CharacterStyle::default(),
        }]);
    document.sync_compat_from_first_section();
    document
}

fn test_docx_image(id: usize, layout_mode: ImageLayoutMode, wrap_mode: WrapMode) -> DocumentImage {
    DocumentImage {
        id,
        bytes: b"\x89PNG\r\n\x1a\nfake".to_vec(),
        alt_text: format!("image-{id}"),
        width_points: 72.0,
        height_points: 36.0,
        lock_aspect_ratio: true,
        opacity: 1.0,
        layout_mode,
        wrap_mode,
        rendering: ImageRendering::Smooth,
        horizontal_position: Default::default(),
        vertical_position: Default::default(),
        distance_from_text: Default::default(),
        z_index: 1,
        move_with_text: true,
        allow_overlap: true,
    }
}

fn zip_text(archive: &mut ZipArchive<Cursor<Vec<u8>>>, path: &str) -> String {
    let mut file = archive
        .by_name(path)
        .unwrap_or_else(|_| panic!("missing {path}"));
    let mut text = String::new();
    file.read_to_string(&mut text)
        .unwrap_or_else(|_| panic!("failed to read {path}"));
    text
}
