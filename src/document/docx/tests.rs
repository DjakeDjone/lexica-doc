use std::collections::HashMap;

use super::numbering::parse_numbering_xml;
use super::styles::parse_styles_xml;
use super::styles::parse_theme_xml;
use super::{parse_document_relationships, parse_document_xml};
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
