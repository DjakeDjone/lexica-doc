use std::{
    fs,
    time::{SystemTime, UNIX_EPOCH},
};

use super::{
    empty_header_footer_runs, export::plain_text_from_runs, text_format, CharacterStyle,
    DocumentImage, DocumentState, FontChoice, HeaderFooterKind, HeaderFooterStory,
    HeaderFooterVariant, ImageLayoutMode, ImageRendering, ListKind, ParagraphStyle, TextRun,
    VerticalAlign, WrapMode, DOCX_BODY_BOLD, DOCX_CARLITO_BOLD, DOCX_LIBERATION_MONO_BOLD,
    OBJECT_REPLACEMENT_CHAR,
};

fn test_image(id: usize) -> DocumentImage {
    DocumentImage {
        id,
        bytes: vec![1, 2, 3],
        alt_text: format!("image-{id}"),
        width_points: 120.0,
        height_points: 60.0,
        lock_aspect_ratio: true,
        opacity: 1.0,
        layout_mode: ImageLayoutMode::Inline,
        wrap_mode: WrapMode::Inline,
        rendering: ImageRendering::Smooth,
        horizontal_position: Default::default(),
        vertical_position: Default::default(),
        distance_from_text: Default::default(),
        z_index: 7,
        move_with_text: true,
        allow_overlap: false,
    }
}

#[test]
fn bold_text_uses_registered_bold_font_faces() {
    let body = text_format(
        CharacterStyle {
            bold: true,
            ..CharacterStyle::default()
        },
        1.0,
    );
    assert_eq!(
        body.font_id.family,
        eframe::egui::FontFamily::Name(DOCX_BODY_BOLD.into())
    );
    assert_eq!(
        body.font_id.size,
        CharacterStyle::default().font_size_points
    );

    let monospace = text_format(
        CharacterStyle {
            bold: true,
            font_choice: FontChoice::LiberationMono,
            ..CharacterStyle::default()
        },
        1.0,
    );
    assert_eq!(
        monospace.font_id.family,
        eframe::egui::FontFamily::Name(DOCX_LIBERATION_MONO_BOLD.into())
    );

    let imported_family = text_format(
        CharacterStyle {
            bold: true,
            font_family_name: Some("docx-carlito"),
            ..CharacterStyle::default()
        },
        1.0,
    );
    assert_eq!(
        imported_family.font_id.family,
        eframe::egui::FontFamily::Name(DOCX_CARLITO_BOLD.into())
    );
}

#[test]
fn character_style_defaults_to_baseline_vertical_align() {
    assert_eq!(
        CharacterStyle::default().vertical_align,
        VerticalAlign::Baseline
    );
}

#[test]
fn vertical_align_text_format_uses_smaller_raised_or_lowered_text() {
    let baseline = text_format(CharacterStyle::default(), 1.0);
    assert_eq!(baseline.font_id.size, 12.0);
    assert_eq!(baseline.valign, eframe::egui::Align::BOTTOM);

    let superscript = text_format(
        CharacterStyle {
            vertical_align: VerticalAlign::Superscript,
            ..CharacterStyle::default()
        },
        1.0,
    );
    assert!((superscript.font_id.size - 7.8).abs() < 0.001);
    assert_eq!(superscript.valign, eframe::egui::Align::TOP);

    let subscript = text_format(
        CharacterStyle {
            vertical_align: VerticalAlign::Subscript,
            ..CharacterStyle::default()
        },
        1.0,
    );
    assert!((subscript.font_id.size - 7.8).abs() < 0.001);
    assert_eq!(subscript.valign, eframe::egui::Align::BOTTOM);
}

#[test]
fn selected_style_uses_last_selected_character_at_run_boundary() {
    let mut document = DocumentState::bootstrap();
    document.replace_with_runs(
        "Test".to_owned(),
        vec![
            TextRun {
                text: "Bold".to_owned(),
                style: CharacterStyle {
                    bold: true,
                    ..CharacterStyle::default()
                },
            },
            TextRun {
                text: " plain".to_owned(),
                style: CharacterStyle::default(),
            },
        ],
    );

    assert!(document.selection_style_at(0..4).bold);
    assert!(!document.selection_style_at(0..5).bold);
}

#[test]
fn exports_vertical_align_to_html_pdf_and_markdown() {
    let mut document = DocumentState::bootstrap();
    document.replace_with_runs(
        "Test".to_owned(),
        vec![
            TextRun {
                text: "super".to_owned(),
                style: CharacterStyle {
                    vertical_align: VerticalAlign::Superscript,
                    ..CharacterStyle::default()
                },
            },
            TextRun {
                text: "sub".to_owned(),
                style: CharacterStyle {
                    vertical_align: VerticalAlign::Subscript,
                    ..CharacterStyle::default()
                },
            },
        ],
    );

    let html = document.to_html();
    assert!(html.contains("vertical-align:super;font-size:65%;"));
    assert!(html.contains("vertical-align:sub;font-size:65%;"));

    let pdf_html = document.to_pdf_html();
    assert!(pdf_html.contains("vertical-align:super;font-size:65%;"));
    assert!(pdf_html.contains("vertical-align:sub;font-size:65%;"));

    let markdown = document.to_markdown();
    assert!(markdown.contains("<sup>super</sup>"));
    assert!(markdown.contains("<sub>sub</sub>"));
}

#[test]
fn inserts_page_break_between_split_paragraphs() {
    let mut document = DocumentState::bootstrap();
    document.replace_with_runs(
        "Test".to_owned(),
        vec![TextRun {
            text: "alpha beta".to_owned(),
            style: CharacterStyle::default(),
        }],
    );

    let cursor = document.insert_page_break(6);
    let paragraphs = document.paragraphs();

    assert_eq!(cursor, paragraphs[1].range.start);
    assert_eq!(paragraphs.len(), 2);
    assert_eq!(plain_text_from_runs(&paragraphs[0].runs), "alpha ");
    assert_eq!(plain_text_from_runs(&paragraphs[1].runs), "beta");
    assert!(paragraphs[1].style.page_break_before);
}

#[test]
fn inserts_block_image_as_its_own_paragraph() {
    let mut document = DocumentState::bootstrap();
    document.replace_with_runs(
        "Test".to_owned(),
        vec![TextRun {
            text: "alpha beta".to_owned(),
            style: CharacterStyle::default(),
        }],
    );

    let cursor = document.insert_image(
        6,
        DocumentImage {
            id: 1,
            bytes: vec![1, 2, 3],
            alt_text: "diagram".to_owned(),
            width_points: 120.0,
            height_points: 60.0,
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
        },
    );
    let paragraphs = document.paragraphs();

    assert_eq!(cursor, paragraphs[1].range.end);
    assert_eq!(paragraphs.len(), 3);
    assert_eq!(
        paragraphs[1]
            .image
            .as_ref()
            .map(|image| image.alt_text.as_str()),
        Some("diagram")
    );
    assert_eq!(paragraphs[1].style.list_kind, ListKind::None);
    assert_eq!(
        document.plain_text(),
        format!("alpha \n{OBJECT_REPLACEMENT_CHAR}\nbeta")
    );
}

#[test]
fn formats_empty_table_cell_and_uses_style_for_inserted_text() {
    let mut document = DocumentState::bootstrap();
    document.replace_with_runs("Test".to_owned(), Vec::new());
    document.insert_table(0, 1, 1);
    let table_id = document
        .paragraph_tables
        .iter()
        .flatten()
        .next()
        .unwrap()
        .id;

    document.apply_style_to_table_cell(table_id, 0, 0, |style| {
        style.bold = true;
        style.font_size_points = 18.0;
    });
    let active_style = document.table_cell_typing_style(table_id, 0, 0).unwrap();
    document.append_table_cell_text(table_id, 0, 0, "Styled", active_style);

    let cell = &document.table_by_id(table_id).unwrap().rows[0][0];
    assert_eq!(cell.plain_text(), "Styled");
    assert!(cell.runs[0].style.bold);
    assert_eq!(cell.runs[0].style.font_size_points, 18.0);
}

#[test]
fn inserts_image_into_table_cell() {
    let mut document = DocumentState::bootstrap();
    document.replace_with_runs("Test".to_owned(), Vec::new());
    document.insert_table(0, 1, 1);
    let table_id = document
        .paragraph_tables
        .iter()
        .flatten()
        .next()
        .unwrap()
        .id;

    document.insert_table_cell_image(table_id, 0, 0, test_image(42), CharacterStyle::default());

    let cell = &document.table_by_id(table_id).unwrap().rows[0][0];
    assert_eq!(cell.images.len(), 1);
    assert_eq!(cell.plain_text(), OBJECT_REPLACEMENT_CHAR.to_string());
    assert!(document
        .to_markdown()
        .contains("![image-42](embedded-image)"));
}

#[test]
fn moves_image_paragraph_later_without_extra_blank_lines() {
    let image = test_image(7);
    let mut document = DocumentState {
        title: "Test".to_owned(),
        runs: vec![TextRun {
            text: format!("alpha\n{OBJECT_REPLACEMENT_CHAR}\nbeta\ngamma"),
            style: CharacterStyle::default(),
        }],
        paragraph_styles: vec![
            Default::default(),
            ParagraphStyle {
                page_break_before: true,
                ..Default::default()
            },
            Default::default(),
            Default::default(),
        ],
        paragraph_images: vec![None, Some(image.clone()), None, None],
        paragraph_tables: vec![None; 4],
        page_size: super::PageSize::a4(),
        margins: super::PageMargins::standard(),
        header_text: String::new(),
        footer_text: String::new(),
        first_page_header_text: String::new(),
        first_page_footer_text: String::new(),
        even_page_header_text: String::new(),
        even_page_footer_text: String::new(),
        header_runs: empty_header_footer_runs(),
        footer_runs: empty_header_footer_runs(),
        first_page_header_runs: empty_header_footer_runs(),
        first_page_footer_runs: empty_header_footer_runs(),
        even_page_header_runs: empty_header_footer_runs(),
        even_page_footer_runs: empty_header_footer_runs(),
        different_first_page: false,
        different_odd_even_pages: false,
        page_number_start: 1,
        sections: vec![super::Section::first(super::PageSetup::standard())],
    };

    let cursor = document
        .move_image_paragraph_to_cursor(7, document.total_chars())
        .expect("image should move");
    let paragraphs = document.paragraphs();

    assert_eq!(
        document.plain_text(),
        format!("alpha\nbeta\ngamma\n{OBJECT_REPLACEMENT_CHAR}")
    );
    assert_eq!(cursor, paragraphs[3].range.start);
    assert_eq!(paragraphs[3].image.as_ref().map(|image| image.id), Some(7));
    assert_eq!(paragraphs[3].image.as_ref().unwrap().z_index, image.z_index);
    assert!(paragraphs[3].style.page_break_before);
    assert_eq!(
        document
            .plain_text()
            .chars()
            .filter(|ch| *ch == '\n')
            .count(),
        3
    );
}

#[test]
fn moves_image_paragraph_earlier_without_losing_metadata() {
    let mut image = test_image(8);
    image.layout_mode = ImageLayoutMode::Floating;
    image.wrap_mode = WrapMode::Square;
    image.horizontal_position.offset_points = 42.0;
    let mut document = DocumentState {
        title: "Test".to_owned(),
        runs: vec![TextRun {
            text: format!("alpha\nbeta\n{OBJECT_REPLACEMENT_CHAR}\ngamma"),
            style: CharacterStyle::default(),
        }],
        paragraph_styles: vec![
            Default::default(),
            Default::default(),
            Default::default(),
            Default::default(),
        ],
        paragraph_images: vec![None, None, Some(image.clone()), None],
        paragraph_tables: vec![None; 4],
        page_size: super::PageSize::a4(),
        margins: super::PageMargins::standard(),
        header_text: String::new(),
        footer_text: String::new(),
        first_page_header_text: String::new(),
        first_page_footer_text: String::new(),
        even_page_header_text: String::new(),
        even_page_footer_text: String::new(),
        header_runs: empty_header_footer_runs(),
        footer_runs: empty_header_footer_runs(),
        first_page_header_runs: empty_header_footer_runs(),
        first_page_footer_runs: empty_header_footer_runs(),
        even_page_header_runs: empty_header_footer_runs(),
        even_page_footer_runs: empty_header_footer_runs(),
        different_first_page: false,
        different_odd_even_pages: false,
        page_number_start: 1,
        sections: vec![super::Section::first(super::PageSetup::standard())],
    };

    let cursor = document
        .move_image_paragraph_to_cursor(8, 0)
        .expect("image should move");
    let paragraphs = document.paragraphs();

    assert_eq!(
        document.plain_text(),
        format!("{OBJECT_REPLACEMENT_CHAR}\nalpha\nbeta\ngamma")
    );
    assert_eq!(cursor, 0);
    let moved = paragraphs[0].image.as_ref().expect("moved image");
    assert_eq!(moved.id, 8);
    assert_eq!(moved.layout_mode, ImageLayoutMode::Floating);
    assert_eq!(moved.wrap_mode, WrapMode::Square);
    assert_eq!(moved.horizontal_position.offset_points, 42.0);
}

#[test]
fn exports_html_with_styled_runs() {
    let mut document = DocumentState::bootstrap();
    document.header_text = "Draft {page}".to_owned();
    document.footer_text = "Page {page} of {pagecount}".to_owned();
    document.replace_with_runs(
        "Styled".to_owned(),
        vec![
            TextRun {
                text: "Bold".to_owned(),
                style: CharacterStyle {
                    bold: true,
                    ..CharacterStyle::default()
                },
            },
            TextRun {
                text: " + ".to_owned(),
                style: CharacterStyle::default(),
            },
            TextRun {
                text: "Mono".to_owned(),
                style: CharacterStyle {
                    font_choice: super::FontChoice::Monospace,
                    ..CharacterStyle::default()
                },
            },
        ],
    );

    let html = document.to_html();
    assert!(html.contains("<!doctype html>"));
    assert!(html.contains("font-weight:700;"));
    assert!(html.contains("Bold"));
    assert!(html.contains("Mono"));
    assert!(html.contains("Draft 1"));
    assert!(html.contains("Page 1 of 1"));
}

#[test]
fn renders_page_fields_with_total_page_aliases() {
    let mut document = DocumentState::bootstrap();
    document.page_number_start = 3;
    assert_eq!(
        document.render_page_field(
            "Page {page} / { PAGE } of {pagecount} / { NUMPAGES } / { SECTIONPAGES }",
            2,
            7
        ),
        "Page 4 / 4 of 7 / 7 / 7"
    );
}

#[test]
fn resolves_linked_section_header_to_previous_section() {
    let mut document = DocumentState::bootstrap();
    document.sections[0].header_footer.header_default.story =
        HeaderFooterStory::from_runs(vec![TextRun {
            text: "Section 1".to_owned(),
            style: CharacterStyle::default(),
        }]);
    let section_id = document.insert_section_break_before_paragraph(1);

    let resolved = document.resolve_header_footer_slot(
        section_id,
        HeaderFooterKind::Header,
        HeaderFooterVariant::Default,
    );

    assert!(resolved.inherited);
    assert_eq!(resolved.source_section_id, 1);
    assert_eq!(resolved.story.plain_text(), "Section 1");
}

#[test]
fn editing_linked_header_materializes_local_copy() {
    let mut document = DocumentState::bootstrap();
    document.sections[0].header_footer.header_default.story =
        HeaderFooterStory::from_runs(vec![TextRun {
            text: "Inherited".to_owned(),
            style: CharacterStyle::default(),
        }]);
    let section_id = document.insert_section_break_before_paragraph(1);
    let story = document
        .header_footer_story_mut_materialized(
            section_id,
            HeaderFooterKind::Header,
            HeaderFooterVariant::Default,
        )
        .expect("section story");
    story.runs[0].text = "Local".to_owned();

    assert!(!document.header_footer_linked(
        section_id,
        HeaderFooterKind::Header,
        HeaderFooterVariant::Default
    ));
    assert_eq!(
        document
            .resolve_header_footer_slot(
                section_id,
                HeaderFooterKind::Header,
                HeaderFooterVariant::Default
            )
            .story
            .plain_text(),
        "Local"
    );
    assert_eq!(
        document.sections[0]
            .header_footer
            .header_default
            .story
            .plain_text(),
        "Inherited"
    );
}

#[test]
fn section_variants_and_page_fields_resolve_per_section() {
    let mut document = DocumentState::bootstrap();
    document.different_odd_even_pages = true;
    document.sections[0].different_first_page = true;
    document.sections[0].page_setup.page_number_start = Some(3);

    assert_eq!(
        document.header_footer_variant_for_page(1, 0, HeaderFooterKind::Header),
        HeaderFooterVariant::First
    );
    assert_eq!(
        document.header_footer_variant_for_page(1, 1, HeaderFooterKind::Header),
        HeaderFooterVariant::Even
    );
    assert_eq!(
        document.render_page_field_for_section_page(
            "Page { PAGE } of { NUMPAGES } / { SECTIONPAGES }",
            1,
            1,
            1,
            9,
            4
        ),
        "Page 4 of 9 / 4"
    );
}

#[test]
fn exports_pdf_html_with_pdf_friendly_css() {
    let mut document = DocumentState::bootstrap();
    document.replace_with_runs(
        "Styled".to_owned(),
        vec![TextRun {
            text: "Bold".to_owned(),
            style: CharacterStyle {
                bold: true,
                ..CharacterStyle::default()
            },
        }],
    );

    let html = document.to_pdf_html();
    assert!(html.contains("font-family: 'LiberationSans-Regular'"));
    assert!(html.contains("font-size:"));
    assert!(html.contains("px"));
    assert!(html.contains("font-weight:bold;"));
    assert!(!html.contains("box-shadow"));
}

#[test]
fn saves_pdf_extension() {
    let mut path = std::env::temp_dir();
    let stamp = SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .expect("clock")
        .as_nanos();
    path.push(format!("wors-export-{stamp}.pdf"));

    let document = DocumentState::bootstrap();
    document
        .save_to_path(&path)
        .expect("pdf save should succeed");

    let bytes = fs::read(&path).expect("pdf should be readable");
    assert!(bytes.starts_with(b"%PDF"));

    let _ = fs::remove_file(path);
}

#[test]
fn test_print_html_css() {
    let mut doc = DocumentState::bootstrap();
    let style = CharacterStyle {
        bold: true,
        italic: false,
        underline: false,
        strikethrough: false,
        vertical_align: VerticalAlign::Baseline,
        font_size_points: 11.5,
        font_choice: FontChoice::Proportional,
        font_family_name: Some("Carlito"),
        text_color: eframe::egui::Color32::from_rgb(36, 39, 46),
        highlight_color: eframe::egui::Color32::TRANSPARENT,
    };
    doc.runs.push(TextRun {
        text: "Einzug: ".to_string(),
        style,
    });
    doc.paragraph_styles.insert(0, ParagraphStyle::default());
    let html = doc.to_pdf_html();
    println!("DUMP HTML:\n{}", html);
}
