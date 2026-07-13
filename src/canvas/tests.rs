use super::*;
use crate::canvas::layout::DocumentLayout;
use crate::{
    app::{CanvasState, ZoomMode},
    document::{
        CharacterStyle, DocumentImage, DocumentState, ImageLayoutMode, ImageRendering, LineSpacing,
        LineSpacingKind, ListKind, PageMargins, PageSize, ParagraphAlignment, ParagraphStyle,
        TextRun, WrapMode, OBJECT_REPLACEMENT_CHAR,
    },
    layout::fit_page_zoom,
};
use std::sync::Arc;

/// Run `layout_document` inside a headless egui context and return
/// the layout result for assertion.
fn run_headless_layout(
    document: &DocumentState,
    canvas: &CanvasState,
    wrap_width: f32,
) -> DocumentLayout {
    let ctx = egui::Context::default();
    let mut layout = None;

    let _ = ctx.run_ui(egui::RawInput::default(), |ui| {
        layout = Some(layout_document(ui, document, canvas, wrap_width));
    });

    layout.expect("layout_document should have been called inside the egui frame")
}

#[test]
fn unchanged_document_layout_is_reused() {
    let ctx = egui::Context::default();
    let mut document = make_document(
        vec![TextRun {
            text: "cached layout".to_owned(),
            style: CharacterStyle::default(),
        }],
        vec![ParagraphStyle::default()],
        vec![None],
    );
    let mut canvas = CanvasState::default();
    let mut layouts = Vec::new();

    for _ in 0..2 {
        let _ = ctx.run_ui(egui::RawInput::default(), |ui| {
            layouts.push(cached_layout_document(ui, &document, &canvas, 500.0));
        });
    }

    assert!(Arc::ptr_eq(&layouts[0], &layouts[1]));

    canvas.zoom = 1.25;
    let _ = ctx.run_ui(egui::RawInput::default(), |ui| {
        layouts.push(cached_layout_document(ui, &document, &canvas, 500.0));
    });
    assert!(!Arc::ptr_eq(&layouts[1], &layouts[2]));

    document.insert_text(0, "updated ", CharacterStyle::default());
    let _ = ctx.run_ui(egui::RawInput::default(), |ui| {
        layouts.push(cached_layout_document(ui, &document, &canvas, 500.0));
    });
    assert!(!Arc::ptr_eq(&layouts[2], &layouts[3]));
}

#[test]
fn first_line_and_hanging_indents_move_only_the_first_wrapped_row() {
    let mut document = make_document(
        vec![TextRun {
            text: "one two three four five six seven eight nine ten eleven twelve".to_owned(),
            style: CharacterStyle::default(),
        }],
        vec![ParagraphStyle {
            left_indent_points: 24.0,
            right_indent_points: 24.0,
            first_line_indent_points: 18.0,
            ..ParagraphStyle::default()
        }],
        vec![None],
    );
    let canvas = CanvasState::default();

    let first_line = run_headless_layout(&document, &canvas, 180.0);
    assert!(first_line.galley.rows.len() > 1);
    assert!(
        first_line.galley.rows[0].rect_without_leading_space().min.x
            > first_line.galley.rows[1].rect_without_leading_space().min.x
    );

    document.paragraph_styles[0].first_line_indent_points = -18.0;
    let hanging = run_headless_layout(&document, &canvas, 180.0);
    assert!(
        hanging.galley.rows[0].rect_without_leading_space().min.x
            < hanging.galley.rows[1].rect_without_leading_space().min.x
    );
}

fn make_document(
    runs: Vec<TextRun>,
    paragraph_styles: Vec<ParagraphStyle>,
    paragraph_images: Vec<Option<DocumentImage>>,
) -> DocumentState {
    let paragraph_count = paragraph_styles.len();
    DocumentState {
        title: "Test".to_owned(),
        runs,
        paragraph_styles,
        paragraph_images,
        paragraph_tables: vec![None; paragraph_count],
        page_size: PageSize::a4(),
        margins: PageMargins::standard(),
        header_text: String::new(),
        footer_text: String::new(),
        first_page_header_text: String::new(),
        first_page_footer_text: String::new(),
        even_page_header_text: String::new(),
        even_page_footer_text: String::new(),
        header_runs: vec![TextRun {
            text: String::new(),
            style: CharacterStyle::default(),
        }],
        footer_runs: vec![TextRun {
            text: String::new(),
            style: CharacterStyle::default(),
        }],
        first_page_header_runs: vec![TextRun {
            text: String::new(),
            style: CharacterStyle::default(),
        }],
        first_page_footer_runs: vec![TextRun {
            text: String::new(),
            style: CharacterStyle::default(),
        }],
        even_page_header_runs: vec![TextRun {
            text: String::new(),
            style: CharacterStyle::default(),
        }],
        even_page_footer_runs: vec![TextRun {
            text: String::new(),
            style: CharacterStyle::default(),
        }],
        different_first_page: false,
        different_odd_even_pages: false,
        page_number_start: 1,
        sections: vec![crate::document::Section::first(
            crate::document::PageSetup::standard(),
        )],
        source_docx: None,
    }
}

fn make_test_image(id: usize, width: f32, height: f32, wrap_mode: WrapMode) -> DocumentImage {
    DocumentImage {
        id,
        bytes: vec![].into(),
        alt_text: "test".to_owned(),
        width_points: width,
        height_points: height,
        lock_aspect_ratio: true,
        opacity: 1.0,
        layout_mode: ImageLayoutMode::Inline,
        wrap_mode,
        rendering: ImageRendering::Smooth,
        horizontal_position: Default::default(),
        vertical_position: Default::default(),
        distance_from_text: Default::default(),
        z_index: 0,
        move_with_text: true,
        allow_overlap: false,
    }
}

#[test]
fn single_paragraph_produces_at_least_one_row() {
    let document = make_document(
        vec![TextRun {
            text: "Hello world".to_owned(),
            style: CharacterStyle::default(),
        }],
        vec![ParagraphStyle::default()],
        vec![None],
    );
    let canvas = CanvasState::default();
    let layout = run_headless_layout(&document, &canvas, 600.0);

    assert!(
        !layout.galley.rows.is_empty(),
        "galley should have at least one row"
    );
    assert!(
        layout.manual_page_break_rows.is_empty(),
        "no manual page breaks expected"
    );
    assert!(layout.images.is_empty(), "no images expected");
}

#[test]
fn long_text_wraps_into_multiple_rows() {
    let long_text = "Lorem ipsum dolor sit amet, consectetur adipiscing elit. ".repeat(20);
    let document = make_document(
        vec![TextRun {
            text: long_text,
            style: CharacterStyle::default(),
        }],
        vec![ParagraphStyle::default()],
        vec![None],
    );
    let canvas = CanvasState::default();
    let layout = run_headless_layout(&document, &canvas, 300.0);

    assert!(
        layout.galley.rows.len() > 1,
        "long text at narrow width should produce multiple rows, got {}",
        layout.galley.rows.len()
    );
}

#[test]
fn manual_page_break_is_recorded() {
    let document = make_document(
        vec![TextRun {
            text: "First paragraph\nSecond paragraph".to_owned(),
            style: CharacterStyle::default(),
        }],
        vec![
            ParagraphStyle::default(),
            ParagraphStyle {
                page_break_before: true,
                ..ParagraphStyle::default()
            },
        ],
        vec![None, None],
    );
    let canvas = CanvasState::default();
    let layout = run_headless_layout(&document, &canvas, 600.0);

    assert!(
        !layout.manual_page_break_rows.is_empty(),
        "should record a manual page break"
    );
    // The page break should be at the row index where the second paragraph starts.
    assert!(
        layout.manual_page_break_rows[0] > 0,
        "page break row index should be > 0"
    );
}

#[test]
fn block_image_paragraph_produces_image_layout() {
    let document = make_document(
        vec![TextRun {
            text: format!("{OBJECT_REPLACEMENT_CHAR}"),
            style: CharacterStyle::default(),
        }],
        vec![ParagraphStyle::default()],
        vec![Some(make_test_image(1, 200.0, 100.0, WrapMode::Inline))],
    );
    let canvas = CanvasState::default();
    let layout = run_headless_layout(&document, &canvas, 600.0);

    assert_eq!(layout.images.len(), 1, "should have one image layout");
    assert_eq!(layout.images[0].row_index, 0);
    assert!(
        layout.images[0].size.x > 0.0 && layout.images[0].size.y > 0.0,
        "image size should be positive"
    );
}

#[test]
fn wide_image_display_size_can_exceed_wrap_width() {
    let document = make_document(
        vec![TextRun {
            text: format!("{OBJECT_REPLACEMENT_CHAR}"),
            style: CharacterStyle::default(),
        }],
        vec![ParagraphStyle::default()],
        vec![Some(make_test_image(31, 900.0, 300.0, WrapMode::Inline))],
    );
    let canvas = CanvasState::default();
    let layout = run_headless_layout(&document, &canvas, 600.0);

    assert_eq!(layout.images.len(), 1);
    assert!(
        layout.images[0].size.x > 600.0,
        "wide images should keep stored size instead of being clamped to wrap width"
    );
    assert!((layout.images[0].size.x - 900.0).abs() < 1.0);
    assert!((layout.images[0].size.y - 300.0).abs() < 1.0);
}

#[test]
fn resize_geometry_non_locked_handles_keep_opposite_edges_stable() {
    let cases = [
        (ResizeHandle::NW, (90.0, 65.0, 10.0, 15.0)),
        (ResizeHandle::N, (100.0, 65.0, 0.0, 15.0)),
        (ResizeHandle::NE, (110.0, 65.0, 0.0, 15.0)),
        (ResizeHandle::E, (110.0, 80.0, 0.0, 0.0)),
        (ResizeHandle::SE, (110.0, 95.0, 0.0, 0.0)),
        (ResizeHandle::S, (100.0, 95.0, 0.0, 0.0)),
        (ResizeHandle::SW, (90.0, 95.0, 10.0, 0.0)),
        (ResizeHandle::W, (90.0, 80.0, 10.0, 0.0)),
    ];

    for (handle, expected) in cases {
        let actual =
            crate::canvas::image::resized_image_geometry(handle, 100.0, 80.0, 10.0, 15.0, false);
        assert_geometry_close(actual, expected, handle);
    }
}

#[test]
fn resize_geometry_clamps_to_minimum_size() {
    let west = crate::canvas::image::resized_image_geometry(
        ResizeHandle::W,
        100.0,
        80.0,
        200.0,
        0.0,
        false,
    );
    assert_geometry_close(west, (1.0, 80.0, 99.0, 0.0), ResizeHandle::W);

    let north = crate::canvas::image::resized_image_geometry(
        ResizeHandle::N,
        100.0,
        80.0,
        0.0,
        200.0,
        false,
    );
    assert_geometry_close(north, (100.0, 1.0, 0.0, 79.0), ResizeHandle::N);
}

#[test]
fn resize_geometry_locked_ratio_keeps_anchors_stable() {
    let nw = crate::canvas::image::resized_image_geometry(
        ResizeHandle::NW,
        100.0,
        50.0,
        20.0,
        0.0,
        true,
    );
    assert_geometry_close(nw, (80.0, 40.0, 20.0, 10.0), ResizeHandle::NW);

    let east =
        crate::canvas::image::resized_image_geometry(ResizeHandle::E, 100.0, 50.0, 20.0, 0.0, true);
    assert_geometry_close(east, (120.0, 60.0, 0.0, -5.0), ResizeHandle::E);

    let south =
        crate::canvas::image::resized_image_geometry(ResizeHandle::S, 100.0, 50.0, 0.0, 20.0, true);
    assert_geometry_close(south, (140.0, 70.0, -20.0, 0.0), ResizeHandle::S);
}

#[test]
fn image_drag_preview_keeps_image_and_handles_aligned() {
    let original = egui::Rect::from_min_size(egui::pos2(20.0, 30.0), egui::vec2(100.0, 60.0));
    let preview = crate::canvas::image::image_drag_preview_rect(
        7,
        original,
        Some((7, original, egui::vec2(35.0, -10.0))),
    );

    assert_eq!(preview.min, egui::pos2(55.0, 20.0));
    assert_eq!(preview.size(), original.size());
    let handles = crate::canvas::image::resize_handle_rects(preview);
    assert_eq!(handles[0].1.center(), preview.left_top());
    assert_eq!(handles[4].1.center(), preview.right_bottom());
}

#[test]
fn thin_imported_image_keeps_its_size_and_can_shrink() {
    let image = make_test_image(32, 453.6, 20.35, WrapMode::Inline);
    let size = crate::canvas::image::image_display_size(&image, 600.0, 1.0);
    assert!((size.x - 453.6).abs() < 0.01);
    assert!((size.y - 20.35).abs() < 0.01);

    let resized = crate::canvas::image::resized_image_geometry(
        ResizeHandle::E,
        453.6,
        20.35,
        -100.0,
        0.0,
        true,
    );
    assert!((resized.0 - 353.6).abs() < 0.01);
    assert!(
        resized.1 < 20.35,
        "locked resize should not be blocked at 24pt"
    );
}

#[test]
fn unselected_image_does_not_expose_resize_handles() {
    let rect = egui::Rect::from_min_size(egui::pos2(20.0, 30.0), egui::vec2(100.0, 20.0));
    let mut canvas = CanvasState::default();
    canvas.image_rects.push((7, rect));

    assert_eq!(
        crate::canvas::image::image_handle_hit(&canvas, rect.left_top(), 6.0),
        None
    );
    canvas.selected_image_id = Some(7);
    assert!(crate::canvas::image::image_handle_hit(&canvas, rect.left_top(), 6.0).is_some());
}

fn assert_geometry_close(
    actual: (f32, f32, f32, f32),
    expected: (f32, f32, f32, f32),
    handle: ResizeHandle,
) {
    assert!(
        (actual.0 - expected.0).abs() < 0.01
            && (actual.1 - expected.1).abs() < 0.01
            && (actual.2 - expected.2).abs() < 0.01
            && (actual.3 - expected.3).abs() < 0.01,
        "{handle:?}: expected {expected:?}, got {actual:?}"
    );
}

#[test]
fn centered_alignment_offsets_row_positions() {
    let document = make_document(
        vec![TextRun {
            text: "Short".to_owned(),
            style: CharacterStyle::default(),
        }],
        vec![ParagraphStyle {
            alignment: ParagraphAlignment::Center,
            ..ParagraphStyle::default()
        }],
        vec![None],
    );
    let canvas = CanvasState::default();
    let layout = run_headless_layout(&document, &canvas, 600.0);

    assert!(!layout.galley.rows.is_empty());
    let row = &layout.galley.rows[0];
    // Centered text on a 600px wrap should have a positive x offset.
    assert!(
        row.pos.x > 0.0,
        "centered text should be offset from left edge, got x={}",
        row.pos.x
    );
}

#[test]
fn multiple_paragraphs_with_varying_styles() {
    let document = make_document(
        vec![TextRun {
            text: "Left aligned\nCenter aligned\nRight aligned".to_owned(),
            style: CharacterStyle::default(),
        }],
        vec![
            ParagraphStyle {
                alignment: ParagraphAlignment::Left,
                ..ParagraphStyle::default()
            },
            ParagraphStyle {
                alignment: ParagraphAlignment::Center,
                ..ParagraphStyle::default()
            },
            ParagraphStyle {
                alignment: ParagraphAlignment::Right,
                ..ParagraphStyle::default()
            },
        ],
        vec![None, None, None],
    );
    let canvas = CanvasState::default();
    let layout = run_headless_layout(&document, &canvas, 600.0);

    // Should have at least 3 rows (one per paragraph).
    assert!(
        layout.galley.rows.len() >= 3,
        "expected at least 3 rows for 3 paragraphs, got {}",
        layout.galley.rows.len()
    );
}

#[test]
fn paragraph_spacing_offsets_following_rows() {
    let document = make_document(
        vec![TextRun {
            text: "First\nSecond".to_owned(),
            style: CharacterStyle::default(),
        }],
        vec![
            ParagraphStyle {
                spacing_after_points: 24,
                ..ParagraphStyle::default()
            },
            ParagraphStyle::default(),
        ],
        vec![None, None],
    );
    let canvas = CanvasState::default();
    let layout = run_headless_layout(&document, &canvas, 600.0);

    assert!(layout.galley.rows.len() >= 2);
    let first_row = &layout.galley.rows[0];
    let second_row = &layout.galley.rows[1];
    assert!(
        second_row.pos.y - first_row.pos.y > first_row.rect().height(),
        "paragraph spacing should increase the gap between rows"
    );
}

#[test]
fn auto_multiplier_line_spacing_offsets_following_line() {
    let document = make_document(
        vec![TextRun {
            text: "First line in paragraph that wraps on purpose because the width is narrow"
                .to_owned(),
            style: CharacterStyle::default(),
        }],
        vec![ParagraphStyle {
            line_spacing: LineSpacing {
                kind: LineSpacingKind::AutoMultiplier,
                value: 1.5,
            },
            ..ParagraphStyle::default()
        }],
        vec![None],
    );
    let canvas = CanvasState::default();
    let layout = run_headless_layout(&document, &canvas, 240.0);

    assert!(layout.galley.rows.len() >= 2);
    let first_row = &layout.galley.rows[0];
    let second_row = &layout.galley.rows[1];
    let default_gap = first_row.row.height();
    let actual_gap = second_row.pos.y - first_row.pos.y;
    assert!(
        actual_gap > default_gap * 1.45,
        "1.5x line spacing should enlarge row advance, got actual_gap={actual_gap}, default_gap={default_gap}"
    );
}

#[test]
fn exact_line_spacing_uses_requested_row_advance() {
    let document = make_document(
        vec![TextRun {
            text: "First line in paragraph that wraps on purpose because the width is narrow"
                .to_owned(),
            style: CharacterStyle::default(),
        }],
        vec![ParagraphStyle {
            line_spacing: LineSpacing {
                kind: LineSpacingKind::ExactPoints,
                value: 24.0,
            },
            ..ParagraphStyle::default()
        }],
        vec![None],
    );
    let canvas = CanvasState::default();
    let layout = run_headless_layout(&document, &canvas, 240.0);

    assert!(layout.galley.rows.len() >= 2);
    let first_row = &layout.galley.rows[0];
    let second_row = &layout.galley.rows[1];
    let actual_gap = second_row.pos.y - first_row.pos.y;
    assert!(
        (actual_gap - 24.0).abs() < 1.5,
        "exact line spacing should follow the requested advance, got {actual_gap}"
    );
}

#[test]
fn fit_page_zoom_uses_manual_override_rules() {
    let viewport = egui::Rect::from_min_size(egui::Pos2::ZERO, egui::vec2(400.0, 500.0));
    let fit = fit_page_zoom(viewport, PageSize::a4());
    assert!(
        fit < 1.0,
        "fit zoom should shrink an A4 page in a small viewport"
    );

    let mut canvas = CanvasState::default();
    canvas.imported_docx_view = true;
    canvas.zoom_mode = ZoomMode::FitPage;
    canvas.zoom = fit;
    canvas.zoom_mode = ZoomMode::Manual;
    canvas.zoom = (canvas.zoom * 1.1).clamp(0.5, 3.0);
    assert_eq!(canvas.zoom_mode, ZoomMode::Manual);
    assert!(canvas.zoom > fit);
}

#[test]
fn view_scaling_accumulates_small_deltas_before_reflowing() {
    let mut canvas = CanvasState::default();

    canvas.scale_view(1.002);
    canvas.scale_view(1.002);
    assert_eq!(canvas.zoom, 1.0);

    canvas.scale_view(1.002);
    assert_eq!(canvas.zoom, 1.01);
}

#[test]
fn zoom_gesture_defers_exact_layout_until_it_settles() {
    let ctx = egui::Context::default();
    let mut document = make_document(
        vec![TextRun {
            text: "large document preview".to_owned(),
            style: CharacterStyle::default(),
        }],
        vec![ParagraphStyle::default()],
        vec![None],
    );
    let mut canvas = CanvasState::default();
    canvas.zoom = 1.5;
    canvas.last_zoom_input_time = 10.0;
    let mut history = ChangeHistory::default();

    let mut input = egui::RawInput::default();
    input.time = Some(10.0);
    let _ = ctx.run_ui(input, |ui| {
        paint_document_canvas(
            ui,
            &mut document,
            &mut canvas,
            ThemeMode::Light,
            &mut history,
            &[],
        );
    });
    assert_eq!(canvas.layout_zoom, 1.0);

    let mut input = egui::RawInput::default();
    input.time = Some(10.2);
    let _ = ctx.run_ui(input, |ui| {
        paint_document_canvas(
            ui,
            &mut document,
            &mut canvas,
            ThemeMode::Light,
            &mut history,
            &[],
        );
    });
    assert_eq!(canvas.layout_zoom, 1.5);
}

#[test]
fn image_with_page_break_in_multi_paragraph_document() {
    let document = make_document(
        vec![TextRun {
            text: format!("Intro paragraph\n{OBJECT_REPLACEMENT_CHAR}\nClosing paragraph"),
            style: CharacterStyle::default(),
        }],
        vec![
            ParagraphStyle::default(),
            ParagraphStyle {
                page_break_before: true,
                ..ParagraphStyle::default()
            },
            ParagraphStyle::default(),
        ],
        vec![
            None,
            Some(make_test_image(2, 400.0, 300.0, WrapMode::Inline)),
            None,
        ],
    );
    let canvas = CanvasState::default();
    let layout = run_headless_layout(&document, &canvas, 600.0);

    assert_eq!(layout.images.len(), 1, "should have one image");
    assert!(
        !layout.manual_page_break_rows.is_empty(),
        "should have a manual page break"
    );
    // Image row should match the page break row.
    assert_eq!(
        layout.images[0].row_index, layout.manual_page_break_rows[0],
        "image should be on the page-break row"
    );
}

#[test]
fn list_markers_are_produced_for_bullet_paragraphs() {
    let document = make_document(
        vec![TextRun {
            text: "Item one\nItem two".to_owned(),
            style: CharacterStyle::default(),
        }],
        vec![
            ParagraphStyle {
                list_kind: ListKind::Bullet,
                ..ParagraphStyle::default()
            },
            ParagraphStyle {
                list_kind: ListKind::Bullet,
                ..ParagraphStyle::default()
            },
        ],
        vec![None, None],
    );
    let canvas = CanvasState::default();
    let layout = run_headless_layout(&document, &canvas, 600.0);

    assert_eq!(layout.list_markers.len(), 2, "should have two list markers");
    assert_eq!(layout.list_markers[0].text, "•");
    assert_eq!(layout.list_markers[1].text, "•");
}

#[test]
fn wrap_modes_change_image_row_geometry() {
    let make_doc = |wrap_mode| {
        make_document(
            vec![TextRun {
                text: format!("{OBJECT_REPLACEMENT_CHAR}"),
                style: CharacterStyle::default(),
            }],
            vec![ParagraphStyle::default()],
            vec![Some(make_test_image(10, 180.0, 90.0, wrap_mode))],
        )
    };

    let canvas = CanvasState::default();
    let wrap_width = 600.0;

    let inline = run_headless_layout(&make_doc(WrapMode::Inline), &canvas, wrap_width);
    let square = run_headless_layout(&make_doc(WrapMode::Square), &canvas, wrap_width);
    let tight = run_headless_layout(&make_doc(WrapMode::Tight), &canvas, wrap_width);
    let through = run_headless_layout(&make_doc(WrapMode::Through), &canvas, wrap_width);
    let top_bottom = run_headless_layout(&make_doc(WrapMode::TopAndBottom), &canvas, wrap_width);

    let inline_row = &inline.galley.rows[inline.images[0].row_index];
    let square_row = &square.galley.rows[square.images[0].row_index];
    let tight_row = &tight.galley.rows[tight.images[0].row_index];
    let through_row = &through.galley.rows[through.images[0].row_index];
    let top_bottom_row = &top_bottom.galley.rows[top_bottom.images[0].row_index];

    assert!(
        square_row.size.x > tight_row.size.x,
        "square wrap should reserve more horizontal space than tight"
    );
    assert!(
        square_row.size.y > tight_row.size.y,
        "square wrap should reserve more vertical space than tight"
    );
    assert!(
        through_row.size.x <= 1.0 && through_row.size.y <= 1.0,
        "through wrap should not reserve layout space"
    );
    assert!(
        (top_bottom_row.size.x - wrap_width).abs() < 1.0,
        "top-and-bottom wrap should reserve the full row width"
    );
    assert!(
        top_bottom.images[0].offset.x > through.images[0].offset.x,
        "top-and-bottom should center image while through should anchor without horizontal reservation"
    );
    assert!(
        inline_row.size.x < top_bottom_row.size.x,
        "inline wrap should keep a tighter row than top-and-bottom"
    );
}

#[test]
fn tight_wrap_chooses_side_from_image_position() {
    let make_doc = |offset_x: f32| {
        let mut img = make_test_image(21, 180.0, 90.0, WrapMode::Tight);
        img.horizontal_position.offset_points = offset_x;
        make_document(
            vec![TextRun {
                text: format!(
                    "{OBJECT_REPLACEMENT_CHAR}\nthis paragraph should flow beside the image and expose side choice"
                ),
                style: CharacterStyle::default(),
            }],
            vec![ParagraphStyle::default(), ParagraphStyle::default()],
            vec![
                Some(img),
                None,
            ],
        )
    };

    let canvas = CanvasState::default();
    let wrap_width = 600.0;
    let left_placed = run_headless_layout(&make_doc(0.0), &canvas, wrap_width);
    let right_placed = run_headless_layout(&make_doc(220.0), &canvas, wrap_width);

    let left_row_x = left_placed.galley.rows[1].pos.x;
    let right_row_x = right_placed.galley.rows[1].pos.x;

    assert!(
        left_row_x > right_row_x,
        "left-placed image should push text to the right; got left_row_x={left_row_x}, right_row_x={right_row_x}"
    );
}

#[test]
fn side_wrap_modes_place_short_rows_beside_image() {
    for wrap_mode in [WrapMode::Square, WrapMode::Tight, WrapMode::Through] {
        let document = make_document(
            vec![TextRun {
                text: format!("{OBJECT_REPLACEMENT_CHAR}\nshort side text"),
                style: CharacterStyle::default(),
            }],
            vec![ParagraphStyle::default(), ParagraphStyle::default()],
            vec![Some(make_test_image(41, 180.0, 90.0, wrap_mode)), None],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 600.0);
        let image = &layout.images[0];
        let image_row = &layout.galley.rows[image.row_index];
        let text_row = &layout.galley.rows[image.row_index + 1];
        let image_right = image_row.pos.x + image.offset.x + image.size.x;

        assert!(
            text_row.pos.x > image_right,
            "{wrap_mode:?} should place short text beside the image: text_x={}, image_right={}, text_width={}, row_size={}",
            text_row.pos.x,
            image_right,
            text_row.size.x,
            text_row.row.size.x,
        );
    }
}

#[test]
fn tight_wrap_reflows_wrappable_paragraph_into_side_column() {
    let document = make_document(
        vec![TextRun {
            text: format!(
                "{OBJECT_REPLACEMENT_CHAR}\n{}",
                "bon above to change bold italic underline strike through text size font family text color and highlight"
            ),
            style: CharacterStyle::default(),
        }],
        vec![ParagraphStyle::default(), ParagraphStyle::default()],
        vec![Some(make_test_image(43, 175.0, 130.0, WrapMode::Tight)), None],
    );
    let canvas = CanvasState::default();
    let layout = run_headless_layout(&document, &canvas, 451.0);
    let image = &layout.images[0];
    let image_row = &layout.galley.rows[image.row_index];
    let image_right = image_row.pos.x + image.offset.x + image.size.x;
    let side_rows: Vec<_> = layout
        .galley
        .rows
        .iter()
        .skip(image.row_index + 1)
        .take_while(|row| row.min_y() < image_row.pos.y + image.offset.y + image.size.y)
        .collect();

    assert!(
        side_rows.len() >= 2,
        "tight wrap should reflow the paragraph into multiple side-column rows"
    );
    assert!(
        side_rows.iter().all(|row| row.pos.x > image_right),
        "all rows beside the image should start to the right of the image"
    );
}

#[test]
fn side_wrap_modes_push_wide_rows_below_image() {
    let wide_word = "x".repeat(140);
    for wrap_mode in [WrapMode::Square, WrapMode::Tight, WrapMode::Through] {
        let document = make_document(
            vec![TextRun {
                text: format!("{OBJECT_REPLACEMENT_CHAR}\n{wide_word}"),
                style: CharacterStyle::default(),
            }],
            vec![ParagraphStyle::default(), ParagraphStyle::default()],
            vec![Some(make_test_image(42, 540.0, 90.0, wrap_mode)), None],
        );
        let canvas = CanvasState::default();
        let layout = run_headless_layout(&document, &canvas, 600.0);
        let image = &layout.images[0];
        let image_row = &layout.galley.rows[image.row_index];
        let text_row = &layout.galley.rows[image.row_index + 1];
        let image_bottom = image_row.pos.y + image.offset.y + image.size.y;

        assert!(
            text_row.pos.y >= image_bottom,
            "{wrap_mode:?} should push text below the image when no side fits: text_y={}, image_bottom={}, text_x={}, glyph_width={}, row_width={}",
            text_row.pos.y,
            image_bottom,
            text_row.pos.x,
            text_row.row.size.x,
            text_row.size.x
        );
    }
}

#[test]
fn tight_wrap_vertical_offset_delays_text_wrapping() {
    let make_doc = |offset_y: f32| {
        let mut img = make_test_image(22, 180.0, 90.0, WrapMode::Tight);
        img.vertical_position.offset_points = offset_y;
        make_document(
            vec![TextRun {
                text: format!(
                    "{OBJECT_REPLACEMENT_CHAR}\nthis paragraph should start unwrapped when image is moved down enough"
                ),
                style: CharacterStyle::default(),
            }],
            vec![ParagraphStyle::default(), ParagraphStyle::default()],
            vec![
                Some(img),
                None,
            ],
        )
    };

    let canvas = CanvasState::default();
    let wrap_width = 600.0;
    let normal = run_headless_layout(&make_doc(0.0), &canvas, wrap_width);
    let moved_down = run_headless_layout(&make_doc(220.0), &canvas, wrap_width);

    let normal_row_x = normal.galley.rows[1].pos.x;
    let moved_down_row_x = moved_down.galley.rows[1].pos.x;

    assert!(
        normal_row_x > moved_down_row_x + 40.0,
        "moving image down should reduce early side-wrapping; got normal_row_x={normal_row_x}, moved_down_row_x={moved_down_row_x}"
    );
}
