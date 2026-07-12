use egui::{pos2, vec2, Rect, Vec2};

use crate::document::{PageMargins, PageSetup, PageSize};

pub fn viewport_scale(pixels_per_point: f32, zoom: f32) -> f32 {
    pixels_per_point * zoom
}

pub fn document_points_to_screen_points(document_points: f32, zoom: f32) -> f32 {
    document_points * zoom
}

pub fn document_points_to_pixels(document_points: f32, pixels_per_point: f32, zoom: f32) -> f32 {
    document_points * viewport_scale(pixels_per_point, zoom)
}

pub fn quantize_zoom(zoom: f32) -> f32 {
    (zoom * 100.0).round() / 100.0
}

pub fn fit_page_zoom(viewport: Rect, page_size: PageSize) -> f32 {
    let fit_width = viewport.width() / page_size.width_points.max(1.0);
    let fit_height = viewport.height() / page_size.height_points.max(1.0);
    quantize_zoom((fit_width.min(fit_height) * 0.92).clamp(0.25, 3.0))
}

pub fn page_size_in_screen_points(page_size: PageSize, zoom: f32) -> Vec2 {
    vec2(
        document_points_to_screen_points(page_size.width_points, zoom),
        document_points_to_screen_points(page_size.height_points, zoom),
    )
}

pub fn centered_page_rect(viewport: Rect, page_size: PageSize, zoom: f32, pan: Vec2) -> Rect {
    let page_size = page_size_in_screen_points(page_size, zoom);
    let origin = pos2(
        viewport.center().x - page_size.x * 0.5 + pan.x,
        viewport.center().y - page_size.y * 0.5 + pan.y,
    );

    Rect::from_min_size(origin, page_size)
}

pub fn page_content_rect(page_rect: Rect, margins: PageMargins, zoom: f32) -> Rect {
    let left = document_points_to_screen_points(margins.left_points, zoom);
    let right = document_points_to_screen_points(margins.right_points, zoom);
    let top = document_points_to_screen_points(margins.top_points, zoom);
    let bottom = document_points_to_screen_points(margins.bottom_points, zoom);

    Rect::from_min_max(
        pos2(page_rect.left() + left, page_rect.top() + top),
        pos2(page_rect.right() - right, page_rect.bottom() - bottom),
    )
}

pub fn section_page_content_rect(
    page_rect: Rect,
    setup: PageSetup,
    header_story_height_points: f32,
    footer_story_height_points: f32,
    zoom: f32,
) -> Rect {
    let left = document_points_to_screen_points(setup.margins.left_points, zoom);
    let right = document_points_to_screen_points(setup.margins.right_points, zoom);
    let header_occupied = setup
        .margins
        .top_points
        .max(setup.header_from_top_points + header_story_height_points);
    let footer_occupied = setup
        .margins
        .bottom_points
        .max(setup.footer_from_bottom_points + footer_story_height_points);
    let top = document_points_to_screen_points(header_occupied, zoom);
    let bottom = document_points_to_screen_points(footer_occupied, zoom);
    let min_height = document_points_to_screen_points(12.0, zoom).max(1.0);
    let max_y = (page_rect.bottom() - bottom).max(page_rect.top() + top + min_height);

    Rect::from_min_max(
        pos2(page_rect.left() + left, page_rect.top() + top),
        pos2(page_rect.right() - right, max_y),
    )
}

pub fn screen_points_to_document_points(screen_points: f32, zoom: f32) -> f32 {
    screen_points / zoom
}
