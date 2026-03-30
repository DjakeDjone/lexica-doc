use egui::{pos2, vec2, Rect, Vec2};

use crate::document::{PageMargins, PageSize};

pub fn viewport_scale(pixels_per_point: f32, zoom: f32) -> f32 {
    pixels_per_point * zoom
}

pub fn document_points_to_screen_points(document_points: f32, zoom: f32) -> f32 {
    document_points * zoom
}

pub fn document_points_to_pixels(document_points: f32, pixels_per_point: f32, zoom: f32) -> f32 {
    document_points * viewport_scale(pixels_per_point, zoom)
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
