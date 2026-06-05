use std::sync::Arc;
use eframe::egui::{self, Rect};

use crate::document::{DocumentImage, WrapMode, OBJECT_REPLACEMENT_CHAR};
use crate::layout::document_points_to_screen_points;

pub struct ActiveSideWrapFlow {
    pub pending_top_height: f32,
    pub remaining_height: f32,
    pub text_start_x: f32,
    pub text_width: f32,
}

pub struct TightWrapZone {
    pub row_index: usize,
    pub top: f32,
    pub bottom: f32,
    pub text_start_x: f32,
    pub text_width: f32,
}

#[derive(Clone)]
pub struct ImageLayout {
    pub row_index: usize,
    pub size: egui::Vec2,
    pub offset: egui::Vec2,
    pub image: DocumentImage,
}

pub struct BlockImageReservation {
    pub row_size: egui::Vec2,
    pub image_offset: egui::Vec2,
}

pub fn block_image_reservation(
    wrap_mode: WrapMode,
    image_size: egui::Vec2,
    wrap_width: f32,
    zoom: f32,
) -> BlockImageReservation {
    let square_pad = square_wrap_pad(zoom);
    let tight_pad = tight_wrap_pad(zoom);

    match wrap_mode {
        WrapMode::Inline => BlockImageReservation {
            row_size: image_size,
            image_offset: egui::Vec2::ZERO,
        },
        WrapMode::Square => {
            let row_width = (image_size.x + square_pad * 2.0).min(wrap_width);
            let row_height = (square_pad * 2.0).max(tight_wrap_row_height(zoom));
            BlockImageReservation {
                row_size: egui::vec2(row_width, row_height),
                image_offset: egui::vec2(square_pad, square_pad),
            }
        }
        WrapMode::Tight => {
            let row_width = (image_size.x + tight_pad * 2.0).min(wrap_width);
            let row_height = tight_wrap_row_height(zoom);
            BlockImageReservation {
                row_size: egui::vec2(row_width, row_height),
                image_offset: egui::vec2(tight_pad, 0.0),
            }
        }
        WrapMode::Through | WrapMode::BehindText | WrapMode::InFrontOfText => {
            BlockImageReservation {
                row_size: egui::Vec2::ZERO,
                image_offset: egui::Vec2::ZERO,
            }
        }
        WrapMode::TopAndBottom => BlockImageReservation {
            row_size: egui::vec2(wrap_width, image_size.y),
            image_offset: egui::vec2(((wrap_width - image_size.x) * 0.5).max(0.0), 0.0),
        },
    }
}

pub fn tight_wrap_pad(zoom: f32) -> f32 {
    document_points_to_screen_points(4.0, zoom)
}

pub fn square_wrap_pad(zoom: f32) -> f32 {
    document_points_to_screen_points(12.0, zoom)
}

pub fn side_wrap_pad(wrap_mode: WrapMode, zoom: f32) -> f32 {
    match wrap_mode {
        WrapMode::Square => square_wrap_pad(zoom),
        WrapMode::Tight | WrapMode::Through => tight_wrap_pad(zoom),
        _ => 0.0,
    }
}

pub fn wrap_uses_side_flow(wrap_mode: WrapMode) -> bool {
    matches!(
        wrap_mode,
        WrapMode::Square | WrapMode::Tight | WrapMode::Through
    )
}

pub fn tight_wrap_row_height(zoom: f32) -> f32 {
    document_points_to_screen_points(14.0, zoom).max(1.0)
}

pub fn apply_tight_wrap_row_offsets(
    galley: &mut Arc<egui::Galley>,
    images: &[ImageLayout],
    zoom: f32,
    wrap_width: f32,
) {
    let zones = tight_wrap_zones(galley, images, zoom, wrap_width);
    if zones.is_empty() {
        return;
    }

    let galley = Arc::make_mut(galley);
    let mut min_rect = galley.rect.min;
    let mut max_rect = galley.rect.max;
    let mut mesh_bounds = egui::Rect::NOTHING;
    let mut cumulative_shift = 0.0;

    for (index, row) in galley.rows.iter_mut().enumerate() {
        let row_text = row.row.text();
        if !row_text.chars().all(|ch| ch == OBJECT_REPLACEMENT_CHAR) {
            row.pos.y += cumulative_shift;

            loop {
                let Some(zone) = zones.iter().find(|zone| {
                    zone.row_index != index && row.max_y() > zone.top && row.min_y() < zone.bottom
                }) else {
                    break;
                };

                // Check if the row's actual glyph content fits in the side column.
                // Use the intrinsic row width (row.row.size.x) to decide if the row
                // can fit beside the image, and reposition it into the side column.
                let row_content_width = row.row.size.x;
                let min_side_width = document_points_to_screen_points(36.0, zoom);
                if row_content_width <= zone.text_width && zone.text_width >= min_side_width {
                    row.pos.x = zone.text_start_x;
                    break;
                }

                let offset_y = zone.bottom - row.pos.y;
                if offset_y <= 0.0 {
                    break;
                }
                row.pos.y += offset_y;
                cumulative_shift += offset_y;
            }
        } else if cumulative_shift > 0.0 {
            row.pos.y += cumulative_shift;
        }

        let row_rect = row.rect();
        min_rect.x = min_rect.x.min(row_rect.min.x);
        min_rect.y = min_rect.y.min(row_rect.min.y);
        max_rect.x = max_rect.x.max(row_rect.max.x);
        max_rect.y = max_rect.y.max(row_rect.max.y);
        mesh_bounds |= row.visuals.mesh_bounds.translate(row.pos.to_vec2());
    }

    galley.rect = Rect::from_min_max(min_rect, max_rect);
    galley.mesh_bounds = mesh_bounds;
}

pub fn tight_wrap_zones(
    galley: &egui::Galley,
    images: &[ImageLayout],
    zoom: f32,
    wrap_width: f32,
) -> Vec<TightWrapZone> {
    let mut zones = Vec::new();

    for image in images
        .iter()
        .filter(|image| wrap_uses_side_flow(image.image.wrap_mode))
    {
        let Some(row) = galley.rows.get(image.row_index) else {
            continue;
        };

        let image_left = row.pos.x
            + image.offset.x
            + document_points_to_screen_points(image.image.offset_x_points(), zoom);
        let image_top = row.pos.y
            + image.offset.y
            + document_points_to_screen_points(image.image.offset_y_points(), zoom);
        let pad = side_wrap_pad(image.image.wrap_mode, zoom);

        let image_right = image_left + image.size.x;
        let image_bottom = image_top + image.size.y;
        let left_width = (image_left - pad).max(0.0);
        let right_start = (image_right + pad).max(0.0);
        let right_width = (wrap_width - right_start).max(0.0);
        let (text_start_x, text_width) = if right_width >= left_width {
            (right_start, right_width)
        } else {
            (0.0, left_width)
        };

        zones.push(TightWrapZone {
            row_index: image.row_index,
            top: image_top - pad,
            bottom: image_bottom + pad,
            text_start_x,
            text_width,
        });
    }

    zones
}

pub fn reserve_block_image_space(galley: &mut Arc<egui::Galley>, row_size: egui::Vec2) -> bool {
    let galley = Arc::make_mut(galley);
    let Some(placed_row) = galley.rows.first_mut() else {
        return false;
    };
    let row = Arc::make_mut(&mut placed_row.row);
    if row.glyphs.len() != 1 || row.glyphs[0].chr != OBJECT_REPLACEMENT_CHAR {
        return false;
    }

    row.glyphs[0].advance_width = row_size.x;
    row.glyphs[0].line_height = row_size.y;
    row.glyphs[0].font_ascent = row_size.y;
    row.glyphs[0].font_height = row_size.y;
    row.glyphs[0].font_face_ascent = row_size.y;
    row.glyphs[0].font_face_height = row_size.y;
    row.size = row_size;
    row.visuals = Default::default();

    galley.rect = Rect::from_min_size(galley.rect.min, row_size);
    galley.mesh_bounds = Rect::NOTHING;
    true
}
