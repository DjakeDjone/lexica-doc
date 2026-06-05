use std::collections::HashMap;
use std::path::Path;
use eframe::egui::Color32;
use quick_xml::events::BytesStart;

use crate::document::{
    CharacterStyle, DistanceFromText, DocumentImage, LineSpacing, LineSpacingKind, PageMargins,
    PageSize, ParagraphAlignment, ParagraphStyle, TextRun,
};

pub(crate) fn append_plain(runs: &mut Vec<TextRun>, text: &str, style: CharacterStyle) {
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

pub(crate) fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

pub(crate) fn attr_value(event: &BytesStart<'_>, key: &[u8]) -> Option<String> {
    event
        .attributes()
        .flatten()
        .find(|attr| local_name(attr.key.as_ref()) == key)
        .and_then(|attr| String::from_utf8(attr.value.into_owned()).ok())
}

pub(crate) fn docx_flag(event: &BytesStart<'_>, default: bool) -> bool {
    match attr_value(event, b"val").as_deref() {
        Some("0" | "false") => false,
        Some("1" | "true") => true,
        Some(_) => default,
        None => default,
    }
}

pub(crate) fn parse_hex_color(value: &str) -> Option<Color32> {
    if value.len() != 6 {
        return None;
    }

    let red = u8::from_str_radix(&value[0..2], 16).ok()?;
    let green = u8::from_str_radix(&value[2..4], 16).ok()?;
    let blue = u8::from_str_radix(&value[4..6], 16).ok()?;
    Some(Color32::from_rgb(red, green, blue))
}

pub(crate) fn highlight_color(value: &str) -> Color32 {
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

pub(crate) fn paragraph_alignment_for(value: &str) -> ParagraphAlignment {
    match value {
        "center" => ParagraphAlignment::Center,
        "right" => ParagraphAlignment::Right,
        "both" | "distribute" => ParagraphAlignment::Justify,
        _ => ParagraphAlignment::Left,
    }
}

pub(crate) fn apply_spacing(event: &BytesStart<'_>, paragraph_style: &mut ParagraphStyle) {
    if let Some(value) = attr_value(event, b"before")
        .and_then(|value| value.parse::<f32>().ok())
        .map(twips_to_points)
    {
        paragraph_style.spacing_before_points = value.round().clamp(0.0, u16::MAX as f32) as u16;
    }
    if let Some(value) = attr_value(event, b"after")
        .and_then(|value| value.parse::<f32>().ok())
        .map(twips_to_points)
    {
        paragraph_style.spacing_after_points = value.round().clamp(0.0, u16::MAX as f32) as u16;
    }
    if let Some(line_spacing) = parse_line_spacing(event) {
        paragraph_style.line_spacing = line_spacing;
    }
}

pub(crate) fn parse_line_spacing(event: &BytesStart<'_>) -> Option<LineSpacing> {
    let line = attr_value(event, b"line")?.parse::<f32>().ok()?;
    let line_rule = attr_value(event, b"lineRule").unwrap_or_else(|| "auto".to_owned());
    Some(match line_rule.as_str() {
        "atLeast" => LineSpacing {
            kind: LineSpacingKind::AtLeastPoints,
            value: twips_to_points(line),
        },
        "exact" => LineSpacing {
            kind: LineSpacingKind::ExactPoints,
            value: twips_to_points(line),
        },
        _ => LineSpacing {
            kind: LineSpacingKind::AutoMultiplier,
            value: line / 240.0,
        },
    })
}

pub(crate) fn parse_page_size(event: &BytesStart<'_>) -> Option<PageSize> {
    let width_twips = attr_value(event, b"w")?.parse::<f32>().ok()?;
    let height_twips = attr_value(event, b"h")?.parse::<f32>().ok()?;
    Some(PageSize {
        width_points: twips_to_points(width_twips),
        height_points: twips_to_points(height_twips),
    })
}

pub(crate) fn parse_page_margins(event: &BytesStart<'_>) -> Option<PageMargins> {
    Some(PageMargins {
        top_points: twips_to_points(attr_value(event, b"top")?.parse::<f32>().ok()?),
        right_points: twips_to_points(attr_value(event, b"right")?.parse::<f32>().ok()?),
        bottom_points: twips_to_points(attr_value(event, b"bottom")?.parse::<f32>().ok()?),
        left_points: twips_to_points(attr_value(event, b"left")?.parse::<f32>().ok()?),
    })
}

pub(crate) fn parse_emu_extent(event: &BytesStart<'_>) -> Option<(f32, f32)> {
    let width = attr_value(event, b"cx")?.parse::<f32>().ok()?;
    let height = attr_value(event, b"cy")?.parse::<f32>().ok()?;
    Some((emu_to_points(width), emu_to_points(height)))
}

pub(crate) fn twips_to_points(value: f32) -> f32 {
    value / 20.0
}

pub(crate) fn emu_to_points(value: f32) -> f32 {
    value / 12_700.0
}

pub(crate) fn parse_anchor_distance(event: &BytesStart<'_>) -> Option<DistanceFromText> {
    let top = attr_value(event, b"distT")
        .and_then(|v| v.parse::<f32>().ok())
        .map(emu_to_points)
        .unwrap_or(0.0);
    let bottom = attr_value(event, b"distB")
        .and_then(|v| v.parse::<f32>().ok())
        .map(emu_to_points)
        .unwrap_or(0.0);
    let left = attr_value(event, b"distL")
        .and_then(|v| v.parse::<f32>().ok())
        .map(emu_to_points)
        .unwrap_or(8.0);
    let right = attr_value(event, b"distR")
        .and_then(|v| v.parse::<f32>().ok())
        .map(emu_to_points)
        .unwrap_or(8.0);
    Some(DistanceFromText {
        top_points: top,
        right_points: right,
        bottom_points: bottom,
        left_points: left,
    })
}

pub(crate) fn normalize_relationship_target(target: &str) -> String {
    if target.starts_with('/') {
        target.trim_start_matches('/').to_owned()
    } else if target.starts_with("word/") {
        target.to_owned()
    } else {
        format!("word/{target}")
    }
}

pub(crate) fn resolve_drawing(
    drawing: Option<super::DrawingState>,
    relationships: &super::DocumentRelationships,
    media: &HashMap<String, Vec<u8>>,
    next_image_id: &mut usize,
) -> Option<DocumentImage> {
    let drawing = drawing?;
    let rel_id = drawing.rel_id?;
    let target = relationships.image_targets.get(&rel_id)?;
    let bytes = media.get(target)?.clone();

    let (width_points, height_points) = drawing.size_points.unwrap_or((240.0, 180.0));
    let alt_text = drawing.alt_text.unwrap_or_else(|| {
        Path::new(target)
            .file_stem()
            .and_then(|value| value.to_str())
            .unwrap_or("Image")
            .to_owned()
    });

    let layout_mode = if drawing.is_anchor {
        crate::document::ImageLayoutMode::Floating
    } else {
        crate::document::ImageLayoutMode::Inline
    };
    let wrap_mode = drawing.wrap_mode.unwrap_or(if drawing.is_anchor {
        crate::document::WrapMode::Square
    } else {
        crate::document::WrapMode::Inline
    });
    let distance_from_text = drawing.distance_from_text.unwrap_or_default();

    let image = DocumentImage {
        id: *next_image_id,
        bytes,
        alt_text,
        width_points,
        height_points,
        lock_aspect_ratio: true,
        opacity: 1.0,
        layout_mode,
        wrap_mode,
        rendering: crate::document::ImageRendering::Smooth,
        horizontal_position: Default::default(),
        vertical_position: Default::default(),
        distance_from_text,
        z_index: 0,
        move_with_text: true,
        allow_overlap: false,
    };
    *next_image_id += 1;
    Some(image)
}
