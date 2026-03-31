use eframe::egui::Color32;

use crate::app::ThemeMode;

#[derive(Clone, Copy)]
pub(super) struct CanvasPalette {
    pub(super) canvas_bg: Color32,
    pub(super) page_bg: Color32,
    pub(super) page_border: Color32,
    pub(super) page_focus: Color32,
    pub(super) page_shadow: Color32,
    pub(super) footer_bg: Color32,
    pub(super) footer_stroke: Color32,
    pub(super) footer_text: Color32,
}

pub(super) fn canvas_palette(mode: ThemeMode) -> CanvasPalette {
    match mode {
        ThemeMode::Light => CanvasPalette {
            canvas_bg: Color32::from_rgb(217, 219, 223),
            page_bg: Color32::from_rgb(255, 255, 255),
            page_border: Color32::from_rgb(186, 193, 204),
            page_focus: Color32::from_rgb(79, 129, 209),
            page_shadow: Color32::from_black_alpha(24),
            footer_bg: Color32::from_rgba_premultiplied(241, 244, 248, 220),
            footer_stroke: Color32::from_rgba_premultiplied(173, 183, 198, 200),
            footer_text: Color32::from_rgb(88, 96, 108),
        },
        ThemeMode::Dark => CanvasPalette {
            canvas_bg: Color32::from_rgb(45, 49, 56),
            page_bg: Color32::from_rgb(250, 250, 252),
            page_border: Color32::from_rgb(124, 132, 147),
            page_focus: Color32::from_rgb(126, 171, 236),
            page_shadow: Color32::from_black_alpha(54),
            footer_bg: Color32::from_rgba_premultiplied(29, 33, 40, 220),
            footer_stroke: Color32::from_rgba_premultiplied(98, 108, 126, 200),
            footer_text: Color32::from_rgb(196, 206, 224),
        },
    }
}
