use std::collections::HashMap;

use eframe::egui;

use crate::document::{
    CharacterStyle, HeaderFooterKind, HeaderFooterVariant, ParagraphStyle, SectionId,
};

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ResizeHandle {
    NW,
    N,
    NE,
    E,
    SE,
    S,
    SW,
    W,
}

pub struct ImageResizeDrag {
    pub image_id: usize,
    pub handle: ResizeHandle,
    pub start_ptr: egui::Pos2,
    pub start_width_points: f32,
    pub start_height_points: f32,
    pub start_x_points: f32,
    pub start_y_points: f32,
}

pub struct ImageMoveDrag {
    pub image_id: usize,
    pub start_ptr: egui::Pos2,
    pub current_ptr: egui::Pos2,
    pub start_rect: egui::Rect,
    pub start_x_points: f32,
    pub start_y_points: f32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum TableResizeKind {
    Column { left_col: usize },
    Row { top_row: usize },
}

#[derive(Clone, Copy, Debug)]
pub struct TableResizeHandleRect {
    pub table_id: usize,
    pub kind: TableResizeKind,
    pub rect: egui::Rect,
}

pub struct TableResizeDrag {
    pub table_id: usize,
    pub kind: TableResizeKind,
    pub start_ptr: egui::Pos2,
    pub first_points: f32,
    pub second_points: f32,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ZoomMode {
    Manual,
    FitPage,
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub struct ActiveHeaderFooter {
    pub kind: HeaderFooterKind,
    pub section_id: SectionId,
    pub variant: HeaderFooterVariant,
    pub page_number: usize,
}

pub struct CanvasState {
    pub zoom: f32,
    pub(crate) zoom_target: f32,
    pub(crate) layout_zoom: f32,
    pub(crate) last_zoom_input_time: f64,
    pub zoom_mode: ZoomMode,
    pub imported_docx_view: bool,
    pub pan: egui::Vec2,
    pub scroll_range: egui::Vec2,
    pub scroll_velocity: egui::Vec2,
    pub scrollbar_drag_offset: f32,
    pub selection: egui::text_selection::CCursorRange,
    pub active_style: CharacterStyle,
    pub active_paragraph_style: ParagraphStyle,
    pub last_interaction_time: f64,
    pub image_textures: HashMap<usize, egui::TextureHandle>,
    pub selected_image_id: Option<usize>,
    pub image_rects: Vec<(usize, egui::Rect)>,
    pub resize_drag: Option<ImageResizeDrag>,
    pub move_drag: Option<ImageMoveDrag>,
    pub active_table_cell: Option<(usize, usize, usize)>,
    pub table_cell_rects: Vec<(usize, usize, usize, egui::Rect)>,
    pub table_cell_content_rects: Vec<(usize, usize, usize, egui::Rect)>,
    pub table_cell_selection: egui::text_selection::CCursorRange,
    pub table_resize_handles: Vec<TableResizeHandleRect>,
    pub table_resize_drag: Option<TableResizeDrag>,
    pub active_header_footer: Option<ActiveHeaderFooter>,
    pub active_header_footer_cursor: usize,
    pub active_header_footer_selection: egui::text_selection::CCursorRange,
    pub ai_completion: Option<String>,
    pub ai_working: bool,
}

impl Default for CanvasState {
    fn default() -> Self {
        Self {
            zoom: 1.0,
            zoom_target: 1.0,
            layout_zoom: 1.0,
            last_zoom_input_time: f64::NEG_INFINITY,
            zoom_mode: ZoomMode::Manual,
            imported_docx_view: false,
            pan: egui::Vec2::ZERO,
            scroll_range: egui::Vec2::ZERO,
            scroll_velocity: egui::Vec2::ZERO,
            scrollbar_drag_offset: 0.0,
            selection: egui::text_selection::CCursorRange::default(),
            active_style: CharacterStyle::default(),
            active_paragraph_style: ParagraphStyle::default(),
            last_interaction_time: 0.0,
            image_textures: HashMap::new(),
            selected_image_id: None,
            image_rects: Vec::new(),
            resize_drag: None,
            move_drag: None,
            active_table_cell: None,
            table_cell_rects: Vec::new(),
            table_cell_content_rects: Vec::new(),
            table_cell_selection: egui::text_selection::CCursorRange::default(),
            table_resize_handles: Vec::new(),
            table_resize_drag: None,
            active_header_footer: None,
            active_header_footer_cursor: 0,
            active_header_footer_selection: egui::text_selection::CCursorRange::default(),
            ai_completion: None,
            ai_working: false,
        }
    }
}

impl CanvasState {
    pub fn scale_view(&mut self, delta: f32) {
        if !delta.is_finite() || delta <= 0.0 {
            return;
        }

        if (crate::layout::quantize_zoom(self.zoom_target) - self.zoom).abs() > 0.001 {
            self.zoom_target = self.zoom;
        }
        self.zoom_target = (self.zoom_target * delta).clamp(0.5, 3.0);
        self.zoom = crate::layout::quantize_zoom(self.zoom_target);
    }
}
