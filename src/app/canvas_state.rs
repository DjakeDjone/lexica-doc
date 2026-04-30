use std::collections::HashMap;

use eframe::egui;

use crate::document::{CharacterStyle, ParagraphStyle};

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

pub struct CanvasState {
    pub zoom: f32,
    pub zoom_mode: ZoomMode,
    pub imported_docx_view: bool,
    pub pan: egui::Vec2,
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
}

impl Default for CanvasState {
    fn default() -> Self {
        Self {
            zoom: 1.0,
            zoom_mode: ZoomMode::Manual,
            imported_docx_view: false,
            pan: egui::Vec2::ZERO,
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
        }
    }
}
