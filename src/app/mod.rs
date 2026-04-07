mod actions;
mod chrome;
mod palette;

use std::collections::HashMap;
use std::path::PathBuf;

use eframe::{egui, App, CreationContext, Frame};

use crate::{
    canvas::paint_document_canvas,
    document::{CharacterStyle, DocumentState, ParagraphStyle},
};

use actions::handle_global_shortcuts;
use chrome::{paint_ribbon, paint_status_bar, paint_tab_row, paint_title_bar, RibbonTab};
use palette::{configure_theme, theme_palette};

pub use palette::ThemeMode;

const HISTORY_LIMIT: usize = 200;

pub struct ChangeHistory {
    undo_stack: Vec<DocumentState>,
    redo_stack: Vec<DocumentState>,
    last_checkpoint_time: f64,
}

impl ChangeHistory {
    pub fn new() -> Self {
        Self {
            undo_stack: Vec::new(),
            redo_stack: Vec::new(),
            last_checkpoint_time: f64::NEG_INFINITY,
        }
    }

    fn push_snapshot(&mut self, document: &DocumentState) {
        self.undo_stack.push(document.clone());
        self.redo_stack.clear();
        if self.undo_stack.len() > HISTORY_LIMIT {
            self.undo_stack.remove(0);
        }
    }

    /// Always checkpoint — use before discrete actions (button clicks).
    pub fn checkpoint(&mut self, document: &DocumentState, now: f64) {
        self.push_snapshot(document);
        self.last_checkpoint_time = now;
    }

    /// Checkpoint only if enough time has elapsed — use before continuous controls (drag values).
    pub fn checkpoint_coalesced(&mut self, document: &DocumentState, now: f64) {
        if now - self.last_checkpoint_time > 0.75 {
            self.push_snapshot(document);
            self.last_checkpoint_time = now;
        }
    }

    pub fn undo(&mut self, document: &mut DocumentState) -> bool {
        if let Some(prev) = self.undo_stack.pop() {
            self.redo_stack.push(document.clone());
            if self.redo_stack.len() > HISTORY_LIMIT {
                self.redo_stack.remove(0);
            }
            *document = prev;
            true
        } else {
            false
        }
    }

    pub fn redo(&mut self, document: &mut DocumentState) -> bool {
        if let Some(next) = self.redo_stack.pop() {
            self.undo_stack.push(document.clone());
            if self.undo_stack.len() > HISTORY_LIMIT {
                self.undo_stack.remove(0);
            }
            *document = next;
            true
        } else {
            false
        }
    }

    pub fn can_undo(&self) -> bool {
        !self.undo_stack.is_empty()
    }

    pub fn can_redo(&self) -> bool {
        !self.redo_stack.is_empty()
    }

    pub fn clear(&mut self) {
        self.undo_stack.clear();
        self.redo_stack.clear();
    }
}

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub enum ResizeHandle {
    NW, N, NE, E, SE, S, SW, W,
}

pub struct ImageResizeDrag {
    pub image_id: usize,
    pub handle: ResizeHandle,
    pub start_ptr: egui::Pos2,
    pub start_width_points: f32,
    pub start_height_points: f32,
}

pub struct CanvasState {
    pub zoom: f32,
    pub pan: egui::Vec2,
    pub selection: egui::text_selection::CCursorRange,
    pub active_style: CharacterStyle,
    pub active_paragraph_style: ParagraphStyle,
    pub last_interaction_time: f64,
    pub image_textures: HashMap<usize, egui::TextureHandle>,
    pub selected_image_id: Option<usize>,
    pub image_rects: Vec<(usize, egui::Rect)>,
    pub resize_drag: Option<ImageResizeDrag>,
}

impl Default for CanvasState {
    fn default() -> Self {
        Self {
            zoom: 1.0,
            pan: egui::Vec2::ZERO,
            selection: egui::text_selection::CCursorRange::default(),
            active_style: CharacterStyle::default(),
            active_paragraph_style: ParagraphStyle::default(),
            last_interaction_time: 0.0,
            image_textures: HashMap::new(),
            selected_image_id: None,
            image_rects: Vec::new(),
            resize_drag: None,
        }
    }
}

pub struct WorsApp {
    document: DocumentState,
    canvas: CanvasState,
    history: ChangeHistory,
    active_tab: RibbonTab,
    theme_mode: ThemeMode,
    status_message: String,
    current_path: Option<PathBuf>,
    logo_texture: egui::TextureHandle,
}

const LOGO_BYTES: &[u8] = include_bytes!("../../assets/logo.png");

impl WorsApp {
    pub fn new(cc: &CreationContext<'_>) -> Self {
        cc.egui_ctx
            .set_pixels_per_point(cc.egui_ctx.pixels_per_point());

        let theme_mode = ThemeMode::Light;
        configure_theme(&cc.egui_ctx, theme_mode, theme_palette(theme_mode));

        let logo_texture = {
            let img = ::image::load_from_memory(LOGO_BYTES).expect("Failed to load logo");
            let rgba = img.to_rgba8();
            let size = [rgba.width() as usize, rgba.height() as usize];
            let color_image = egui::ColorImage::from_rgba_unmultiplied(size, rgba.as_raw());
            cc.egui_ctx.load_texture("app-logo", color_image, egui::TextureOptions::LINEAR)
        };

        Self {
            document: DocumentState::bootstrap(),
            canvas: CanvasState::default(),
            history: ChangeHistory::new(),
            active_tab: RibbonTab::Home,
            theme_mode,
            status_message: "Ready".to_owned(),
            current_path: None,
            logo_texture,
        }
    }
}

impl App for WorsApp {
    fn ui(&mut self, ui: &mut egui::Ui, _frame: &mut Frame) {
        handle_global_shortcuts(
            ui,
            &mut self.document,
            &mut self.canvas,
            &mut self.history,
            &mut self.current_path,
            &mut self.status_message,
        );

        let palette = theme_palette(self.theme_mode);
        let status_line = self.status_message.clone();
        configure_theme(ui.ctx(), self.theme_mode, palette);

        egui::Panel::top("title_bar")
            .frame(egui::Frame::new().fill(palette.title_bg))
            .show_inside(ui, |ui| {
                paint_title_bar(
                    ui,
                    &mut self.document,
                    &mut self.canvas,
                    &self.current_path,
                    &status_line,
                    &mut self.theme_mode,
                    &mut self.status_message,
                    &mut self.history,
                    palette,
                    &self.logo_texture,
                );
            });

        egui::Panel::top("tabs_bar")
            .frame(egui::Frame::new().fill(palette.tab_bg))
            .show_inside(ui, |ui| {
                paint_tab_row(ui, &mut self.active_tab, self.canvas.selected_image_id, palette);
            });

        egui::Panel::top("ribbon")
            .frame(
                egui::Frame::new()
                    .fill(palette.ribbon_bg)
                    .stroke(egui::Stroke::new(1.0, palette.border)),
            )
            .show_inside(ui, |ui| {
                paint_ribbon(
                    ui,
                    &mut self.document,
                    &mut self.canvas,
                    &mut self.active_tab,
                    &mut self.status_message,
                    &mut self.current_path,
                    &mut self.theme_mode,
                    &mut self.history,
                    palette,
                );
            });

        egui::CentralPanel::default()
            .frame(egui::Frame::new().fill(palette.workspace_bg))
            .show_inside(ui, |ui| {
                paint_document_canvas(ui, &mut self.document, &mut self.canvas, self.theme_mode, &mut self.history);
            });

        // Auto-switch to Picture contextual tab when an image is selected
        match (self.canvas.selected_image_id, self.active_tab) {
            (Some(_), tab) if tab != RibbonTab::Picture => {
                self.active_tab = RibbonTab::Picture;
            }
            (None, RibbonTab::Picture) => {
                self.active_tab = RibbonTab::Home;
            }
            _ => {}
        }

        egui::Panel::bottom("status")
            .frame(
                egui::Frame::new()
                    .fill(palette.status_bg)
                    .stroke(egui::Stroke::new(1.0, palette.border))
                    .inner_margin(egui::Margin::symmetric(10, 6)),
            )
            .show_inside(ui, |ui| {
                paint_status_bar(
                    ui,
                    &self.document,
                    &self.canvas,
                    &self.status_message,
                    palette,
                );
            });
    }
}
