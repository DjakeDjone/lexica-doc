mod actions;
mod chrome;
mod palette;

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

pub struct CanvasState {
    pub zoom: f32,
    pub pan: egui::Vec2,
    pub selection: egui::text_selection::CCursorRange,
    pub active_style: CharacterStyle,
    pub active_paragraph_style: ParagraphStyle,
    pub last_interaction_time: f64,
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
        }
    }
}

pub struct WorsApp {
    document: DocumentState,
    canvas: CanvasState,
    active_tab: RibbonTab,
    theme_mode: ThemeMode,
    status_message: String,
    current_path: Option<PathBuf>,
}

impl WorsApp {
    pub fn new(cc: &CreationContext<'_>) -> Self {
        cc.egui_ctx
            .set_pixels_per_point(cc.egui_ctx.pixels_per_point());

        let theme_mode = ThemeMode::Light;
        configure_theme(&cc.egui_ctx, theme_mode, theme_palette(theme_mode));

        Self {
            document: DocumentState::bootstrap(),
            canvas: CanvasState::default(),
            active_tab: RibbonTab::Home,
            theme_mode,
            status_message: "Ready".to_owned(),
            current_path: None,
        }
    }
}

impl App for WorsApp {
    fn ui(&mut self, ui: &mut egui::Ui, _frame: &mut Frame) {
        handle_global_shortcuts(
            ui,
            &mut self.document,
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
                    &self.document,
                    &self.current_path,
                    &status_line,
                    &mut self.theme_mode,
                    &mut self.status_message,
                    palette,
                );
            });

        egui::Panel::top("tabs_bar")
            .frame(egui::Frame::new().fill(palette.tab_bg))
            .show_inside(ui, |ui| {
                paint_tab_row(ui, &mut self.active_tab, palette);
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
                    palette,
                );
            });

        egui::CentralPanel::default()
            .frame(egui::Frame::new().fill(palette.workspace_bg))
            .show_inside(ui, |ui| {
                paint_document_canvas(ui, &mut self.document, &mut self.canvas, self.theme_mode);
            });

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
