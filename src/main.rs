#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

mod app;
mod canvas;
mod document;
mod layout;

use app::WorsApp;

fn main() -> eframe::Result<()> {
    let native_options = eframe::NativeOptions {
        renderer: eframe::Renderer::Wgpu,
        viewport: egui::ViewportBuilder::default()
            .with_title("wors")
            .with_inner_size([1400.0, 1000.0])
            .with_min_inner_size([960.0, 720.0]),
        ..Default::default()
    };

    eframe::run_native(
        "wors",
        native_options,
        Box::new(|cc| Ok(Box::new(WorsApp::new(cc)))),
    )
}
