#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

use std::sync::Arc;

use wors::app::WorsApp;

const LOGO_BYTES: &[u8] = include_bytes!("../assets/logo.png");

fn load_icon() -> egui::viewport::IconData {
    let img = image::load_from_memory(LOGO_BYTES).expect("Failed to load app icon");
    let rgba = img.to_rgba8();
    let (width, height) = (rgba.width(), rgba.height());
    egui::viewport::IconData {
        rgba: rgba.into_raw(),
        width,
        height,
    }
}

fn main() -> eframe::Result<()> {
    let native_options = eframe::NativeOptions {
        renderer: eframe::Renderer::Wgpu,
        viewport: egui::ViewportBuilder::default()
            .with_title("wors")
            .with_inner_size([1400.0, 1000.0])
            .with_min_inner_size([960.0, 720.0])
            .with_icon(Arc::new(load_icon())),
        ..Default::default()
    };

    eframe::run_native(
        "wors",
        native_options,
        Box::new(|cc| Ok(Box::new(WorsApp::new(cc)))),
    )
}
