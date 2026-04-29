#![cfg_attr(
    all(not(debug_assertions), not(target_arch = "wasm32")),
    windows_subsystem = "windows"
)]

#[cfg(not(target_arch = "wasm32"))]
use std::sync::Arc;

use wors::app::WorsApp;

#[cfg(not(target_arch = "wasm32"))]
const LOGO_BYTES: &[u8] = include_bytes!("../assets/logo.png");

#[cfg(not(target_arch = "wasm32"))]
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

#[cfg(not(target_arch = "wasm32"))]
fn main() -> eframe::Result<()> {
    let native_options = eframe::NativeOptions {
        renderer: eframe::Renderer::Wgpu,
        viewport: egui::ViewportBuilder::default()
            .with_title("wors")
            .with_inner_size([1400.0, 1000.0])
            .with_min_inner_size([960.0, 720.0])
            .with_decorations(false)
            .with_icon(Arc::new(load_icon())),
        ..Default::default()
    };

    eframe::run_native(
        "wors",
        native_options,
        Box::new(|cc| Ok(Box::new(WorsApp::new(cc)))),
    )
}

#[cfg(target_arch = "wasm32")]
fn main() {}

#[cfg(target_arch = "wasm32")]
#[wasm_bindgen::prelude::wasm_bindgen(start)]
pub async fn start() -> Result<(), wasm_bindgen::JsValue> {
    use wasm_bindgen::JsCast as _;

    eframe::WebLogger::init(log::LevelFilter::Debug).ok();

    let window = web_sys::window().ok_or("window is unavailable")?;
    let document = window.document().ok_or("document is unavailable")?;
    let canvas = document
        .get_element_by_id("wors-canvas")
        .ok_or("missing #wors-canvas")?
        .dyn_into::<web_sys::HtmlCanvasElement>()?;

    eframe::WebRunner::new()
        .start(
            canvas,
            eframe::WebOptions::default(),
            Box::new(|cc| Ok(Box::new(WorsApp::new(cc)))),
        )
        .await
}
