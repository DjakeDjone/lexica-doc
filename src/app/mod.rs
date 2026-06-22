pub(crate) mod actions;
mod backstage;
mod canvas_state;
mod chrome;
pub(crate) mod find_replace;
mod grammar;
mod history;
mod palette;
mod recent_files;
mod settings;

use std::path::PathBuf;
#[cfg(not(target_arch = "wasm32"))]
use std::process::Child;

#[cfg(not(target_arch = "wasm32"))]
use crate::grammar::task::{GrammarRequest, GrammarTaskResult};
use crate::{
    canvas::{paint_document_canvas, CanvasOutput},
    document::{
        DocumentState, DOCX_BODY_BOLD, DOCX_CALADEA_BOLD, DOCX_CARLITO_BOLD, DOCX_COMIC_SANS_BOLD,
        DOCX_LIBERATION_MONO_BOLD, DOCX_LIBERATION_SANS_BOLD, DOCX_LIBERATION_SERIF_BOLD,
        DOCX_MONOSPACE_BOLD,
    },
    grammar::{GrammarConfig, GrammarError, GrammarStatus},
};
use eframe::{egui, App, CreationContext, Frame};
#[cfg(not(target_arch = "wasm32"))]
use tokio::runtime::{Builder as RuntimeBuilder, Runtime};
#[cfg(not(target_arch = "wasm32"))]
use tokio::sync::mpsc;

#[cfg(not(target_arch = "wasm32"))]
use actions::open_document_from_path;
use actions::{handle_global_shortcuts, open_document, save_document, save_document_as_with_name};
use backstage::{paint_backstage, BackstageOutput, BackstageState};
pub use canvas_state::{
    ActiveHeaderFooter, CanvasState, ImageMoveDrag, ImageResizeDrag, ResizeHandle, TableResizeDrag,
    TableResizeHandleRect, TableResizeKind, ZoomMode,
};
use chrome::{paint_ribbon, paint_status_bar, paint_tab_row, paint_title_bar, RibbonTab};
use find_replace::{paint_find_replace_window, FindReplaceState};
pub use history::ChangeHistory;
use palette::{configure_theme, theme_palette};
use recent_files::{load_recent_files, remember_recent_file};
use settings::{load_settings, save_settings, AppSettings, OllamaSettings};

pub use palette::ThemeMode;

const DOCX_CARLITO: &str = "docx-carlito";
const DOCX_CALADEA: &str = "docx-caladea";
const DOCX_LIBERATION_SANS: &str = "docx-liberation-sans";
const DOCX_LIBERATION_SERIF: &str = "docx-liberation-serif";
const DOCX_LIBERATION_MONO: &str = "docx-liberation-mono";
const DOCX_COMIC_SANS: &str = "docx-comic-sans";
#[cfg(not(target_arch = "wasm32"))]
use grammar::GrammarDownloadResult;
use grammar::GrammarDownloadStatus;

#[cfg(not(target_arch = "wasm32"))]
pub enum DialogAction {
    OpenDocument(PathBuf),
    SaveDocument(PathBuf),
    InsertImage(PathBuf),
}

pub struct WorsApp {
    document: DocumentState,
    canvas: CanvasState,
    history: ChangeHistory,
    active_tab: RibbonTab,
    theme_mode: ThemeMode,
    backstage: BackstageState,
    find_replace: FindReplaceState,
    status_message: String,
    current_path: Option<PathBuf>,
    recent_files: Vec<PathBuf>,
    logo_texture: egui::TextureHandle,
    grammar_config: GrammarConfig,
    grammar_errors: Vec<GrammarError>,
    grammar_status: GrammarStatus,
    #[cfg(not(target_arch = "wasm32"))]
    grammar_tx: Option<mpsc::Sender<GrammarRequest>>,
    #[cfg(not(target_arch = "wasm32"))]
    grammar_results_rx: Option<mpsc::Receiver<GrammarTaskResult>>,
    #[cfg(not(target_arch = "wasm32"))]
    _grammar_runtime: Option<Runtime>,
    #[cfg(not(target_arch = "wasm32"))]
    grammar_process: Option<Child>,
    grammar_warning_message: Option<String>,
    show_grammar_warning: bool,
    grammar_download_status: GrammarDownloadStatus,
    #[cfg(not(target_arch = "wasm32"))]
    grammar_download_rx: Option<mpsc::UnboundedReceiver<GrammarDownloadResult>>,
    grammar_auto_check: bool,
    tracked_path: Option<PathBuf>,
    last_cursor_index: usize,
    pub(crate) ai_config: OllamaSettings,
    #[cfg(not(target_arch = "wasm32"))]
    ai_tx: Option<mpsc::Sender<crate::ai::task::AiRequest>>,
    #[cfg(not(target_arch = "wasm32"))]
    ai_rx: Option<mpsc::Receiver<crate::ai::task::AiTaskResult>>,
    #[cfg(not(target_arch = "wasm32"))]
    _ai_runtime: Option<Runtime>,
    #[cfg(not(target_arch = "wasm32"))]
    dialog_tx: std::sync::mpsc::Sender<DialogAction>,
    #[cfg(not(target_arch = "wasm32"))]
    dialog_rx: std::sync::mpsc::Receiver<DialogAction>,
    #[cfg(not(target_arch = "wasm32"))]
    api_rx: Option<tokio::sync::mpsc::Receiver<crate::http_server::ApiRequest>>,
}

const LOGO_BYTES: &[u8] = include_bytes!("../../assets/logo.png");

impl WorsApp {
    #[allow(unused_variables)]
    pub fn new(cc: &CreationContext<'_>, file_to_open: Option<PathBuf>) -> Self {
        cc.egui_ctx
            .set_pixels_per_point(cc.egui_ctx.pixels_per_point());
        configure_docx_fonts(&cc.egui_ctx);

        let mut theme_mode = ThemeMode::Light;
        let mut canvas = CanvasState::default();
        let mut ai_config = OllamaSettings::default();
        if let Some(settings) = load_settings(cc.storage) {
            settings.apply(&mut theme_mode, &mut canvas, &mut ai_config);
        }
        configure_theme(&cc.egui_ctx, theme_mode, theme_palette(theme_mode));

        let logo_texture = {
            let img = ::image::load_from_memory(LOGO_BYTES).expect("Failed to load logo");
            let rgba = img.to_rgba8();
            let size = [rgba.width() as usize, rgba.height() as usize];
            let color_image = egui::ColorImage::from_rgba_unmultiplied(size, rgba.as_raw());
            cc.egui_ctx
                .load_texture("app-logo", color_image, egui::TextureOptions::LINEAR)
        };

        let grammar_config = GrammarConfig::default();
        #[cfg(not(target_arch = "wasm32"))]
        let mut grammar_status = GrammarStatus::Idle;
        #[cfg(target_arch = "wasm32")]
        let grammar_status = GrammarStatus::Unavailable(
            "Grammar checking is not available in the web build".to_owned(),
        );
        #[cfg(not(target_arch = "wasm32"))]
        let mut grammar_warning_message = None;
        #[cfg(target_arch = "wasm32")]
        let grammar_warning_message =
            Some("Grammar checking is not available in the web build".to_owned());
        #[cfg(not(target_arch = "wasm32"))]
        let mut show_grammar_warning = false;
        #[cfg(target_arch = "wasm32")]
        let show_grammar_warning = false;
        #[cfg(not(target_arch = "wasm32"))]
        let grammar_runtime = match RuntimeBuilder::new_multi_thread().enable_all().build() {
            Ok(runtime) => Some(runtime),
            Err(error) => {
                grammar_status =
                    GrammarStatus::Unavailable(format!("Failed to start grammar runtime: {error}"));
                None
            }
        };

        #[cfg(not(target_arch = "wasm32"))]
        let (ai_tx, ai_rx, _ai_runtime) = {
            let rt = RuntimeBuilder::new_multi_thread().enable_all().build().ok();
            if let Some(rt) = rt {
                let (tx, rx) = tokio::sync::mpsc::channel(8);
                let (res_tx, res_rx) = tokio::sync::mpsc::channel(8);
                rt.spawn(crate::ai::task::run_ai_task(rx, res_tx));
                (Some(tx), Some(res_rx), Some(rt))
            } else {
                (None, None, None)
            }
        };

        #[cfg(not(target_arch = "wasm32"))]
        if !grammar_config.lt_jar_path.exists() {
            let message = format!(
                "LanguageTool JAR not found at {}",
                grammar_config.lt_jar_path.display()
            );
            grammar_status = GrammarStatus::Unavailable(message.clone());
            grammar_warning_message = Some(message);
            show_grammar_warning = true;
        }

        #[cfg(not(target_arch = "wasm32"))]
        let (dialog_tx, dialog_rx) = std::sync::mpsc::channel();

        #[cfg(not(target_arch = "wasm32"))]
        let mut api_rx = None;
        #[cfg(not(target_arch = "wasm32"))]
        {
            if let Some(rt) = &_ai_runtime {
                let (tx, rx) = tokio::sync::mpsc::channel(32);
                rt.spawn(async move {
                    if let Err(e) = crate::http_server::start_server(tx).await {
                        eprintln!("Failed to start API server: {}", e);
                    }
                });
                api_rx = Some(rx);
            }
        }

        #[cfg_attr(target_arch = "wasm32", allow(unused_mut))]
        let mut app = Self {
            document: DocumentState::bootstrap(),
            canvas,
            history: ChangeHistory::new(),
            active_tab: RibbonTab::Home,
            theme_mode,
            backstage: BackstageState::default(),
            find_replace: FindReplaceState::default(),
            status_message: "Ready".to_owned(),
            current_path: None,
            recent_files: load_recent_files(),
            logo_texture,
            grammar_config,
            grammar_errors: Vec::new(),
            grammar_status,
            #[cfg(not(target_arch = "wasm32"))]
            grammar_tx: None,
            #[cfg(not(target_arch = "wasm32"))]
            grammar_results_rx: None,
            #[cfg(not(target_arch = "wasm32"))]
            _grammar_runtime: grammar_runtime,
            #[cfg(not(target_arch = "wasm32"))]
            grammar_process: None,
            grammar_warning_message,
            show_grammar_warning,
            grammar_download_status: GrammarDownloadStatus::Idle,
            #[cfg(not(target_arch = "wasm32"))]
            grammar_download_rx: None,
            grammar_auto_check: true,
            tracked_path: None,
            last_cursor_index: 0,
            ai_config,
            #[cfg(not(target_arch = "wasm32"))]
            ai_tx,
            #[cfg(not(target_arch = "wasm32"))]
            ai_rx,
            #[cfg(not(target_arch = "wasm32"))]
            _ai_runtime,
            #[cfg(not(target_arch = "wasm32"))]
            dialog_tx,
            #[cfg(not(target_arch = "wasm32"))]
            dialog_rx,
            #[cfg(not(target_arch = "wasm32"))]
            api_rx,
        };

        #[cfg(not(target_arch = "wasm32"))]
        if app.grammar_config.lt_jar_path.exists() {
            if let Err(message) = app.start_grammar_service() {
                app.grammar_status = GrammarStatus::Unavailable(message);
            }
        }

        #[cfg(not(target_arch = "wasm32"))]
        if let Some(path) = file_to_open {
            open_document_from_path(
                &mut app.document,
                &mut app.canvas,
                &mut app.status_message,
                &mut app.current_path,
                &mut app.history,
                &path,
            );
        }

        app
    }

    fn remember_recent_file(&mut self, path: PathBuf) {
        remember_recent_file(&mut self.recent_files, path);
    }

    fn poll_ai_results(&mut self) {
        #[cfg(not(target_arch = "wasm32"))]
        if let Some(rx) = &mut self.ai_rx {
            while let Ok(result) = rx.try_recv() {
                match result {
                    crate::ai::task::AiTaskResult::Started => {
                        self.canvas.ai_working = true;
                    }
                    crate::ai::task::AiTaskResult::Completed(text) => {
                        self.canvas.ai_working = false;
                        if !text.is_empty() {
                            self.canvas.ai_completion = Some(text);
                        } else {
                            self.canvas.ai_completion = None;
                        }
                    }
                    crate::ai::task::AiTaskResult::Unavailable(err) => {
                        self.canvas.ai_working = false;
                        self.status_message = format!("AI Error: {}", err);
                    }
                }
            }
        }
    }

    #[cfg(not(target_arch = "wasm32"))]
    fn poll_api_requests(&mut self) {
        if let Some(rx) = &mut self.api_rx {
            while let Ok(request) = rx.try_recv() {
                crate::http_server::handle_api_request(
                    request,
                    &mut self.document,
                    &mut self.canvas,
                    &mut self.history,
                    &mut self.status_message,
                    &self.current_path,
                );
            }
        }
    }
}

fn configure_docx_fonts(ctx: &egui::Context) {
    let mut fonts = egui::FontDefinitions::default();
    register_font(
        &mut fonts,
        DOCX_CARLITO,
        include_bytes!("../../assets/fonts/Carlito-Regular.ttf"),
    );
    register_font(
        &mut fonts,
        DOCX_CALADEA,
        include_bytes!("../../assets/fonts/Caladea-Regular.ttf"),
    );
    register_font(
        &mut fonts,
        DOCX_LIBERATION_SANS,
        include_bytes!("../../assets/fonts/LiberationSans-Regular.ttf"),
    );
    register_font(
        &mut fonts,
        DOCX_LIBERATION_SERIF,
        include_bytes!("../../assets/fonts/LiberationSerif-Regular.ttf"),
    );
    register_font(
        &mut fonts,
        DOCX_LIBERATION_MONO,
        include_bytes!("../../assets/fonts/LiberationMono-Regular.ttf"),
    );
    register_font(
        &mut fonts,
        DOCX_COMIC_SANS,
        include_bytes!("../../assets/fonts/ComicNeue-Regular.ttf"),
    );
    register_font_with_fallback(
        &mut fonts,
        DOCX_BODY_BOLD,
        include_bytes!("../../assets/fonts/LiberationSans-Bold.ttf"),
        &[DOCX_LIBERATION_SANS],
    );
    register_font_with_fallback(
        &mut fonts,
        DOCX_MONOSPACE_BOLD,
        include_bytes!("../../assets/fonts/LiberationMono-Bold.ttf"),
        &[DOCX_LIBERATION_MONO],
    );
    register_font_with_fallback(
        &mut fonts,
        DOCX_CARLITO_BOLD,
        include_bytes!("../../assets/fonts/Carlito-Bold.ttf"),
        &[DOCX_CARLITO],
    );
    register_font_with_fallback(
        &mut fonts,
        DOCX_CALADEA_BOLD,
        include_bytes!("../../assets/fonts/Caladea-Bold.ttf"),
        &[DOCX_CALADEA],
    );
    register_font_with_fallback(
        &mut fonts,
        DOCX_LIBERATION_SANS_BOLD,
        include_bytes!("../../assets/fonts/LiberationSans-Bold.ttf"),
        &[DOCX_LIBERATION_SANS],
    );
    register_font_with_fallback(
        &mut fonts,
        DOCX_LIBERATION_SERIF_BOLD,
        include_bytes!("../../assets/fonts/LiberationSerif-Bold.ttf"),
        &[DOCX_LIBERATION_SERIF],
    );
    register_font_with_fallback(
        &mut fonts,
        DOCX_LIBERATION_MONO_BOLD,
        include_bytes!("../../assets/fonts/LiberationMono-Bold.ttf"),
        &[DOCX_LIBERATION_MONO],
    );
    register_font_with_fallback(
        &mut fonts,
        DOCX_COMIC_SANS_BOLD,
        include_bytes!("../../assets/fonts/ComicNeue-Bold.ttf"),
        &[DOCX_COMIC_SANS],
    );
    ctx.set_fonts(fonts);
}

fn register_font(fonts: &mut egui::FontDefinitions, name: &str, bytes: &'static [u8]) {
    fonts
        .font_data
        .insert(name.to_owned(), egui::FontData::from_static(bytes).into());
    fonts
        .families
        .insert(egui::FontFamily::Name(name.into()), vec![name.to_owned()]);
}

fn register_font_with_fallback(
    fonts: &mut egui::FontDefinitions,
    name: &str,
    bytes: &'static [u8],
    fallback_names: &[&str],
) {
    fonts
        .font_data
        .insert(name.to_owned(), egui::FontData::from_static(bytes).into());
    let mut family = vec![name.to_owned()];
    family.extend(fallback_names.iter().map(|fallback| fallback.to_string()));
    fonts
        .families
        .insert(egui::FontFamily::Name(name.into()), family);
}

impl App for WorsApp {
    fn save(&mut self, storage: &mut dyn eframe::Storage) {
        save_settings(
            storage,
            AppSettings::from_state(self.theme_mode, &self.canvas, self.ai_config.clone()),
        );
    }

    fn ui(&mut self, ui: &mut egui::Ui, frame: &mut Frame) {
        #[cfg(not(target_arch = "wasm32"))]
        handle_window_resize(ui.ctx());

        self.poll_grammar_results();
        self.poll_grammar_download();
        self.poll_ai_results();
        #[cfg(not(target_arch = "wasm32"))]
        self.poll_api_requests();

        #[cfg(not(target_arch = "wasm32"))]
        if let Ok(action) = self.dialog_rx.try_recv() {
            match action {
                DialogAction::OpenDocument(path) => {
                    if crate::app::actions::open_document_from_path(
                        &mut self.document,
                        &mut self.canvas,
                        &mut self.status_message,
                        &mut self.current_path,
                        &mut self.history,
                        &path,
                    ) {
                        self.remember_recent_file(path);
                        self.backstage
                            .open_save_as(&self.document, &self.current_path);
                        self.backstage.visible = false;
                    }
                }
                DialogAction::SaveDocument(path) => match self.document.save_to_path(&path) {
                    Ok(()) => {
                        self.current_path = Some(path.clone());
                        self.status_message = format!(
                            "Saved {}",
                            path.file_name()
                                .and_then(|name| name.to_str())
                                .unwrap_or("document")
                        );
                        self.remember_recent_file(path);
                        self.backstage.visible = false;
                    }
                    Err(error) => {
                        self.status_message = error;
                    }
                },
                DialogAction::InsertImage(path) => {
                    crate::app::actions::insert::finish_insert_image(
                        &mut self.document,
                        &mut self.canvas,
                        &mut self.status_message,
                        &mut self.history,
                        &path,
                    );
                }
            }
        }

        let initial_settings =
            AppSettings::from_state(self.theme_mode, &self.canvas, self.ai_config.clone());

        let shortcut_changed = handle_global_shortcuts(
            ui,
            &mut self.document,
            &mut self.canvas,
            &mut self.history,
            &mut self.current_path,
            &mut self.status_message,
            #[cfg(not(target_arch = "wasm32"))]
            &self.dialog_tx,
        );

        if ui.input_mut(|input| {
            input.consume_key(
                egui::Modifiers::COMMAND | egui::Modifiers::SHIFT,
                egui::Key::O,
            )
        }) {
            self.ai_config.enable = !self.ai_config.enable;
            self.status_message = format!(
                "AI completions {}",
                if self.ai_config.enable {
                    "enabled"
                } else {
                    "disabled"
                }
            );
        }
        if ui.input_mut(|input| input.consume_key(egui::Modifiers::COMMAND, egui::Key::F)) {
            self.find_replace.visible = true;
        }
        if ui.input_mut(|input| input.consume_key(egui::Modifiers::COMMAND, egui::Key::H)) {
            self.find_replace.visible = true;
        }

        let palette = theme_palette(self.theme_mode);
        let status_line = self.status_message.clone();
        #[cfg(not(target_arch = "wasm32"))]
        let grammar_download_available = self._grammar_runtime.is_some();
        #[cfg(target_arch = "wasm32")]
        let grammar_download_available = false;
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

        if !self.backstage.visible {
            let mut file_requested = false;
            egui::Panel::top("tabs_bar")
                .frame(egui::Frame::new().fill(palette.tab_bg))
                .show_inside(ui, |ui| {
                    file_requested = paint_tab_row(
                        ui,
                        &mut self.active_tab,
                        self.canvas.selected_image_id,
                        self.canvas.active_table_cell,
                        self.canvas.active_header_footer.is_some(),
                        palette,
                    );
                });
            if file_requested {
                self.backstage
                    .open_save_as(&self.document, &self.current_path);
            }
        }
        if self.backstage.visible
            && ui.input_mut(|input| input.consume_key(egui::Modifiers::NONE, egui::Key::Escape))
        {
            self.backstage.visible = false;
        }

        let mut grammar_ribbon_output = chrome::GrammarRibbonOutput::default();
        let mut canvas_output = CanvasOutput::default();
        if self.backstage.visible {
            let mut backstage_output = BackstageOutput::default();
            egui::CentralPanel::default()
                .frame(egui::Frame::new().fill(palette.workspace_bg))
                .show_inside(ui, |ui| {
                    backstage_output = paint_backstage(
                        ui,
                        &mut self.backstage,
                        &self.document,
                        &self.current_path,
                        &self.recent_files,
                        &mut self.ai_config,
                        palette,
                    );
                });

            if backstage_output.close_requested {
                self.backstage.visible = false;
            }
            if backstage_output.save_requested {
                #[cfg(not(target_arch = "wasm32"))]
                save_document(
                    &self.document,
                    &mut self.status_message,
                    &mut self.current_path,
                    &self.dialog_tx,
                );
                #[cfg(target_arch = "wasm32")]
                save_document(
                    &self.document,
                    &mut self.status_message,
                    &mut self.current_path,
                );
            }
            if backstage_output.save_as_requested {
                #[cfg(not(target_arch = "wasm32"))]
                save_document_as_with_name(
                    &self.document,
                    &mut self.status_message,
                    &mut self.current_path,
                    &self.backstage.file_name,
                    self.backstage.format.extension(),
                    &self.dialog_tx,
                );
                #[cfg(target_arch = "wasm32")]
                save_document_as_with_name(
                    &self.document,
                    &mut self.status_message,
                    &mut self.current_path,
                    &self.backstage.file_name,
                    self.backstage.format.extension(),
                );
            }
            if backstage_output.open_requested {
                #[cfg(not(target_arch = "wasm32"))]
                open_document(
                    &mut self.document,
                    &mut self.canvas,
                    &mut self.status_message,
                    &mut self.current_path,
                    &mut self.history,
                    &self.dialog_tx,
                );
                #[cfg(target_arch = "wasm32")]
                open_document(
                    &mut self.document,
                    &mut self.canvas,
                    &mut self.status_message,
                    &mut self.current_path,
                    &mut self.history,
                );
            }
            if let Some(path) = backstage_output.recent_open_requested {
                #[cfg(not(target_arch = "wasm32"))]
                {
                    if open_document_from_path(
                        &mut self.document,
                        &mut self.canvas,
                        &mut self.status_message,
                        &mut self.current_path,
                        &mut self.history,
                        &path,
                    ) {
                        self.remember_recent_file(path);
                        self.backstage
                            .open_save_as(&self.document, &self.current_path);
                        self.backstage.visible = false;
                    }
                }
                #[cfg(target_arch = "wasm32")]
                {
                    let _ = path;
                    self.status_message =
                        "Opening recent files is not available in the web build yet".to_owned();
                }
            }
        } else {
            egui::Panel::top("ribbon")
                .frame(
                    egui::Frame::new()
                        .fill(palette.ribbon_bg)
                        .stroke(egui::Stroke::new(1.0, palette.border)),
                )
                .show_inside(ui, |ui| {
                    grammar_ribbon_output = paint_ribbon(
                        ui,
                        &mut self.document,
                        &mut self.canvas,
                        &mut self.active_tab,
                        &mut self.status_message,
                        &mut self.current_path,
                        &mut self.theme_mode,
                        &mut self.history,
                        &mut self.find_replace,
                        &mut self.grammar_config,
                        &self.grammar_status,
                        &mut self.grammar_auto_check,
                        grammar_download_available,
                        #[cfg(not(target_arch = "wasm32"))]
                        &self.dialog_tx,
                        palette,
                    );
                });

            egui::CentralPanel::default()
                .frame(egui::Frame::new().fill(palette.workspace_bg))
                .show_inside(ui, |ui| {
                    canvas_output = paint_document_canvas(
                        ui,
                        &mut self.document,
                        &mut self.canvas,
                        self.theme_mode,
                        &mut self.history,
                        &self.grammar_errors,
                    );
                });
        }

        if grammar_ribbon_output.download_requested {
            self.start_grammar_download();
        }
        if grammar_ribbon_output.restart_requested {
            self.restart_grammar_service();
        }
        if grammar_ribbon_output.manual_check_requested {
            self.request_grammar_check(true);
        }
        if grammar_ribbon_output.settings_changed {
            self.status_message = "Grammar settings updated".to_owned();
            if self.grammar_auto_check {
                self.request_grammar_check(false);
            }
        }
        if shortcut_changed || canvas_output.text_changed {
            self.request_grammar_check(false);
        }

        let current_cursor = self.canvas.selection.primary.index;
        if canvas_output.text_changed || self.last_cursor_index != current_cursor {
            self.last_cursor_index = current_cursor;
            self.canvas.ai_completion = None;
            self.canvas.ai_working = false;
            #[cfg(not(target_arch = "wasm32"))]
            if self.ai_config.enable {
                if let Some(tx) = &self.ai_tx {
                    let cursor_idx = self.canvas.selection.primary.index;
                    let text = self.document.plain_text();
                    let text_before = text.chars().take(cursor_idx).collect::<String>();
                    let text_before_cursor = text_before
                        .chars()
                        .rev()
                        .take(1000)
                        .collect::<Vec<_>>()
                        .into_iter()
                        .rev()
                        .collect();
                    let text_after_cursor =
                        text.chars().skip(cursor_idx).take(1000).collect::<String>();

                    let _ = tx.try_send(crate::ai::task::AiRequest {
                        text_before_cursor,
                        text_after_cursor,
                        endpoint: self.ai_config.endpoint.clone(),
                        model: self.ai_config.model.clone(),
                    });
                }
            }
        }

        if self.current_path != self.tracked_path {
            self.tracked_path = self.current_path.clone();
            if let Some(ref path) = self.current_path {
                self.remember_recent_file(path.clone());
            }
        }

        // Auto-switch to contextual tabs when an object is selected.
        match (
            self.canvas.active_header_footer,
            self.canvas.selected_image_id,
            self.canvas.active_table_cell,
            self.active_tab,
        ) {
            (Some(_), _, _, tab) if tab != RibbonTab::HeaderFooter => {
                self.active_tab = RibbonTab::HeaderFooter;
            }
            (None, Some(_), _, tab) if tab != RibbonTab::Picture => {
                self.active_tab = RibbonTab::Picture;
            }
            (None, None, Some(_), tab) if tab != RibbonTab::Table => {
                self.active_tab = RibbonTab::Table;
            }
            (None, None, None, RibbonTab::HeaderFooter | RibbonTab::Picture | RibbonTab::Table) => {
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
                    &self.grammar_status,
                    self.grammar_errors.len(),
                    &mut self.ai_config,
                    palette,
                );
            });

        let find_replace_changed = paint_find_replace_window(
            ui.ctx(),
            &mut self.find_replace,
            &mut self.document,
            &mut self.canvas,
            &mut self.history,
            &mut self.status_message,
        );
        if find_replace_changed {
            self.request_grammar_check(false);
        }

        if self.show_grammar_warning {
            let warning_message = self.grammar_warning_message.clone();
            if let Some(message) = warning_message.as_deref() {
                egui::Window::new("Grammar Checker Unavailable")
                    .collapsible(false)
                    .resizable(false)
                    .anchor(egui::Align2::CENTER_TOP, egui::vec2(0.0, 16.0))
                    .show(ui.ctx(), |ui| {
                        ui.label(message);
                        if self.grammar_download_status == GrammarDownloadStatus::Downloading {
                            ui.horizontal(|ui| {
                                ui.spinner();
                                ui.label("Downloading LanguageTool…");
                            });
                            ui.ctx().request_repaint();
                        } else {
                            let can_download = grammar_download_available;
                            if ui
                                .add_enabled(
                                    can_download,
                                    egui::Button::new("Download LanguageTool (~240 MB)"),
                                )
                                .clicked()
                            {
                                self.start_grammar_download();
                            }
                            if !can_download {
                                ui.label("Download unavailable: runtime failed to initialize.");
                            }
                        }
                        if ui.button("Dismiss").clicked() {
                            self.show_grammar_warning = false;
                        }
                    });
            }
        }

        let current_settings =
            AppSettings::from_state(self.theme_mode, &self.canvas, self.ai_config.clone());
        if current_settings != initial_settings {
            if let Some(storage) = frame.storage_mut() {
                save_settings(storage, current_settings);
            }
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn handle_window_resize(ctx: &egui::Context) {
    #[allow(deprecated)]
    let rect = ctx.screen_rect();
    let resize_margin = 6.0;

    if let Some(pos) = ctx.pointer_hover_pos() {
        let left = pos.x < rect.min.x + resize_margin;
        let right = pos.x > rect.max.x - resize_margin;
        let top = pos.y < rect.min.y + resize_margin;
        let bottom = pos.y > rect.max.y - resize_margin;

        let direction = match (top, bottom, left, right) {
            (true, false, true, false) => Some(egui::ResizeDirection::NorthWest),
            (true, false, false, true) => Some(egui::ResizeDirection::NorthEast),
            (false, true, true, false) => Some(egui::ResizeDirection::SouthWest),
            (false, true, false, true) => Some(egui::ResizeDirection::SouthEast),
            (true, false, false, false) => Some(egui::ResizeDirection::North),
            (false, true, false, false) => Some(egui::ResizeDirection::South),
            (false, false, true, false) => Some(egui::ResizeDirection::West),
            (false, false, false, true) => Some(egui::ResizeDirection::East),
            _ => None,
        };

        if let Some(dir) = direction {
            ctx.set_cursor_icon(match dir {
                egui::ResizeDirection::NorthWest | egui::ResizeDirection::SouthEast => {
                    egui::CursorIcon::ResizeNwSe
                }
                egui::ResizeDirection::NorthEast | egui::ResizeDirection::SouthWest => {
                    egui::CursorIcon::ResizeNeSw
                }
                egui::ResizeDirection::North | egui::ResizeDirection::South => {
                    egui::CursorIcon::ResizeVertical
                }
                egui::ResizeDirection::West | egui::ResizeDirection::East => {
                    egui::CursorIcon::ResizeHorizontal
                }
            });

            if ctx.input(|i| i.pointer.any_pressed()) {
                ctx.send_viewport_cmd(egui::ViewportCommand::BeginResize(dir));
            }
        }
    }
}
