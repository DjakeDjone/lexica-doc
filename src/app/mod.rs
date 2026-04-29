mod actions;
mod chrome;
mod palette;

use std::collections::HashMap;
use std::path::PathBuf;
#[cfg(not(target_arch = "wasm32"))]
use std::process::Child;
#[cfg(not(target_arch = "wasm32"))]
use std::{env, fs};

use eframe::{egui, App, CreationContext, Frame};
#[cfg(not(target_arch = "wasm32"))]
use tokio::runtime::{Builder as RuntimeBuilder, Runtime};
#[cfg(not(target_arch = "wasm32"))]
use tokio::sync::mpsc;

#[cfg(not(target_arch = "wasm32"))]
use crate::grammar::{
    download::{download_languagetool_server_jar, LT_STABLE_ZIP_URL},
    process::{kill_languagetool, spawn_languagetool},
    task::{run_grammar_task, GrammarRequest, GrammarTaskResult},
    GrammarChecker,
};
use crate::{
    canvas::{paint_document_canvas, CanvasOutput},
    document::{CharacterStyle, DocumentState, ParagraphStyle},
    grammar::{GrammarConfig, GrammarError, GrammarStatus},
};

#[cfg(not(target_arch = "wasm32"))]
use actions::open_document_from_path;
use actions::{handle_global_shortcuts, open_document, save_document, save_document_as_with_name};
use chrome::{
    paint_backstage, paint_ribbon, paint_status_bar, paint_tab_row, paint_title_bar,
    BackstageState, RibbonTab,
};
use palette::{configure_theme, theme_palette};

pub use palette::ThemeMode;

const HISTORY_LIMIT: usize = 200;
const DOCX_CARLITO: &str = "docx-carlito";
const DOCX_CALADEA: &str = "docx-caladea";
const DOCX_LIBERATION_SANS: &str = "docx-liberation-sans";
const DOCX_LIBERATION_SERIF: &str = "docx-liberation-serif";
const DOCX_LIBERATION_MONO: &str = "docx-liberation-mono";
const DOCX_COMIC_SANS: &str = "docx-comic-sans";
#[cfg(not(target_arch = "wasm32"))]
const GRAMMAR_QUEUE_CAPACITY: usize = 8;
#[cfg(not(target_arch = "wasm32"))]
const RECENT_FILES_LIMIT: usize = 12;

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum GrammarDownloadStatus {
    Idle,
    Downloading,
}

#[derive(Debug)]
#[cfg(not(target_arch = "wasm32"))]
enum GrammarDownloadResult {
    Ready(PathBuf),
    Failed(String),
}

#[cfg(not(target_arch = "wasm32"))]
fn recent_files_path() -> Option<PathBuf> {
    #[cfg(target_os = "windows")]
    {
        env::var_os("APPDATA")
            .map(PathBuf::from)
            .map(|path| path.join("wors").join("recent-files.json"))
    }
    #[cfg(target_os = "macos")]
    {
        env::var_os("HOME").map(PathBuf::from).map(|path| {
            path.join("Library")
                .join("Application Support")
                .join("wors")
                .join("recent-files.json")
        })
    }
    #[cfg(all(not(target_os = "windows"), not(target_os = "macos")))]
    {
        if let Some(config_home) = env::var_os("XDG_CONFIG_HOME") {
            Some(
                PathBuf::from(config_home)
                    .join("wors")
                    .join("recent-files.json"),
            )
        } else {
            env::var_os("HOME")
                .map(PathBuf::from)
                .map(|path| path.join(".config").join("wors").join("recent-files.json"))
        }
    }
}

#[cfg(not(target_arch = "wasm32"))]
fn normalize_recent_path(path: PathBuf) -> PathBuf {
    path.canonicalize().unwrap_or(path)
}

#[cfg(target_arch = "wasm32")]
fn normalize_recent_path(path: PathBuf) -> PathBuf {
    path
}

#[cfg(not(target_arch = "wasm32"))]
fn load_recent_files() -> Vec<PathBuf> {
    let Some(path) = recent_files_path() else {
        return Vec::new();
    };
    let Ok(source) = fs::read_to_string(path) else {
        return Vec::new();
    };
    let Ok(raw_paths) = serde_json::from_str::<Vec<String>>(&source) else {
        return Vec::new();
    };
    let mut recent_files = Vec::new();
    for raw_path in raw_paths {
        let path = normalize_recent_path(PathBuf::from(raw_path));
        if path.exists() && !recent_files.contains(&path) {
            recent_files.push(path);
        }
        if recent_files.len() >= RECENT_FILES_LIMIT {
            break;
        }
    }
    recent_files
}

#[cfg(target_arch = "wasm32")]
fn load_recent_files() -> Vec<PathBuf> {
    Vec::new()
}

#[cfg(not(target_arch = "wasm32"))]
fn save_recent_files(recent_files: &[PathBuf]) {
    let Some(path) = recent_files_path() else {
        return;
    };
    let Some(parent) = path.parent() else {
        return;
    };
    if fs::create_dir_all(parent).is_err() {
        return;
    }
    let raw_paths: Vec<String> = recent_files
        .iter()
        .filter(|path| path.exists())
        .map(|path| path.display().to_string())
        .collect();
    if let Ok(json) = serde_json::to_string_pretty(&raw_paths) {
        let _ = fs::write(path, json);
    }
}

#[cfg(target_arch = "wasm32")]
fn save_recent_files(_recent_files: &[PathBuf]) {}

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

impl Default for ChangeHistory {
    fn default() -> Self {
        Self::new()
    }
}

pub struct WorsApp {
    document: DocumentState,
    canvas: CanvasState,
    history: ChangeHistory,
    active_tab: RibbonTab,
    theme_mode: ThemeMode,
    backstage: BackstageState,
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
}

const LOGO_BYTES: &[u8] = include_bytes!("../../assets/logo.png");

impl WorsApp {
    pub fn new(cc: &CreationContext<'_>) -> Self {
        cc.egui_ctx
            .set_pixels_per_point(cc.egui_ctx.pixels_per_point());
        configure_docx_fonts(&cc.egui_ctx);

        let theme_mode = ThemeMode::Light;
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
        if !grammar_config.lt_jar_path.exists() {
            let message = format!(
                "LanguageTool JAR not found at {}",
                grammar_config.lt_jar_path.display()
            );
            grammar_status = GrammarStatus::Unavailable(message.clone());
            grammar_warning_message = Some(message);
            show_grammar_warning = true;
        }

        #[cfg_attr(target_arch = "wasm32", allow(unused_mut))]
        let mut app = Self {
            document: DocumentState::bootstrap(),
            canvas: CanvasState::default(),
            history: ChangeHistory::new(),
            active_tab: RibbonTab::Home,
            theme_mode,
            backstage: BackstageState::default(),
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
        };

        #[cfg(not(target_arch = "wasm32"))]
        if app.grammar_config.lt_jar_path.exists() {
            if let Err(message) = app.start_grammar_service() {
                app.grammar_status = GrammarStatus::Unavailable(message);
            }
        }

        app
    }

    fn remember_recent_file(&mut self, path: PathBuf) {
        if path.as_os_str().is_empty() {
            return;
        }
        let path = normalize_recent_path(path);
        if self.recent_files.first() == Some(&path) {
            return;
        }
        self.recent_files.retain(|recent| recent != &path);
        self.recent_files.insert(0, path.clone());
        self.recent_files.truncate(12);
        save_recent_files(&self.recent_files);
    }

    #[cfg(not(target_arch = "wasm32"))]
    fn start_grammar_service(&mut self) -> Result<(), String> {
        self.stop_grammar_service();

        if !self.grammar_config.lt_jar_path.exists() {
            return Err(format!(
                "LanguageTool JAR not found at {}",
                self.grammar_config.lt_jar_path.display()
            ));
        }

        let Some(runtime) = self._grammar_runtime.as_ref() else {
            return Err("Grammar runtime is unavailable".to_owned());
        };

        let child = spawn_languagetool(&self.grammar_config)
            .map_err(|error| format!("Grammar unavailable: {error}"))?;
        let (tx, rx) = mpsc::channel(GRAMMAR_QUEUE_CAPACITY);
        let (results_tx, results_rx) = mpsc::channel(GRAMMAR_QUEUE_CAPACITY);

        runtime.spawn(run_grammar_task(
            rx,
            results_tx,
            GrammarChecker::new(self.grammar_config.port),
            self.grammar_config.port,
        ));

        self.grammar_process = Some(child);
        self.grammar_tx = Some(tx);
        self.grammar_results_rx = Some(results_rx);
        self.grammar_status = GrammarStatus::Idle;
        Ok(())
    }

    #[cfg(target_arch = "wasm32")]
    fn start_grammar_service(&mut self) -> Result<(), String> {
        Err("Grammar checking is not available in the web build".to_owned())
    }

    #[cfg(not(target_arch = "wasm32"))]
    fn stop_grammar_service(&mut self) {
        self.grammar_tx = None;
        self.grammar_results_rx = None;
        if let Some(child) = self.grammar_process.as_mut() {
            kill_languagetool(child);
        }
        self.grammar_process = None;
    }

    #[cfg(target_arch = "wasm32")]
    fn stop_grammar_service(&mut self) {}

    fn restart_grammar_service(&mut self) {
        match self.start_grammar_service() {
            Ok(()) => {
                self.grammar_warning_message = None;
                self.show_grammar_warning = false;
                self.status_message = "Grammar server restarted".to_owned();
            }
            Err(message) => {
                self.grammar_status = GrammarStatus::Unavailable(message.clone());
                self.grammar_warning_message = Some(message);
                self.show_grammar_warning = true;
            }
        }
    }

    #[cfg(not(target_arch = "wasm32"))]
    fn poll_grammar_results(&mut self) {
        let Some(results_rx) = self.grammar_results_rx.as_mut() else {
            return;
        };

        while let Ok(message) = results_rx.try_recv() {
            match message {
                GrammarTaskResult::Completed(errors) => {
                    self.grammar_errors = errors;
                    self.grammar_status = GrammarStatus::Done;
                }
                GrammarTaskResult::Unavailable(message) => {
                    self.grammar_status = GrammarStatus::Unavailable(message);
                }
            }
        }
    }

    #[cfg(target_arch = "wasm32")]
    fn poll_grammar_results(&mut self) {}

    #[cfg(not(target_arch = "wasm32"))]
    fn start_grammar_download(&mut self) {
        if self.grammar_download_status == GrammarDownloadStatus::Downloading {
            return;
        }

        let Some(runtime) = self._grammar_runtime.as_ref() else {
            self.grammar_status = GrammarStatus::Unavailable(
                "Cannot download LanguageTool because runtime is unavailable".to_owned(),
            );
            return;
        };

        let target_path = self.grammar_config.lt_jar_path.clone();
        let (tx, rx) = mpsc::unbounded_channel::<GrammarDownloadResult>();
        runtime.spawn(async move {
            let result = match download_languagetool_server_jar(target_path.clone()).await {
                Ok(()) => GrammarDownloadResult::Ready(target_path),
                Err(error) => GrammarDownloadResult::Failed(error.to_string()),
            };
            let _ = tx.send(result);
        });

        self.grammar_download_rx = Some(rx);
        self.grammar_download_status = GrammarDownloadStatus::Downloading;
        self.show_grammar_warning = true;
        self.grammar_warning_message = Some(format!(
            "Downloading LanguageTool from {LT_STABLE_ZIP_URL}. This can take a while."
        ));
    }

    #[cfg(target_arch = "wasm32")]
    fn start_grammar_download(&mut self) {
        self.grammar_status = GrammarStatus::Unavailable(
            "Grammar downloads are not available in the web build".to_owned(),
        );
        self.status_message = "Grammar download unavailable on web".to_owned();
    }

    #[cfg(not(target_arch = "wasm32"))]
    fn poll_grammar_download(&mut self) {
        let mut drained = Vec::new();
        if let Some(rx) = self.grammar_download_rx.as_mut() {
            while let Ok(message) = rx.try_recv() {
                drained.push(message);
            }
        }
        if drained.is_empty() {
            return;
        }

        for message in drained {
            match message {
                GrammarDownloadResult::Ready(path) => {
                    self.grammar_download_status = GrammarDownloadStatus::Idle;
                    self.grammar_download_rx = None;
                    self.grammar_warning_message =
                        Some(format!("LanguageTool downloaded to {}", path.display()));
                    self.show_grammar_warning = false;
                    self.status_message = "LanguageTool downloaded".to_owned();
                    if let Err(error_message) = self.start_grammar_service() {
                        self.grammar_status = GrammarStatus::Unavailable(error_message);
                        self.show_grammar_warning = true;
                    } else {
                        self.grammar_status = GrammarStatus::Idle;
                        self.request_grammar_check(true);
                    }
                }
                GrammarDownloadResult::Failed(error_message) => {
                    self.grammar_download_status = GrammarDownloadStatus::Idle;
                    self.grammar_download_rx = None;
                    self.grammar_status = GrammarStatus::Unavailable(format!(
                        "LanguageTool download failed: {error_message}"
                    ));
                    self.show_grammar_warning = true;
                    self.grammar_warning_message =
                        Some(format!("LanguageTool download failed: {error_message}"));
                }
            }
        }
    }

    #[cfg(target_arch = "wasm32")]
    fn poll_grammar_download(&mut self) {}

    #[cfg(not(target_arch = "wasm32"))]
    fn request_grammar_check(&mut self, force: bool) {
        if !force && !self.grammar_auto_check {
            return;
        }
        if self.grammar_download_status == GrammarDownloadStatus::Downloading {
            self.status_message = "Grammar download in progress".to_owned();
            return;
        }

        if self.grammar_tx.is_none() {
            if let Err(message) = self.start_grammar_service() {
                self.grammar_status = GrammarStatus::Unavailable(message.clone());
                self.grammar_warning_message = Some(message);
                self.show_grammar_warning = true;
                return;
            }
        }

        let text = self.document.plain_text();
        let language = self
            .grammar_config
            .language
            .to_languagetool_code(&text)
            .to_owned();
        let request = GrammarRequest { text, language };

        let Some(tx) = self.grammar_tx.clone() else {
            return;
        };

        match tx.try_send(request.clone()) {
            Ok(()) => {
                self.grammar_status = GrammarStatus::Checking;
            }
            Err(mpsc::error::TrySendError::Full(_)) => {
                self.grammar_status = GrammarStatus::Checking;
            }
            Err(mpsc::error::TrySendError::Closed(_)) => {
                if let Err(message) = self.start_grammar_service() {
                    self.grammar_status = GrammarStatus::Unavailable(message.clone());
                    self.grammar_warning_message = Some(message);
                    self.show_grammar_warning = true;
                    return;
                }

                if let Some(restarted_tx) = self.grammar_tx.clone() {
                    match restarted_tx.try_send(request) {
                        Ok(()) | Err(mpsc::error::TrySendError::Full(_)) => {
                            self.grammar_status = GrammarStatus::Checking;
                        }
                        Err(mpsc::error::TrySendError::Closed(_)) => {
                            self.grammar_status = GrammarStatus::Unavailable(
                                "Grammar worker channel closed unexpectedly".to_owned(),
                            );
                        }
                    }
                }
            }
        }
    }

    #[cfg(target_arch = "wasm32")]
    fn request_grammar_check(&mut self, _force: bool) {}
}

impl Drop for WorsApp {
    fn drop(&mut self) {
        self.stop_grammar_service();
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

impl App for WorsApp {
    fn ui(&mut self, ui: &mut egui::Ui, _frame: &mut Frame) {
        self.poll_grammar_results();
        self.poll_grammar_download();

        let shortcut_changed = handle_global_shortcuts(
            ui,
            &mut self.document,
            &mut self.canvas,
            &mut self.history,
            &mut self.current_path,
            &mut self.status_message,
        );

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
            let mut backstage_output = chrome::BackstageOutput::default();
            egui::CentralPanel::default()
                .frame(egui::Frame::new().fill(palette.workspace_bg))
                .show_inside(ui, |ui| {
                    backstage_output = paint_backstage(
                        ui,
                        &mut self.backstage,
                        &self.document,
                        &self.current_path,
                        &self.recent_files,
                        palette,
                    );
                });

            if backstage_output.close_requested {
                self.backstage.visible = false;
            }
            if backstage_output.save_requested {
                if let Some(path) = save_document(
                    &self.document,
                    &mut self.status_message,
                    &mut self.current_path,
                ) {
                    self.remember_recent_file(path);
                }
            }
            if backstage_output.save_as_requested {
                if let Some(path) = save_document_as_with_name(
                    &self.document,
                    &mut self.status_message,
                    &mut self.current_path,
                    &self.backstage.file_name,
                    self.backstage.format.extension(),
                ) {
                    self.remember_recent_file(path);
                }
            }
            if backstage_output.open_requested {
                if let Some(path) = open_document(
                    &mut self.document,
                    &mut self.canvas,
                    &mut self.status_message,
                    &mut self.current_path,
                    &mut self.history,
                ) {
                    self.remember_recent_file(path);
                }
                self.backstage
                    .open_save_as(&self.document, &self.current_path);
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
                        &mut self.grammar_config,
                        &self.grammar_status,
                        &mut self.grammar_auto_check,
                        grammar_download_available,
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

        if let Some(path) = self.current_path.clone() {
            self.remember_recent_file(path);
        }

        // Auto-switch to contextual tabs when an object is selected.
        match (
            self.canvas.selected_image_id,
            self.canvas.active_table_cell,
            self.active_tab,
        ) {
            (Some(_), _, tab) if tab != RibbonTab::Picture => {
                self.active_tab = RibbonTab::Picture;
            }
            (None, Some(_), tab) if tab != RibbonTab::Table => {
                self.active_tab = RibbonTab::Table;
            }
            (None, None, RibbonTab::Picture | RibbonTab::Table) => {
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
                    palette,
                );
            });

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
    }
}
