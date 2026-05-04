use eframe::Storage;
use serde::{Deserialize, Serialize};

use super::{CanvasState, ThemeMode, ZoomMode};

const SETTINGS_KEY: &str = "wors-app-settings";

#[derive(Clone, Copy, Debug, PartialEq)]
pub(super) struct AppSettings {
    pub(super) theme_mode: ThemeMode,
    pub(super) zoom: f32,
    pub(super) zoom_mode: ZoomMode,
}

#[derive(Clone, Copy, Debug, Deserialize, Serialize)]
struct StoredAppSettings {
    theme_mode: StoredThemeMode,
    zoom: f32,
    zoom_mode: StoredZoomMode,
}

#[derive(Clone, Copy, Debug, Deserialize, Serialize)]
enum StoredThemeMode {
    Light,
    Dark,
}

#[derive(Clone, Copy, Debug, Deserialize, Serialize)]
enum StoredZoomMode {
    Manual,
    FitPage,
}

impl AppSettings {
    pub(super) fn from_state(theme_mode: ThemeMode, canvas: &CanvasState) -> Self {
        Self {
            theme_mode,
            zoom: canvas.zoom.clamp(0.5, 3.0),
            zoom_mode: canvas.zoom_mode,
        }
    }

    pub(super) fn apply(self, theme_mode: &mut ThemeMode, canvas: &mut CanvasState) {
        *theme_mode = self.theme_mode;
        canvas.zoom = self.zoom.clamp(0.5, 3.0);
        canvas.zoom_mode = self.zoom_mode;
    }

    fn stored(self) -> StoredAppSettings {
        StoredAppSettings {
            theme_mode: self.theme_mode.into(),
            zoom: self.zoom.clamp(0.5, 3.0),
            zoom_mode: self.zoom_mode.into(),
        }
    }
}

pub(super) fn load_settings(storage: Option<&dyn Storage>) -> Option<AppSettings> {
    let source = storage?.get_string(SETTINGS_KEY)?;
    serde_json::from_str::<StoredAppSettings>(&source)
        .ok()
        .map(AppSettings::from)
}

pub(super) fn save_settings(storage: &mut dyn Storage, settings: AppSettings) {
    if let Ok(json) = serde_json::to_string(&settings.stored()) {
        storage.set_string(SETTINGS_KEY, json);
        storage.flush();
    }
}

impl From<StoredAppSettings> for AppSettings {
    fn from(settings: StoredAppSettings) -> Self {
        Self {
            theme_mode: settings.theme_mode.into(),
            zoom: settings.zoom.clamp(0.5, 3.0),
            zoom_mode: settings.zoom_mode.into(),
        }
    }
}

impl From<ThemeMode> for StoredThemeMode {
    fn from(mode: ThemeMode) -> Self {
        match mode {
            ThemeMode::Light => Self::Light,
            ThemeMode::Dark => Self::Dark,
        }
    }
}

impl From<StoredThemeMode> for ThemeMode {
    fn from(mode: StoredThemeMode) -> Self {
        match mode {
            StoredThemeMode::Light => Self::Light,
            StoredThemeMode::Dark => Self::Dark,
        }
    }
}

impl From<ZoomMode> for StoredZoomMode {
    fn from(mode: ZoomMode) -> Self {
        match mode {
            ZoomMode::Manual => Self::Manual,
            ZoomMode::FitPage => Self::FitPage,
        }
    }
}

impl From<StoredZoomMode> for ZoomMode {
    fn from(mode: StoredZoomMode) -> Self {
        match mode {
            StoredZoomMode::Manual => Self::Manual,
            StoredZoomMode::FitPage => Self::FitPage,
        }
    }
}
