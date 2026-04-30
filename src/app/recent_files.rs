use std::path::PathBuf;

#[cfg(not(target_arch = "wasm32"))]
use std::{env, fs};

const RECENT_FILES_LIMIT: usize = 12;

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
pub fn normalize_recent_path(path: PathBuf) -> PathBuf {
    path.canonicalize().unwrap_or(path)
}

#[cfg(target_arch = "wasm32")]
pub fn normalize_recent_path(path: PathBuf) -> PathBuf {
    path
}

#[cfg(not(target_arch = "wasm32"))]
pub fn load_recent_files() -> Vec<PathBuf> {
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
pub fn load_recent_files() -> Vec<PathBuf> {
    Vec::new()
}

#[cfg(not(target_arch = "wasm32"))]
pub fn save_recent_files(recent_files: &[PathBuf]) {
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
pub fn save_recent_files(_recent_files: &[PathBuf]) {}

pub fn remember_recent_file(recent_files: &mut Vec<PathBuf>, path: PathBuf) {
    if path.as_os_str().is_empty() {
        return;
    }
    let path = normalize_recent_path(path);
    if recent_files.first() == Some(&path) {
        return;
    }
    recent_files.retain(|recent| recent != &path);
    recent_files.insert(0, path);
    recent_files.truncate(RECENT_FILES_LIMIT);
    save_recent_files(recent_files);
}
