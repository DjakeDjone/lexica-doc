use std::path::{Path, PathBuf};
#[cfg(not(target_arch = "wasm32"))]
use std::{
    fs,
    time::{Duration, SystemTime},
};

use eframe::egui;

use crate::app::palette::ThemePalette;
use super::BackstageOutput;

#[cfg(not(target_arch = "wasm32"))]
pub(crate) struct FolderEntry {
    pub name: String,
    pub path: PathBuf,
    pub is_dir: bool,
    pub modified: String,
}

pub fn paint_recent_files(
    ui: &mut egui::Ui,
    recent_files: &[PathBuf],
    output: &mut BackstageOutput,
    width: f32,
    palette: ThemePalette,
) {
    folder_header(ui, width, palette);
    if recent_files.is_empty() {
        ui.add_space(12.0);
        ui.label(
            egui::RichText::new("No recent files yet")
                .size(12.0)
                .color(palette.text_muted),
        );
        return;
    }

    let width = width.min(ui.available_width()).max(360.0);
    egui::ScrollArea::vertical()
        .auto_shrink([false, false])
        .max_height((ui.available_height() - 8.0).max(120.0))
        .show(ui, |ui| {
            for path in recent_files {
                let name = path
                    .file_name()
                    .and_then(|name| name.to_str())
                    .unwrap_or("document");
                let detail = path.parent().map_or_else(
                    || path.display().to_string(),
                    |parent| parent.display().to_string(),
                );
                if recent_file_row(ui, name, &detail, width, palette).clicked() {
                    output.recent_open_requested = Some(path.clone());
                }
            }
        });
}

pub fn folder_header(ui: &mut egui::Ui, width: f32, palette: ThemePalette) {
    let width = width.min(ui.available_width()).max(360.0);
    let (rect, _) = ui.allocate_exact_size(egui::vec2(width, 32.0), egui::Sense::hover());
    let date_x = (rect.right() - 220.0).max(rect.left() + 280.0);
    ui.painter().text(
        rect.left_center() + egui::vec2(34.0, 0.0),
        egui::Align2::LEFT_CENTER,
        "Name",
        egui::FontId::proportional(12.0),
        palette.text_muted,
    );
    ui.painter().text(
        egui::pos2(date_x, rect.center().y),
        egui::Align2::LEFT_CENTER,
        "Date Modified",
        egui::FontId::proportional(12.0),
        palette.text_muted,
    );
    ui.painter().line_segment(
        [rect.left_bottom(), rect.right_bottom()],
        egui::Stroke::new(1.0, palette.border),
    );
}

#[cfg(not(target_arch = "wasm32"))]
pub fn folder_row(
    ui: &mut egui::Ui,
    name: &str,
    detail: &str,
    is_dir: bool,
    width: f32,
    palette: ThemePalette,
) -> egui::Response {
    let width = width.min(ui.available_width()).max(360.0);
    let (rect, response) = ui.allocate_exact_size(egui::vec2(width, 32.0), egui::Sense::click());
    let fill = if response.hovered() {
        palette.accent.gamma_multiply(0.08)
    } else {
        egui::Color32::TRANSPARENT
    };
    ui.painter().rect_filled(rect, 0.0, fill);

    let icon_rect = egui::Rect::from_min_size(
        rect.left_center() + egui::vec2(10.0, -7.0),
        egui::vec2(16.0, 14.0),
    );
    if is_dir {
        paint_folder_icon(ui.painter(), icon_rect, palette);
    } else {
        paint_file_icon(ui.painter(), icon_rect, palette);
    }

    let date_x = (rect.right() - 220.0).max(rect.left() + 280.0);
    ui.painter().text(
        rect.left_center() + egui::vec2(34.0, 0.0),
        egui::Align2::LEFT_CENTER,
        name,
        egui::FontId::proportional(12.5),
        palette.text_primary,
    );
    ui.painter().text(
        egui::pos2(date_x, rect.center().y),
        egui::Align2::LEFT_CENTER,
        detail,
        egui::FontId::proportional(12.0),
        palette.text_muted,
    );
    response
}

#[cfg(not(target_arch = "wasm32"))]
pub fn paint_folder_icon(painter: &egui::Painter, rect: egui::Rect, palette: ThemePalette) {
    let stroke = egui::Stroke::new(1.2, palette.text_muted);
    let tab = egui::Rect::from_min_size(rect.min + egui::vec2(1.0, 0.0), egui::vec2(7.0, 4.0));
    let body = egui::Rect::from_min_max(
        rect.min + egui::vec2(1.0, 3.0),
        rect.max - egui::vec2(1.0, 1.0),
    );
    painter.rect_stroke(tab, 1.0, stroke, egui::StrokeKind::Inside);
    painter.rect_stroke(body, 1.0, stroke, egui::StrokeKind::Inside);
}

pub fn paint_file_icon(painter: &egui::Painter, rect: egui::Rect, palette: ThemePalette) {
    let stroke = egui::Stroke::new(1.2, palette.text_muted);
    let page = egui::Rect::from_min_max(
        rect.min + egui::vec2(3.0, 1.0),
        rect.max - egui::vec2(3.0, 1.0),
    );
    painter.rect_stroke(page, 1.0, stroke, egui::StrokeKind::Inside);
    painter.line_segment(
        [
            page.left_top() + egui::vec2(3.0, 5.0),
            page.right_top() + egui::vec2(-3.0, 5.0),
        ],
        stroke,
    );
}

fn recent_file_row(
    ui: &mut egui::Ui,
    name: &str,
    detail: &str,
    width: f32,
    palette: ThemePalette,
) -> egui::Response {
    let width = width.min(ui.available_width()).max(360.0);
    let (rect, response) = ui.allocate_exact_size(egui::vec2(width, 42.0), egui::Sense::click());
    let fill = if response.hovered() {
        palette.accent.gamma_multiply(0.08)
    } else {
        egui::Color32::TRANSPARENT
    };
    ui.painter().rect_filled(rect, 0.0, fill);

    let icon_rect = egui::Rect::from_min_size(
        rect.left_center() + egui::vec2(10.0, -7.0),
        egui::vec2(16.0, 14.0),
    );
    paint_file_icon(ui.painter(), icon_rect, palette);

    ui.painter().text(
        rect.left_top() + egui::vec2(34.0, 7.0),
        egui::Align2::LEFT_TOP,
        name,
        egui::FontId::proportional(12.5),
        palette.text_primary,
    );
    ui.painter().text(
        rect.left_top() + egui::vec2(34.0, 24.0),
        egui::Align2::LEFT_TOP,
        detail,
        egui::FontId::proportional(10.5),
        palette.text_muted,
    );
    response
}

#[cfg(not(target_arch = "wasm32"))]
pub fn folder_entries(dir: &Path) -> Vec<FolderEntry> {
    let Ok(read_dir) = fs::read_dir(dir) else {
        return Vec::new();
    };
    let mut entries: Vec<_> = read_dir
        .filter_map(Result::ok)
        .filter_map(|entry| {
            let metadata = entry.metadata().ok()?;
            let name = entry.file_name().to_string_lossy().into_owned();
            if name.starts_with('.') {
                return None;
            }
            let modified = metadata
                .modified()
                .ok()
                .map(modified_label)
                .unwrap_or_else(|| "Unknown".to_owned());
            Some(FolderEntry {
                name,
                path: entry.path(),
                is_dir: metadata.is_dir(),
                modified,
            })
        })
        .collect();
    entries.sort_by(|left, right| {
        right
            .is_dir
            .cmp(&left.is_dir)
            .then_with(|| left.name.to_lowercase().cmp(&right.name.to_lowercase()))
    });
    entries
}

#[cfg(not(target_arch = "wasm32"))]
fn modified_label(modified: SystemTime) -> String {
    match SystemTime::now().duration_since(modified) {
        Ok(elapsed) if elapsed < Duration::from_secs(60) => "Just now".to_owned(),
        Ok(elapsed) if elapsed < Duration::from_secs(60 * 60) => {
            format!("{} min ago", elapsed.as_secs() / 60)
        }
        Ok(elapsed) if elapsed < Duration::from_secs(60 * 60 * 24) => {
            format!("{} hours ago", elapsed.as_secs() / 3600)
        }
        Ok(elapsed) => format!("{} days ago", elapsed.as_secs() / 86_400),
        Err(_) => "In the future".to_owned(),
    }
}
