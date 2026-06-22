use eframe::egui;

pub(crate) const TITLE_BAR_HEIGHT: f32 = 28.0;
pub(crate) const RIBBON_HEIGHT: f32 = 88.0;
pub(crate) const TAB_ROW_HEIGHT: f32 = 28.0;

pub(crate) fn title_action_width(total_width: f32) -> f32 {
    if total_width >= 760.0 {
        340.0
    } else if total_width >= 560.0 {
        260.0
    } else if total_width >= 360.0 {
        96.0
    } else {
        0.0
    }
}

pub(crate) fn clipped_child_ui(
    ui: &mut egui::Ui,
    id_salt: &'static str,
    rect: egui::Rect,
    layout: egui::Layout,
) -> egui::Ui {
    let mut child = ui.new_child(
        egui::UiBuilder::new()
            .id_salt(id_salt)
            .max_rect(rect)
            .layout(layout),
    );
    child.set_clip_rect(rect);
    child
}
