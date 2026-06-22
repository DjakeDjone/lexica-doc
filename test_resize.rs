use eframe::egui;

fn handle_window_resize(ctx: &egui::Context) {
    let rect = ctx.screen_rect();
    let resize_margin = 6.0;

    let pointer = ctx.pointer_hover_pos();
    if let Some(pos) = pointer {
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
