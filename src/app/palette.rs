use eframe::egui;

#[derive(Clone, Copy, PartialEq, Eq)]
pub enum ThemeMode {
    Light,
    Dark,
}

impl ThemeMode {
    pub const fn label(self) -> &'static str {
        match self {
            Self::Light => "Light",
            Self::Dark => "Dark",
        }
    }
}

#[derive(Clone, Copy)]
pub(super) struct ThemePalette {
    pub(super) title_bg: egui::Color32,
    pub(super) title_fg: egui::Color32,
    pub(super) title_muted: egui::Color32,
    pub(super) tab_bg: egui::Color32,
    pub(super) tab_fg: egui::Color32,
    pub(super) tab_active_bg: egui::Color32,
    pub(super) tab_active_fg: egui::Color32,
    pub(super) ribbon_bg: egui::Color32,
    pub(super) ribbon_group_bg: egui::Color32,
    pub(super) border: egui::Color32,
    pub(super) text_primary: egui::Color32,
    pub(super) text_muted: egui::Color32,
    pub(super) workspace_bg: egui::Color32,
    pub(super) status_bg: egui::Color32,
    pub(super) accent: egui::Color32,
}

pub(super) fn theme_palette(mode: ThemeMode) -> ThemePalette {
    match mode {
        ThemeMode::Light => ThemePalette {
            title_bg: egui::Color32::from_rgb(43, 87, 154),
            title_fg: egui::Color32::from_rgb(247, 250, 255),
            title_muted: egui::Color32::from_rgb(214, 227, 247),
            tab_bg: egui::Color32::from_rgb(43, 87, 154),
            tab_fg: egui::Color32::from_rgb(239, 246, 255),
            tab_active_bg: egui::Color32::from_rgb(245, 248, 252),
            tab_active_fg: egui::Color32::from_rgb(31, 64, 115),
            ribbon_bg: egui::Color32::from_rgb(244, 246, 249),
            ribbon_group_bg: egui::Color32::from_rgb(251, 252, 254),
            border: egui::Color32::from_rgb(202, 210, 224),
            text_primary: egui::Color32::from_rgb(30, 34, 40),
            text_muted: egui::Color32::from_rgb(94, 101, 114),
            workspace_bg: egui::Color32::from_rgb(215, 217, 220),
            status_bg: egui::Color32::from_rgb(235, 238, 243),
            accent: egui::Color32::from_rgb(54, 109, 193),
        },
        ThemeMode::Dark => ThemePalette {
            title_bg: egui::Color32::from_rgb(28, 34, 47),
            title_fg: egui::Color32::from_rgb(236, 241, 251),
            title_muted: egui::Color32::from_rgb(156, 170, 197),
            tab_bg: egui::Color32::from_rgb(28, 34, 47),
            tab_fg: egui::Color32::from_rgb(213, 222, 240),
            tab_active_bg: egui::Color32::from_rgb(65, 79, 105),
            tab_active_fg: egui::Color32::from_rgb(241, 247, 255),
            ribbon_bg: egui::Color32::from_rgb(49, 55, 66),
            ribbon_group_bg: egui::Color32::from_rgb(57, 64, 77),
            border: egui::Color32::from_rgb(84, 94, 112),
            text_primary: egui::Color32::from_rgb(233, 238, 248),
            text_muted: egui::Color32::from_rgb(172, 181, 197),
            workspace_bg: egui::Color32::from_rgb(50, 52, 56),
            status_bg: egui::Color32::from_rgb(43, 49, 59),
            accent: egui::Color32::from_rgb(109, 157, 228),
        },
    }
}

pub(super) fn configure_theme(ctx: &egui::Context, mode: ThemeMode, palette: ThemePalette) {
    let mut style = (*ctx.global_style()).clone();
    style.spacing.item_spacing = egui::vec2(8.0, 8.0);
    style.spacing.button_padding = egui::vec2(8.0, 5.0);
    style.spacing.combo_width = 130.0;
    style.visuals = match mode {
        ThemeMode::Light => egui::Visuals::light(),
        ThemeMode::Dark => egui::Visuals::dark(),
    };
    style.visuals.interact_cursor = Some(egui::CursorIcon::PointingHand);
    style.visuals.override_text_color = Some(palette.text_primary);
    style.visuals.widgets.inactive.bg_fill = palette.ribbon_group_bg;
    style.visuals.widgets.inactive.weak_bg_fill = palette.status_bg;
    style.visuals.widgets.inactive.bg_stroke = egui::Stroke::new(1.0, palette.border);
    style.visuals.widgets.inactive.fg_stroke = egui::Stroke::new(1.0, palette.text_primary);
    style.visuals.widgets.hovered.bg_fill = palette.ribbon_group_bg;
    style.visuals.widgets.hovered.weak_bg_fill = palette.tab_active_bg;
    style.visuals.widgets.hovered.bg_stroke = egui::Stroke::new(1.0, palette.accent);
    style.visuals.widgets.hovered.fg_stroke = egui::Stroke::new(1.0, palette.text_primary);
    style.visuals.widgets.active.bg_fill = palette.tab_active_bg;
    style.visuals.widgets.active.weak_bg_fill = palette.tab_active_bg;
    style.visuals.widgets.active.bg_stroke = egui::Stroke::new(1.0, palette.accent);
    style.visuals.widgets.active.fg_stroke = egui::Stroke::new(1.0, palette.text_primary);
    style.visuals.widgets.open = style.visuals.widgets.active;
    style.visuals.selection.bg_fill = palette.accent.gamma_multiply(0.35);
    style.visuals.selection.stroke = egui::Stroke::new(1.0, palette.accent);
    style.visuals.text_cursor.stroke = egui::Stroke::new(2.0, egui::Color32::BLACK);
    style.visuals.panel_fill = palette.ribbon_bg;
    style.visuals.window_fill = palette.ribbon_group_bg;
    ctx.set_global_style(style);
}

pub(super) fn theme_switch(
    ui: &mut egui::Ui,
    theme_mode: &mut ThemeMode,
    palette: ThemePalette,
    dark_surface: bool,
) -> bool {
    let original = *theme_mode;
    let _ = dark_surface;
    let switch_bg = palette.ribbon_group_bg;
    let inactive_text = palette.text_primary;
    let active_fill = palette.status_bg;
    let active_text = palette.text_primary;

    egui::Frame::new()
        .fill(switch_bg)
        .inner_margin(egui::Margin::symmetric(2, 2))
        .stroke(egui::Stroke::new(1.0, palette.border))
        .corner_radius(8.0)
        .show(ui, |ui| {
            ui.horizontal(|ui| {
                ui.spacing_mut().item_spacing.x = 2.0;
                for (mode, icon) in [(ThemeMode::Light, "☀"), (ThemeMode::Dark, "🌙")] {
                    let selected = *theme_mode == mode;
                    let button = egui::Button::new(
                        egui::RichText::new(icon).size(13.0).color(if selected {
                            active_text
                        } else {
                            inactive_text
                        }),
                    )
                    .min_size(egui::vec2(34.0, 22.0))
                    .fill(if selected {
                        active_fill
                    } else {
                        egui::Color32::TRANSPARENT
                    })
                    .stroke(if selected {
                        egui::Stroke::new(1.0, palette.accent)
                    } else {
                        egui::Stroke::NONE
                    })
                    .corner_radius(6.0);
                    let response = ui.add(button).on_hover_text(mode.label());
                    if response.clicked() {
                        *theme_mode = mode;
                    }
                }
            });
        });
    *theme_mode != original
}
