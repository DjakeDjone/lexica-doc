pub(crate) mod header_layout;
pub(crate) mod ribbon;
pub(crate) mod status_bar;
pub(crate) mod tab_row;
pub(crate) mod title_bar;

use crate::app::CanvasState;
use crate::document::DocumentState;

pub(super) use ribbon::paint_ribbon;
pub(super) use ribbon::GrammarRibbonOutput;
pub(super) use status_bar::paint_status_bar;
pub(super) use tab_row::paint_tab_row;
pub(super) use title_bar::paint_title_bar;

#[derive(Clone, Copy, PartialEq, Eq)]
pub(super) enum RibbonTab {
    Home,
    Insert,
    Design,
    Layout,
    View,
    Grammar,
    HeaderFooter,
    Picture,
    Table,
}

impl RibbonTab {
    pub(crate) const ALL: [Self; 6] = [
        Self::Home,
        Self::Insert,
        Self::Design,
        Self::Layout,
        Self::View,
        Self::Grammar,
    ];

    pub(crate) const fn label(self) -> &'static str {
        match self {
            Self::Home => "Home",
            Self::Insert => "Insert",
            Self::Design => "Design",
            Self::Layout => "Layout",
            Self::View => "View",
            Self::Grammar => "Grammar",
            Self::HeaderFooter => "Header & Footer",
            Self::Picture => "Picture Format",
            Self::Table => "Table Format",
        }
    }
}

pub(crate) fn current_section_id(
    document: &DocumentState,
    canvas: &CanvasState,
) -> crate::document::SectionId {
    if let Some(active) = canvas.active_header_footer {
        return active.section_id;
    }
    let paragraph_index = document
        .paragraphs()
        .iter()
        .position(|paragraph| {
            paragraph.range.contains(&canvas.selection.primary.index)
                || paragraph.range.start == canvas.selection.primary.index
        })
        .unwrap_or(0);
    document.section_at_paragraph(paragraph_index).id
}

#[cfg(test)]
fn layout_tab_command_labels() -> &'static [&'static str] {
    &[
        "Margins",
        "Size",
        "Orientation",
        "Columns",
        "Breaks",
        "Line Numbers",
        "Header",
        "Footer",
        "Page #",
        "Page Setup",
    ]
}

#[cfg(test)]
fn layout_tab_removed_labels() -> &'static [&'static str] {
    &[
        "Zoom",
        "Dark",
        "Header from Top",
        "Footer from Bottom",
        "Remove Header",
        "Remove Footer",
        "Close Header and Footer",
    ]
}

#[cfg(test)]
fn header_footer_contextual_tab_visible(canvas: &CanvasState) -> bool {
    canvas.active_header_footer.is_some()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::app::ActiveHeaderFooter;
    use crate::document::{HeaderFooterKind, HeaderFooterVariant};

    #[test]
    fn layout_tab_contract_lists_only_page_layout_commands() {
        assert_eq!(
            layout_tab_command_labels(),
            &[
                "Margins",
                "Size",
                "Orientation",
                "Columns",
                "Breaks",
                "Line Numbers",
                "Header",
                "Footer",
                "Page #",
                "Page Setup",
            ]
        );
        assert!(!layout_tab_command_labels()
            .iter()
            .any(|label| layout_tab_removed_labels().contains(label)));
    }

    #[test]
    fn header_footer_tab_visibility_tracks_editing_state() {
        let mut canvas = CanvasState::default();
        assert!(!header_footer_contextual_tab_visible(&canvas));

        canvas.active_header_footer = Some(ActiveHeaderFooter {
            kind: HeaderFooterKind::Header,
            section_id: 1,
            variant: HeaderFooterVariant::Default,
            page_number: 1,
        });
        assert!(header_footer_contextual_tab_visible(&canvas));
    }

    #[cfg(not(target_arch = "wasm32"))]
    #[test]
    fn unix_epoch_formats_as_civil_date() {
        assert_eq!(status_bar::civil_from_days(0), (1970, 1, 1));
    }
}
