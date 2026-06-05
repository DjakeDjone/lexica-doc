pub(crate) mod rendering;
pub(crate) mod editor;

pub(super) use rendering::{header_footer_hit, paint_page_header_footer};
pub(super) use editor::paint_active_header_footer_editor;
