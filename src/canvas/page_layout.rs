use eframe::egui::{self, epaint::text::cursor::CCursor, Rect};

use crate::{
    app::CanvasState,
    document::DocumentState,
    layout::{centered_page_rect, document_points_to_screen_points, page_content_rect},
};

pub(super) struct PageSlice {
    pub(super) page_rect: Rect,
    pub(super) content_rect: Rect,
    pub(super) start_y: f32,
    pub(super) end_y: f32,
}

pub(super) struct PageLayout {
    pub(super) pages: Vec<PageSlice>,
}

impl PageLayout {
    pub(super) fn document_pos(&self, pointer_pos: egui::Pos2) -> Option<egui::Vec2> {
        self.pages.iter().find_map(|page| {
            if page.content_rect.contains(pointer_pos) {
                let local_y = pointer_pos.y - page.content_rect.top();
                Some(egui::vec2(
                    pointer_pos.x - page.content_rect.left(),
                    (page.start_y + local_y).clamp(page.start_y, page.end_y),
                ))
            } else {
                None
            }
        })
    }

    pub(super) fn caret_rect(
        &self,
        galley: &egui::Galley,
        cursor: CCursor,
        height: f32,
    ) -> Option<Rect> {
        let document_rect = caret_rect(galley, cursor, height);
        self.pages.iter().find_map(|page| {
            if document_rect.center().y >= page.start_y && document_rect.center().y <= page.end_y {
                Some(
                    document_rect
                        .translate(page.content_rect.min.to_vec2() - egui::vec2(0.0, page.start_y)),
                )
            } else {
                None
            }
        })
    }
}

pub(super) fn layout_page_stack(
    viewport: Rect,
    document: &DocumentState,
    canvas: &CanvasState,
    galley: &egui::Galley,
    manual_page_break_rows: &[usize],
) -> PageLayout {
    let page_gap = document_points_to_screen_points(24.0, canvas.zoom);
    let base_page_rect =
        centered_page_rect(viewport, document.page_size, canvas.zoom, egui::Vec2::ZERO);
    let page_size = base_page_rect.size();
    let content_height = page_content_rect(base_page_rect, document.margins, canvas.zoom).height();
    let page_ranges = compute_page_ranges(galley, content_height, manual_page_break_rows);
    let page_count = page_ranges.len().max(1);
    let stack_height =
        page_count as f32 * page_size.y + (page_count.saturating_sub(1) as f32 * page_gap);

    let top = if stack_height < viewport.height() {
        viewport.center().y - stack_height * 0.5 + canvas.pan.y
    } else {
        viewport.top() + document_points_to_screen_points(24.0, canvas.zoom) + canvas.pan.y
    };
    let left = viewport.center().x - page_size.x * 0.5 + canvas.pan.x;

    let mut pages = Vec::with_capacity(page_count);
    for (index, (start_y, end_y)) in page_ranges.into_iter().enumerate() {
        let min = egui::pos2(left, top + index as f32 * (page_size.y + page_gap));
        let page_rect = Rect::from_min_size(min, page_size);
        let content_rect = page_content_rect(page_rect, document.margins, canvas.zoom);
        pages.push(PageSlice {
            page_rect,
            content_rect,
            start_y,
            end_y,
        });
    }

    PageLayout { pages }
}

fn caret_rect(galley: &egui::Galley, cursor: CCursor, height: f32) -> Rect {
    let layout_cursor = galley.layout_from_cursor(cursor);
    let mut rect = galley.pos_from_cursor(cursor);
    if let Some(row) = galley.rows.get(layout_cursor.row) {
        let row_min = row.min_y();
        let row_max = row.max_y();
        let height = height.clamp(1.0, row_max - row_min);
        rect.max.y = row_max;
        rect.min.y = (row_max - height).max(row_min);
    }
    rect.expand2(egui::vec2(0.75, 0.75))
}

fn compute_page_ranges(
    galley: &egui::Galley,
    page_height: f32,
    manual_page_break_rows: &[usize],
) -> Vec<(f32, f32)> {
    if galley.rows.is_empty() {
        return vec![(0.0, page_height)];
    }

    let mut pages = Vec::new();
    let mut page_start: f32 = 0.0;
    let mut last_row_end: f32 = 0.0;
    let mut break_rows = manual_page_break_rows.iter().copied().peekable();

    for (row_index, row) in galley.rows.iter().enumerate() {
        let row_start = row.pos.y;
        let row_end = row.pos.y + row.row.height();

        while break_rows
            .peek()
            .copied()
            .is_some_and(|break_row| break_row == row_index)
        {
            if row_start > page_start {
                pages.push((page_start, last_row_end.max(page_start)));
            } else if pages.is_empty() {
                pages.push((page_start, page_start));
            }
            page_start = row_start;
            break_rows.next();
        }

        if row_end - page_start > page_height && row_start > page_start {
            pages.push((page_start, last_row_end.max(page_start)));
            page_start = row_start;
        }

        last_row_end = row_end;
    }

    pages.push((page_start, last_row_end.max(page_start)));
    pages
}
