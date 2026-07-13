use eframe::egui::{self, epaint::text::cursor::CCursor, Rect};

use crate::{
    app::CanvasState,
    document::{DocumentState, SectionId},
    layout::{centered_page_rect, document_points_to_screen_points, section_page_content_rect},
};

use super::layout::TableLayout;

pub(crate) struct PageSlice {
    pub(super) page_rect: Rect,
    pub(super) content_rect: Rect,
    pub(super) section_id: SectionId,
    pub(super) page_index_within_section: usize,
    pub(super) section_page_count: usize,
    pub(super) start_y: f32,
    pub(super) end_y: f32,
}

pub(crate) struct PageLayout {
    pub(super) pages: Vec<PageSlice>,
}

impl PageLayout {
    pub(super) fn current_page(&self, galley: &egui::Galley, cursor: CCursor) -> usize {
        let y = galley.pos_from_cursor(cursor).center().y;
        self.pages
            .iter()
            .position(|page| y >= page.start_y && y <= page.end_y)
            .map_or(1, |index| index + 1)
    }

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

    pub(super) fn clamped_document_pos(&self, pointer_pos: egui::Pos2) -> Option<egui::Vec2> {
        self.pages.iter().find_map(|page| {
            page.page_rect.contains(pointer_pos).then(|| {
                egui::vec2(
                    (pointer_pos.x - page.content_rect.left())
                        .clamp(0.0, page.content_rect.width()),
                    (page.start_y
                        + (pointer_pos.y - page.content_rect.top())
                            .clamp(0.0, page.content_rect.height()))
                    .clamp(page.start_y, page.end_y),
                )
            })
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
    canvas: &mut CanvasState,
    galley: &egui::Galley,
    manual_page_break_rows: &[usize],
    paragraph_start_rows: &[usize],
    tables: &[TableLayout],
) -> PageLayout {
    let page_gap = document_points_to_screen_points(24.0, canvas.zoom);
    let page_setup = document.default_page_setup();
    let base_page_rect = centered_page_rect(
        viewport,
        page_setup.page_size,
        canvas.zoom,
        egui::Vec2::ZERO,
    );
    let page_size = base_page_rect.size();
    let content_height =
        section_page_content_rect(base_page_rect, page_setup, 14.0, 14.0, canvas.zoom).height();
    let page_ranges = compute_page_ranges(galley, content_height, manual_page_break_rows, tables);
    let page_count = page_ranges.len().max(1);
    let stack_height =
        page_count as f32 * page_size.y + (page_count.saturating_sub(1) as f32 * page_gap);

    let pan_before_clamp = canvas.pan;
    canvas.scroll_range = clamp_pan(
        viewport,
        page_size.x,
        stack_height,
        page_gap,
        &mut canvas.pan,
    );
    for axis in 0..2 {
        if canvas.pan[axis] != pan_before_clamp[axis] {
            canvas.scroll_velocity[axis] = 0.0;
        }
    }

    let top = if stack_height < viewport.height() {
        viewport.center().y - stack_height * 0.5 + canvas.pan.y
    } else {
        viewport.top() + document_points_to_screen_points(24.0, canvas.zoom) + canvas.pan.y
    };
    let left = viewport.center().x - page_size.x * 0.5 + canvas.pan.x;

    let mut pages = Vec::with_capacity(page_count);
    let section_starts = section_start_positions(document, galley, paragraph_start_rows);
    let mut page_sections = Vec::with_capacity(page_ranges.len());
    for (start_y, _) in &page_ranges {
        page_sections.push(section_for_y(&section_starts, *start_y));
    }
    let mut section_totals = std::collections::HashMap::<SectionId, usize>::new();
    for section_id in &page_sections {
        *section_totals.entry(*section_id).or_insert(0) += 1;
    }
    let mut section_seen = std::collections::HashMap::<SectionId, usize>::new();

    for (index, (start_y, end_y)) in page_ranges.into_iter().enumerate() {
        let section_id = page_sections[index];
        let page_index_within_section = section_seen.entry(section_id).or_insert(0);
        let local_page_index = *page_index_within_section;
        *page_index_within_section += 1;
        let section_page_count = section_totals.get(&section_id).copied().unwrap_or(1);
        let setup = document
            .section_by_id(section_id)
            .map(|section| section.page_setup)
            .unwrap_or(page_setup);
        let min = egui::pos2(left, top + index as f32 * (page_size.y + page_gap));
        let page_rect = Rect::from_min_size(min, page_size);
        let content_rect = section_page_content_rect(page_rect, setup, 14.0, 14.0, canvas.zoom);
        pages.push(PageSlice {
            page_rect,
            content_rect,
            section_id,
            page_index_within_section: local_page_index,
            section_page_count,
            start_y,
            end_y,
        });
    }

    PageLayout { pages }
}

fn clamp_pan(
    viewport: Rect,
    page_width: f32,
    stack_height: f32,
    margin: f32,
    pan: &mut egui::Vec2,
) -> egui::Vec2 {
    let horizontal_overflow = ((page_width - viewport.width()) * 0.5).max(0.0);
    pan.x = pan.x.clamp(-horizontal_overflow, horizontal_overflow);

    let vertical_overflow = if stack_height < viewport.height() {
        0.0
    } else {
        stack_height + margin * 2.0 - viewport.height()
    };
    pan.y = pan.y.clamp(-vertical_overflow, 0.0);
    egui::vec2(horizontal_overflow, vertical_overflow)
}

#[cfg(test)]
mod tests {
    use super::clamp_pan;

    #[test]
    fn pan_stays_within_page_stack() {
        let viewport = egui::Rect::from_min_size(egui::Pos2::ZERO, egui::vec2(800.0, 600.0));
        let mut pan = egui::vec2(500.0, -2_000.0);

        let range = clamp_pan(viewport, 1_000.0, 1_200.0, 24.0, &mut pan);

        assert_eq!(pan, egui::vec2(100.0, -648.0));
        assert_eq!(range, egui::vec2(100.0, 648.0));

        let range = clamp_pan(viewport, 500.0, 400.0, 24.0, &mut pan);
        assert_eq!(pan, egui::Vec2::ZERO);
        assert_eq!(range, egui::Vec2::ZERO);
    }
}

fn section_start_positions(
    document: &DocumentState,
    galley: &egui::Galley,
    paragraph_start_rows: &[usize],
) -> Vec<(f32, SectionId)> {
    let mut starts: Vec<(f32, SectionId)> = document
        .sections
        .iter()
        .map(|section| {
            let row = paragraph_start_rows
                .get(section.starts_at_paragraph)
                .copied()
                .unwrap_or(0);
            let y = galley.rows.get(row).map(|row| row.pos.y).unwrap_or(0.0);
            (y, section.id)
        })
        .collect();
    starts.sort_by(|left, right| {
        left.0
            .partial_cmp(&right.0)
            .unwrap_or(std::cmp::Ordering::Equal)
    });
    starts
}

fn section_for_y(starts: &[(f32, SectionId)], y: f32) -> SectionId {
    starts
        .iter()
        .rev()
        .find(|(start_y, _)| *start_y <= y + 0.5)
        .map(|(_, id)| *id)
        .or_else(|| starts.first().map(|(_, id)| *id))
        .unwrap_or(1)
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
    tables: &[TableLayout],
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

        if let Some(table) = tables.iter().find(|table| table.row_index == row_index) {
            let mut table_row_start = row_start;
            for height in &table.row_heights {
                let table_row_end = table_row_start + height;
                if table_row_end - page_start > page_height {
                    if table_row_start > page_start {
                        pages.push((page_start, table_row_start));
                        page_start = table_row_start;
                    }
                    while table_row_end - page_start > page_height {
                        pages.push((page_start, page_start + page_height));
                        page_start += page_height;
                    }
                }
                last_row_end = table_row_end;
                table_row_start = table_row_end;
            }
            last_row_end = last_row_end.max(row_end);
            continue;
        }

        if row_end - page_start > page_height {
            if row_start > page_start {
                pages.push((page_start, last_row_end.max(page_start)));
                page_start = row_start;
            }
            while row_end - page_start > page_height {
                pages.push((page_start, page_start + page_height));
                page_start += page_height;
            }
        }

        last_row_end = row_end;
    }

    pages.push((page_start, last_row_end.max(page_start)));
    pages
}

#[cfg(test)]
mod pagination_tests {
    use super::{compute_page_ranges, PageLayout, PageSlice};
    use crate::{canvas::layout::TableLayout, document::DocumentTable};

    #[test]
    fn drag_selection_reaches_line_ends_from_the_page_margin() {
        let layout = PageLayout {
            pages: vec![PageSlice {
                page_rect: egui::Rect::from_min_size(egui::Pos2::ZERO, egui::vec2(200.0, 200.0)),
                content_rect: egui::Rect::from_min_max(
                    egui::pos2(20.0, 20.0),
                    egui::pos2(180.0, 180.0),
                ),
                section_id: 1,
                page_index_within_section: 0,
                section_page_count: 1,
                start_y: 100.0,
                end_y: 150.0,
            }],
        };

        assert_eq!(
            layout.clamped_document_pos(egui::pos2(195.0, 190.0)),
            Some(egui::vec2(160.0, 150.0))
        );
    }

    #[test]
    fn table_rows_move_intact_to_following_pages() {
        let mut job = egui::epaint::text::LayoutJob::simple_singleline(
            "x".to_owned(),
            egui::FontId::default(),
            egui::Color32::BLACK,
        );
        job.wrap.max_width = 100.0;
        let ctx = egui::Context::default();
        let mut galley = None;
        let _ = ctx.run_ui(egui::RawInput::default(), |ui| {
            galley = Some(ui.painter().layout_job(job.clone()));
        });
        let mut galley = galley.unwrap();
        let row = std::sync::Arc::make_mut(&mut std::sync::Arc::make_mut(&mut galley).rows[0].row);
        row.size.y = 120.0;
        let table = TableLayout {
            row_index: 0,
            height: 120.0,
            row_heights: vec![40.0, 40.0, 40.0],
            table: DocumentTable::new(1, 3, 1, 100.0),
        };

        assert_eq!(
            compute_page_ranges(&galley, 100.0, &[], &[table]),
            vec![(0.0, 80.0), (80.0, 120.0)]
        );
    }
}
