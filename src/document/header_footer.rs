use crate::document::types::{
    plain_text_from_runs, DocumentState, HeaderFooterKind, HeaderFooterStory, HeaderFooterVariant,
    PageMargins, PageSetup, PageSize, ResolvedHeaderFooter, Section, SectionId, TextRun,
};

impl DocumentState {
    pub fn render_page_field(
        &self,
        template: &str,
        page_number: usize,
        page_count: usize,
    ) -> String {
        let displayed_page_number = self
            .page_number_start
            .saturating_add(page_number.saturating_sub(1));
        template
            .replace("{ SECTIONPAGES }", &page_count.to_string())
            .replace("{ SECTIONPAGES}", &page_count.to_string())
            .replace("{SECTIONPAGES }", &page_count.to_string())
            .replace("{SECTIONPAGES}", &page_count.to_string())
            .replace("{ NUMPAGES }", &page_count.to_string())
            .replace("{ NUMPAGES}", &page_count.to_string())
            .replace("{NUMPAGES }", &page_count.to_string())
            .replace("{NUMPAGES}", &page_count.to_string())
            .replace("{ PAGE }", &displayed_page_number.to_string())
            .replace("{ PAGE}", &displayed_page_number.to_string())
            .replace("{PAGE }", &displayed_page_number.to_string())
            .replace("{PAGE}", &displayed_page_number.to_string())
            .replace("{pagecount}", &page_count.to_string())
            .replace("{pages}", &page_count.to_string())
            .replace("{sectionpages}", &page_count.to_string())
            .replace("{numpages}", &page_count.to_string())
            .replace("{page}", &displayed_page_number.to_string())
    }

    pub fn default_page_setup(&self) -> PageSetup {
        self.sections
            .first()
            .map(|section| section.page_setup)
            .unwrap_or(PageSetup {
                page_size: self.page_size,
                margins: self.margins,
                header_from_top_points: 36.0,
                footer_from_bottom_points: 36.0,
                page_number_start: Some(self.page_number_start),
            })
    }

    pub fn active_page_size(&self) -> PageSize {
        self.default_page_setup().page_size
    }

    pub fn active_margins(&self) -> PageMargins {
        self.default_page_setup().margins
    }

    pub fn section_at_paragraph(&self, paragraph_index: usize) -> &Section {
        self.sections
            .iter()
            .rev()
            .find(|section| section.starts_at_paragraph <= paragraph_index)
            .or_else(|| self.sections.first())
            .expect("document always has at least one section")
    }

    pub fn section_at_paragraph_mut(&mut self, paragraph_index: usize) -> &mut Section {
        let id = self.section_at_paragraph(paragraph_index).id;
        self.section_by_id_mut(id)
            .expect("section id from document should exist")
    }

    pub fn section_by_id(&self, id: SectionId) -> Option<&Section> {
        self.sections.iter().find(|section| section.id == id)
    }

    pub fn section_by_id_mut(&mut self, id: SectionId) -> Option<&mut Section> {
        self.sections.iter_mut().find(|section| section.id == id)
    }

    pub fn first_section_id(&self) -> SectionId {
        self.sections.first().map(|section| section.id).unwrap_or(1)
    }

    pub fn insert_section_break_before_paragraph(&mut self, paragraph_index: usize) -> SectionId {
        self.ensure_paragraph_style_count();
        let paragraph_index = paragraph_index.min(self.paragraph_count().saturating_sub(1));
        if let Some(existing) = self
            .sections
            .iter()
            .find(|section| section.starts_at_paragraph == paragraph_index)
        {
            return existing.id;
        }

        let previous = self
            .section_at_paragraph(paragraph_index.saturating_sub(1))
            .clone();
        let next_id = self
            .sections
            .iter()
            .map(|section| section.id)
            .max()
            .unwrap_or(0)
            + 1;
        self.sections
            .push(Section::linked_from(next_id, paragraph_index, &previous));
        self.sections
            .sort_by_key(|section| section.starts_at_paragraph);
        next_id
    }

    pub fn resolve_header_footer_slot(
        &self,
        section_id: SectionId,
        kind: HeaderFooterKind,
        variant: HeaderFooterVariant,
    ) -> ResolvedHeaderFooter<'_> {
        let Some(section_index) = self
            .sections
            .iter()
            .position(|section| section.id == section_id)
        else {
            let section = self.sections.first().expect("document has section");
            return ResolvedHeaderFooter {
                section_id: section.id,
                source_section_id: section.id,
                variant,
                story: section.header_footer.slot(kind, variant).story_ref(),
                inherited: false,
            };
        };

        let section = &self.sections[section_index];
        let slot = section.header_footer.slot(kind, variant);
        if slot.linked_to_previous && section_index > 0 {
            let previous_id = self.sections[section_index - 1].id;
            let mut resolved = self.resolve_header_footer_slot(previous_id, kind, variant);
            resolved.section_id = section_id;
            resolved.inherited = true;
            return resolved;
        }

        ResolvedHeaderFooter {
            section_id,
            source_section_id: section.id,
            variant,
            story: slot.story_ref(),
            inherited: false,
        }
    }

    pub fn header_footer_variant_for_page(
        &self,
        section_id: SectionId,
        page_index_within_section: usize,
        _kind: HeaderFooterKind,
    ) -> HeaderFooterVariant {
        let Some(section) = self.section_by_id(section_id) else {
            return HeaderFooterVariant::Default;
        };
        if section.different_first_page && page_index_within_section == 0 {
            HeaderFooterVariant::First
        } else if self.different_odd_even_pages && (page_index_within_section + 1) % 2 == 0 {
            HeaderFooterVariant::Even
        } else {
            HeaderFooterVariant::Default
        }
    }

    pub fn render_page_field_for_section_page(
        &self,
        text: &str,
        section_id: SectionId,
        page_index_within_section: usize,
        _absolute_page_index: usize,
        absolute_page_count: usize,
        section_page_count: usize,
    ) -> String {
        let displayed_page_number =
            self.displayed_page_number(section_id, page_index_within_section);
        text.replace("{ SECTIONPAGES }", &section_page_count.to_string())
            .replace("{ SECTIONPAGES}", &section_page_count.to_string())
            .replace("{SECTIONPAGES }", &section_page_count.to_string())
            .replace("{SECTIONPAGES}", &section_page_count.to_string())
            .replace("{ NUMPAGES }", &absolute_page_count.to_string())
            .replace("{ NUMPAGES}", &absolute_page_count.to_string())
            .replace("{NUMPAGES }", &absolute_page_count.to_string())
            .replace("{NUMPAGES}", &absolute_page_count.to_string())
            .replace("{ PAGE }", &displayed_page_number.to_string())
            .replace("{ PAGE}", &displayed_page_number.to_string())
            .replace("{PAGE }", &displayed_page_number.to_string())
            .replace("{PAGE}", &displayed_page_number.to_string())
            .replace("{pagecount}", &absolute_page_count.to_string())
            .replace("{pages}", &absolute_page_count.to_string())
            .replace("{sectionpages}", &section_page_count.to_string())
            .replace("{numpages}", &absolute_page_count.to_string())
            .replace("{page}", &displayed_page_number.to_string())
    }

    pub fn displayed_page_number(
        &self,
        section_id: SectionId,
        page_index_within_section: usize,
    ) -> usize {
        let Some(section_index) = self
            .sections
            .iter()
            .position(|section| section.id == section_id)
        else {
            return page_index_within_section + 1;
        };
        if let Some(start) = self.sections[section_index].page_setup.page_number_start {
            return start + page_index_within_section;
        }

        let mut current = 1usize;
        for section in self.sections.iter().take(section_index + 1) {
            if let Some(start) = section.page_setup.page_number_start {
                current = start;
            }
            if section.id == section_id {
                return current + page_index_within_section;
            }
        }
        page_index_within_section + 1
    }

    pub fn header_footer_story(
        &self,
        section_id: SectionId,
        kind: HeaderFooterKind,
        variant: HeaderFooterVariant,
    ) -> Option<&HeaderFooterStory> {
        self.section_by_id(section_id)
            .map(|section| section.header_footer.slot(kind, variant).story_ref())
    }

    pub fn header_footer_story_mut_materialized(
        &mut self,
        section_id: SectionId,
        kind: HeaderFooterKind,
        variant: HeaderFooterVariant,
    ) -> Option<&mut HeaderFooterStory> {
        let inherited_story = self
            .resolve_header_footer_slot(section_id, kind, variant)
            .story
            .clone();
        let section = self.section_by_id_mut(section_id)?;
        let slot = section.header_footer.slot_mut(kind, variant);
        if slot.linked_to_previous {
            slot.story = inherited_story;
            slot.linked_to_previous = false;
        }
        Some(&mut slot.story)
    }

    pub fn set_header_footer_link(
        &mut self,
        section_id: SectionId,
        kind: HeaderFooterKind,
        variant: HeaderFooterVariant,
        linked: bool,
    ) {
        if let Some(section_index) = self
            .sections
            .iter()
            .position(|section| section.id == section_id)
        {
            if section_index == 0 && linked {
                return;
            }
            let slot = self.sections[section_index]
                .header_footer
                .slot_mut(kind, variant);
            slot.linked_to_previous = linked;
        }
    }

    pub fn header_footer_linked(
        &self,
        section_id: SectionId,
        kind: HeaderFooterKind,
        variant: HeaderFooterVariant,
    ) -> bool {
        self.section_by_id(section_id)
            .map(|section| section.header_footer.slot(kind, variant).linked_to_previous)
            .unwrap_or(false)
    }

    pub fn clear_header_footer_slot(
        &mut self,
        section_id: SectionId,
        kind: HeaderFooterKind,
        variant: HeaderFooterVariant,
    ) {
        if let Some(section) = self.section_by_id_mut(section_id) {
            let slot = section.header_footer.slot_mut(kind, variant);
            slot.story = HeaderFooterStory::empty();
            slot.linked_to_previous = false;
        }
    }

    pub fn sync_compat_from_first_section(&mut self) {
        let Some(section) = self.sections.first() else {
            return;
        };
        self.page_size = section.page_setup.page_size;
        self.margins = section.page_setup.margins;
        self.different_first_page = section.different_first_page;
        self.page_number_start = section.page_setup.page_number_start.unwrap_or(1);
        self.header_runs = section.header_footer.header_default.story.runs.clone();
        self.footer_runs = section.header_footer.footer_default.story.runs.clone();
        self.first_page_header_runs = section.header_footer.header_first.story.runs.clone();
        self.first_page_footer_runs = section.header_footer.footer_first.story.runs.clone();
        self.even_page_header_runs = section.header_footer.header_even.story.runs.clone();
        self.even_page_footer_runs = section.header_footer.footer_even.story.runs.clone();
        self.header_text = plain_text_from_runs(&self.header_runs);
        self.footer_text = plain_text_from_runs(&self.footer_runs);
        self.first_page_header_text = plain_text_from_runs(&self.first_page_header_runs);
        self.first_page_footer_text = plain_text_from_runs(&self.first_page_footer_runs);
        self.even_page_header_text = plain_text_from_runs(&self.even_page_header_runs);
        self.even_page_footer_text = plain_text_from_runs(&self.even_page_footer_runs);
    }

    pub fn header_template_for_page(&self, page_number: usize) -> &str {
        if self.different_first_page && page_number == 1 {
            &self.first_page_header_text
        } else if self.different_odd_even_pages && page_number % 2 == 0 {
            &self.even_page_header_text
        } else {
            &self.header_text
        }
    }

    pub fn header_runs_for_page(&self, page_number: usize) -> &[TextRun] {
        if self.different_first_page && page_number == 1 {
            &self.first_page_header_runs
        } else if self.different_odd_even_pages && page_number % 2 == 0 {
            &self.even_page_header_runs
        } else {
            &self.header_runs
        }
    }

    pub fn footer_template_for_page(&self, page_number: usize) -> &str {
        if self.different_first_page && page_number == 1 {
            &self.first_page_footer_text
        } else if self.different_odd_even_pages && page_number % 2 == 0 {
            &self.even_page_footer_text
        } else {
            &self.footer_text
        }
    }

    pub fn footer_runs_for_page(&self, page_number: usize) -> &[TextRun] {
        if self.different_first_page && page_number == 1 {
            &self.first_page_footer_runs
        } else if self.different_odd_even_pages && page_number % 2 == 0 {
            &self.even_page_footer_runs
        } else {
            &self.footer_runs
        }
    }
}
