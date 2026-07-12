pub mod docx;
mod editing;
mod export;
mod formatting;
mod header_footer;
mod markdown;
mod odt;
mod table;
mod text;
mod types;

pub use table::{DocumentTable, TableBorders, TableCell};
pub use types::*;

impl DocumentState {
    pub fn bootstrap() -> Self {
        Self {
            title: "Untitled".to_owned(),
            runs: vec![
                TextRun {
                    text: "wors".to_owned(),
                    style: CharacterStyle {
                        bold: true,
                        font_size_points: 22.0,
                        ..CharacterStyle::default()
                    },
                },
                TextRun {
                    text: " now edits text on a custom painter-backed page.\n\n".to_owned(),
                    style: CharacterStyle {
                        font_size_points: 13.0,
                        ..CharacterStyle::default()
                    },
                },
                TextRun {
                    text: "Use the ribbon above to change".to_owned(),
                    style: CharacterStyle {
                        bold: true,
                        ..CharacterStyle::default()
                    },
                },
                TextRun {
                    text: " bold, italic, underline, strike-through, text size, font family, text color, and highlight.".to_owned(),
                    style: CharacterStyle::default(),
                },
            ],
            paragraph_styles: vec![ParagraphStyle::default(); 3],
            paragraph_images: vec![None; 3],
            paragraph_tables: vec![None; 3],
            page_size: PageSize::a4(),
            margins: PageMargins::standard(),
            header_text: String::new(),
            footer_text: String::new(),
            first_page_header_text: String::new(),
            first_page_footer_text: String::new(),
            even_page_header_text: String::new(),
            even_page_footer_text: String::new(),
            header_runs: empty_header_footer_runs(),
            footer_runs: empty_header_footer_runs(),
            first_page_header_runs: empty_header_footer_runs(),
            first_page_footer_runs: empty_header_footer_runs(),
            even_page_header_runs: empty_header_footer_runs(),
            even_page_footer_runs: empty_header_footer_runs(),
            different_first_page: false,
            different_odd_even_pages: false,
            page_number_start: 1,
            sections: vec![Section::first(PageSetup::standard())],
            source_docx: None,
        }
    }
}

#[cfg(test)]
mod tests;
