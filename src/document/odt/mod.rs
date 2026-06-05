pub mod export;
pub mod import;

pub use export::document_to_odt;
pub use import::odt_to_document;

#[cfg(test)]
mod tests {
    use super::{document_to_odt, odt_to_document};
    use crate::document::{
        CharacterStyle, DocumentImage, DocumentState, DocumentTable, ImageLayoutMode,
        ImageRendering, ParagraphAlignment, ParagraphStyle, TableCell, TextRun, WrapMode,
        OBJECT_REPLACEMENT_CHAR,
    };

    const PNG_1X1: &[u8] = &[
        137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82, 0, 0, 0, 1, 0, 0, 0, 1, 8, 6,
        0, 0, 0, 31, 21, 196, 137, 0, 0, 0, 10, 73, 68, 65, 84, 120, 156, 99, 0, 1, 0, 0, 5, 0, 1,
        13, 10, 45, 180, 0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130,
    ];

    #[test]
    fn exports_and_imports_basic_odt_package() {
        let mut document = DocumentState::bootstrap();
        document.title = "Roundtrip".to_owned();
        document.replace_with_runs(
            "Roundtrip".to_owned(),
            vec![TextRun {
                text: "Hello ODT".to_owned(),
                style: CharacterStyle {
                    bold: true,
                    italic: true,
                    ..CharacterStyle::default()
                },
            }],
        );
        document.paragraph_styles[0].alignment = ParagraphAlignment::Center;

        let bytes = document_to_odt(&document).expect("odt should export");
        assert!(bytes.starts_with(b"PK"));

        let imported = odt_to_document(&bytes).expect("odt should import");
        assert_eq!(imported.runs[0].text, "Hello ODT");
        assert!(imported.runs[0].style.bold);
        assert!(imported.runs[0].style.italic);
        assert_eq!(
            imported.paragraph_styles[0].alignment,
            ParagraphAlignment::Center
        );
    }

    #[test]
    fn exports_and_imports_odt_images_and_tables() {
        let mut document = DocumentState::bootstrap();
        let table = DocumentTable {
            id: 1,
            rows: vec![vec![TableCell::new("Cell")]],
            col_widths_points: vec![144.0],
            row_heights_points: vec![24.0],
            borders: Default::default(),
        };
        document.runs = vec![
            TextRun {
                text: OBJECT_REPLACEMENT_CHAR.to_string(),
                style: CharacterStyle::default(),
            },
            TextRun {
                text: "\n".to_owned(),
                style: CharacterStyle::default(),
            },
            TextRun {
                text: OBJECT_REPLACEMENT_CHAR.to_string(),
                style: CharacterStyle::default(),
            },
        ];
        document.paragraph_styles = vec![ParagraphStyle::default(), ParagraphStyle::default()];
        document.paragraph_images = vec![Some(test_image()), None];
        document.paragraph_tables = vec![None, Some(table)];

        let bytes = document_to_odt(&document).expect("odt should export");
        let imported = odt_to_document(&bytes).expect("odt should import");

        let image = imported.paragraph_images[0]
            .as_ref()
            .expect("image imports");
        assert_eq!(image.bytes, PNG_1X1);
        let table = imported.paragraph_tables[1]
            .as_ref()
            .expect("table imports");
        assert_eq!(table.rows[0][0].plain_text().trim(), "Cell");
    }

    fn test_image() -> DocumentImage {
        DocumentImage {
            id: 1,
            bytes: PNG_1X1.to_vec(),
            alt_text: "pixel".to_owned(),
            width_points: 12.0,
            height_points: 12.0,
            lock_aspect_ratio: true,
            opacity: 1.0,
            layout_mode: ImageLayoutMode::Inline,
            wrap_mode: WrapMode::Inline,
            rendering: ImageRendering::Smooth,
            horizontal_position: Default::default(),
            vertical_position: Default::default(),
            distance_from_text: Default::default(),
            z_index: 0,
            move_with_text: true,
            allow_overlap: false,
        }
    }
}
