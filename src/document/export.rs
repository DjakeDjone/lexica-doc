use std::{collections::BTreeMap, fmt::Write as _, fs, path::Path};

use base64::{engine::general_purpose::STANDARD as BASE64_STANDARD, Engine as _};
use eframe::egui::Color32;
use printpdf::{
    Base64OrRaw, BuiltinFont, Color, GeneratePdfOptions, Op, PdfDocument, PdfFontHandle, PdfPage,
    PdfSaveOptions, Point, Pt, TextItem,
};

use super::{
    docx::docx_to_document,
    markdown::markdown_to_runs,
    odt::{document_to_odt, odt_to_document},
    CharacterStyle, DocumentImage, DocumentState, DocumentTable, FontChoice, HeaderFooterKind,
    ImageLayoutMode, LineSpacing, LineSpacingKind, ListKind, ParagraphAlignment, TableCell,
    TextRun, OBJECT_REPLACEMENT_CHAR,
};

impl DocumentState {
    pub fn load_from_path(path: &Path) -> Result<Self, String> {
        let title = path
            .file_stem()
            .and_then(|name| name.to_str())
            .unwrap_or("Imported")
            .to_owned();

        let extension = path
            .extension()
            .and_then(|ext| ext.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();

        let runs = match extension.as_str() {
            "docx" => {
                let imported = docx_to_document(
                    &fs::read(path)
                        .map_err(|error| format!("failed to read {}: {error}", path.display()))?,
                )?;

                let mut document = Self::bootstrap();
                document.title = title;
                document.runs = imported.runs;
                document.paragraph_styles = imported.paragraph_styles;
                document.paragraph_images = imported.paragraph_images;
                document.paragraph_tables = imported.paragraph_tables;
                if let Some(page_size) = imported.page_size {
                    document.page_size = page_size;
                    if let Some(section) = document.sections.first_mut() {
                        section.page_setup.page_size = page_size;
                    }
                }
                if let Some(margins) = imported.margins {
                    document.margins = margins;
                    if let Some(section) = document.sections.first_mut() {
                        section.page_setup.margins = margins;
                    }
                }
                document.different_odd_even_pages = imported.different_odd_even_pages;
                if !imported.sections.is_empty() {
                    document.sections = imported.sections;
                    document.sync_compat_from_first_section();
                }
                document.normalize_runs();
                document.ensure_paragraph_style_count();
                return Ok(document);
            }
            "odt" => {
                let imported = odt_to_document(
                    &fs::read(path)
                        .map_err(|error| format!("failed to read {}: {error}", path.display()))?,
                )?;

                let mut document = Self::bootstrap();
                document.title = title;
                document.runs = imported.runs;
                document.paragraph_styles = imported.paragraph_styles;
                document.paragraph_images = imported.paragraph_images;
                document.paragraph_tables = imported.paragraph_tables;
                if let Some(page_size) = imported.page_size {
                    document.page_size = page_size;
                    if let Some(section) = document.sections.first_mut() {
                        section.page_setup.page_size = page_size;
                    }
                }
                if let Some(margins) = imported.margins {
                    document.margins = margins;
                    if let Some(section) = document.sections.first_mut() {
                        section.page_setup.margins = margins;
                    }
                }
                document.sync_compat_from_first_section();
                document.normalize_runs();
                document.ensure_paragraph_style_count();
                return Ok(document);
            }
            "md" | "markdown" => {
                let source = fs::read_to_string(path)
                    .map_err(|error| format!("failed to read {}: {error}", path.display()))?;
                markdown_to_runs(&source)
            }
            _ => {
                let source = fs::read_to_string(path)
                    .map_err(|error| format!("failed to read {}: {error}", path.display()))?;
                vec![TextRun {
                    text: source,
                    style: CharacterStyle::default(),
                }]
            }
        };

        let mut document = Self::bootstrap();
        document.replace_with_runs(title, runs);
        Ok(document)
    }

    pub fn save_to_path(&self, path: &Path) -> Result<(), String> {
        let extension = path
            .extension()
            .and_then(|ext| ext.to_str())
            .unwrap_or_default()
            .to_ascii_lowercase();
        let bytes = self.export_bytes_for_extension(&extension)?;

        fs::write(path, bytes)
            .map_err(|error| format!("failed to save {}: {error}", path.display()))
    }

    pub fn export_bytes_for_extension(&self, extension: &str) -> Result<Vec<u8>, String> {
        match extension
            .trim_start_matches('.')
            .to_ascii_lowercase()
            .as_str()
        {
            "md" | "markdown" => Ok(self.to_markdown().into_bytes()),
            "txt" | "" => Ok(self.to_plain_text_export().into_bytes()),
            "html" | "htm" => Ok(self.to_html().into_bytes()),
            "pdf" => self.to_pdf_bytes(),
            "odt" => document_to_odt(self),
            other => Err(format!(
                "saving .{other} is not supported yet; use .txt, .md, .html, .odt, or .pdf"
            )),
        }
    }

    pub(super) fn to_plain_text_export(&self) -> String {
        self.paragraphs()
            .into_iter()
            .map(|paragraph| {
                if let Some(table) = &paragraph.table {
                    return table_to_plain_text(table);
                }
                let mut text = plain_text_from_runs(&paragraph.runs);
                text.retain(|ch| ch != OBJECT_REPLACEMENT_CHAR);
                if paragraph.style.page_break_before {
                    if text.is_empty() {
                        text.push('\u{000C}');
                    } else {
                        text.insert(0, '\u{000C}');
                    }
                }
                if let Some(marker) = paragraph.list_marker {
                    if text.is_empty() {
                        marker
                    } else {
                        text.insert_str(0, &format!("{marker} "));
                        text
                    }
                } else {
                    text
                }
            })
            .collect::<Vec<_>>()
            .join("\n")
    }

    pub(super) fn to_markdown(&self) -> String {
        self.paragraphs()
            .into_iter()
            .map(|paragraph| {
                if let Some(table) = &paragraph.table {
                    return table_to_markdown(table);
                }
                let mut text = markdown_text_from_runs(&paragraph.runs);
                if paragraph.style.page_break_before {
                    let break_marker = "<div style=\"page-break-before: always\"></div>";
                    text = if text.is_empty() {
                        break_marker.to_owned()
                    } else {
                        format!("{break_marker}\n\n{text}")
                    };
                }
                if paragraph.image.is_some() {
                    let alt = paragraph
                        .image
                        .as_ref()
                        .map(|image| image.alt_text.as_str())
                        .filter(|alt| !alt.is_empty())
                        .unwrap_or("Image");
                    if text.is_empty() {
                        text = format!("![{alt}](embedded-image)");
                    } else {
                        text = format!("{text}\n\n![{alt}](embedded-image)");
                    }
                }
                if let Some(marker) = paragraph.list_marker.as_deref() {
                    let prefix = match paragraph.style.list_kind {
                        ListKind::Bullet => "- ".to_owned(),
                        ListKind::Ordered => format!("{marker} "),
                        ListKind::None => String::new(),
                    };
                    text = format!("{prefix}{text}");
                }

                match paragraph.style.alignment {
                    ParagraphAlignment::Left => text,
                    ParagraphAlignment::Center => format!("<div align=\"center\">{text}</div>"),
                    ParagraphAlignment::Right => format!("<div align=\"right\">{text}</div>"),
                    ParagraphAlignment::Justify => format!("<div align=\"justify\">{text}</div>"),
                }
            })
            .collect::<Vec<_>>()
            .join("\n")
    }

    pub(super) fn to_html(&self) -> String {
        let mut html = String::new();
        let _ = write!(
            html,
            "<!doctype html>\
<html lang=\"en\">\
<head>\
<meta charset=\"utf-8\">\
<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\
<title>{}</title>\
<style>\
body {{ margin: 0; padding: 18pt; background: #e7ebf0; }}\
.page {{ box-sizing: border-box; margin: 0 auto; width: {}pt; min-height: {}pt; padding: {}pt {}pt {}pt {}pt; background: #ffffff; color: #24272e; box-shadow: 0 1px 5px rgba(0, 0, 0, 0.18); }}\
.page-header, .page-footer {{ color: #58606c; font: 9pt Helvetica, Arial, sans-serif; text-align: center; }}\
.page-header {{ margin-top: -{}pt; margin-bottom: {}pt; }}\
.page-footer {{ margin-top: {}pt; margin-bottom: -{}pt; }}\
.paragraph {{ margin: 0; white-space: pre-wrap; }}\
.page-break {{ break-before: page; page-break-before: always; height: 0; }}\
.image-block {{ display: block; max-width: 100%; }}\
@media print {{ body {{ background: transparent; padding: 0; }} .page {{ box-shadow: none; width: auto; min-height: auto; }} }}\
</style>\
</head>\
<body>\
<div class=\"page\">",
            html_escape(&self.title),
            self.page_size.width_points,
            self.page_size.height_points,
            self.margins.top_points,
            self.margins.right_points,
            self.margins.bottom_points,
            self.margins.left_points,
            (self.margins.top_points * 0.65).max(0.0),
            (self.margins.top_points * 0.45).max(0.0),
            (self.margins.bottom_points * 0.45).max(0.0),
            (self.margins.bottom_points * 0.65).max(0.0)
        );

        let first_header = self.header_template_for_page(1);
        if !first_header.trim().is_empty() {
            let _ = write!(
                html,
                "<div class=\"page-header\">{}</div>",
                html_escape(&self.render_page_field(first_header, 1, 1))
            );
        }

        for paragraph in self.paragraphs() {
            if paragraph.style.page_break_before {
                html.push_str("<div class=\"page-break\"></div>");
            }

            if let Some(table) = &paragraph.table {
                html.push_str(&table_to_html(table));
                continue;
            }

            let _ = write!(
                html,
                "<p class=\"paragraph\" style=\"text-align:{};margin-top:{}pt;margin-bottom:{}pt;{}\">",
                paragraph_alignment_css(paragraph.style.alignment),
                paragraph.style.spacing_before_points,
                paragraph.style.spacing_after_points,
                line_spacing_css(paragraph.style.line_spacing)
            );

            if let Some(marker) = paragraph.list_marker {
                let prefix = match paragraph.style.list_kind {
                    ListKind::Bullet | ListKind::Ordered => format!("{marker} "),
                    ListKind::None => String::new(),
                };
                html.push_str(&html_escape(&prefix));
            }

            for run in paragraph.runs {
                let text: String = run
                    .text
                    .chars()
                    .filter(|ch| *ch != OBJECT_REPLACEMENT_CHAR)
                    .collect();
                if text.is_empty() {
                    continue;
                }

                let _ = write!(
                    html,
                    "<span style=\"{}\">{}</span>",
                    run_style_css(run.style),
                    html_escape(&text)
                );
            }

            if let Some(image) = paragraph.image.as_ref() {
                if let Some(mime_type) = image_mime_type(&image.bytes) {
                    let _ = write!(
                        html,
                        "<img class=\"image-block\" alt=\"{}\" src=\"data:{};base64,{}\" style=\"width:{}pt;height:{}pt;opacity:{:.3};{}\" />",
                        html_escape(&image.alt_text),
                        mime_type,
                        BASE64_STANDARD.encode(&image.bytes),
                        image.width_points,
                        image.height_points,
                        image.opacity.clamp(0.0, 1.0),
                        image_position_css(image)
                    );
                }
            }

            html.push_str("</p>");
        }

        let first_footer = self.footer_template_for_page(1);
        if !first_footer.trim().is_empty() {
            let _ = write!(
                html,
                "<div class=\"page-footer\">{}</div>",
                html_escape(&self.render_page_field(first_footer, 1, 1))
            );
        }

        html.push_str("</div></body></html>");
        html
    }

    pub(super) fn to_pdf_bytes(&self) -> Result<Vec<u8>, String> {
        let html = self.to_pdf_html();
        let options = GeneratePdfOptions {
            page_width: Some(points_to_mm(self.page_size.width_points)),
            page_height: Some(points_to_mm(self.page_size.height_points)),
            margin_top: Some(points_to_mm(self.margins.top_points)),
            margin_right: Some(points_to_mm(self.margins.right_points)),
            margin_bottom: Some(points_to_mm(self.margins.bottom_points)),
            margin_left: Some(points_to_mm(self.margins.left_points)),
            ..GeneratePdfOptions::default()
        };
        let images: BTreeMap<String, Base64OrRaw> = BTreeMap::new();
        let fonts: BTreeMap<String, Base64OrRaw> = BTreeMap::new();

        let mut warnings = Vec::new();
        let mut rendered = PdfDocument::from_html(&html, &images, &fonts, &options, &mut warnings)
            .map_err(|error| format!("failed to render PDF: {error}"))?;
        rendered.metadata.info.document_title = self.title.clone();
        rendered.metadata.info.conformance = Default::default();
        if rendered.pages.is_empty() {
            return Ok(self.to_plain_text_pdf_bytes());
        }
        self.stamp_pdf_header_footer(&mut rendered);

        Ok(rendered.save(&PdfSaveOptions::default(), &mut warnings))
    }

    pub(super) fn to_pdf_html(&self) -> String {
        let mut html = String::new();
        let _ = write!(
            html,
            "<!doctype html>\
<html lang=\"en\">\
<head>\
<meta charset=\"utf-8\">\
<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\
<title>{}</title>\
<style>\
body {{ margin: 0; padding: 0; color: #24272e; font-family: Helvetica, Arial, sans-serif; }}\
p {{ margin: 0; white-space: pre-wrap; }}\
.page-break {{ break-before: page; page-break-before: always; height: 0; }}\
.image-block {{ display: block; max-width: 100%; }}\
</style>\
</head>\
<body>",
            html_escape(&self.title)
        );

        for paragraph in self.paragraphs() {
            if paragraph.style.page_break_before {
                html.push_str("<div class=\"page-break\"></div>");
            }

            let _ = write!(
                html,
                "<p style=\"text-align:{};margin-top:{:.2}px;margin-bottom:{:.2}px;{}\">",
                paragraph_alignment_css(paragraph.style.alignment),
                points_to_css_px(paragraph.style.spacing_before_points as f32),
                points_to_css_px(paragraph.style.spacing_after_points as f32),
                line_spacing_css_pdf(paragraph.style.line_spacing)
            );

            if let Some(marker) = paragraph.list_marker {
                let prefix = match paragraph.style.list_kind {
                    ListKind::Bullet | ListKind::Ordered => format!("{marker} "),
                    ListKind::None => String::new(),
                };
                html.push_str(&html_escape(&prefix));
            }

            for run in paragraph.runs {
                let text: String = run
                    .text
                    .chars()
                    .filter(|ch| *ch != OBJECT_REPLACEMENT_CHAR)
                    .collect();
                if text.is_empty() {
                    continue;
                }

                let escaped = html_escape(&text);
                let mut run_html = format!(
                    "<span style=\"{}\">{escaped}</span>",
                    run_style_css_pdf(run.style)
                );
                if run.style.bold {
                    run_html = format!("<strong>{run_html}</strong>");
                }
                if run.style.italic {
                    run_html = format!("<em>{run_html}</em>");
                }
                if run.style.underline {
                    run_html =
                        format!("<span style=\"text-decoration:underline;\">{run_html}</span>");
                }
                if run.style.strikethrough {
                    run_html =
                        format!("<span style=\"text-decoration:line-through;\">{run_html}</span>");
                }
                html.push_str(&run_html);
            }

            if let Some(image) = paragraph.image.as_ref() {
                if let Some(mime_type) = image_mime_type(&image.bytes) {
                    let _ = write!(
                        html,
                        "<img class=\"image-block\" alt=\"{}\" src=\"data:{};base64,{}\" style=\"width:{:.2}px;height:{:.2}px;opacity:{:.3};{}\" />",
                        html_escape(&image.alt_text),
                        mime_type,
                        BASE64_STANDARD.encode(&image.bytes),
                        points_to_css_px(image.width_points),
                        points_to_css_px(image.height_points),
                        image.opacity.clamp(0.0, 1.0),
                        image_position_css_pdf(image)
                    );
                }
            }

            html.push_str("</p>");
        }

        html.push_str("</body></html>");
        html
    }

    pub(super) fn to_plain_text_pdf_bytes(&self) -> Vec<u8> {
        let page_width_mm = points_to_mm(self.page_size.width_points);
        let page_height_mm = points_to_mm(self.page_size.height_points);
        let left = self.margins.left_points.max(18.0);
        let top = self.margins.top_points.max(18.0);
        let bottom = self.margins.bottom_points.max(18.0);

        let font_size = 11.0_f32;
        let line_height = 14.0_f32;
        let max_lines =
            (((self.page_size.height_points - top - bottom) / line_height).floor() as usize).max(1);

        let mut logical_lines = Vec::new();
        for line in self
            .to_plain_text_export()
            .replace('\u{000C}', "\n\n----- Page Break -----\n\n")
            .lines()
        {
            let wrapped = wrap_text_for_pdf(line, 100);
            if wrapped.is_empty() {
                logical_lines.push(String::new());
            } else {
                logical_lines.extend(wrapped);
            }
        }
        if logical_lines.is_empty() {
            logical_lines.push(String::new());
        }

        let mut pages = Vec::new();
        for chunk in logical_lines.chunks(max_lines) {
            let mut y = self.page_size.height_points - top - font_size;
            let mut ops = vec![
                Op::StartTextSection,
                Op::SetFont {
                    font: PdfFontHandle::Builtin(BuiltinFont::Helvetica),
                    size: Pt(font_size),
                },
                Op::SetLineHeight {
                    lh: Pt(line_height),
                },
                Op::SetTextCursor {
                    pos: Point {
                        x: Pt(left),
                        y: Pt(y),
                    },
                },
            ];

            for (i, line) in chunk.iter().enumerate() {
                ops.push(Op::ShowText {
                    items: vec![TextItem::Text(line.clone())],
                });
                if i + 1 < chunk.len() {
                    ops.push(Op::AddLineBreak);
                    y -= line_height;
                    if y <= bottom {
                        break;
                    }
                }
            }

            ops.push(Op::EndTextSection);
            pages.push(PdfPage::new(
                printpdf::Mm(page_width_mm),
                printpdf::Mm(page_height_mm),
                ops,
            ));
        }

        let mut document = PdfDocument::new(&self.title);
        self.stamp_pdf_pages_header_footer(&mut pages);
        let document = document.with_pages(pages);
        let mut warnings = Vec::new();
        document.save(&PdfSaveOptions::default(), &mut warnings)
    }

    fn stamp_pdf_header_footer(&self, document: &mut PdfDocument) {
        self.stamp_pdf_pages_header_footer(&mut document.pages);
    }

    fn stamp_pdf_pages_header_footer(&self, pages: &mut [PdfPage]) {
        let Some(section) = self.sections.first() else {
            return;
        };

        let page_count = pages.len();
        for (index, page) in pages.iter_mut().enumerate() {
            let section_page_index = index;
            let header_variant = self.header_footer_variant_for_page(
                section.id,
                section_page_index,
                HeaderFooterKind::Header,
            );
            let footer_variant = self.header_footer_variant_for_page(
                section.id,
                section_page_index,
                HeaderFooterKind::Footer,
            );
            let header_story = self.resolve_header_footer_slot(
                section.id,
                HeaderFooterKind::Header,
                header_variant,
            );
            let footer_story = self.resolve_header_footer_slot(
                section.id,
                HeaderFooterKind::Footer,
                footer_variant,
            );
            let header_template = header_story.story.plain_text();
            let footer_template = footer_story.story.plain_text();
            if header_template.trim().is_empty() && footer_template.trim().is_empty() {
                continue;
            }
            let header = self.render_page_field_for_section_page(
                &header_template,
                section.id,
                section_page_index,
                index,
                page_count,
                page_count,
            );
            let footer = self.render_page_field_for_section_page(
                &footer_template,
                section.id,
                section_page_index,
                index,
                page_count,
                page_count,
            );
            append_pdf_page_field_ops(
                &mut page.ops,
                &header,
                self.page_size.width_points,
                self.margins.left_points,
                self.margins.right_points,
                self.page_size.height_points - (self.margins.top_points * 0.5).max(12.0),
            );
            append_pdf_page_field_ops(
                &mut page.ops,
                &footer,
                self.page_size.width_points,
                self.margins.left_points,
                self.margins.right_points,
                (self.margins.bottom_points * 0.5).max(12.0),
            );
        }
    }
}

fn append_pdf_page_field_ops(
    ops: &mut Vec<Op>,
    text: &str,
    page_width: f32,
    left_margin: f32,
    right_margin: f32,
    y: f32,
) {
    if text.trim().is_empty() {
        return;
    }

    let segments: Vec<&str> = text.split('\t').collect();
    append_pdf_text_segment(
        ops,
        segments.first().copied().unwrap_or_default(),
        left_margin,
        y,
    );
    if let Some(center) = segments.get(1).filter(|segment| !segment.is_empty()) {
        let width = estimate_pdf_text_width(center, 9.0);
        append_pdf_text_segment(ops, center, page_width * 0.5 - width * 0.5, y);
    }
    if segments.len() > 2 {
        let right = segments[2..].join(" ");
        let width = estimate_pdf_text_width(&right, 9.0);
        append_pdf_text_segment(ops, &right, page_width - right_margin - width, y);
    }
}

fn append_pdf_text_segment(ops: &mut Vec<Op>, text: &str, x: f32, y: f32) {
    if text.is_empty() {
        return;
    }

    ops.extend([
        Op::StartTextSection,
        Op::SetFillColor {
            col: Color::Rgb(printpdf::Rgb::new(0.35, 0.38, 0.43, None)),
        },
        Op::SetFont {
            font: PdfFontHandle::Builtin(BuiltinFont::Helvetica),
            size: Pt(9.0),
        },
        Op::SetTextCursor {
            pos: Point { x: Pt(x), y: Pt(y) },
        },
        Op::ShowText {
            items: vec![TextItem::Text(text.to_owned())],
        },
        Op::EndTextSection,
    ]);
}

fn estimate_pdf_text_width(text: &str, font_size: f32) -> f32 {
    text.chars().count() as f32 * font_size * 0.5
}

pub(super) fn plain_text_from_runs(runs: &[TextRun]) -> String {
    runs.iter().map(|run| run.text.as_str()).collect()
}

fn table_to_plain_text(table: &DocumentTable) -> String {
    let mut lines = Vec::new();
    for row in &table.rows {
        let cells: Vec<String> = row
            .iter()
            .map(|cell| {
                cell.plain_text()
                    .chars()
                    .filter(|ch| *ch != OBJECT_REPLACEMENT_CHAR)
                    .collect()
            })
            .collect();
        lines.push(cells.join("\t"));
    }
    lines.join("\n")
}

fn table_to_markdown(table: &DocumentTable) -> String {
    if table.rows.is_empty() {
        return String::new();
    }
    let num_cols = table.num_cols();
    let mut lines = Vec::new();

    // Header row (first row)
    let header: Vec<String> = table
        .rows
        .first()
        .map(|row| row.iter().map(markdown_text_from_table_cell).collect())
        .unwrap_or_default();
    lines.push(format!("| {} |", header.join(" | ")));

    // Separator
    let separator: Vec<&str> = (0..num_cols).map(|_| "---").collect();
    lines.push(format!("| {} |", separator.join(" | ")));

    // Data rows
    for row in table.rows.iter().skip(1) {
        let cells: Vec<String> = row.iter().map(markdown_text_from_table_cell).collect();
        // Pad if row has fewer cells
        let mut padded = cells;
        while padded.len() < num_cols {
            padded.push(String::new());
        }
        lines.push(format!("| {} |", padded.join(" | ")));
    }

    lines.join("\n")
}

fn table_to_html(table: &DocumentTable) -> String {
    let mut html = String::new();
    let border_color = css_color(table.borders.color);
    let border_width = table.borders.width_points;
    let _ = write!(
        html,
        "<table style=\"border-collapse:collapse;margin:4pt 0;width:100%;\">"
    );
    for (row_idx, row) in table.rows.iter().enumerate() {
        html.push_str("<tr>");
        for (col_idx, cell) in row.iter().enumerate() {
            let col_width = table
                .col_widths_points
                .get(col_idx)
                .copied()
                .unwrap_or(72.0);
            let tag = if row_idx == 0 { "th" } else { "td" };
            let _ = write!(
                html,
                "<{tag} style=\"border:{border_width:.1}pt solid {border_color};padding:4pt 6pt;width:{col_width:.1}pt;\">"
            );
            for run in &cell.runs {
                let text: String = run
                    .text
                    .chars()
                    .filter(|ch| *ch != OBJECT_REPLACEMENT_CHAR)
                    .collect();
                if !text.is_empty() {
                    let _ = write!(
                        html,
                        "<span style=\"{}\">{}</span>",
                        run_style_css(run.style),
                        html_escape(&text)
                    );
                }
            }
            for image in &cell.images {
                if let Some(mime_type) = image_mime_type(&image.bytes) {
                    let _ = write!(
                        html,
                        "<img class=\"image-block\" alt=\"{}\" src=\"data:{};base64,{}\" style=\"width:{}pt;height:{}pt;opacity:{:.3};\" />",
                        html_escape(&image.alt_text),
                        mime_type,
                        BASE64_STANDARD.encode(&image.bytes),
                        image.width_points,
                        image.height_points,
                        image.opacity.clamp(0.0, 1.0)
                    );
                }
            }
            let _ = write!(html, "</{tag}>");
        }
        html.push_str("</tr>");
    }
    html.push_str("</table>");
    html
}

fn markdown_text_from_table_cell(cell: &TableCell) -> String {
    let mut text = markdown_text_from_runs(&cell.runs);
    for image in &cell.images {
        let alt = if image.alt_text.is_empty() {
            "Image"
        } else {
            &image.alt_text
        };
        if !text.is_empty() {
            text.push(' ');
        }
        text.push_str(&format!("![{alt}](embedded-image)"));
    }
    text
}

fn markdown_text_from_runs(runs: &[TextRun]) -> String {
    let mut output = String::new();
    for run in runs {
        let mut text: String = run
            .text
            .chars()
            .filter(|ch| *ch != OBJECT_REPLACEMENT_CHAR)
            .collect();
        if text.is_empty() {
            continue;
        }
        if FontChoice::from_style(run.style).is_monospace() {
            text = format!("`{text}`");
        }
        if run.style.bold {
            text = format!("**{text}**");
        }
        if run.style.italic {
            text = format!("*{text}*");
        }
        if run.style.strikethrough {
            text = format!("~~{text}~~");
        }
        if run.style.underline {
            text = format!("<u>{text}</u>");
        }
        output.push_str(&text);
    }
    output
}

fn points_to_mm(points: f32) -> f32 {
    points * (25.4 / 72.0)
}

fn points_to_css_px(points: f32) -> f32 {
    points * (96.0 / 72.0)
}

fn paragraph_alignment_css(alignment: ParagraphAlignment) -> &'static str {
    match alignment {
        ParagraphAlignment::Left => "left",
        ParagraphAlignment::Center => "center",
        ParagraphAlignment::Right => "right",
        ParagraphAlignment::Justify => "justify",
    }
}

fn line_spacing_css(line_spacing: LineSpacing) -> String {
    match line_spacing.kind {
        LineSpacingKind::AutoMultiplier => {
            format!("line-height:{:.3};", line_spacing.value.max(0.1))
        }
        LineSpacingKind::AtLeastPoints | LineSpacingKind::ExactPoints => {
            format!("line-height:{:.3}pt;", line_spacing.value.max(1.0))
        }
    }
}

fn line_spacing_css_pdf(line_spacing: LineSpacing) -> String {
    match line_spacing.kind {
        LineSpacingKind::AutoMultiplier => {
            format!("line-height:{:.3};", line_spacing.value.max(0.1))
        }
        LineSpacingKind::AtLeastPoints | LineSpacingKind::ExactPoints => {
            format!(
                "line-height:{:.2}px;",
                points_to_css_px(line_spacing.value.max(1.0))
            )
        }
    }
}

fn run_style_css(style: CharacterStyle) -> String {
    let mut css = format!(
        "font-family:{};font-size:{:.2}pt;color:{};",
        css_font_family(style),
        style.font_size_points.max(1.0),
        css_color(style.text_color)
    );
    if style.bold {
        css.push_str("font-weight:700;");
    }
    if style.italic {
        css.push_str("font-style:italic;");
    }
    if style.highlight_color != Color32::TRANSPARENT {
        let _ = write!(
            css,
            "background-color:{};",
            css_color(style.highlight_color)
        );
    }
    let decoration = text_decoration_css(style);
    if !decoration.is_empty() {
        let _ = write!(css, "text-decoration:{};", decoration);
    }
    css
}

fn run_style_css_pdf(style: CharacterStyle) -> String {
    let font_points = if style.bold {
        style.font_size_points + 0.8
    } else {
        style.font_size_points
    };

    let mut css = format!(
        "font-family:{};font-size:{:.2}px;color:{};",
        css_font_family(style),
        points_to_css_px(font_points.max(1.0)),
        css_color_rgb(style.text_color)
    );
    if style.italic {
        css.push_str("font-style:italic;");
    }
    if style.highlight_color != Color32::TRANSPARENT {
        let _ = write!(
            css,
            "background-color:{};",
            css_color_rgb(style.highlight_color)
        );
    }
    let decoration = text_decoration_css(style);
    if !decoration.is_empty() {
        let _ = write!(css, "text-decoration:{};", decoration);
    }
    css
}

fn text_decoration_css(style: CharacterStyle) -> &'static str {
    match (style.underline, style.strikethrough) {
        (true, true) => "underline line-through",
        (true, false) => "underline",
        (false, true) => "line-through",
        (false, false) => "",
    }
}

fn css_font_family(style: CharacterStyle) -> String {
    match style.font_family_name {
        Some("docx-carlito") => "Carlito, Calibri, sans-serif".to_owned(),
        Some("docx-caladea") => "Caladea, Cambria, serif".to_owned(),
        Some("docx-liberation-sans") => "\"Liberation Sans\", Arial, sans-serif".to_owned(),
        Some("docx-liberation-serif") => {
            "\"Liberation Serif\", \"Times New Roman\", serif".to_owned()
        }
        Some("docx-liberation-mono") => {
            "\"Liberation Mono\", \"Courier New\", Consolas, monospace".to_owned()
        }
        Some("docx-comic-sans") => "\"Comic Neue\", \"Comic Sans MS\", cursive".to_owned(),
        Some(name) => format!("\"{}\", sans-serif", name.replace('"', "\\\"")),
        None => match style.font_choice {
            FontChoice::Proportional => "sans-serif".to_owned(),
            FontChoice::Monospace => "monospace".to_owned(),
            FontChoice::Carlito => "Carlito, Calibri, sans-serif".to_owned(),
            FontChoice::Caladea => "Caladea, Cambria, serif".to_owned(),
            FontChoice::LiberationSans => "\"Liberation Sans\", Arial, sans-serif".to_owned(),
            FontChoice::LiberationSerif => {
                "\"Liberation Serif\", \"Times New Roman\", serif".to_owned()
            }
            FontChoice::LiberationMono => {
                "\"Liberation Mono\", \"Courier New\", Consolas, monospace".to_owned()
            }
            FontChoice::ComicSans => "\"Comic Neue\", \"Comic Sans MS\", cursive".to_owned(),
        },
    }
}

fn css_color(color: Color32) -> String {
    format!(
        "rgba({}, {}, {}, {:.3})",
        color.r(),
        color.g(),
        color.b(),
        (color.a() as f32 / 255.0).clamp(0.0, 1.0)
    )
}

fn css_color_rgb(color: Color32) -> String {
    format!("rgb({}, {}, {})", color.r(), color.g(), color.b())
}

fn image_mime_type(bytes: &[u8]) -> Option<&'static str> {
    match image::guess_format(bytes) {
        Ok(image::ImageFormat::Png) => Some("image/png"),
        Ok(image::ImageFormat::Jpeg) => Some("image/jpeg"),
        Ok(image::ImageFormat::Gif) => Some("image/gif"),
        Ok(image::ImageFormat::Bmp) => Some("image/bmp"),
        _ => None,
    }
}

fn wrap_text_for_pdf(text: &str, max_chars: usize) -> Vec<String> {
    if text.is_empty() {
        return Vec::new();
    }

    let mut out = Vec::new();
    let mut current = String::new();
    for word in text.split_whitespace() {
        let current_len = current.chars().count();
        let word_len = word.chars().count();
        let projected = if current.is_empty() {
            word_len
        } else {
            current_len + 1 + word_len
        };

        if projected > max_chars && !current.is_empty() {
            out.push(std::mem::take(&mut current));
        }

        if !current.is_empty() {
            current.push(' ');
        }
        current.push_str(word);
    }

    if !current.is_empty() {
        out.push(current);
    }
    out
}

fn image_position_css(image: &DocumentImage) -> String {
    if image.layout_mode == ImageLayoutMode::Floating {
        format!(
            "position:relative;left:{:.2}pt;top:{:.2}pt;z-index:{};",
            image.offset_x_points(),
            image.offset_y_points(),
            image.z_index
        )
    } else {
        String::new()
    }
}

fn image_position_css_pdf(image: &DocumentImage) -> String {
    if image.layout_mode == ImageLayoutMode::Floating {
        format!(
            "position:relative;left:{:.2}px;top:{:.2}px;z-index:{};",
            points_to_css_px(image.offset_x_points()),
            points_to_css_px(image.offset_y_points()),
            image.z_index
        )
    } else {
        String::new()
    }
}

fn html_escape(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&#39;"),
            _ => out.push(ch),
        }
    }
    out
}
