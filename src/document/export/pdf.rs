use std::{collections::BTreeMap, fmt::Write as _};

use base64::{engine::general_purpose::STANDARD as BASE64_STANDARD, Engine as _};
use printpdf::{
    Base64OrRaw, BuiltinFont, Color, GeneratePdfOptions, Op, PdfDocument, PdfFontHandle, PdfPage,
    PdfSaveOptions, Point, Pt, TextItem,
};

use crate::document::{
    DocumentState, HeaderFooterKind, ListKind, OBJECT_REPLACEMENT_CHAR,
};
use super::{
    html_escape, image_mime_type, image_position_css_pdf, line_spacing_css_pdf,
    paragraph_alignment_css, points_to_css_px, points_to_mm, run_style_css_pdf,
    wrap_text_for_pdf,
};

impl DocumentState {
    pub(crate) fn to_pdf_bytes(&self) -> Result<Vec<u8>, String> {
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

    pub(crate) fn to_pdf_html(&self) -> String {
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

    pub(crate) fn to_plain_text_pdf_bytes(&self) -> Vec<u8> {
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
