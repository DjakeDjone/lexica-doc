use std::fmt::Write as _;

use base64::{engine::general_purpose::STANDARD as BASE64_STANDARD, Engine as _};

use crate::document::{
    DocumentState, DocumentTable, ListKind, OBJECT_REPLACEMENT_CHAR,
};
use super::{
    css_color, html_escape, image_mime_type, image_position_css, line_spacing_css,
    paragraph_alignment_css, run_style_css,
};

impl DocumentState {
    pub(crate) fn to_html(&self) -> String {
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
