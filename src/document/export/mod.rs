pub mod html;
pub mod markdown;
pub mod pdf;
pub mod text;

#[allow(unused_imports)]
pub(crate) use text::plain_text_from_runs;

use std::{fmt::Write as _, fs, path::Path};

use eframe::egui::Color32;

use super::{
    docx::{document_to_docx, docx_to_document},
    markdown::import_markdown,
    odt::{document_to_odt, odt_to_document},
    CharacterStyle, DocumentImage, DocumentState, FontChoice, ImageLayoutMode, LineSpacing,
    LineSpacingKind, ParagraphAlignment, TextRun, VerticalAlign,
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
                let source_bytes = fs::read(path)
                    .map_err(|error| format!("failed to read {}: {error}", path.display()))?;
                let imported = docx_to_document(&source_bytes)?;

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
                document.remember_source_docx(source_bytes)?;
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
                let imported = import_markdown(&source);
                let mut document = Self::bootstrap();
                document.replace_with_runs(title, imported.runs);
                document.paragraph_tables = imported.paragraph_tables;
                document.ensure_paragraph_style_count();
                return Ok(document);
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
            "docx" => document_to_docx(self),
            "odt" => document_to_odt(self),
            other => Err(format!(
                "saving .{other} is not supported yet; use .txt, .md, .html, .docx, .odt, or .pdf"
            )),
        }
    }
}

pub(crate) fn points_to_mm(points: f32) -> f32 {
    points * (25.4 / 72.0)
}

pub(crate) fn points_to_css_px(points: f32) -> f32 {
    points * (96.0 / 72.0)
}

pub(crate) fn paragraph_alignment_css(alignment: ParagraphAlignment) -> &'static str {
    match alignment {
        ParagraphAlignment::Left => "left",
        ParagraphAlignment::Center => "center",
        ParagraphAlignment::Right => "right",
        ParagraphAlignment::Justify => "justify",
    }
}

pub(crate) fn line_spacing_css(line_spacing: LineSpacing) -> String {
    match line_spacing.kind {
        LineSpacingKind::AutoMultiplier => {
            format!("line-height:{:.3};", line_spacing.value.max(0.1))
        }
        LineSpacingKind::AtLeastPoints | LineSpacingKind::ExactPoints => {
            format!("line-height:{:.3}pt;", line_spacing.value.max(1.0))
        }
    }
}

pub(crate) fn line_spacing_css_pdf(line_spacing: LineSpacing) -> String {
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

pub(crate) fn run_style_css(style: CharacterStyle) -> String {
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
    match style.vertical_align {
        VerticalAlign::Baseline => {}
        VerticalAlign::Superscript => css.push_str("vertical-align:super;font-size:65%;"),
        VerticalAlign::Subscript => css.push_str("vertical-align:sub;font-size:65%;"),
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

pub(crate) fn run_style_css_pdf(style: CharacterStyle) -> String {
    let mut css = format!(
        "white-space:pre-wrap;font-family:'{}';font-size:{:.2}px;color:{};",
        css_font_family_pdf(style),
        points_to_css_px(style.font_size_points.max(1.0)),
        css_color_rgb(style.text_color)
    );
    if style.bold {
        css.push_str("font-weight:bold;");
    }
    if style.italic {
        css.push_str("font-style:italic;");
    }
    match style.vertical_align {
        VerticalAlign::Baseline => {}
        VerticalAlign::Superscript => css.push_str("vertical-align:super;font-size:65%;"),
        VerticalAlign::Subscript => css.push_str("vertical-align:sub;font-size:65%;"),
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

pub(crate) fn text_decoration_css(style: CharacterStyle) -> &'static str {
    match (style.underline, style.strikethrough) {
        (true, true) => "underline line-through",
        (true, false) => "underline",
        (false, true) => "line-through",
        (false, false) => "",
    }
}

pub(crate) fn css_font_family(style: CharacterStyle) -> String {
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

pub(crate) fn css_font_family_pdf(style: CharacterStyle) -> String {
    let suffix = if style.bold { "-Bold" } else { "-Regular" };
    match style.font_family_name {
        Some("docx-carlito") => format!("Carlito{}", suffix),
        Some("docx-caladea") => format!("Caladea{}", suffix),
        Some("docx-liberation-sans") => format!("LiberationSans{}", suffix),
        Some("docx-liberation-serif") => format!("LiberationSerif{}", suffix),
        Some("docx-liberation-mono") => format!("LiberationMono{}", suffix),
        Some("docx-comic-sans") => format!("ComicNeue{}", suffix),
        Some(_) => format!("LiberationSans{}", suffix),
        None => match style.font_choice {
            FontChoice::Monospace => format!("LiberationMono{}", suffix),
            FontChoice::Carlito => format!("Carlito{}", suffix),
            FontChoice::Caladea => format!("Caladea{}", suffix),
            FontChoice::LiberationSerif => format!("LiberationSerif{}", suffix),
            FontChoice::LiberationMono => format!("LiberationMono{}", suffix),
            FontChoice::ComicSans => format!("ComicNeue{}", suffix),
            _ => format!("LiberationSans{}", suffix),
        },
    }
}

pub(crate) fn css_color(color: Color32) -> String {
    format!(
        "rgba({}, {}, {}, {:.3})",
        color.r(),
        color.g(),
        color.b(),
        (color.a() as f32 / 255.0).clamp(0.0, 1.0)
    )
}

pub(crate) fn css_color_rgb(color: Color32) -> String {
    format!("rgb({}, {}, {})", color.r(), color.g(), color.b())
}

pub(crate) fn image_mime_type(bytes: &[u8]) -> Option<&'static str> {
    match image::guess_format(bytes) {
        Ok(image::ImageFormat::Png) => Some("image/png"),
        Ok(image::ImageFormat::Jpeg) => Some("image/jpeg"),
        Ok(image::ImageFormat::Gif) => Some("image/gif"),
        Ok(image::ImageFormat::Bmp) => Some("image/bmp"),
        _ => None,
    }
}

pub(crate) fn wrap_text_for_pdf(text: &str, max_chars: usize) -> Vec<String> {
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

pub(crate) fn image_position_css(image: &DocumentImage) -> String {
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

pub(crate) fn image_position_css_pdf(image: &DocumentImage) -> String {
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

pub(crate) fn html_escape(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&#39;"),
            c if c.is_ascii() => out.push(c),
            c => {
                let _ = write!(out, "&#{};", c as u32);
            }
        }
    }
    out
}

/// HTML escape for PDF export: keeps non-ASCII characters as raw UTF-8
/// because printpdf's HTML parser doesn't properly decode numeric HTML entities,
/// but the embedded TTF fonts handle Unicode glyphs directly.
pub(crate) fn html_escape_pdf(text: &str) -> String {
    let mut out = String::with_capacity(text.len());
    for ch in text.chars() {
        match ch {
            '&' => out.push_str("&amp;"),
            '<' => out.push_str("&lt;"),
            '>' => out.push_str("&gt;"),
            '"' => out.push_str("&quot;"),
            '\'' => out.push_str("&#39;"),
            c => out.push(c),
        }
    }
    out
}
