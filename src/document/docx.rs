use std::io::{Cursor, Read};

use eframe::egui::Color32;
use quick_xml::{events::Event as XmlEvent, Reader};
use zip::ZipArchive;

use crate::document::{CharacterStyle, TextRun};

pub(super) fn docx_to_runs(bytes: &[u8]) -> Result<Vec<TextRun>, String> {
    let cursor = Cursor::new(bytes);
    let mut archive =
        ZipArchive::new(cursor).map_err(|error| format!("invalid .docx archive: {error}"))?;
    let mut document_xml = String::new();
    archive
        .by_name("word/document.xml")
        .map_err(|error| format!("missing word/document.xml: {error}"))?
        .read_to_string(&mut document_xml)
        .map_err(|error| format!("failed to read word/document.xml: {error}"))?;

    let mut reader = Reader::from_str(&document_xml);
    reader.config_mut().trim_text(false);

    let mut runs = Vec::new();
    let mut run_style = CharacterStyle::default();
    let mut in_text = false;
    let mut paragraph_count = 0usize;

    loop {
        match reader.read_event() {
            Ok(XmlEvent::Start(event)) => match local_name(event.name().as_ref()) {
                b"p" => {
                    if paragraph_count > 0 {
                        append_plain(&mut runs, "\n\n", CharacterStyle::default());
                    }
                    paragraph_count += 1;
                }
                b"r" => {
                    run_style = CharacterStyle::default();
                }
                b"t" => in_text = true,
                b"br" => append_plain(&mut runs, "\n", run_style),
                b"tab" => append_plain(&mut runs, "\t", run_style),
                b"b" => run_style.bold = docx_flag(&event, true),
                b"i" => run_style.italic = docx_flag(&event, true),
                b"u" => {
                    run_style.underline =
                        !matches!(attr_value(&event, b"val").as_deref(), Some("none"))
                }
                b"strike" | b"dstrike" => run_style.strikethrough = docx_flag(&event, true),
                b"sz" => {
                    if let Some(value) = attr_value(&event, b"val") {
                        if let Ok(half_points) = value.parse::<f32>() {
                            run_style.font_size_points = (half_points / 2.0).clamp(8.0, 72.0);
                        }
                    }
                }
                b"color" => {
                    if let Some(value) = attr_value(&event, b"val") {
                        if let Some(color) = parse_hex_color(&value) {
                            run_style.text_color = color;
                        }
                    }
                }
                b"highlight" => {
                    if let Some(value) = attr_value(&event, b"val") {
                        run_style.highlight_color = highlight_color(&value);
                    }
                }
                _ => {}
            },
            Ok(XmlEvent::Empty(event)) => match local_name(event.name().as_ref()) {
                b"br" => append_plain(&mut runs, "\n", run_style),
                b"tab" => append_plain(&mut runs, "\t", run_style),
                b"b" => run_style.bold = docx_flag(&event, true),
                b"i" => run_style.italic = docx_flag(&event, true),
                b"u" => {
                    run_style.underline =
                        !matches!(attr_value(&event, b"val").as_deref(), Some("none"))
                }
                b"strike" | b"dstrike" => run_style.strikethrough = docx_flag(&event, true),
                b"sz" => {
                    if let Some(value) = attr_value(&event, b"val") {
                        if let Ok(half_points) = value.parse::<f32>() {
                            run_style.font_size_points = (half_points / 2.0).clamp(8.0, 72.0);
                        }
                    }
                }
                b"color" => {
                    if let Some(value) = attr_value(&event, b"val") {
                        if let Some(color) = parse_hex_color(&value) {
                            run_style.text_color = color;
                        }
                    }
                }
                b"highlight" => {
                    if let Some(value) = attr_value(&event, b"val") {
                        run_style.highlight_color = highlight_color(&value);
                    }
                }
                _ => {}
            },
            Ok(XmlEvent::Text(text)) => {
                if in_text {
                    let decoded = text
                        .xml_content()
                        .map_err(|error| format!("failed to decode document text: {error}"))?;
                    append_plain(&mut runs, decoded.as_ref(), run_style);
                }
            }
            Ok(XmlEvent::End(event)) => {
                if local_name(event.name().as_ref()) == b"t" {
                    in_text = false;
                }
            }
            Ok(XmlEvent::Eof) => break,
            Err(error) => return Err(format!("failed to parse word/document.xml: {error}")),
            _ => {}
        }
    }

    if runs.is_empty() {
        runs.push(TextRun {
            text: String::new(),
            style: CharacterStyle::default(),
        });
    }

    Ok(runs)
}

fn append_plain(runs: &mut Vec<TextRun>, text: &str, style: CharacterStyle) {
    if text.is_empty() {
        return;
    }

    if let Some(last) = runs.last_mut() {
        if last.style == style {
            last.text.push_str(text);
            return;
        }
    }

    runs.push(TextRun {
        text: text.to_owned(),
        style,
    });
}

fn local_name(name: &[u8]) -> &[u8] {
    name.rsplit(|byte| *byte == b':').next().unwrap_or(name)
}

fn attr_value(event: &quick_xml::events::BytesStart<'_>, key: &[u8]) -> Option<String> {
    event
        .attributes()
        .flatten()
        .find(|attr| local_name(attr.key.as_ref()) == key)
        .and_then(|attr| String::from_utf8(attr.value.into_owned()).ok())
}

fn docx_flag(event: &quick_xml::events::BytesStart<'_>, default: bool) -> bool {
    match attr_value(event, b"val").as_deref() {
        Some("0" | "false") => false,
        Some("1" | "true") => true,
        Some(_) => default,
        None => default,
    }
}

fn parse_hex_color(value: &str) -> Option<Color32> {
    if value.len() != 6 {
        return None;
    }

    let red = u8::from_str_radix(&value[0..2], 16).ok()?;
    let green = u8::from_str_radix(&value[2..4], 16).ok()?;
    let blue = u8::from_str_radix(&value[4..6], 16).ok()?;
    Some(Color32::from_rgb(red, green, blue))
}

fn highlight_color(value: &str) -> Color32 {
    match value {
        "yellow" => Color32::from_rgb(255, 242, 129),
        "green" => Color32::from_rgb(187, 232, 172),
        "cyan" => Color32::from_rgb(163, 231, 240),
        "magenta" => Color32::from_rgb(244, 188, 231),
        "blue" => Color32::from_rgb(177, 205, 252),
        "red" => Color32::from_rgb(248, 188, 188),
        "darkYellow" => Color32::from_rgb(215, 185, 90),
        "darkGreen" => Color32::from_rgb(104, 170, 112),
        "darkBlue" => Color32::from_rgb(99, 129, 207),
        _ => Color32::TRANSPARENT,
    }
}
