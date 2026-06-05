use std::collections::{HashMap, HashSet};

use eframe::egui::Color32;
use quick_xml::{events::Event as XmlEvent, Reader};

use crate::document::{
    CharacterStyle, FontChoice, ParagraphStyle, ParagraphAlignment, LineSpacing,
};
use super::{
    attr_value, local_name, docx_flag, parse_hex_color, highlight_color,
    twips_to_points, parse_line_spacing,
};

pub(crate) const DOCX_CARLITO: &str = "docx-carlito";
pub(crate) const DOCX_CALADEA: &str = "docx-caladea";
pub(crate) const DOCX_LIBERATION_SANS: &str = "docx-liberation-sans";
pub(crate) const DOCX_LIBERATION_SERIF: &str = "docx-liberation-serif";
pub(crate) const DOCX_LIBERATION_MONO: &str = "docx-liberation-mono";
pub(crate) const DOCX_COMIC_SANS: &str = "docx-comic-sans";

#[derive(Clone, Copy, Debug, PartialEq, Eq)]
pub(crate) struct ResolvedFont {
    pub(crate) font_family_name: Option<&'static str>,
    pub(crate) font_choice: FontChoice,
}

#[derive(Default)]
pub(crate) struct ThemeFonts {
    pub(crate) major_latin: Option<String>,
    pub(crate) minor_latin: Option<String>,
}

#[derive(Clone, Copy, Default)]
pub(crate) struct CharacterStylePatch {
    pub(crate) bold: Option<bool>,
    pub(crate) italic: Option<bool>,
    pub(crate) underline: Option<bool>,
    pub(crate) strikethrough: Option<bool>,
    pub(crate) font_size_points: Option<f32>,
    pub(crate) font: Option<ResolvedFont>,
    pub(crate) text_color: Option<Color32>,
    pub(crate) highlight_color: Option<Color32>,
}

impl CharacterStylePatch {
    pub(crate) fn apply(self, mut style: CharacterStyle) -> CharacterStyle {
        if let Some(value) = self.bold {
            style.bold = value;
        }
        if let Some(value) = self.italic {
            style.italic = value;
        }
        if let Some(value) = self.underline {
            style.underline = value;
        }
        if let Some(value) = self.strikethrough {
            style.strikethrough = value;
        }
        if let Some(value) = self.font_size_points {
            style.font_size_points = value;
        }
        if let Some(value) = self.font {
            style.font_family_name = value.font_family_name;
            style.font_choice = value.font_choice;
        }
        if let Some(value) = self.text_color {
            style.text_color = value;
        }
        if let Some(value) = self.highlight_color {
            style.highlight_color = value;
        }
        style
    }

    pub(crate) fn overlay(&mut self, other: Self) {
        if other.bold.is_some() {
            self.bold = other.bold;
        }
        if other.italic.is_some() {
            self.italic = other.italic;
        }
        if other.underline.is_some() {
            self.underline = other.underline;
        }
        if other.strikethrough.is_some() {
            self.strikethrough = other.strikethrough;
        }
        if other.font_size_points.is_some() {
            self.font_size_points = other.font_size_points;
        }
        if other.font.is_some() {
            self.font = other.font;
        }
        if other.text_color.is_some() {
            self.text_color = other.text_color;
        }
        if other.highlight_color.is_some() {
            self.highlight_color = other.highlight_color;
        }
    }
}

#[derive(Clone, Copy, Default)]
pub(crate) struct ParagraphStylePatch {
    pub(crate) alignment: Option<ParagraphAlignment>,
    pub(crate) page_break_before: Option<bool>,
    pub(crate) spacing_before_points: Option<u16>,
    pub(crate) spacing_after_points: Option<u16>,
    pub(crate) line_spacing: Option<LineSpacing>,
}

impl ParagraphStylePatch {
    pub(crate) fn apply(self, mut style: ParagraphStyle) -> ParagraphStyle {
        if let Some(value) = self.alignment {
            style.alignment = value;
        }
        if let Some(value) = self.page_break_before {
            style.page_break_before = value;
        }
        if let Some(value) = self.spacing_before_points {
            style.spacing_before_points = value;
        }
        if let Some(value) = self.spacing_after_points {
            style.spacing_after_points = value;
        }
        if let Some(value) = self.line_spacing {
            style.line_spacing = value;
        }
        style
    }

    pub(crate) fn overlay(&mut self, other: Self) {
        if other.alignment.is_some() {
            self.alignment = other.alignment;
        }
        if other.page_break_before.is_some() {
            self.page_break_before = other.page_break_before;
        }
        if other.spacing_before_points.is_some() {
            self.spacing_before_points = other.spacing_before_points;
        }
        if other.spacing_after_points.is_some() {
            self.spacing_after_points = other.spacing_after_points;
        }
        if other.line_spacing.is_some() {
            self.line_spacing = other.line_spacing;
        }
    }
}

#[derive(Clone, Default)]
struct RawParagraphStyleDefinition {
    based_on: Option<String>,
    paragraph: ParagraphStylePatch,
    run: CharacterStylePatch,
}

#[derive(Clone, Default)]
struct RawCharacterStyleDefinition {
    based_on: Option<String>,
    run: CharacterStylePatch,
}

#[derive(Clone, Copy, Default)]
struct ResolvedParagraphStyle {
    paragraph: ParagraphStylePatch,
    run: CharacterStylePatch,
}

#[derive(Default)]
struct RawDocxStyles {
    default_paragraph: ParagraphStylePatch,
    default_run: CharacterStylePatch,
    paragraph_styles: HashMap<String, RawParagraphStyleDefinition>,
    character_styles: HashMap<String, RawCharacterStyleDefinition>,
}

#[derive(Default)]
pub(crate) struct DocxStyles {
    default_paragraph: ParagraphStyle,
    default_run: CharacterStyle,
    paragraph_styles: HashMap<String, ResolvedParagraphStyle>,
    character_styles: HashMap<String, CharacterStylePatch>,
}

impl DocxStyles {
    pub(crate) fn default_paragraph_style(&self) -> ParagraphStyle {
        self.default_paragraph
    }

    pub(crate) fn default_run_style(&self) -> CharacterStyle {
        self.default_run
    }

    pub(crate) fn apply_paragraph_style(&self, style_id: &str, style: &mut ParagraphStyle) {
        if let Some(resolved) = self.paragraph_styles.get(style_id) {
            *style = resolved.paragraph.apply(*style);
        }
    }

    pub(crate) fn run_style_for_paragraph(&self, style_id: &str) -> CharacterStyle {
        self.paragraph_styles
            .get(style_id)
            .map(|resolved| resolved.run.apply(self.default_run_style()))
            .unwrap_or_else(|| self.default_run_style())
    }

    pub(crate) fn apply_run_style(&self, style_id: &str, base: CharacterStyle) -> CharacterStyle {
        self.character_styles
            .get(style_id)
            .copied()
            .map(|style| style.apply(base))
            .unwrap_or(base)
    }
}

pub(crate) fn parse_theme_xml(theme_xml: &str) -> Result<ThemeFonts, String> {
    let mut reader = Reader::from_str(theme_xml);
    reader.config_mut().trim_text(false);

    let mut theme_fonts = ThemeFonts::default();
    let mut in_major_font = false;
    let mut in_minor_font = false;

    loop {
        match reader.read_event() {
            Ok(XmlEvent::Start(event)) | Ok(XmlEvent::Empty(event)) => {
                match local_name(event.name().as_ref()) {
                    b"majorFont" => in_major_font = true,
                    b"minorFont" => in_minor_font = true,
                    b"latin" => {
                        let typeface = attr_value(&event, b"typeface")
                            .filter(|value| !value.trim().is_empty());
                        if in_major_font {
                            theme_fonts.major_latin = typeface;
                        } else if in_minor_font {
                            theme_fonts.minor_latin = typeface;
                        }
                    }
                    _ => {}
                }
            }
            Ok(XmlEvent::End(event)) => match local_name(event.name().as_ref()) {
                b"majorFont" => in_major_font = false,
                b"minorFont" => in_minor_font = false,
                _ => {}
            },
            Ok(XmlEvent::Eof) => break,
            Err(error) => return Err(format!("failed to parse word/theme/theme1.xml: {error}")),
            _ => {}
        }
    }

    Ok(theme_fonts)
}

pub(crate) fn apply_resolved_font(style: &mut CharacterStyle, font: ResolvedFont) {
    style.font_family_name = font.font_family_name;
    style.font_choice = font.font_choice;
}

pub(crate) fn resolve_rfonts(
    event: &quick_xml::events::BytesStart<'_>,
    theme_fonts: &ThemeFonts,
) -> Option<ResolvedFont> {
    for key in [b"ascii".as_slice(), b"hAnsi", b"cs", b"eastAsia"] {
        if let Some(value) = attr_value(event, key).filter(|value| !value.trim().is_empty()) {
            return Some(resolve_font_name(&value));
        }
    }

    for key in [
        b"asciiTheme".as_slice(),
        b"hAnsiTheme",
        b"csTheme",
        b"eastAsiaTheme",
    ] {
        if let Some(value) = attr_value(event, key).filter(|value| !value.trim().is_empty()) {
            return resolve_theme_font(&value, theme_fonts);
        }
    }

    None
}

fn resolve_theme_font(slot: &str, theme_fonts: &ThemeFonts) -> Option<ResolvedFont> {
    let font_name = match slot {
        "majorAscii" | "majorHAnsi" | "majorBidi" | "majorEastAsia" => {
            theme_fonts.major_latin.as_deref()
        }
        "minorAscii" | "minorHAnsi" | "minorBidi" | "minorEastAsia" => {
            theme_fonts.minor_latin.as_deref()
        }
        _ => None,
    }?;
    Some(resolve_font_name(font_name))
}

pub(crate) fn resolve_font_name(name: &str) -> ResolvedFont {
    let normalized = name.trim().to_ascii_lowercase();
    let family_name = match normalized.as_str() {
        "calibri" | "calibri light" | "aptos" | "aptos display" => Some(DOCX_CARLITO),
        "cambria" => Some(DOCX_CALADEA),
        "arial" => Some(DOCX_LIBERATION_SANS),
        "times new roman" => Some(DOCX_LIBERATION_SERIF),
        "courier new" | "consolas" => Some(DOCX_LIBERATION_MONO),
        "comic sans" | "comic sans ms" | "comic neue" => Some(DOCX_COMIC_SANS),
        _ => None,
    };

    let monospace = matches!(
        normalized.as_str(),
        "courier new" | "consolas" | "menlo" | "monaco" | "source code pro"
    ) || normalized.contains("mono");

    ResolvedFont {
        font_family_name: family_name,
        font_choice: if monospace {
            FontChoice::Monospace
        } else {
            FontChoice::Proportional
        },
    }
}

pub(crate) fn parse_styles_xml(styles_xml: &str, theme_fonts: &ThemeFonts) -> Result<DocxStyles, String> {
    let mut reader = Reader::from_str(styles_xml);
    reader.config_mut().trim_text(false);

    let mut raw = RawDocxStyles::default();
    let mut current_style_id = None::<String>;
    let mut current_style_type = None::<String>;
    let mut current_paragraph_style = RawParagraphStyleDefinition::default();
    let mut current_character_style = RawCharacterStyleDefinition::default();
    let mut in_doc_defaults = false;
    let mut in_doc_defaults_paragraph = false;
    let mut in_doc_defaults_run = false;
    let mut in_style_paragraph = false;
    let mut in_style_run = false;

    loop {
        match reader.read_event() {
            Ok(XmlEvent::Start(event)) => match local_name(event.name().as_ref()) {
                b"docDefaults" => in_doc_defaults = true,
                b"style" => {
                    current_style_id = attr_value(&event, b"styleId");
                    current_style_type = attr_value(&event, b"type");
                    current_paragraph_style = RawParagraphStyleDefinition::default();
                    current_character_style = RawCharacterStyleDefinition::default();
                }
                b"basedOn" => {
                    let based_on = attr_value(&event, b"val");
                    match current_style_type.as_deref() {
                        Some("paragraph") => current_paragraph_style.based_on = based_on,
                        Some("character") => current_character_style.based_on = based_on,
                        _ => {}
                    }
                }
                b"pPr" => {
                    if in_doc_defaults {
                        in_doc_defaults_paragraph = true;
                    } else if current_style_type.as_deref() == Some("paragraph") {
                        in_style_paragraph = true;
                    }
                }
                b"rPr" => {
                    if in_doc_defaults {
                        in_doc_defaults_run = true;
                    } else if current_style_id.is_some() {
                        in_style_run = true;
                    }
                }
                name => {
                    if in_doc_defaults_paragraph {
                        apply_paragraph_style_patch_event(name, &event, &mut raw.default_paragraph);
                    }
                    if in_doc_defaults_run {
                        apply_run_style_patch_event(
                            name,
                            &event,
                            &mut raw.default_run,
                            theme_fonts,
                        );
                    }
                    if in_style_run {
                        match current_style_type.as_deref() {
                            Some("paragraph") => apply_run_style_patch_event(
                                name,
                                &event,
                                &mut current_paragraph_style.run,
                                theme_fonts,
                            ),
                            Some("character") => apply_run_style_patch_event(
                                name,
                                &event,
                                &mut current_character_style.run,
                                theme_fonts,
                            ),
                            _ => {}
                        }
                    }
                    if in_style_paragraph {
                        apply_paragraph_style_patch_event(
                            name,
                            &event,
                            &mut current_paragraph_style.paragraph,
                        );
                    }
                }
            },
            Ok(XmlEvent::Empty(event)) => match local_name(event.name().as_ref()) {
                b"basedOn" => {
                    let based_on = attr_value(&event, b"val");
                    match current_style_type.as_deref() {
                        Some("paragraph") => current_paragraph_style.based_on = based_on,
                        Some("character") => current_character_style.based_on = based_on,
                        _ => {}
                    }
                }
                name => {
                    if in_doc_defaults && name == b"rPr" {
                        continue;
                    }
                    if in_doc_defaults_paragraph {
                        apply_paragraph_style_patch_event(name, &event, &mut raw.default_paragraph);
                    }
                    if in_doc_defaults_run {
                        apply_run_style_patch_event(
                            name,
                            &event,
                            &mut raw.default_run,
                            theme_fonts,
                        );
                    }
                    if in_style_run {
                        match current_style_type.as_deref() {
                            Some("paragraph") => apply_run_style_patch_event(
                                name,
                                &event,
                                &mut current_paragraph_style.run,
                                theme_fonts,
                            ),
                            Some("character") => apply_run_style_patch_event(
                                name,
                                &event,
                                &mut current_character_style.run,
                                theme_fonts,
                            ),
                            _ => {}
                        }
                    }
                    if in_style_paragraph {
                        apply_paragraph_style_patch_event(
                            name,
                            &event,
                            &mut current_paragraph_style.paragraph,
                        );
                    }
                }
            },
            Ok(XmlEvent::End(event)) => match local_name(event.name().as_ref()) {
                b"docDefaults" => in_doc_defaults = false,
                b"rPr" => {
                    in_doc_defaults_run = false;
                    in_style_run = false;
                }
                b"pPr" => {
                    in_doc_defaults_paragraph = false;
                    in_style_paragraph = false;
                }
                b"style" => {
                    if let Some(style_id) = current_style_id.take() {
                        match current_style_type.as_deref() {
                            Some("paragraph") => {
                                raw.paragraph_styles
                                    .insert(style_id, current_paragraph_style.clone());
                            }
                            Some("character") => {
                                raw.character_styles
                                    .insert(style_id, current_character_style.clone());
                            }
                            _ => {}
                        }
                    }
                    current_style_type = None;
                }
                _ => {}
            },
            Ok(XmlEvent::Eof) => break,
            Err(error) => return Err(format!("failed to parse word/styles.xml: {error}")),
            _ => {}
        }
    }

    Ok(resolve_styles(raw))
}

fn resolve_styles(raw: RawDocxStyles) -> DocxStyles {
    let mut paragraph_styles = HashMap::new();
    let mut character_styles = HashMap::new();

    for style_id in raw.paragraph_styles.keys() {
        let mut active = HashSet::new();
        let resolved = resolve_paragraph_style(style_id, &raw, &mut active);
        paragraph_styles.insert(style_id.clone(), resolved);
    }

    for style_id in raw.character_styles.keys() {
        let mut active = HashSet::new();
        let resolved = resolve_character_style(style_id, &raw, &mut active);
        character_styles.insert(style_id.clone(), resolved);
    }

    DocxStyles {
        default_paragraph: raw.default_paragraph.apply(ParagraphStyle::default()),
        default_run: raw.default_run.apply(CharacterStyle::default()),
        paragraph_styles,
        character_styles,
    }
}

fn resolve_paragraph_style(
    style_id: &str,
    raw: &RawDocxStyles,
    active: &mut HashSet<String>,
) -> ResolvedParagraphStyle {
    if !active.insert(style_id.to_owned()) {
        return ResolvedParagraphStyle::default();
    }

    let Some(style) = raw.paragraph_styles.get(style_id) else {
        active.remove(style_id);
        return ResolvedParagraphStyle::default();
    };

    let mut resolved = if let Some(parent) = style.based_on.as_deref() {
        resolve_paragraph_style(parent, raw, active)
    } else {
        ResolvedParagraphStyle::default()
    };
    resolved.paragraph.overlay(style.paragraph);
    resolved.run.overlay(style.run);
    active.remove(style_id);
    resolved
}

fn resolve_character_style(
    style_id: &str,
    raw: &RawDocxStyles,
    active: &mut HashSet<String>,
) -> CharacterStylePatch {
    if !active.insert(style_id.to_owned()) {
        return CharacterStylePatch::default();
    }

    let Some(style) = raw.character_styles.get(style_id) else {
        active.remove(style_id);
        return CharacterStylePatch::default();
    };

    let mut resolved = if let Some(parent) = style.based_on.as_deref() {
        resolve_character_style(parent, raw, active)
    } else {
        CharacterStylePatch::default()
    };
    resolved.overlay(style.run);
    active.remove(style_id);
    resolved
}

pub(crate) fn apply_run_style_patch_event(
    name: &[u8],
    event: &quick_xml::events::BytesStart<'_>,
    patch: &mut CharacterStylePatch,
    theme_fonts: &ThemeFonts,
) {
    match name {
        b"rFonts" => patch.font = resolve_rfonts(event, theme_fonts),
        b"b" => patch.bold = Some(docx_flag(event, true)),
        b"i" => patch.italic = Some(docx_flag(event, true)),
        b"u" => {
            patch.underline = Some(!matches!(
                attr_value(event, b"val").as_deref(),
                Some("none")
            ))
        }
        b"strike" | b"dstrike" => patch.strikethrough = Some(docx_flag(event, true)),
        b"sz" => {
            if let Some(value) = attr_value(event, b"val") {
                if let Ok(half_points) = value.parse::<f32>() {
                    patch.font_size_points = Some((half_points / 2.0).clamp(8.0, 72.0));
                }
            }
        }
        b"color" => {
            if let Some(value) = attr_value(event, b"val") {
                patch.text_color = parse_hex_color(&value);
            }
        }
        b"highlight" => {
            if let Some(value) = attr_value(event, b"val") {
                patch.highlight_color = Some(highlight_color(&value));
            }
        }
        _ => {}
    }
}

pub(crate) fn apply_paragraph_style_patch_event(
    name: &[u8],
    event: &quick_xml::events::BytesStart<'_>,
    patch: &mut ParagraphStylePatch,
) {
    match name {
        b"jc" => {
            patch.alignment = Some(super::paragraph_alignment_for(
                attr_value(event, b"val").as_deref().unwrap_or_default(),
            ));
        }
        b"spacing" => apply_spacing_patch(event, patch),
        b"pageBreakBefore" => patch.page_break_before = Some(docx_flag(event, true)),
        _ => {}
    }
}

fn apply_spacing_patch(event: &quick_xml::events::BytesStart<'_>, patch: &mut ParagraphStylePatch) {
    if let Some(value) = attr_value(event, b"before")
        .and_then(|value| value.parse::<f32>().ok())
        .map(twips_to_points)
    {
        patch.spacing_before_points = Some(value.round().clamp(0.0, u16::MAX as f32) as u16);
    }
    if let Some(value) = attr_value(event, b"after")
        .and_then(|value| value.parse::<f32>().ok())
        .map(twips_to_points)
    {
        patch.spacing_after_points = Some(value.round().clamp(0.0, u16::MAX as f32) as u16);
    }
    if let Some(line_spacing) = parse_line_spacing(event) {
        patch.line_spacing = Some(line_spacing);
    }
}

pub(crate) fn resolve_font_from_event_without_theme(
    event: &quick_xml::events::BytesStart<'_>,
) -> Option<ResolvedFont> {
    for key in [b"ascii".as_slice(), b"hAnsi", b"cs", b"eastAsia"] {
        if let Some(value) = attr_value(event, key).filter(|value| !value.trim().is_empty()) {
            return Some(resolve_font_name(&value));
        }
    }
    None
}
