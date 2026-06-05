use std::collections::HashMap;

use quick_xml::{events::Event as XmlEvent, Reader};

use crate::document::ListKind;
use super::{attr_value, local_name};

#[derive(Default)]
pub(crate) struct NumberingDefinitions {
    pub num_to_abstract: HashMap<String, String>,
    pub level_kinds: HashMap<(String, String), ListKind>,
}

impl NumberingDefinitions {
    pub(crate) fn lookup(&self, num_id: Option<&str>, ilvl: Option<&str>) -> ListKind {
        let Some(num_id) = num_id else {
            return ListKind::None;
        };
        if num_id == "0" {
            return ListKind::None;
        }

        let Some(abstract_id) = self.num_to_abstract.get(num_id) else {
            return ListKind::None;
        };
        let level = ilvl.unwrap_or("0");
        self.level_kinds
            .get(&(abstract_id.clone(), level.to_owned()))
            .copied()
            .or_else(|| {
                self.level_kinds
                    .get(&(abstract_id.clone(), "0".to_owned()))
                    .copied()
            })
            .unwrap_or(ListKind::None)
    }
}

pub(crate) fn parse_numbering_xml(numbering_xml: &str) -> Result<NumberingDefinitions, String> {
    let mut reader = Reader::from_str(numbering_xml);
    reader.config_mut().trim_text(false);

    let mut numbering = NumberingDefinitions::default();
    let mut current_abstract = None::<String>;
    let mut current_level = None::<String>;
    let mut current_num = None::<String>;

    loop {
        match reader.read_event() {
            Ok(XmlEvent::Start(event)) => match local_name(event.name().as_ref()) {
                b"abstractNum" => current_abstract = attr_value(&event, b"abstractNumId"),
                b"lvl" => current_level = attr_value(&event, b"ilvl"),
                b"num" => current_num = attr_value(&event, b"numId"),
                b"numFmt" => {
                    if let (Some(abstract_id), Some(level), Some(value)) = (
                        current_abstract.as_ref(),
                        current_level.as_ref(),
                        attr_value(&event, b"val"),
                    ) {
                        numbering.level_kinds.insert(
                            (abstract_id.clone(), level.clone()),
                            list_kind_for_numbering(&value),
                        );
                    }
                }
                b"abstractNumId" => {
                    if let (Some(num_id), Some(abstract_id)) =
                        (current_num.as_ref(), attr_value(&event, b"val"))
                    {
                        numbering
                            .num_to_abstract
                            .insert(num_id.clone(), abstract_id);
                    }
                }
                _ => {}
            },
            Ok(XmlEvent::Empty(event)) => match local_name(event.name().as_ref()) {
                b"abstractNum" => current_abstract = attr_value(&event, b"abstractNumId"),
                b"lvl" => current_level = attr_value(&event, b"ilvl"),
                b"num" => current_num = attr_value(&event, b"numId"),
                b"numFmt" => {
                    if let (Some(abstract_id), Some(level), Some(value)) = (
                        current_abstract.as_ref(),
                        current_level.as_ref(),
                        attr_value(&event, b"val"),
                    ) {
                        numbering.level_kinds.insert(
                            (abstract_id.clone(), level.clone()),
                            list_kind_for_numbering(&value),
                        );
                    }
                }
                b"abstractNumId" => {
                    if let (Some(num_id), Some(abstract_id)) =
                        (current_num.as_ref(), attr_value(&event, b"val"))
                    {
                        numbering
                            .num_to_abstract
                            .insert(num_id.clone(), abstract_id);
                    }
                }
                _ => {}
            },
            Ok(XmlEvent::End(event)) => match local_name(event.name().as_ref()) {
                b"abstractNum" => current_abstract = None,
                b"lvl" => current_level = None,
                b"num" => current_num = None,
                _ => {}
            },
            Ok(XmlEvent::Eof) => break,
            Err(error) => return Err(format!("failed to parse word/numbering.xml: {error}")),
            _ => {}
        }
    }

    Ok(numbering)
}

fn list_kind_for_numbering(value: &str) -> ListKind {
    match value {
        "bullet" => ListKind::Bullet,
        "none" => ListKind::None,
        _ => ListKind::Ordered,
    }
}
