use std::ops::Range;

pub(super) fn line_char_range(text: &str, char_index: usize) -> Range<usize> {
    let total_chars = text.chars().count();
    let target = char_index.min(total_chars);
    let mut start = 0;
    let mut end = total_chars;

    for (index, ch) in text.chars().enumerate() {
        if index < target && ch == '\n' {
            start = index + 1;
        }
        if index >= target && ch == '\n' {
            end = index;
            break;
        }
    }

    start..end
}

pub(super) fn char_to_byte_index(text: &str, char_index: usize) -> usize {
    text.char_indices()
        .nth(char_index)
        .map(|(idx, _)| idx)
        .unwrap_or(text.len())
}

pub(super) fn slice_char_range(text: &str, range: Range<usize>) -> &str {
    let start = char_to_byte_index(text, range.start);
    let end = char_to_byte_index(text, range.end);
    &text[start..end]
}
