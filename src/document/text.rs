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

pub(super) fn word_char_range(text: &str, char_index: usize) -> Option<Range<usize>> {
    let chars: Vec<char> = text.chars().collect();
    let total_chars = chars.len();
    let target = char_index.min(total_chars);

    let seed = if target < total_chars && is_word_char(chars[target]) {
        target
    } else if target > 0 && is_word_char(chars[target - 1]) {
        target - 1
    } else {
        return None;
    };

    let mut start = seed;
    while start > 0 && is_word_char(chars[start - 1]) {
        start -= 1;
    }

    let mut end = seed + 1;
    while end < total_chars && is_word_char(chars[end]) {
        end += 1;
    }

    Some(start..end)
}

fn is_word_char(ch: char) -> bool {
    ch.is_alphanumeric() || ch == '_'
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

#[cfg(test)]
mod tests {
    use super::word_char_range;

    #[test]
    fn selects_word_when_cursor_is_inside_it() {
        assert_eq!(word_char_range("alpha beta", 2), Some(0..5));
        assert_eq!(word_char_range("alpha beta", 8), Some(6..10));
    }

    #[test]
    fn selects_adjacent_word_at_a_boundary() {
        assert_eq!(word_char_range("alpha beta", 5), Some(0..5));
        assert_eq!(word_char_range("alpha beta", 6), Some(6..10));
        assert_eq!(word_char_range("alpha", 5), Some(0..5));
    }

    #[test]
    fn skips_non_word_characters() {
        assert_eq!(word_char_range("alpha - beta", 6), None);
        assert_eq!(word_char_range("", 0), None);
    }
}
