// Word-compatible counting of words and characters.
//
// The rules reproduce Microsoft Word's own statistics, measured on Word for Mac
// 16.113.2 (docs/superpowers/specs/2026-09-26-ppt-word-count-design.md, section 3).
// Counting walks Unicode code points: String.length, Intl.Segmenter and regex
// splitting all disagree with Word on everyday text (hyphens, URLs, CJK, emoji).

/** @typedef {{ words: number, characters: number, charactersWithSpaces: number }} Counts */

/** @type {Readonly<Counts>} */
export const ZERO = Object.freeze({ words: 0, characters: 0, charactersWithSpaces: 0 });

// Paragraph breaks end a word and are never counted.
const PARAGRAPH = new Set([0x0a, 0x0c, 0x0d, 0x0e, 0x2029]);
// Spaces end a word and are counted only in "characters with spaces".
const SPACE = new Set([0x09, 0x0b, 0x20, 0xa0, 0x2028, 0x3000]);
// En dash, em dash and bullet end a word and count as characters, but are never words themselves.
const DASH_SEPARATOR = new Set([0x2013, 0x2014, 0x2022]);

const HYPHEN = 0x2d;
const ZERO_WIDTH_JOINER = 0x200d;

// Each of these characters is a word on its own, like Word's "Asian characters".
const EAST_ASIAN = new RegExp(
  "[" +
    "\\p{Script=Han}\\p{Script=Hiragana}\\p{Script=Katakana}" +
    "\\u3100-\\u312F" + // Bopomofo (Bopomofo Extended is not East Asian to Word)
    "\\u2024\\u2025\\u2027\\u203B" + // one/two dot leaders, hyphenation point, reference mark
    "\\u2160-\\u217F" + // Roman numerals
    "\\u2460-\\u24B5\\u24D0-\\u24FE" + // circled and parenthesized numbers, circled small letters
    "\\u2E80-\\u2FDF" + // CJK and Kangxi radicals
    "\\u3001-\\u303F" + // CJK symbols and punctuation
    "\\u3099-\\u309C\\u30A0\\u30FB\\u30FC" + // kana marks, middle dot, prolonged sound mark
    "\\u3190-\\u319F" + // Kanbun
    "\\u3200-\\u33FF" + // enclosed CJK letters, CJK compatibility
    "\\uFE30-\\uFE4F" + // CJK compatibility forms
    "\\uFE50-\\uFE6B" + // small form variants
    "\\uFF01-\\uFFEF" + // half-width and full-width forms
    "]",
  "u",
);
const HANGUL = /\p{Script=Hangul}/u;
// A zero-width joiner after one of these belongs to an emoji sequence and is counted.
const EMOJI_BASE = /[\p{Extended_Pictographic}\p{Emoji_Modifier}]/u;

// Invisible format characters that are neither counted nor word breaks.
function isIgnorable(cp) {
  return (
    cp === 0x200b || // zero-width space
    cp === 0x200c || // zero-width non-joiner
    cp === ZERO_WIDTH_JOINER ||
    cp === 0xfeff || // byte order mark
    cp === 0x1f || // Word's optional hyphen
    (cp >= 0xfe00 && cp <= 0xfe0f) || // variation selectors
    (cp >= 0xe0100 && cp <= 0xe01ef) || // variation selectors supplement
    (cp >= 0xe0000 && cp <= 0xe007f) // tags
  );
}

/**
 * Counts one piece of text the way Microsoft Word does.
 * @param {string} text
 * @param {{ hangulPerSyllable?: boolean }} [options] Count each Hangul syllable as a
 *   word (Word's behaviour for text not tagged as Korean). Default: space-delimited Korean words.
 * @returns {Counts}
 */
export function countText(text, { hangulPerSyllable = false } = {}) {
  let words = 0;
  let withSpaces = 0;
  let spaces = 0;
  let inWord = false;
  let breakAfterHyphens = false; // after "--", the next non-hyphen starts a new word
  let previous = -1; // previous code point in the current word
  let previousVisible = -1; // previous non-ignorable code point, for the joiner rule

  const endWord = () => {
    inWord = false;
    breakAfterHyphens = false;
    previous = -1;
  };

  for (const ch of text) {
    const cp = ch.codePointAt(0);
    if (cp === ZERO_WIDTH_JOINER && previousVisible !== -1 && EMOJI_BASE.test(String.fromCodePoint(previousVisible))) {
      withSpaces++;
      continue;
    }
    if (isIgnorable(cp)) continue;
    previousVisible = cp;

    if (PARAGRAPH.has(cp)) {
      endWord();
      continue;
    }
    withSpaces++;
    if (SPACE.has(cp)) {
      spaces++;
      endWord();
      continue;
    }
    if (DASH_SEPARATOR.has(cp)) {
      endWord();
      continue;
    }
    if (EAST_ASIAN.test(ch) || (hangulPerSyllable && HANGUL.test(ch))) {
      words++;
      endWord();
      continue;
    }
    if (breakAfterHyphens && cp !== HYPHEN) endWord();
    if (!inWord) {
      words++;
      inWord = true;
    }
    if (cp === HYPHEN && previous === HYPHEN) breakAfterHyphens = true;
    previous = cp;
  }

  return { words, characters: withSpaces - spaces, charactersWithSpaces: withSpaces };
}

/**
 * Adds counts together.
 * @param {Counts[]} list
 * @returns {Counts}
 */
export function sumCounts(list) {
  const total = { ...ZERO };
  for (const c of list) {
    total.words += c.words;
    total.characters += c.characters;
    total.charactersWithSpaces += c.charactersWithSpaces;
  }
  return total;
}

/**
 * Counts several separate pieces of text (text boxes, table cells...) and adds the
 * results. Pieces are never joined, so words can't run together across them.
 * @param {string[]} texts
 * @param {{ hangulPerSyllable?: boolean }} [options]
 * @returns {Counts}
 */
export function countChunks(texts, options) {
  return sumCounts(texts.map((text) => countText(text, options)));
}
