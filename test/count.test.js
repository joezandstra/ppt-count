import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { test } from "node:test";

import { countChunks, countText, sumCounts, ZERO } from "../src/count.js";

const fixture = JSON.parse(readFileSync(new URL("./fixtures/word-measurements.json", import.meta.url), "utf8"));

const counts = (words, characters, charactersWithSpaces) => ({ words, characters, charactersWithSpaces });

// [input, expected, why]. All verified against Microsoft Word for Mac 16.113.2
// unless marked "deviation".
const VECTORS = [
  ["Hello world", counts(2, 10, 11), "basic"],
  ["Hello  world", counts(2, 10, 12), "each space counts"],
  ["  Hello world  ", counts(2, 10, 15), "leading/trailing spaces count"],
  ["a\tb", counts(2, 2, 3), "tab is a space"],
  ["a\rb", counts(2, 2, 2), "paragraph mark not counted"],
  ["a\nb", counts(2, 2, 2), "LF is a paragraph"],
  ["a\r\nb", counts(2, 2, 2), "CRLF is one paragraph"],
  ["a\u000Bb", counts(2, 2, 3), "Shift+Enter line break counts as a space"],
  ["word word", counts(2, 8, 9), "NBSP separates and counts as a space"],
  ["a b", counts(1, 3, 3), "thin space is not a space for Word"],
  ["a　b", counts(2, 2, 3), "ideographic space is a space"],
  ["word - word", counts(3, 9, 11), "spaced hyphen is a word"],
  ["word & word", counts(3, 9, 11), "standalone symbol is a word"],
  ["word — word", counts(2, 9, 11), "spaced em dash is not a word but is a character"],
  ["word—word", counts(2, 9, 9), "em dash splits words"],
  ["word–word", counts(2, 9, 9), "en dash splits words"],
  ["well-known", counts(1, 10, 10), "hyphen joins"],
  ["e-mail", counts(1, 6, 6), "hyphen joins"],
  ["don’t", counts(1, 5, 5), "apostrophe joins"],
  ["3.14", counts(1, 4, 4), "decimal"],
  ["U.S.A.", counts(1, 6, 6), "abbreviation"],
  ["https://example.com/a/b?x=1&y=2", counts(1, 31, 31), "URL is one word"],
  ["a/b", counts(1, 3, 3), "slash joins"],
  [". . .", counts(3, 3, 5), "each dot is a word"],
  ["The court … concluded", counts(4, 18, 21), "ellipsis character is a word"],
  ["—", counts(0, 1, 1), "lone em dash"],
  ["word--word", counts(2, 10, 10), "double hyphen breaks"],
  ["--word", counts(2, 6, 6), "leading double hyphen"],
  ["• item", counts(1, 5, 6), "bullet character is a separator"],
  ["word−word", counts(1, 9, 9), "minus sign joins"],
  ["word​word", counts(1, 8, 8), "zero-width space ignored"],
  ["word­word", counts(1, 9, 9), "literal soft hyphen counted"],
  ["\"Remember, we—\"", counts(3, 14, 15), "closing quote after em dash is a word"],
  ["世界", counts(2, 2, 2), "each CJK character is a word"],
  ["Hello世界", counts(3, 7, 7), "CJK breaks a Latin run"],
  ["你好，世界。", counts(6, 6, 6), "full-width punctuation are words"],
  ["Hi， this is a test", counts(6, 14, 18), "full-width comma"],
  ["世界,", counts(3, 3, 3), "ASCII comma after CJK is a word"],
  ["日本語123", counts(4, 6, 6), "digit run is one word"],
  ["ｶﾀｶﾅ", counts(4, 4, 4), "half-width katakana per character"],
  ["ＡＢＣ", counts(3, 3, 3), "full-width Latin per character"],
  ["안녕하세요 세계", counts(2, 7, 8), "Korean words by default"],
  ["สวัสดีครับ", counts(1, 10, 10), "Thai is not segmented"],
  ["\u{1F44D}", counts(1, 1, 1), "emoji is one character"],
  ["\u{1F44D}\u{1F3FD}", counts(1, 2, 2), "skin tone adds one"],
  ["\u{1F468}‍\u{1F469}‍\u{1F467}", counts(1, 5, 5), "joiners inside emoji count"],
  ["\u{1F1FA}\u{1F1F8}", counts(1, 2, 2), "flag is two"],
  ["❤️", counts(1, 1, 1), "variation selector ignored"],
  ["Hi\u{1F44D}", counts(1, 3, 3), "emoji joins a word"],
  ["café", counts(1, 5, 5), "combining accent counted"],
  [" ", counts(0, 0, 1), "lone space"],
  ["​", counts(0, 0, 0), "lone zero-width space"],
  ["", counts(0, 0, 0), "empty"],
  ["The quick brown fox — it jumped – over 3.5 lazy dogs.", counts(10, 42, 53), "mixed dashes"],
  ["\t•\tBullet text", counts(2, 11, 14), "typed bullet with tabs"],
  ["Title\rBody text here", counts(4, 17, 19), "paragraphs"],
  ["Line one\u000BLine two", counts(4, 14, 17), "soft line break"],
  ["\"This — as expected — irritated.\"", counts(4, 28, 33), "quoted with em dashes"],
  ["a b", counts(2, 2, 3), "deviation: U+2028 is a line break (Word: 3/3/3)"],
  ["a b", counts(2, 2, 2), "deviation: U+2029 is a paragraph break (Word: 3/3/3)"],
  ["a※b", counts(3, 3, 3), "reference mark is East Asian (fresh-document Word)"],
];

for (const [input, expected, why] of VECTORS) {
  test(`countText ${JSON.stringify(input)} (${why})`, () => {
    assert.deepEqual(countText(input), expected);
  });
}

test("counts Hangul per syllable when asked (Word's untagged behaviour)", () => {
  assert.deepEqual(countText("안녕하세요 세계", { hangulPerSyllable: true }), counts(7, 7, 8));
});

test("word count for a sentence with dash variants matches Word", () => {
  const text = "How about some dash stuff–things like hyphen-dashes and em-dashes. 1-248-434-5508.";
  assert.equal(countText(text).words, 11);
});

test(`matches all ${fixture.rows.length} Word-measured rows`, () => {
  const failures = [];
  for (const row of fixture.rows) {
    const actual = countText(row.text, { hangulPerSyllable: row.hangulPerSyllable === true });
    const expected = counts(row.words, row.characters, row.charactersWithSpaces);
    if (JSON.stringify(actual) !== JSON.stringify(expected)) {
      failures.push(`${row.id} ${JSON.stringify(row.text)}: expected ${JSON.stringify(expected)}, got ${JSON.stringify(actual)}`);
    }
  }
  assert.deepEqual(failures, []);
});

test("each deliberate deviation really differs from Word (so the list stays honest)", () => {
  for (const row of fixture.deviations) {
    const actual = countText(row.text, { hangulPerSyllable: row.hangulPerSyllable === true });
    assert.notDeepEqual(actual, counts(row.words, row.characters, row.charactersWithSpaces), row.id);
  }
});

test("sumCounts adds field by field and returns a new object", () => {
  const total = sumCounts([counts(1, 2, 3), counts(10, 20, 30)]);
  assert.deepEqual(total, counts(11, 22, 33));
  assert.deepEqual(sumCounts([]), ZERO);
  assert.notEqual(sumCounts([]), ZERO);
});

test("countChunks never lets words run together across chunks", () => {
  assert.deepEqual(countChunks(["Hello", "world"]), counts(2, 10, 10));
  assert.deepEqual(countChunks([]), ZERO);
});

test("ZERO is frozen", () => {
  assert.ok(Object.isFrozen(ZERO));
});
