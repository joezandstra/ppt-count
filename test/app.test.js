import assert from "node:assert/strict";
import { test } from "node:test";

import { describeSelection, summarize, themeFrom, unsupportedNote } from "../src/app.js";

const summary = (overrides) => ({
  source: "shapes",
  selectedCount: 1,
  items: [],
  unsupported: [],
  total: { words: 0, characters: 0, charactersWithSpaces: 0 },
  ...overrides,
});
const item = (name, kind = "text") => ({ id: name, name, kind, counts: { words: 1, characters: 1, charactersWithSpaces: 1 } });

test("summarize counts each item and the total", () => {
  const result = summarize({
    source: "shapes",
    selectedCount: 2,
    items: [
      { id: "1", name: "TextBox 1", kind: "text", texts: ["Hello world"] },
      { id: "2", name: "Table 1", kind: "table", texts: ["a", "b c"] },
    ],
    unsupported: [],
  });
  assert.deepEqual(result.items[0], { id: "1", name: "TextBox 1", kind: "text", counts: { words: 2, characters: 10, charactersWithSpaces: 11 } });
  assert.deepEqual(result.items[1].counts, { words: 3, characters: 3, charactersWithSpaces: 4 });
  assert.deepEqual(result.total, { words: 5, characters: 13, charactersWithSpaces: 15 });
  assert.equal(result.source, "shapes");
  assert.equal(result.selectedCount, 2);
});

test("describe: nothing selected", () => {
  assert.deepEqual(describeSelection(summary({ source: "none", selectedCount: 0 })), {
    title: "Nothing selected",
    detail: "Select a text box, table or group, or highlight some text.",
    note: "",
    empty: true,
  });
});

test("describe: highlighted text, with and without a shape name", () => {
  assert.deepEqual(describeSelection(summary({ source: "highlight", items: [item("TextBox 3", "highlight")] })), {
    title: "Highlighted text",
    detail: "in TextBox 3",
    note: "",
    empty: false,
  });
  assert.equal(describeSelection(summary({ source: "highlight", selectedCount: 0, items: [item("", "highlight")] })).detail, "");
});

test("describe: one shape uses its name; tables and groups say what's counted", () => {
  assert.deepEqual(describeSelection(summary({ items: [item("TextBox 3")] })), { title: "TextBox 3", detail: "", note: "", empty: false });
  assert.equal(describeSelection(summary({ items: [item("Table 2", "table")] })).detail, "Whole table");
  assert.equal(describeSelection(summary({ items: [item("Group 5", "group")] })).detail, "Whole group");
});

test("describe: several shapes", () => {
  assert.equal(describeSelection(summary({ selectedCount: 3, items: [item("A"), item("B")] })).title, "3 shapes selected");
});

test("describe: selected shapes without text", () => {
  assert.deepEqual(describeSelection(summary({ selectedCount: 1 })), {
    title: "No text to count",
    detail: "The selected shape doesn't contain text.",
    note: "",
    empty: true,
  });
  assert.equal(describeSelection(summary({ selectedCount: 2 })).detail, "The selected shapes don't contain text.");
});

test("describe: notes objects that can't be counted next to ones that can", () => {
  const unsupported = [{ id: "c", name: "Chart 1", type: "Chart" }];
  assert.equal(describeSelection(summary({ selectedCount: 2, items: [item("A")], unsupported })).note, unsupportedNote(1));
});

test("describe: only a chart or SmartArt selected says so once, without claiming it has no text", () => {
  const unsupported = [{ id: "c", name: "SmartArt 1", type: "SmartArt" }];
  assert.deepEqual(describeSelection(summary({ selectedCount: 1, unsupported })), {
    title: "Can't count this",
    detail: "PowerPoint doesn't let add-ins read the text in charts, SmartArt and some other objects.",
    note: "",
    empty: true,
  });
});

test("unsupportedNote wording", () => {
  assert.equal(unsupportedNote(0), "");
  assert.equal(unsupportedNote(1), "1 object isn't included: PowerPoint doesn't let add-ins read the text in charts, SmartArt and some other objects.");
  assert.equal(unsupportedNote(3), "3 objects aren't included: PowerPoint doesn't let add-ins read the text in charts, SmartArt and some other objects.");
});

test("themeFrom uses PowerPoint's colours and dark flag", () => {
  assert.deepEqual(themeFrom({ bodyBackgroundColor: "#1F1F1F", bodyForegroundColor: "#FFFFFF", isDarkTheme: true }), {
    background: "#1f1f1f",
    foreground: "#ffffff",
    dark: true,
  });
});

test("themeFrom works out dark or light from the background when the flag is missing", () => {
  assert.equal(themeFrom({ bodyBackgroundColor: "#202020", bodyForegroundColor: "#F0F0F0" }).dark, true);
  assert.equal(themeFrom({ bodyBackgroundColor: "FFFFFF", bodyForegroundColor: "000000" }).dark, false);
});

test("themeFrom ignores missing, empty or invalid themes", () => {
  assert.equal(themeFrom(undefined), null);
  assert.equal(themeFrom({}), null);
  assert.equal(themeFrom({ bodyBackgroundColor: "red", bodyForegroundColor: "#000000" }), null);
});
