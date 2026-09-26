import assert from "node:assert/strict";
import { test } from "node:test";

import { createFakeHost } from "../dev/fake-powerpoint.js";
import { SCENARIOS } from "../dev/scenarios.js";
import { readSelection, tableCellTexts } from "../src/selection.js";

const SENTENCE = "The quick brown fox jumps over the lazy dog.";

async function read(scenario, options) {
  const { PowerPoint, host } = createFakeHost(scenario);
  const snapshot = await PowerPoint.run((context) => readSelection(context, options));
  return { snapshot, host };
}

const TABLE_TEXT = "Region\r\nQ1\r\nQ2\r\nNorth\r\n1,200 units\r\n1,450 units\r\nSouth \u2014 combined total\r\n";
const selectedText = (text) => async () => text;

test("nothing selected", async () => {
  const { snapshot } = await read(SCENARIOS.nothing);
  assert.deepEqual(snapshot, { source: "none", selectedCount: 0, items: [], unsupported: [] });
});

test("a text box: its whole text, in three syncs", async () => {
  const { snapshot, host } = await read(SCENARIOS.textBox);
  assert.deepEqual(snapshot, {
    source: "shapes",
    selectedCount: 1,
    items: [{ id: "103", name: "TextBox 3", kind: "text", texts: [SENTENCE] }],
    unsupported: [],
  });
  assert.ok(host.stats.syncs <= 3, `used ${host.stats.syncs} syncs`);
});

test("highlighted text counts only the highlight and names its shape", async () => {
  const { snapshot, host } = await read(SCENARIOS.highlight);
  assert.deepEqual(snapshot, {
    source: "highlight",
    selectedCount: 1,
    items: [{ id: "103", name: "TextBox 3", kind: "highlight", texts: ["quick brown fox"] }],
    unsupported: [],
  });
  assert.ok(host.stats.syncs <= 3, `used ${host.stats.syncs} syncs`);
});

test("a whole box that PowerPoint reports as selected text counts as the box (PowerPoint for Mac)", async () => {
  const { snapshot } = await read({ highlight: SENTENCE, selected: SCENARIOS.textBox.selected });
  assert.equal(snapshot.source, "shapes");
  assert.deepEqual(snapshot.items, [{ id: "103", name: "TextBox 3", kind: "text", texts: [SENTENCE] }]);
});

test("a whole table reported as selected text counts as the table, whatever separates its cells", async () => {
  const highlight = "Region\tQ1\tQ2\rNorth\t1,200 units\t1,450 units\rSouth \u2014 combined total";
  const { snapshot } = await read({ highlight, selected: SCENARIOS.table.selected });
  assert.equal(snapshot.source, "shapes");
  assert.equal(snapshot.items[0].kind, "table");
});

test("highlighting all but one word of a box is still a highlight", async () => {
  const { snapshot } = await read({ highlight: "The quick brown fox jumps over the lazy", selected: SCENARIOS.textBox.selected });
  assert.equal(snapshot.source, "highlight");
});

test("a highlight with no selected shape (e.g. speaker notes) has no name", async () => {
  const { snapshot } = await read({ highlight: "notes text", selected: [] });
  assert.equal(snapshot.source, "highlight");
  assert.deepEqual(snapshot.items, [{ id: "", name: "", kind: "highlight", texts: ["notes text"] }]);
});

test("an empty highlight (just a cursor) counts the whole shape", async () => {
  const { snapshot } = await read({ highlight: "", selected: SCENARIOS.textBox.selected });
  assert.equal(snapshot.source, "shapes");
  assert.deepEqual(snapshot.items[0].texts, [SENTENCE]);
});

test("a highlight is ignored when several shapes are selected", async () => {
  const { snapshot } = await read({ highlight: "Quarterly", selected: SCENARIOS.several.selected });
  assert.equal(snapshot.source, "shapes");
});

test("several shapes: pictures are left out but still counted as selected", async () => {
  const { snapshot } = await read(SCENARIOS.several);
  assert.equal(snapshot.selectedCount, 3);
  assert.deepEqual(
    snapshot.items.map((i) => [i.name, i.kind, i.texts.length]),
    [
      ["Title 1", "text", 1],
      ["Content Placeholder 2", "text", 1],
    ],
  );
});

test("a table: every cell once, merged cells not repeated", async () => {
  const { snapshot, host } = await read(SCENARIOS.table);
  assert.deepEqual(snapshot.items, [
    {
      id: "105",
      name: "Table 5",
      kind: "table",
      texts: ["Region", "Q1", "Q2", "North", "1,200 units", "1,450 units", "South — combined total"],
    },
  ]);
  assert.ok(host.stats.syncs <= 3, `used ${host.stats.syncs} syncs`);
});

test("nested groups: every text shape inside, attributed to the top-level group", async () => {
  const { snapshot, host } = await read(SCENARIOS.group);
  assert.deepEqual(snapshot.items, [{ id: "106", name: "Group 6", kind: "group", texts: ["Step one: plan", "Step two: build", ""] }]);
  assert.ok(host.stats.syncs <= 5, `used ${host.stats.syncs} syncs`);
});

test("a group containing a table reads the table's cells", async () => {
  const group = {
    id: "g",
    name: "Group 1",
    type: "Group",
    shapes: [
      { id: "t", name: "Table 1", type: "Table", rows: [["a", "b"]] },
      { id: "x", name: "TextBox 1", type: "TextBox", text: "c" },
    ],
  };
  const { snapshot } = await read({ highlight: null, selected: [group] });
  assert.equal(snapshot.items[0].kind, "group");
  assert.deepEqual([...snapshot.items[0].texts].sort(), ["a", "b", "c"]);
});

test("a shape selected together with its group is counted once", async () => {
  const inner = { id: "t", name: "TextBox 1", type: "TextBox", text: "inside" };
  const other = { id: "u", name: "TextBox 2", type: "TextBox", text: "sibling" };
  const group = { id: "g", name: "Group 1", type: "Group", shapes: [inner, other] };
  const { snapshot } = await read({ highlight: null, selected: [group, inner] });
  assert.deepEqual(snapshot.items.flatMap((i) => i.texts).sort(), ["inside", "sibling"]);
});

test("a chart that PowerPoint reports as a range with no text is treated as the chart", async () => {
  const { snapshot } = await read({ highlight: { text: null }, selected: [{ id: "112", name: "Chart 12", type: "Chart" }] });
  assert.deepEqual(snapshot, { source: "shapes", selectedCount: 1, items: [], unsupported: [{ id: "112", name: "Chart 12", type: "Chart" }] });
});

test("charts are reported as not countable; the rest still counts", async () => {
  const { snapshot } = await read(SCENARIOS.chart);
  assert.deepEqual(snapshot.unsupported, [{ id: "112", name: "Chart 12", type: "Chart" }]);
  assert.deepEqual(
    snapshot.items.map((i) => i.name),
    ["TextBox 3"],
  );
  assert.equal(snapshot.selectedCount, 2);
});

test("a picture alone: nothing to count", async () => {
  const { snapshot } = await read(SCENARIOS.picture);
  assert.deepEqual(snapshot, { source: "shapes", selectedCount: 1, items: [], unsupported: [] });
});

test("an empty placeholder is an item with empty text", async () => {
  const { snapshot } = await read(SCENARIOS.emptyPlaceholder);
  assert.deepEqual(
    snapshot.items.map((i) => i.texts),
    [[""]],
  );
});

test("a table inside a content placeholder is read as a table", async () => {
  const { snapshot } = await read(SCENARIOS.tablePlaceholder);
  assert.deepEqual(snapshot.items, [{ id: "114", name: "Content Placeholder 3", kind: "table", texts: ["Name", "Role", "Ana", "Design lead"] }]);
});

test("a table reported as 'Unsupported' (older builds) is still counted", async () => {
  const { snapshot } = await read({ highlight: null, selected: [{ id: "u", name: "Table 1", type: "Unsupported", rows: [["a", "b"]] }] });
  assert.deepEqual(snapshot.items, [{ id: "u", name: "Table 1", kind: "table", texts: ["a", "b"] }]);
  assert.deepEqual(snapshot.unsupported, []);
});

test("an 'Unsupported' shape that isn't a table is reported, not fatal", async () => {
  const { snapshot } = await read({ highlight: null, selected: [{ id: "z", name: "Zoom 1", type: "Unsupported" }, ...SCENARIOS.textBox.selected] });
  assert.deepEqual(snapshot.unsupported, [{ id: "z", name: "Zoom 1", type: "Unsupported" }]);
  assert.deepEqual(
    snapshot.items.map((i) => i.name),
    ["TextBox 3"],
  );
});

test("one shape failing doesn't hide the others", async () => {
  const broken = { id: "9", name: "Broken 9", type: "TextBox", text: "x", failTextFrame: true };
  const { snapshot } = await read({ highlight: null, selected: [SCENARIOS.textBox.selected[0], broken, SCENARIOS.title.selected[0]] });
  assert.deepEqual(
    snapshot.items.map((i) => i.name),
    ["TextBox 3", "Title 1"],
  );
  assert.deepEqual(snapshot.unsupported, [{ id: "9", name: "Broken 9", type: "TextBox" }]);
});

test("never uses the APIs that throw on the wrong kind of shape", async () => {
  for (const [key, scenario] of Object.entries(SCENARIOS)) {
    const { host } = await read(scenario);
    assert.deepEqual(host.stats.forbidden, [], key);
  }
});

test("a failed first sync is retried once without the text range", async () => {
  const { PowerPoint, host } = createFakeHost(SCENARIOS.textBox);
  host.failNextSyncs(1);
  const snapshot = await PowerPoint.run((context) => readSelection(context));
  assert.deepEqual(snapshot.items[0].texts, ["The quick brown fox jumps over the lazy dog."]);
});

test("if the retry fails too, the read rejects so the caller can retry later", async () => {
  const { PowerPoint, host } = createFakeHost(SCENARIOS.textBox);
  host.failNextSyncs(2);
  await assert.rejects(PowerPoint.run((context) => readSelection(context)), { code: "GeneralException" });
});

test("a value that goes missing after a successful sync rejects, so the whole read is retried", async () => {
  const { PowerPoint, host } = createFakeHost(SCENARIOS.textBox);
  host.trampleSync(2);
  await assert.rejects(PowerPoint.run((context) => readSelection(context)), { code: "PropertyNotLoaded" });
});

test("a later batch failing once is retried shape by shape and still counts everything", async () => {
  const { PowerPoint, host } = createFakeHost(SCENARIOS.several);
  host.failSync(2);
  const snapshot = await PowerPoint.run((context) => readSelection(context));
  assert.deepEqual(
    snapshot.items.map((i) => i.name),
    ["Title 1", "Content Placeholder 2"],
  );
  assert.deepEqual(snapshot.unsupported, []);
});

test("if every shape keeps failing, the read rejects instead of calling them uncountable", async () => {
  const { PowerPoint, host } = createFakeHost(SCENARIOS.several);
  for (const n of [2, 3, 4]) host.failSync(n);
  await assert.rejects(PowerPoint.run((context) => readSelection(context)), { code: "GeneralException" });
});

test("a wrong guess about an 'Unsupported' shape doesn't slow down the other shapes", async () => {
  const boxes = Array.from({ length: 30 }, (_, i) => ({ id: `b${i}`, name: `TextBox ${i}`, type: "TextBox", text: "one two" }));
  const { snapshot, host } = await read({ highlight: null, selected: [...boxes, { id: "z", name: "Zoom 1", type: "Unsupported" }] });
  assert.equal(snapshot.items.length, 30);
  assert.deepEqual(snapshot.unsupported, [{ id: "z", name: "Zoom 1", type: "Unsupported" }]);
  assert.ok(host.stats.syncs <= 5, `used ${host.stats.syncs} syncs`);
});

test("an 'Unsupported' shape selected on its own is reported, not an error", async () => {
  const { snapshot } = await read({ highlight: null, selected: [{ id: "z", name: "Zoom 1", type: "Unsupported" }] });
  assert.deepEqual(snapshot, { source: "shapes", selectedCount: 1, items: [], unsupported: [{ id: "z", name: "Zoom 1", type: "Unsupported" }] });
});

test("text highlighted in a table cell counts only that text", async () => {
  // PowerPoint for Mac: a text range without text, but the plain-text selection has the highlight.
  const { snapshot } = await read({ highlight: { text: null }, selected: SCENARIOS.table.selected }, { selectedText: selectedText("1,200") });
  assert.deepEqual(snapshot, {
    source: "highlight",
    selectedCount: 1,
    items: [{ id: "105", name: "Table 5", kind: "highlight", texts: ["1,200"] }],
    unsupported: [],
  });
});

test("a selected cell counts that cell", async () => {
  const { snapshot } = await read({ highlight: { text: null }, selected: SCENARIOS.table.selected }, { selectedText: selectedText("1,450 units\r\n") });
  assert.equal(snapshot.source, "highlight");
  assert.deepEqual(snapshot.items[0].texts, ["1,450 units\r\n"]);
});

test("a whole selected table still counts every cell", async () => {
  const { snapshot } = await read(SCENARIOS.table, { selectedText: selectedText(TABLE_TEXT) });
  assert.equal(snapshot.source, "shapes");
  assert.equal(snapshot.items[0].kind, "table");
  assert.equal(snapshot.items[0].texts.length, 7);
});

test("no plain-text selection (empty, missing or unavailable) counts the whole table", async () => {
  for (const options of [{ selectedText: selectedText("") }, { selectedText: selectedText(null) }, { selectedText: selectedText("  \r\n") }, {}]) {
    const { snapshot } = await read(SCENARIOS.table, options);
    assert.equal(snapshot.source, "shapes");
  }
});

test("a text range that fails to load doesn't break the read", async () => {
  const { snapshot } = await read({ highlight: { error: "InvalidArgument" }, selected: SCENARIOS.table.selected }, { selectedText: selectedText("Q1") });
  assert.equal(snapshot.source, "highlight");
  assert.deepEqual(snapshot.items[0].texts, ["Q1"]);
});

test("the plain-text selection is only asked for when one table is selected", async () => {
  let calls = 0;
  const counting = async () => {
    calls++;
    return "x";
  };
  for (const key of ["textBox", "several", "group", "chart", "picture", "nothing", "highlight"]) {
    await read(SCENARIOS[key], { selectedText: counting });
  }
  assert.equal(calls, 0);
  await read(SCENARIOS.tablePlaceholder, { selectedText: counting });
  assert.equal(calls, 1);
});

test("tableCellTexts skips cells hidden under a merge", () => {
  assert.deepEqual(
    tableCellTexts([
      { isNullObject: false, text: "a" },
      { isNullObject: true, text: null },
      { isNullObject: false, text: "" },
    ]),
    ["a", ""],
  );
});
