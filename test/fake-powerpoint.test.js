import assert from "node:assert/strict";
import { test } from "node:test";

import { createFakeHost, FakeOfficeError } from "../dev/fake-powerpoint.js";
import { SCENARIOS } from "../dev/scenarios.js";

const textBox = { id: "1", name: "TextBox 1", type: "TextBox", text: "Hello world" };
const table = {
  id: "2",
  name: "Table 1",
  type: "Table",
  rows: [
    ["A", "B"],
    ["merged", ""],
  ],
  merged: [{ row: 1, column: 0, rowCount: 1, columnCount: 2 }],
};

test("properties must be loaded and synced before they're read", async () => {
  const { PowerPoint } = createFakeHost({ highlight: null, selected: [textBox] });
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    assert.throws(() => shapes.items, { code: "PropertyNotLoaded" });
    shapes.load("items/id,items/name,items/type");
    assert.throws(() => shapes.items, { code: "PropertyNotLoaded" });
    await context.sync();
    assert.equal(shapes.items.length, 1);
    assert.equal(shapes.items[0].name, "TextBox 1");
    assert.throws(() => shapes.items[0].level, { code: "PropertyNotLoaded" });
  });
});

test("an invalid navigation rejects the whole batch and applies nothing", async () => {
  const { PowerPoint } = createFakeHost({ highlight: null, selected: [textBox, table] });
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id,items/type");
    await context.sync();
    const [box, tbl] = shapes.items;
    const range = box.getTextFrameOrNullObject().textRange;
    range.load("text");
    tbl.textFrame.load("hasText"); // tables have no text frame
    await assert.rejects(context.sync(), (error) => error instanceof FakeOfficeError && error.code === "InvalidArgument");
    assert.throws(() => range.text, { code: "PropertyNotLoaded" });
    // The context keeps working after a failed batch.
    range.load("text");
    await context.sync();
    assert.equal(range.text, "Hello world");
  });
});

test("group, placeholderFormat and getTable fail on the wrong kind of shape", async () => {
  const { PowerPoint } = createFakeHost({ highlight: null, selected: [textBox] });
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();
    const [box] = shapes.items;
    for (const load of [() => box.group.shapes.load("items/id"), () => box.placeholderFormat.load("containedType"), () => box.getTable().load("rowCount")]) {
      load();
      await assert.rejects(context.sync(), { code: "GeneralException" });
    }
  });
});

test("OrNullObject methods return null objects instead of failing", async () => {
  const picture = { id: "3", name: "Picture 1", type: "Image" };
  const { PowerPoint } = createFakeHost({ highlight: null, selected: [picture] });
  await PowerPoint.run(async (context) => {
    const range = context.presentation.getSelectedTextRangeOrNullObject();
    range.load("text");
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();
    assert.equal(range.isNullObject, true);
    const frame = shapes.items[0].getTextFrameOrNullObject();
    frame.load("hasText");
    await context.sync();
    assert.equal(frame.isNullObject, true);
    // Navigating from a null object is undocumented in Office.js, so the fake refuses.
    frame.textRange.load("text");
    await assert.rejects(context.sync(), { code: "InvalidArgument" });
  });
});

test("table cells hidden under a merge are null objects; values repeats merged text", async () => {
  const { PowerPoint } = createFakeHost({ highlight: null, selected: [table] });
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/type");
    await context.sync();
    const t = shapes.items[0].getTable();
    t.load("rowCount,columnCount,values");
    await context.sync();
    assert.equal(t.rowCount, 2);
    assert.equal(t.columnCount, 2);
    assert.deepEqual(t.values, [
      ["A", "B"],
      ["merged", "merged"],
    ]);
    const anchor = t.getCellOrNullObject(1, 0);
    const hidden = t.getCellOrNullObject(1, 1);
    anchor.load("text");
    hidden.load("text");
    await context.sync();
    assert.equal(anchor.isNullObject, false);
    assert.equal(anchor.text, "merged");
    assert.equal(hidden.isNullObject, true);
  });
});

test("group children are reachable through group.shapes", async () => {
  const group = { id: "9", name: "Group 1", type: "Group", shapes: [textBox, table] };
  const { PowerPoint } = createFakeHost({ highlight: null, selected: [group] });
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id,items/type");
    await context.sync();
    const children = shapes.items[0].group.shapes;
    children.load("items/id,items/name,items/type");
    await context.sync();
    assert.deepEqual(
      children.items.map((s) => s.name),
      ["TextBox 1", "Table 1"],
    );
  });
});

test("proxies can't be used after their PowerPoint.run has finished", async () => {
  const { PowerPoint } = createFakeHost({ highlight: null, selected: [textBox] });
  let leaked;
  await PowerPoint.run(async (context) => {
    leaked = context.presentation.getSelectedShapes();
  });
  await PowerPoint.run(async () => {
    assert.throws(() => leaked.load("items/id"), { code: "InvalidObjectPath" });
  });
});

test("records forbidden APIs and counts runs and syncs", async () => {
  const { PowerPoint, host } = createFakeHost({ highlight: null, selected: [textBox] });
  await PowerPoint.run(async (context) => {
    context.presentation.getSelectedTextRange();
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id");
    await context.sync();
    void shapes.items[0].textFrame;
  });
  assert.deepEqual(host.stats.forbidden, ["Presentation.getSelectedTextRange", "Shape.textFrame"]);
  assert.equal(host.stats.runs, 1);
  assert.equal(host.stats.syncs, 1);
});

test("failNextSyncs makes the next syncs fail like a transient host error", async () => {
  const { PowerPoint, host } = createFakeHost({ highlight: null, selected: [textBox] });
  host.failNextSyncs(1);
  await PowerPoint.run(async (context) => {
    const shapes = context.presentation.getSelectedShapes();
    shapes.load("items/id");
    await assert.rejects(context.sync(), { code: "GeneralException" });
    shapes.load("items/id");
    await context.sync();
    assert.equal(shapes.items.length, 1);
  });
});

test("Office fake: onReady, requirement sets, selection events", async () => {
  const { Office, host } = createFakeHost({ highlight: null, selected: [] }, { platform: "PC", apiVersion: "1.10" });
  assert.deepEqual(await Office.onReady(), { host: "PowerPoint", platform: "PC" });
  assert.equal(Office.context.requirements.isSetSupported("PowerPointApi", "1.10"), true);
  assert.equal(Office.context.requirements.isSetSupported("PowerPointApi", "1.8"), true);
  assert.equal(Office.context.requirements.isSetSupported("PowerPointApi", "1.11"), false);
  let fired = 0;
  let status;
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => fired++, (result) => (status = result.status));
  assert.equal(status, Office.AsyncResultStatus.Succeeded);
  host.fireSelectionChanged();
  assert.equal(fired, 1);
});

test("every scenario has a label and a selection", () => {
  assert.ok(Object.keys(SCENARIOS).length >= 10);
  for (const [key, scenario] of Object.entries(SCENARIOS)) {
    assert.equal(typeof scenario.label, "string", key);
    assert.ok(Array.isArray(scenario.selected), key);
    assert.ok(scenario.highlight === null || typeof scenario.highlight === "string", key);
  }
});
