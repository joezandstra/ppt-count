// Reads the current PowerPoint selection and returns the text to count.
//
// This is the only module that talks to the PowerPoint API. Several Office.js
// calls fail the whole context.sync() when used on the wrong kind of shape
// (textFrame on a table, group on a non-group...), so each step checks the
// shape's type first, and a failed batch is retried one job at a time so one
// bad shape can't hide the rest.
// Design: docs/superpowers/specs/2026-09-26-ppt-word-count-design.md, section 4.3.

/**
 * @typedef {{ id: string, name: string, kind: "highlight" | "text" | "table" | "group", texts: string[] }} SelectionItem
 * @typedef {{ id: string, name: string, type: string }} UnsupportedShape
 * @typedef {{ source: "highlight" | "shapes" | "none", selectedCount: number, items: SelectionItem[], unsupported: UnsupportedShape[] }} SelectionSnapshot
 */

const SHAPE_FIELDS = "items/id,items/name,items/type";
// Shapes that never hold text.
const NO_TEXT = new Set(["Image", "Line", "Media", "Model3D", "Ink", "ContentApp", "Graphic"]);
// Shapes that show text the API can't read.
const UNREADABLE = new Set(["Chart", "SmartArt", "Diagram", "Ole"]);

/**
 * @param {PowerPoint.RequestContext} context
 * @param {{ selectedText?: () => Promise<string | null> }} [options]
 *   `selectedText` returns the selection as plain text through Office's older
 *   common API (Office.context.document.getSelectedDataAsync). It's the only way
 *   to see text highlighted inside a table cell; see readTable.
 * @returns {Promise<SelectionSnapshot>}
 */
export async function readSelection(context, { selectedText } = {}) {
  const presentation = context.presentation;
  let range = presentation.getSelectedTextRangeOrNullObject();
  range.load("text");
  let selection = presentation.getSelectedShapes();
  selection.load(SHAPE_FIELDS);
  try {
    await context.sync();
  } catch {
    // The text range can fail for text inside table cells (PowerPoint for Mac
    // 16.113 throws InvalidArgument for some loads), so read the shapes on their own.
    range = null;
    selection = presentation.getSelectedShapes();
    selection.load(SHAPE_FIELDS);
    await context.sync();
  }

  const shapes = selection.items;
  // With a chart selected, or text highlighted in a table cell, PowerPoint for Mac
  // returns a range whose text is null rather than a null object (seen on 16.113).
  const rangeText = !range || range.isNullObject ? "" : (range.text ?? "");
  // Highlighted text wins unless several shapes are selected (what the range means
  // then is undocumented). Never navigate from the range to its shape: that fails
  // for text inside table cells.
  if (rangeText.length > 0 && shapes.length <= 1) {
    const shape = shapes[0];
    const highlight = {
      source: "highlight",
      selectedCount: shapes.length,
      items: [{ id: shape?.id ?? "", name: shape?.name ?? "", kind: "highlight", texts: [rangeText] }],
      unsupported: [],
    };
    if (!shape) return highlight;
    // PowerPoint for Mac also reports a whole selected box as "selected text"
    // (seen on 16.113). If the range holds all of the shape's text, it's the shape.
    const whole = await readShapes(context, shapes);
    return coversAllText(rangeText, whole) ? whole : highlight;
  }
  if (shapes.length === 0) return { source: "none", selectedCount: 0, items: [], unsupported: [] };
  const snapshot = await readShapes(context, shapes);
  if (shapes.length === 1 && snapshot.items[0]?.kind === "table" && selectedText) {
    return readTable(snapshot, await selectedText());
  }
  return snapshot;
}

// PowerPoint's API has no text range for text highlighted inside a table cell
// (office-js #6906), but the older common API returns the selection as plain text:
// the highlighted text, the selected cells, or every cell when the whole table is
// selected. Anything less than the whole table is counted as a highlight.
function readTable(snapshot, text) {
  if (typeof text !== "string" || !text.trim() || coversAllText(text, snapshot)) return snapshot;
  const [table] = snapshot.items;
  return {
    source: "highlight",
    selectedCount: 1,
    items: [{ id: table.id, name: table.name, kind: "highlight", texts: [text] }],
    unsupported: [],
  };
}

// Compares ignoring whitespace, because how PowerPoint separates paragraphs or
// table cells in a range's text is undocumented.
function coversAllText(rangeText, snapshot) {
  const squash = (text) => text.replace(/\s+/g, "");
  const all = snapshot.items.flatMap((item) => item.texts).join("");
  return squash(all).length > 0 && squash(rangeText) === squash(all);
}

/**
 * Text of each table cell, skipping cells hidden under a merged area. PowerPoint
 * returns a null object for those; only a merge's top-left cell holds its text.
 * @param {Array<{ isNullObject: boolean, text: string }>} cells
 * @returns {string[]}
 */
export function tableCellTexts(cells) {
  return cells.filter((cell) => !cell.isNullObject).map((cell) => cell.text ?? "");
}

// Walks the selected shapes breadth first. Each round is one context.sync(), so a
// selection costs one sync per group-nesting level plus two for reading text.
async function readShapes(context, shapes) {
  const items = [];
  const unsupported = [];
  const seen = new Set();
  let pending = [];

  const reportUnsupported = (shape) => {
    if (!unsupported.some((u) => u.id === shape.id)) unsupported.push({ id: shape.id, name: shape.name, type: shape.type });
  };
  // Queues the work for one shape. Its text goes to `item`, its top-level selected shape.
  const visit = (shape, item) => {
    if (seen.has(shape.id)) return; // a group and its child can both be selected
    seen.add(shape.id);
    if (NO_TEXT.has(shape.type)) return;
    if (UNREADABLE.has(shape.type)) return reportUnsupported(shape);
    if (shape.type === "Group") pending.push(groupJob(shape, item));
    else if (shape.type === "Table") pending.push(tableJob(shape, item));
    else pending.push(textFrameJob(shape, item));
  };

  for (const shape of shapes) {
    const item = { id: shape.id, name: shape.name, kind: kindOf(shape.type), texts: [] };
    items.push(item);
    visit(shape, item);
  }

  while (pending.length > 0) {
    const round = pending;
    pending = [];
    // Guesses (could this odd shape be a table?) get their own batch: a wrong guess
    // fails its sync, and would otherwise force every other job in the round to be
    // retried one at a time.
    const results = [
      ...(await runJobs(context, round.filter((job) => !job.guess), reportUnsupported)),
      ...(await runJobs(context, round.filter((job) => job.guess), reportUnsupported, { guesses: true })),
    ];
    for (const { job, outcome } of results) {
      for (const child of outcome.children ?? []) visit(child, job.item);
      if (outcome.next) pending.push(outcome.next);
      if (outcome.unsupported) reportUnsupported(job.shape);
    }
  }

  return {
    source: "shapes",
    selectedCount: shapes.length,
    items: items.filter((item) => item.texts.length > 0),
    unsupported,
  };
}

// Syncs a round of jobs as one batch. When the batch fails, each job is retried on
// its own, so one shape PowerPoint won't read can't hide the rest; those shapes are
// reported. Two kinds of failure are passed on instead, so the caller retries the
// whole read:
// - every job failing on its own too (the problem isn't any particular shape),
//   unless the jobs are guesses that are expected to fail;
// - read() throwing after a successful sync, which is a host glitch (a loaded value
//   that came back missing, office-js #6363), not something about the shape.
async function runJobs(context, jobs, reportUnsupported, { guesses = false } = {}) {
  if (jobs.length === 0) return [];
  for (const job of jobs) job.queue();
  let batchError = null;
  try {
    await context.sync();
  } catch (error) {
    batchError = error;
  }
  if (!batchError) return jobs.map((job) => ({ job, outcome: job.read() }));

  const done = [];
  const failed = [];
  for (const job of jobs) {
    job.queue();
    try {
      await context.sync();
    } catch {
      failed.push(job);
      continue;
    }
    done.push({ job, outcome: job.read() });
  }
  if (done.length === 0 && !guesses) throw batchError;
  for (const job of failed) reportUnsupported(job.shape);
  return done;
}

function kindOf(type) {
  if (type === "Group") return "group";
  if (type === "Table") return "table";
  return "text";
}

// A job queues loads for one shape (queue) and interprets them after the sync
// (read). read() can return group children to visit, a follow-up job for the
// next round, or `unsupported: true`.

function groupJob(shape, item) {
  let children;
  return {
    shape,
    item,
    queue() {
      children = shape.group.shapes;
      children.load(SHAPE_FIELDS);
    },
    read: () => ({ children: children.items }),
  };
}

// `guess: true` when the shape might not be a table at all (see textFrameJob).
function tableJob(shape, item, { guess = false } = {}) {
  let table;
  return {
    shape,
    item,
    guess,
    queue() {
      table = shape.getTable();
      table.load("rowCount,columnCount");
    },
    read: () => ({ next: tableCellsJob(shape, item, table) }),
  };
}

// Reads cells one by one rather than Table.values: getCellOrNullObject is
// documented to return a null object for cells hidden under a merge, but what
// `values` holds for them is not documented.
function tableCellsJob(shape, item, table) {
  let cells;
  return {
    shape,
    item,
    queue() {
      cells = [];
      for (let row = 0; row < table.rowCount; row++) {
        for (let column = 0; column < table.columnCount; column++) {
          const cell = table.getCellOrNullObject(row, column);
          cell.load("text");
          cells.push(cell);
        }
      }
    },
    read() {
      if (item.id === shape.id) item.kind = "table";
      item.texts.push(...tableCellTexts(cells));
      return {};
    },
  };
}

function textFrameJob(shape, item) {
  let frame;
  let format = null;
  return {
    shape,
    item,
    queue() {
      frame = shape.getTextFrameOrNullObject();
      frame.load("hasText");
      if (shape.type === "Placeholder") {
        format = shape.placeholderFormat;
        format.load("containedType");
      }
    },
    read() {
      if (!frame.isNullObject) return { next: textRangeJob(shape, item, frame) };
      const contained = format?.containedType;
      // Tables in content placeholders, and tables on builds that report their
      // type as "Unsupported", have no text frame but can still be read as tables.
      if (contained === "Table") return { next: tableJob(shape, item) };
      if (shape.type === "Unsupported") return { next: tableJob(shape, item, { guess: true }) };
      if (UNREADABLE.has(contained)) return { unsupported: true };
      return {};
    },
  };
}

// The frame's text is read in a second round: navigating from a frame that might
// be a null object into its textRange is not documented to be safe.
function textRangeJob(shape, item, frame) {
  let range;
  return {
    shape,
    item,
    queue() {
      range = frame.textRange;
      range.load("text");
    },
    read() {
      item.texts.push(range.text ?? "");
      return {};
    },
  };
}
