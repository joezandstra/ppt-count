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
 * @returns {Promise<SelectionSnapshot>}
 */
export async function readSelection(context) {
  const range = context.presentation.getSelectedTextRangeOrNullObject();
  range.load("text");
  const selection = context.presentation.getSelectedShapes();
  selection.load(SHAPE_FIELDS);
  await context.sync();

  const shapes = selection.items;
  // Highlighted text wins unless several shapes are selected (what the range means
  // then is undocumented). Never navigate from the range to its shape: that fails
  // for text inside table cells.
  if (!range.isNullObject && range.text.length > 0 && shapes.length <= 1) {
    const shape = shapes[0];
    return {
      source: "highlight",
      selectedCount: shapes.length,
      items: [{ id: shape?.id ?? "", name: shape?.name ?? "", kind: "highlight", texts: [range.text] }],
      unsupported: [],
    };
  }
  if (shapes.length === 0) return { source: "none", selectedCount: 0, items: [], unsupported: [] };
  return readShapes(context, shapes);
}

/**
 * Text of each table cell, skipping cells hidden under a merged area. PowerPoint
 * returns a null object for those; only a merge's top-left cell holds its text.
 * @param {Array<{ isNullObject: boolean, text: string }>} cells
 * @returns {string[]}
 */
export function tableCellTexts(cells) {
  return cells.filter((cell) => !cell.isNullObject).map((cell) => cell.text);
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
    for (const { job, outcome } of await runJobs(context, round, reportUnsupported)) {
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

// Syncs a round of jobs as one batch. When the batch fails, each job is retried
// on its own; jobs that still fail are reported instead of breaking the count.
async function runJobs(context, jobs, reportUnsupported) {
  for (const job of jobs) job.queue();
  let batchFailed = false;
  try {
    await context.sync();
  } catch {
    batchFailed = true;
  }
  const done = [];
  for (const job of jobs) {
    try {
      if (batchFailed) {
        job.queue();
        await context.sync();
      }
      done.push({ job, outcome: job.read() });
    } catch {
      reportUnsupported(job.shape);
    }
  }
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

function tableJob(shape, item) {
  let table;
  return {
    shape,
    item,
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
      if (contained === "Table" || shape.type === "Unsupported") return { next: tableJob(shape, item) };
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
      item.texts.push(range.text);
      return {};
    },
  };
}
