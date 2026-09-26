// In-memory stand-in for the parts of the Office.js PowerPoint API this add-in
// uses. The unit tests and the browser preview (dev/preview.html) run against it.
//
// It copies the behaviour of the real proxy model that can trip code up:
// - properties must be load()ed and synced before they can be read;
// - navigating to something a shape doesn't have (textFrame on a table, group on
//   a non-group, getTable() on a text box, placeholderFormat on a non-placeholder)
//   is only detected at context.sync(), and it rejects the whole batch;
// - *OrNullObject methods return an object whose isNullObject is true instead;
// - proxies stop working once the PowerPoint.run that created them has finished.
//
// Fake shapes are plain objects: see dev/scenarios.js for examples.

export class FakeOfficeError extends Error {
  constructor(code, message) {
    super(message);
    this.name = "RichApi.Error";
    this.code = code;
  }
}

const NULL = Symbol("null object");
const invalidArgument = () => new FakeOfficeError("InvalidArgument", "The argument is invalid or missing or has an incorrect format.");
const generalException = () => new FakeOfficeError("GeneralException", "There was an internal error while processing the request.");

class FakeProxy {
  constructor(context, resolve, { nullable = false } = {}) {
    this._context = context;
    this._resolve = resolve; // () => data or NULL; throws FakeOfficeError when invalid
    this._nullable = nullable;
    this._values = new Map();
  }

  load(spec = "") {
    this._context._assertOpen();
    const paths = (Array.isArray(spec) ? spec : String(spec).split(",")).map((p) => p.trim()).filter(Boolean);
    const own = [];
    const items = [];
    for (const path of paths) {
      const slash = path.indexOf("/");
      if (slash === -1) own.push(path);
      else if (path.slice(0, slash) === "items" && this instanceof FakeCollection) items.push(path.slice(slash + 1));
      else this[path.slice(0, slash)].load(path.slice(slash + 1));
    }
    this._context._queue.push({ proxy: this, own, items });
    return this;
  }

  get isNullObject() {
    return this._get("isNullObject");
  }

  _get(property) {
    if (!this._values.has(property)) {
      throw new FakeOfficeError(
        "PropertyNotLoaded",
        `The property '${property}' is not available. Before reading the property's value, call the load method on the containing object and call "context.sync()" on the associated request context.`,
      );
    }
    return this._values.get(property);
  }

  _apply(data, own) {
    this._values.set("isNullObject", data === NULL);
    for (const property of own) this._values.set(property, data === NULL ? null : this._read(data, property));
  }

  _read(_data, property) {
    throw new FakeOfficeError("InvalidArgument", `Unknown property '${property}'.`);
  }

  _child(key, create) {
    this[key] ??= create();
    return this[key];
  }
}

class FakeCollection extends FakeProxy {
  constructor(context, resolve, Item) {
    super(context, resolve);
    this._Item = Item;
  }

  get items() {
    return this._get("items");
  }

  _apply(list, own, items = []) {
    super._apply(list, own.filter((p) => p !== "items"));
    const props = own.includes("items") ? this._Item.SCALARS : items;
    this._values.set(
      "items",
      list.map((data) => {
        const item = new this._Item(this._context, () => data);
        item._apply(data, props);
        return item;
      }),
    );
  }
}

function hasTextFrame(shape) {
  return typeof shape.text === "string" && shape.type !== "Table" && shape.type !== "Group";
}

class FakeShape extends FakeProxy {
  static SCALARS = ["id", "name", "type", "level"];

  get id() {
    return this._get("id");
  }
  get name() {
    return this._get("name");
  }
  get type() {
    return this._get("type");
  }
  get level() {
    return this._get("level");
  }

  _read(shape, property) {
    if (property === "level") return shape.level ?? 0;
    if (FakeShape.SCALARS.includes(property)) return shape[property];
    return super._read(shape, property);
  }

  /** Forbidden in the add-in: throws for shapes without a text frame. */
  get textFrame() {
    this._context._host._forbidden("Shape.textFrame");
    return this._child("_textFrame", () =>
      new FakeTextFrame(this._context, () => {
        const shape = this._resolve();
        if (!hasTextFrame(shape)) throw invalidArgument();
        return shape;
      }),
    );
  }

  getTextFrameOrNullObject() {
    return new FakeTextFrame(
      this._context,
      () => {
        const shape = this._resolve();
        if (shape.failTextFrame) throw generalException();
        return hasTextFrame(shape) ? shape : NULL;
      },
      { nullable: true },
    );
  }

  getTable() {
    return new FakeTable(this._context, () => {
      const shape = this._resolve();
      if (!Array.isArray(shape.rows)) throw generalException();
      return shape;
    });
  }

  get group() {
    return this._child("_group", () =>
      new FakeShapeGroup(this._context, () => {
        const shape = this._resolve();
        if (shape.type !== "Group") throw generalException();
        return shape;
      }),
    );
  }

  get placeholderFormat() {
    return this._child("_placeholderFormat", () =>
      new FakePlaceholderFormat(this._context, () => {
        const shape = this._resolve();
        if (shape.type !== "Placeholder") throw generalException();
        return shape;
      }),
    );
  }
}

class FakeTextFrame extends FakeProxy {
  get hasText() {
    return this._get("hasText");
  }

  _read(shape, property) {
    if (property === "hasText") return shape.text.length > 0;
    return super._read(shape, property);
  }

  get textRange() {
    return this._child("_textRange", () =>
      new FakeTextRange(this._context, () => {
        const shape = this._resolve();
        // Office.js doesn't document navigating from a null object, so refuse it.
        if (shape === NULL) throw invalidArgument();
        return { text: shape.text };
      }),
    );
  }
}

class FakeTextRange extends FakeProxy {
  get text() {
    return this._get("text");
  }

  _read(range, property) {
    if (property === "text") return range.text;
    return super._read(range, property);
  }
}

function mergeCovering(shape, row, column) {
  return (shape.merged ?? []).find(
    (m) => row >= m.row && row < m.row + m.rowCount && column >= m.column && column < m.column + m.columnCount,
  );
}

class FakeTable extends FakeProxy {
  get rowCount() {
    return this._get("rowCount");
  }
  get columnCount() {
    return this._get("columnCount");
  }
  get values() {
    return this._get("values");
  }

  _read(shape, property) {
    if (property === "rowCount") return shape.rows.length;
    if (property === "columnCount") return Math.max(0, ...shape.rows.map((r) => r.length));
    if (property === "values") {
      // PowerPoint doesn't document what it reports for cells hidden under a
      // merge. The fake repeats the merged text there (the worst case), so any
      // code that counts from `values` double counts and fails its tests.
      return shape.rows.map((cells, r) =>
        cells.map((text, c) => {
          const merge = mergeCovering(shape, r, c);
          return merge ? shape.rows[merge.row][merge.column] : text;
        }),
      );
    }
    return super._read(shape, property);
  }

  getCellOrNullObject(row, column) {
    return new FakeTableCell(
      this._context,
      () => {
        const shape = this._resolve();
        if (row < 0 || row >= shape.rows.length || column < 0 || column >= (shape.rows[row]?.length ?? 0)) return NULL;
        const merge = mergeCovering(shape, row, column);
        if (merge && (merge.row !== row || merge.column !== column)) return NULL;
        return { text: shape.rows[row][column], rowIndex: row, columnIndex: column };
      },
      { nullable: true },
    );
  }
}

class FakeTableCell extends FakeProxy {
  get text() {
    return this._get("text");
  }

  _read(cell, property) {
    if (property in cell) return cell[property];
    return super._read(cell, property);
  }
}

class FakeShapeGroup extends FakeProxy {
  get id() {
    return this._get("id");
  }

  _read(shape, property) {
    if (property === "id") return shape.id;
    return super._read(shape, property);
  }

  get shapes() {
    return this._child("_shapes", () => new FakeCollection(this._context, () => this._resolve().shapes, FakeShape));
  }
}

class FakePlaceholderFormat extends FakeProxy {
  get containedType() {
    return this._get("containedType");
  }
  get type() {
    return this._get("type");
  }

  _read(shape, property) {
    if (property === "containedType") return shape.containedType ?? null;
    if (property === "type") return shape.placeholderType ?? "Body";
    return super._read(shape, property);
  }
}

class FakePresentation {
  constructor(context) {
    this._context = context;
  }

  getSelectedShapes() {
    return new FakeCollection(this._context, () => this._context._host.scenario.selected, FakeShape);
  }

  getSelectedTextRangeOrNullObject() {
    return new FakeTextRange(
      this._context,
      () => {
        const highlight = this._context._host.scenario.highlight;
        if (highlight == null) return NULL;
        // An object stands for an odd range: { text: null }, as PowerPoint for Mac returns
        // for charts and table cells, or { error: code } for a range that fails to load.
        if (typeof highlight === "object" && highlight.error) throw new FakeOfficeError(highlight.error, "The argument is invalid or missing or has an incorrect format.");
        return typeof highlight === "object" ? highlight : { text: highlight };
      },
      { nullable: true },
    );
  }

  /** Forbidden in the add-in: throws when no text is selected. */
  getSelectedTextRange() {
    this._context._host._forbidden("Presentation.getSelectedTextRange");
    return new FakeTextRange(this._context, () => {
      const highlight = this._context._host.scenario.highlight;
      if (highlight == null) throw generalException();
      return { text: highlight };
    });
  }
}

class FakeRequestContext {
  constructor(host) {
    this._host = host;
    this._queue = [];
    this.presentation = new FakePresentation(this);
  }

  _assertOpen() {
    if (this._closed) {
      throw new FakeOfficeError("InvalidObjectPath", "The object path isn't working for what you're trying to do. Objects can't be used after their PowerPoint.run has finished.");
    }
  }

  async sync() {
    const host = this._host;
    const number = ++host.stats.syncs;
    const queue = this._queue;
    this._queue = [];
    if (host.latencyMs > 0) await new Promise((resolve) => setTimeout(resolve, host.latencyMs));
    if (host._failSyncs.delete(number)) {
      throw new FakeOfficeError("GeneralException", "The property 'items' is not available. Before reading the property's value, call the load method.");
    }
    // A "trampled" sync resolves but loads nothing, like office-js #6363: reading
    // any property loaded in it then throws PropertyNotLoaded.
    if (host._trampleSyncs.delete(number)) return;
    // Resolve everything first: one failure rejects the whole batch and applies nothing.
    const resolved = queue.map((op) => op.proxy._resolve());
    queue.forEach((op, i) => op.proxy._apply(resolved[i], op.own, op.items));
  }
}

function compareVersions(a, b) {
  const pa = String(a).split(".").map(Number);
  const pb = String(b).split(".").map(Number);
  for (let i = 0; i < Math.max(pa.length, pb.length); i++) {
    const diff = (pa[i] ?? 0) - (pb[i] ?? 0);
    if (diff !== 0) return Math.sign(diff);
  }
  return 0;
}

/**
 * Creates a fake Office + PowerPoint pair.
 * @param {{ highlight: string | null, selected: object[] }} scenario
 * @param {{ platform?: string, apiVersion?: string, theme?: object | null, latencyMs?: number }} [options]
 */
export function createFakeHost(scenario = { highlight: null, selected: [] }, { platform = "Mac", apiVersion = "1.10", theme = null, latencyMs = 0 } = {}) {
  const selectionHandlers = [];
  const host = {
    scenario,
    latencyMs,
    stats: { runs: 0, syncs: 0, forbidden: [], selectedDataCalls: 0 },
    hangSelectedData: false,
    _failSyncs: new Set(),
    _trampleSyncs: new Set(),
    _forbidden(api) {
      host.stats.forbidden.push(api);
    },
    setScenario(next) {
      host.scenario = next;
    },
    /** The next `count` syncs reject, like a transient host error. */
    failNextSyncs(count = 1) {
      for (let i = 1; i <= count; i++) host._failSyncs.add(host.stats.syncs + i);
    },
    /** Sync number `n` (counting from the host's first sync, 1-based) rejects. */
    failSync(n) {
      host._failSyncs.add(n);
    },
    /** Sync number `n` resolves without loading anything. */
    trampleSync(n) {
      host._trampleSyncs.add(n);
    },
    fireSelectionChanged() {
      for (const handler of selectionHandlers) handler({ type: "documentSelectionChanged" });
    },
  };

  const PowerPoint = {
    async run(callback) {
      host.stats.runs++;
      const context = new FakeRequestContext(host);
      try {
        return await callback(context);
      } finally {
        context._closed = true;
      }
    },
  };

  const Office = {
    HostType: { PowerPoint: "PowerPoint" },
    PlatformType: { PC: "PC", OfficeOnline: "OfficeOnline", Mac: "Mac", iOS: "iOS", Android: "Android", Universal: "Universal" },
    EventType: { DocumentSelectionChanged: "documentSelectionChanged" },
    CoercionType: { Text: "text" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    onReady: async () => ({ host: "PowerPoint", platform }),
    context: {
      platform,
      officeTheme: theme ?? undefined,
      requirements: {
        isSetSupported: (name, version = "1.1") => name === "PowerPointApi" && compareVersions(apiVersion, version) >= 0,
      },
      document: {
        /**
         * Plain-text selection: the scenario's `selectedText`, or else its highlight
         * (or "" with none). `hangSelectedData` makes it never answer.
         */
        getSelectedDataAsync(coercionType, optionsOrCallback, maybeCallback) {
          const callback = typeof optionsOrCallback === "function" ? optionsOrCallback : maybeCallback;
          host.stats.selectedDataCalls++;
          if (host.hangSelectedData) return;
          const { selectedText, highlight } = host.scenario;
          const value = selectedText ?? (typeof highlight === "string" ? highlight : "");
          setTimeout(() => callback?.({ status: "succeeded", value }), host.latencyMs);
        },
        addHandlerAsync(eventType, handler, optionsOrCallback, maybeCallback) {
          const callback = typeof optionsOrCallback === "function" ? optionsOrCallback : maybeCallback;
          if (eventType === "documentSelectionChanged") selectionHandlers.push(handler);
          callback?.({ status: "succeeded", value: undefined });
        },
      },
    },
  };

  return { Office, PowerPoint, host };
}
