// Named selections for the unit tests and the browser preview (dev/preview.html).
// Shape formats are described in dev/fake-powerpoint.js.

const SENTENCE = "The quick brown fox jumps over the lazy dog.";

const titlePlaceholder = { id: "101", name: "Title 1", type: "Placeholder", text: "Quarterly results" };
const bodyPlaceholder = {
  id: "102",
  name: "Content Placeholder 2",
  type: "Placeholder",
  text: "Revenue grew 12% year over year\rNew customers in three regions\rCosts held flat",
};
const textBox = { id: "103", name: "TextBox 3", type: "TextBox", text: SENTENCE };
const picture = { id: "104", name: "Picture 4", type: "Image" };
const table = {
  id: "105",
  name: "Table 5",
  type: "Table",
  rows: [
    ["Region", "Q1", "Q2"],
    ["North", "1,200 units", "1,450 units"],
    ["South — combined total", "", ""],
  ],
  merged: [{ row: 2, column: 0, rowCount: 1, columnCount: 3 }],
};

export const SCENARIOS = {
  nothing: { label: "Nothing selected", highlight: null, selected: [] },
  textBox: { label: "A text box", highlight: null, selected: [textBox] },
  title: { label: "A title placeholder", highlight: null, selected: [titlePlaceholder] },
  highlight: { label: "Highlighted text", highlight: "quick brown fox", selected: [textBox] },
  cursor: {
    label: "Cursor inside a word (nothing highlighted)",
    // What PowerPoint for Mac reports: the word around the cursor as the range, but no plain-text selection.
    highlight: "quick",
    selectedText: "",
    selected: [textBox],
  },
  several: { label: "Three shapes (one is a picture)", highlight: null, selected: [titlePlaceholder, bodyPlaceholder, picture] },
  table: { label: "Table with merged cells", highlight: null, selected: [table] },
  tableHighlight: {
    label: "Text highlighted in a table cell",
    // What PowerPoint for Mac reports: a text range without text, plus the plain-text selection.
    highlight: null,
    selectedText: "1,450 units",
    selected: [table],
  },
  group: {
    label: "Nested group",
    highlight: null,
    selected: [
      {
        id: "106",
        name: "Group 6",
        type: "Group",
        shapes: [
          { id: "107", name: "Rectangle 7", type: "GeometricShape", text: "Step one: plan" },
          {
            id: "108",
            name: "Group 8",
            type: "Group",
            shapes: [
              { id: "109", name: "Oval 9", type: "GeometricShape", text: "Step two: build" },
              { id: "110", name: "Arrow 10", type: "GeometricShape", text: "" },
              { id: "111", name: "Picture 11", type: "Image" },
            ],
          },
        ],
      },
    ],
  },
  chart: { label: "Chart and a text box", highlight: null, selected: [{ id: "112", name: "Chart 12", type: "Chart" }, textBox] },
  picture: { label: "A picture (no text)", highlight: null, selected: [picture] },
  smartArt: { label: "A SmartArt graphic (can't be read)", highlight: null, selected: [{ id: "116", name: "SmartArt 16", type: "SmartArt" }] },
  emptyPlaceholder: { label: "An empty placeholder", highlight: null, selected: [{ id: "113", name: "Subtitle 2", type: "Placeholder", text: "" }] },
  tablePlaceholder: {
    label: "Table inside a content placeholder",
    highlight: null,
    selected: [{ id: "114", name: "Content Placeholder 3", type: "Placeholder", containedType: "Table", rows: [["Name", "Role"], ["Ana", "Design lead"]] }],
  },
  cjk: {
    label: "Japanese and English",
    highlight: null,
    selected: [{ id: "115", name: "TextBox 15", type: "TextBox", text: "東京で会議があります。 Meeting at 10:00 in Tokyo." }],
  },
  longNames: {
    label: "Many shapes with long names",
    highlight: null,
    selected: Array.from({ length: 6 }, (_, i) => ({
      id: String(200 + i),
      name: `Speaker notes callout with a very long descriptive name ${i + 1}`,
      type: "TextBox",
      text: "Lorem ipsum dolor sit amet ".repeat(i + 1).trim(),
    })),
  },
};
