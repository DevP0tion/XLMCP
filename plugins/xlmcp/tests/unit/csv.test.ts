import { test, expect } from "bun:test";
import { toField, toCsv } from "../../src/tools/cell/export-csv.tool.js";
import { fromField, parseCsv } from "../../src/tools/cell/import-csv.tool.js";

// Excel strips one leading ' from a string written through Value2 and stores the rest as text.
const stored = (c: unknown) => (typeof c === "string" && c.startsWith("'") ? c.slice(1) : c);

test("a constant cell keeps its value and type through toField -> fromField", () => {
  const cells = [
    12, -3.5, 0, 0.1 + 0.2, 1e-7, 1e21, 1234567890123456, true, false, null,
    "abc", "Item 1", "0852", "12", "-3.5", "TRUE", "false", "=1+1", "'q", "1e-7", "1e3", " ", "a,b", 'q"q', "a\nb",
  ];
  for (const v of cells) expect([v, stored(fromField(toField(v, v)))]).toEqual([v, v]);
});

// A cell a dynamic array spills into has a value but no formula text. Exported as a constant it would
// block the spill on import, so it is exported empty and the anchor formula refills it.
test("a cell without formula text (empty, or filled by a spill) exports as an empty field", () => {
  expect([toField(null, ""), toField(4, ""), toField("", "")]).toEqual(["", "", ""]);
});

test("only text that would be read back as another type gets a ' prefix", () => {
  expect(["Item 1", "0852", "12", "TRUE", "=1+1", "'q"].map((v) => toField(v, v))).toEqual([
    "Item 1", "0852", "'12", "'TRUE", "'=1+1", "''q",
  ]);
});

test("a formula cell exports its formula, which imports as a formula", () => {
  expect(toField(24, "=B2*2")).toBe("=B2*2");
  expect(fromField("=B2*2")).toBe("=B2*2");
});

test("CSV text round-trips commas, quotes, line breaks and empty fields", () => {
  const rows = [["a,b", 'q"q', "line1\nline2", ""], ["", " x ", "'12", '=IF(A1="x",1,2)'], ["", "", "", ""]];
  expect(parseCsv(toCsv(rows))).toEqual(rows);
});

test("parseCsv accepts LF line ends, a missing final newline and ragged rows", () => {
  expect(parseCsv('a,b\n"c\r\nd"\n\n1,2,3')).toEqual([["a", "b"], ["c\r\nd"], [""], ["1", "2", "3"]]);
});

test("parseCsv rejects an unterminated quoted field instead of swallowing the rest of the file", () => {
  expect(() => parseCsv('a,"b\nc,d\n')).toThrow(/unterminated/i);
});
