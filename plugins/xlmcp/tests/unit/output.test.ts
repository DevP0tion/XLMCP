import { test, expect } from "bun:test";
import { textContent } from "../../src/services/utils.js";

const text = (data: unknown) => textContent(data).content[0].text;

// P2-1: responses were pretty-printed (2.2x the characters of compact JSON for a 600x29 read).
test("plain results are compact JSON", () => {
  expect(text({ success: true, rows: 2 })).toBe('{"success":true,"rows":2}');
});

test("a Data 2D array is compact with one row per line", () => {
  const data = { Range: "$A$1:$B$2", Rows: 2, Cols: 2, Data: [[1, "a"], [2, null]] };
  expect(text(data)).toBe('{"Range":"$A$1:$B$2","Rows":2,"Cols":2,"Data":[\n[1,"a"],\n[2,null]\n]}');
  expect(JSON.parse(text(data))).toEqual(data);
});

test("empty Data stays valid JSON", () => {
  const data = { Range: "$A$1", Rows: 0, Cols: 0, Data: [] };
  expect(JSON.parse(text(data))).toEqual(data);
});
