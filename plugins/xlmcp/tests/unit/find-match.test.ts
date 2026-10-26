import { test, expect } from "bun:test";
import { matchCells, toRanges } from "../../src/tools/data/find-replace.tool.js";

const whole = { lookAt: "whole", matchCase: false } as const;
const part = { lookAt: "part", matchCase: false } as const;

// P1-2: find-only matches a Value2 array in TS instead of one COM round trip per match.
test("whole vs part, case-insensitive unless matchCase", () => {
  const data = [["apple", "Pineapple"], ["APPLE", null]];
  expect(matchCells(data, "apple", whole)).toEqual([[0, 0], [1, 0]]);
  expect(matchCells(data, "apple", part)).toEqual([[0, 0], [0, 1], [1, 0]]);
  expect(matchCells(data, "apple", { lookAt: "whole", matchCase: true })).toEqual([[0, 0]]);
});

test("numbers and booleans match their raw Value2 text, not formatted text", () => {
  expect(matchCells([[1000, 1.5, true]], "1000", whole)).toEqual([[0, 0]]);
  expect(matchCells([[1000]], "1,000", whole)).toEqual([]);
  expect(matchCells([[true]], "TRUE", { lookAt: "whole", matchCase: true })).toEqual([[0, 0]]);
  expect(matchCells([[null, ""]], "", part)).toEqual([]);
});

test("matches compress into contiguous column runs", () => {
  // range starts at F3 (row 3, column 6)
  expect(toRanges([[0, 0], [1, 0], [2, 0], [5, 0], [0, 2]], 3, 6)).toEqual(["F3:F5", "F8", "H3"]);
});
