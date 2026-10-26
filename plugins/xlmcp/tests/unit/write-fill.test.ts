import { test, expect } from "bun:test";
import { existsSync, unlinkSync } from "fs";
import { buildFill, formulaRuns } from "../../src/tools/cell/write-range.tool.js";
import { runPowerShell } from "../helpers/ps.js";

/** Run write_range's fill fragment in PS 5.1 and return $arr as { value, type } cells. */
export function runFill(data: Parameters<typeof buildFill>[0], formulas = true) {
  const fill = buildFill(data, 0, formulas);
  try {
    const out = runPowerShell(`${fill.script}
"DIMS\`t$($arr.GetLength(0))\`t$($arr.GetLength(1))"
for ($i = 0; $i -lt $arr.GetLength(0); $i++) {
  for ($j = 0; $j -lt $arr.GetLength(1); $j++) {
    $v = $arr[$i,$j]
    if ($v -eq $null) { $t = 'null'; $js = 'null' } else { $t = $v.GetType().Name; $js = ConvertTo-Json -InputObject $v -Compress }
    "$i\`t$j\`t$t\`t$js"
  }
}`);
    const lines = out.trim().split(/\r?\n/);
    const [, r, c] = lines[0].split("\t");
    const values: unknown[][] = Array.from({ length: +r }, () => Array(+c).fill(undefined));
    const types: string[][] = Array.from({ length: +r }, () => Array(+c).fill(""));
    for (const line of lines.slice(1)) {
      const [i, j, t, json] = line.split("\t");
      values[+i][+j] = JSON.parse(json);
      types[+i][+j] = t;
    }
    return { values, types, formulas: fill.formulas };
  } finally {
    if (fill.tmpFile && existsSync(fill.tmpFile)) unlinkSync(fill.tmpFile);
  }
}

// P0-1: a single row must keep all its cells (PS unrolled @(@(...)) to a flat array).
test("1xN row keeps every cell", () => {
  expect(runFill([["A", "B", "C"]]).values).toEqual([["A", "B", "C"]]);
});

test("Nx1 column keeps every cell", () => {
  expect(runFill([["A"], ["B"], ["C"]]).values).toEqual([["A"], ["B"], ["C"]]);
});

test("NxM block keeps every cell", () => {
  expect(runFill([["A", "B"], ["C", "D"]]).values).toEqual([["A", "B"], ["C", "D"]]);
});

// P2-2: strict number rule; strings Excel would parse like typed input are protected with a ' prefix.
test("only plain decimal strings become numbers", () => {
  expect(runFill([["12", "-3.5", "0", "0852", " ", "0x1F", "1e3"]]).values).toEqual([
    [12, -3.5, 0, "'0852", " ", "'0x1F", "'1e3"],
  ]);
});

test("strings Excel would parse are stored as literal text", () => {
  expect(runFill([["1/2", "TRUE", "50%", "#N/A", "'q", "abc", "a\nb", "12345678901234567"]]).values).toEqual([
    ["'1/2", "'TRUE", "'50%", "'#N/A", "''q", "abc", "a\nb", "'12345678901234567"],
  ]);
});

test("typed cells keep their type; null and empty string become empty cells", () => {
  const r = runFill([[7, 1.5, true, false, null, ""]] as never);
  expect(r.values).toEqual([[7, 1.5, true, false, null, null]]);
  expect(r.types[0].slice(2, 4)).toEqual(["Boolean", "Boolean"]);
});

test("formulas: false writes '=' text literally", () => {
  const r = runFill([["=1+1"]], false);
  expect(r.values).toEqual([["'=1+1"]]);
  expect(r.formulas).toEqual([]);
});

// P1-4: formulas are assigned per vertical run, not per cell.
test("formula cells group into vertical runs", () => {
  const f = (rowOffset: number, colOffset: number) => ({ rowOffset, colOffset, formula: `=R${rowOffset}C${colOffset}` });
  expect(formulaRuns([f(0, 1), f(1, 1), f(3, 1), f(0, 2), f(2, 1)])).toEqual([
    { r: 0, c: 1, f: ["=R0C1", "=R1C1", "=R2C1", "=R3C1"] },
    { r: 0, c: 2, f: ["=R0C2"] },
  ]);
});
