import { test, expect, beforeAll, afterAll } from "bun:test";
import { dispose } from "../../src/services/powershell.js";
import { callTool } from "../helpers/tools.js";
import { describeExcel, openScratch, closeScratch, xl, type Scratch } from "../helpers/scratch.js";

const XL_WHOLE = 1;
const XL_PART = 2;
const words = [["apple"], ["pineapple"], ["Apple pie"]];

describeExcel("excel_find_replace (Excel)", () => {
  let s: Scratch;
  const read = async (range: string) =>
    (await callTool("excel_read_range", { workbook: s.name, sheet: "Sheet1", range })).Data;
  // Simulates the user's last Find dialog settings: Excel saves LookAt on every Find/Replace call.
  const primeLookAt = (lookAt: number) =>
    xl(s, `[void]$wb.Worksheets.Item('Sheet1').Range('A1:A3').Find('zzz', [Type]::Missing, -4163, ${lookAt}, 1, 1, $false)`);

  beforeAll(async () => {
    s = await openScratch();
    for (const col of ["A", "C", "D"]) {
      await callTool("excel_write_range", { workbook: s.name, sheet: "Sheet1", startCell: `${col}1`, data: words });
    }
  });
  afterAll(async () => { await closeScratch(s); await dispose(); });

  // P0-5: omitted Find/Replace arguments reused the saved (dialog) settings.
  test("find gives the same count regardless of prior Find settings", async () => {
    const args = { workbook: s.name, sheet: "Sheet1", range: "A1:A3", find: "apple" };
    await primeLookAt(XL_WHOLE);
    const afterWhole = (await callTool("excel_find_replace", args)).Count;
    await primeLookAt(XL_PART);
    const afterPart = (await callTool("excel_find_replace", args)).Count;
    expect(afterPart).toBe(afterWhole);
  });

  test("replace gives the same result regardless of prior Find settings", async () => {
    await primeLookAt(XL_WHOLE);
    await callTool("excel_find_replace", { workbook: s.name, sheet: "Sheet1", range: "C1:C3", find: "apple", replace: "X" });
    await primeLookAt(XL_PART);
    await callTool("excel_find_replace", { workbook: s.name, sheet: "Sheet1", range: "D1:D3", find: "apple", replace: "X" });
    expect(await read("D1:D3")).toEqual(await read("C1:C3"));
  });

  // P0-6: COM Replace returned True on a protected sheet while nothing changed.
  test("replace on a protected sheet fails instead of reporting success", async () => {
    await xl(s, `$p = $wb.Worksheets.Add(); $p.Name = 'Prot'; $p.Range('A1').Value2 = 'apple'; $p.Protect()`);
    const err = await callTool("excel_find_replace", { workbook: s.name, sheet: "Prot", range: "A1", find: "apple", replace: "X" })
      .then(() => null, (e: Error) => e);
    expect(err?.message ?? "resolved").toContain("protected");
    expect((await xl(s, `$wb.Worksheets.Item('Prot').Range('A1').Value2`)).trim()).toBe("apple");
  });

  // P1-2: find-only used FindNext + Address() x3 + Value2 per match (and $results +=).
  test("find-only over 600x29 with thousands of part matches is fast and exact", async () => {
    const data = Array.from({ length: 600 }, (_, r) => Array.from({ length: 29 }, (_, c) => String(r * 29 + c)));
    await xl(s, `$n = $wb.Worksheets.Add(); $n.Name = 'Big'`);
    await callTool("excel_write_range", { workbook: s.name, sheet: "Big", startCell: "A1", data });
    const expected = data.flat().filter((v) => v.includes("1")).length;
    const t = Date.now();
    const res = await callTool("excel_find_replace", { workbook: s.name, sheet: "Big", find: "1", lookAt: "part" });
    expect(res.Count).toBe(expected);
    expect(Date.now() - t).toBeLessThan(5000);
  }, 120_000);

  test("lookIn values matches the raw value, not the displayed text", async () => {
    await xl(s, `$c = $wb.Worksheets.Item('Sheet1').Range('H1'); $c.Value2 = 1000; $c.NumberFormat = '#,##0'`);
    const args = { workbook: s.name, sheet: "Sheet1", range: "H1:H2" };
    expect((await callTool("excel_find_replace", { ...args, find: "1000" })).Count).toBe(1);
    expect((await callTool("excel_find_replace", { ...args, find: "1,000" })).Count).toBe(0);
  });

  test("find-only returns ranges, and cells only when asked", async () => {
    const args = { workbook: s.name, sheet: "Sheet1", range: "A1:A3", find: "apple" };
    const plain = await callTool("excel_find_replace", args);
    expect(plain).toEqual({ Count: 1, Ranges: ["A1"] });
    const withCells = await callTool("excel_find_replace", { ...args, includeCells: true });
    expect(withCells.Cells).toEqual([{ Cell: "A1", Value: "apple" }]);
  });

  test("replace reports the number of cells actually changed", async () => {
    await callTool("excel_write_range", { workbook: s.name, sheet: "Sheet1", startCell: "F1", data: [["a-1"], ["b"], ["a-2"]] });
    const res = await callTool("excel_find_replace", {
      workbook: s.name, sheet: "Sheet1", range: "F1:F3", find: "a-", replace: "z-", lookAt: "part",
    });
    expect(res.Replaced).toBe(2);
  });
});
