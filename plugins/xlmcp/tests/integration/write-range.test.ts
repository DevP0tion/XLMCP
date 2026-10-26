import { test, expect, beforeAll, afterAll } from "bun:test";
import { dispose } from "../../src/services/powershell.js";
import { callTool } from "../helpers/tools.js";
import { describeExcel, openScratch, closeScratch, xl, type Scratch } from "../helpers/scratch.js";

describeExcel("excel_write_range (Excel)", () => {
  let s: Scratch;
  beforeAll(async () => { s = await openScratch(); });
  afterAll(async () => { await closeScratch(s); await dispose(); });

  // P0-1: one-row writes used to write only the first cell.
  test("1x13 row round-trips", async () => {
    const row = Array.from({ length: 13 }, (_, i) => `v${i + 1}`);
    await callTool("excel_write_range", { workbook: s.name, sheet: "Sheet1", startCell: "A1", data: [row] });
    const res = await callTool("excel_read_range", { workbook: s.name, sheet: "Sheet1", range: "A1:M1" });
    expect(res.Data).toEqual([row]);
  });

  // P2-2: Excel parses strings written through Value2 like typed input; they must come back verbatim.
  test("strings keep their exact text and typed cells their type", async () => {
    const row = ["0852", " ", "0x1F", "1e3", "1/2", "TRUE", "50%", "'q", "12345678901234567", "a\nb", 12, -3.5, true, 2147483648];
    await callTool("excel_write_range", { workbook: s.name, sheet: "Sheet1", startCell: "A300", data: [row] });
    const res = await callTool("excel_read_range", { workbook: s.name, sheet: "Sheet1", range: "A300:N300" });
    expect(res.Data).toEqual([row]);
  });

  test('"" leaves a blank cell, even in a Text-formatted column', async () => {
    await xl(s, `$wb.Worksheets.Item('Sheet1').Range('C310:C311').NumberFormat = '@'`);
    await callTool("excel_write_range", { workbook: s.name, sheet: "Sheet1", startCell: "A310", data: [["x", "", "", "y"]] });
    const blank = await xl(s, `$ws = $wb.Worksheets.Item('Sheet1'); "$($ws.Evaluate('ISBLANK(B310)'))|$($ws.Evaluate('ISBLANK(C310)'))"`);
    expect(blank.trim()).toBe("True|True");
  });

  test("null leaves a blank cell", async () => {
    await callTool("excel_write_range", { workbook: s.name, sheet: "Sheet1", startCell: "A311", data: [["x", null, null, "y"]] });
    const blank = await xl(s, `$ws = $wb.Worksheets.Item('Sheet1'); "$($ws.Evaluate('ISBLANK(B311)'))|$($ws.Evaluate('ISBLANK(C311)'))"`);
    expect(blank.trim()).toBe("True|True");
  });

  test("formulas: false writes text that starts with =", async () => {
    await callTool("excel_write_range", { workbook: s.name, sheet: "Sheet1", startCell: "A320", data: [["=1+1"]], formulas: false });
    const res = await xl(s, `$c = $wb.Worksheets.Item('Sheet1').Range('A320'); "$($c.Value2)|$($c.HasFormula)"`);
    expect(res.trim()).toBe("=1+1|False");
  });

  // P2-3: verify reads back Value2 (and .Formula for formula cells) instead of the caller re-reading everything.
  test("verify: a faithful write reports no mismatches", async () => {
    const data = [["0852", 12, "=B330*2", null], ["a\nb", true, "=sum(b331:b331)", "1/2"]];
    const res = await callTool("excel_write_range", { workbook: s.name, sheet: "Sheet1", startCell: "A330", data, verify: true });
    expect(res.verify).toEqual({ mismatches: 0, first: [] });
  });

  test("verify: a value Excel changed is reported briefly with its cell", async () => {
    // Excel silently truncates cell text to 32,767 characters
    const res = await callTool("excel_write_range", {
      workbook: s.name, sheet: "Sheet1", startCell: "A340", data: [["ok", "x".repeat(40000)]], verify: true,
    });
    expect(res.verify.mismatches).toBe(1);
    expect(res.verify.first[0].cell).toBe("B340");
    expect(JSON.stringify(res.verify).length).toBeLessThan(1000);
  });

  // P1-4: formulas were assigned one cell (one COM call + one script line) at a time.
  test("200 formula cells read back as written", async () => {
    const data = Array.from({ length: 200 }, (_, i) => [String(i), `=A${i + 3}*2`]);
    await callTool("excel_write_range", { workbook: s.name, sheet: "Sheet1", startCell: "A3", data });
    const res = await callTool("excel_read_range_formulas", { workbook: s.name, sheet: "Sheet1", range: "B3:B202" });
    expect(res.Data.map((r: unknown[]) => r[0])).toEqual(data.map((r) => r[1]));
  });

  test("2000 formula cells are written in bulk (< 1.5 s)", async () => {
    const data = Array.from({ length: 2000 }, (_, i) => [`=ROW()+${i}`]);
    const t = Date.now();
    await callTool("excel_write_range", { workbook: s.name, sheet: "Sheet1", startCell: "D1", data });
    expect(Date.now() - t).toBeLessThan(1500);
  });
});
