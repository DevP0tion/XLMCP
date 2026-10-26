import { test, expect, beforeAll, afterAll } from "bun:test";
import { dispose } from "../../src/services/powershell.js";
import { callTool } from "../helpers/tools.js";
import { describeExcel, openScratch, closeScratch, type Scratch } from "../helpers/scratch.js";

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
