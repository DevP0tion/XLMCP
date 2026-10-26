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
});
