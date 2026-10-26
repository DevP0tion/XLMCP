import { test, expect, beforeAll, afterAll } from "bun:test";
import { dispose } from "../../src/services/powershell.js";
import { callTool } from "../helpers/tools.js";
import { describeExcel, openScratch, closeScratch, type Scratch } from "../helpers/scratch.js";

// P1-3 follow-up: chunks run sequentially in one script, so many small chunks only add per-chunk
// script overhead (default 30 rows: read 9.2 s / write 6.5 s for 2500x29). Thresholds leave margin.
describeExcel("large range I/O with default chunking (Excel)", () => {
  let s: Scratch;
  const data = Array.from({ length: 2500 }, (_, r) =>
    Array.from({ length: 29 }, (_, c) => (c % 3 === 0 ? `item-${r}-${c}` : String(r * 29 + c + 0.25)))
  );
  beforeAll(async () => { s = await openScratch(); });
  afterAll(async () => { await closeScratch(s); await dispose(); });

  test("2500x29 write and read round-trip within 5 s each", async () => {
    let t = Date.now();
    await callTool("excel_write_range", { workbook: s.name, sheet: "Sheet1", startCell: "A1", data });
    const writeMs = Date.now() - t;
    t = Date.now();
    const res = await callTool("excel_read_range", { workbook: s.name, sheet: "Sheet1", range: "A1:AC2500" });
    const readMs = Date.now() - t;
    expect(res.Data.map((row: unknown[]) => row.map(String))).toEqual(data);
    expect({ writeMs: writeMs < 5000, readMs: readMs < 5000 }).toEqual({ writeMs: true, readMs: true });
  });
});
