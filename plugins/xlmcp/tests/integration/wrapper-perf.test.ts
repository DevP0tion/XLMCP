import { test, expect, beforeAll, afterAll } from "bun:test";
import { dispose } from "../../src/services/powershell.js";
import { describeExcel, openScratch, closeScratch, xl, type Scratch } from "../helpers/scratch.js";

// P1-3 finding: the runPS wrapper's COM cleanup tested `$__obj -ne $null`. For an enumerable COM
// object (an Excel Range) PowerShell's -ne enumerates every cell across the process boundary
// (~1.3 ms/cell), so any script leaving $r on a big range paid ~94 s for A1:AC2500 after its work.
describeExcel("runPS wrapper cleanup cost (Excel)", () => {
  let s: Scratch;
  beforeAll(async () => { s = await openScratch(); });
  afterAll(async () => { await closeScratch(s); await dispose(); });

  test("leaving $r on a 14,500-cell range does not cost per-cell COM calls", async () => {
    const t = Date.now();
    await xl(s, `$r = $wb.Worksheets.Item('Sheet1').Range('A1:AC500'); [void]$r.Rows.Count`);
    expect(Date.now() - t).toBeLessThan(3000);
  });
});
