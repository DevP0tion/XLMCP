import { test, expect, beforeAll, afterAll } from "bun:test";
import { existsSync, readFileSync, unlinkSync, writeFileSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";
import { randomUUID } from "crypto";
import { dispose } from "../../src/services/powershell.js";
import { callTool } from "../helpers/tools.js";
import { describeExcel, openScratch, closeScratch, xl, type Scratch } from "../helpers/scratch.js";

describeExcel("excel_export_csv / excel_import_csv (Excel)", () => {
  let s: Scratch;
  const csv = join(tmpdir(), `xlmcp_it_${randomUUID()}.csv`);
  beforeAll(async () => { s = await openScratch(); });
  afterAll(async () => {
    await closeScratch(s);
    await dispose();
    if (existsSync(csv)) unlinkSync(csv);
  });

  // Not expect(promise).rejects: bun 1.4 waits for it without reading the PowerShell pipe, so a tool call
  // that reaches Excel stalls until the 30 s invoke timeout.
  const failure = (call: Promise<unknown>) => call.then(() => "resolved", (e) => String(e));

  const snapshot = async (sheet: string, range: string) => [
    (await callTool("excel_read_range", { workbook: s.name, sheet, range })).Data,
    (await callTool("excel_read_range_formulas", { workbook: s.name, sheet, range })).Data,
  ];

  test("export then import leaves every cell's value, type and formula unchanged", async () => {
    const at = { workbook: s.name, sheet: "Sheet1" };
    await callTool("excel_write_range", {
      ...at,
      startCell: "B2", // not A1: a full-sheet export must still start at A1 so the default import lands in place
      data: [
        ["name", "qty", "total", "note"],
        ["a,b", 12, "=C3*2", 'say "hi"'],
        ["0852", -3.5, "=SUM(C3:C4)", "line1\nline2"],
        [true, 0.1 + 0.2, null, "Item 1"],
      ],
    });
    // text that looks like another type, and numbers outside the plain-decimal rule: not reachable through write_range
    await xl(s, `
      $ws = $wb.Worksheets.Item('Sheet1')
      $ws.Range('B6').Value2 = "'12"; $ws.Range('C6').Value2 = "'TRUE"; $ws.Range('D6').Value2 = "'=1+1"
      $ws.Range('B7').Value2 = [double]1e-7; $ws.Range('C7').Value2 = [double]1e21; $ws.Range('D7').Value2 = [double]1234567890123456
    `);
    const before = await snapshot("Sheet1", "A1:E7");

    const exported = await callTool("excel_export_csv", { ...at, path: csv });
    expect(exported.Range).toBe("$A$1:$E$7");
    await callTool("excel_clear_range", { ...at, range: "A1:E7", mode: "values" });
    await callTool("excel_import_csv", { ...at, path: csv });

    expect(await snapshot("Sheet1", "A1:E7")).toEqual(before);
    expect(before[0][5].slice(1, 4)).toEqual(["12", "TRUE", "=1+1"]); // the text cells really were text
    expect(before[1][2][3]).toBe("=C3*2");
  });

  test("the exported file is UTF-8 with a BOM and keeps non-ASCII text", async () => {
    const at = { workbook: s.name, sheet: "Sheet1" };
    await callTool("excel_write_range", { ...at, startCell: "H1", data: [["한글", "é"]] });
    await callTool("excel_export_csv", { ...at, range: "H1:I1", path: csv });
    expect(readFileSync(csv, "utf-8")).toBe("﻿한글,é\r\n");
  });

  test("imports a hand-edited file at startCell: ragged rows, ' text prefix, formulas", async () => {
    const at = { workbook: s.name, sheet: "Sheet1" };
    await callTool("excel_write_range", { ...at, startCell: "K1", data: [["old", "old", "old"], ["old", "old", "old"], ["keep", "keep", "keep"]] });
    // trailing blank lines (an editor's doing) must not clear the row below the data
    writeFileSync(csv, `a,"b,c"\n'12,=LEN(K1),extra\n\n\n`);
    await callTool("excel_import_csv", { ...at, path: csv, startCell: "K1" });
    const res = await callTool("excel_read_range", { ...at, range: "K1:M3" });
    expect(res.Data).toEqual([["a", "b,c", null], ["12", 1, "extra"], ["keep", "keep", "keep"]]);
  });

  test("a file that is not UTF-8 is rejected before anything is written", async () => {
    const at = { workbook: s.name, sheet: "Sheet1" };
    writeFileSync(csv, Buffer.from([0xc7, 0xd1, 0xb1, 0xdb])); // "한글" in CP949
    expect(await failure(callTool("excel_import_csv", { ...at, path: csv, startCell: "P1" }))).toMatch(/UTF-8/);
  });

  // Needs dynamic arrays (Excel 365 / 2021+).
  test("a spilled dynamic array is still a live spill after an edit and import", async () => {
    const at = { workbook: s.name, sheet: "Sheet1" };
    await xl(s, `
      $ws = $wb.Worksheets.Item('Sheet1')
      $ws.Range('R1').Value2 = 1; $ws.Range('R2').Value2 = 2; $ws.Range('R3').Value2 = 3
      $ws.Range('T1').Formula2 = '=R1:R3*2'
    `);
    await callTool("excel_export_csv", { ...at, range: "R1:T3", path: csv });
    const text = readFileSync(csv, "utf-8");
    expect(text).toBe("﻿1,,=R1:R3*2\r\n2,,\r\n3,,\r\n"); // spilled cells are derived: not exported as constants
    writeFileSync(csv, text.replace("\r\n2,,", "\r\n10,,"));
    await callTool("excel_import_csv", { ...at, path: csv, startCell: "R1" });
    const res = await callTool("excel_read_range", { ...at, range: "T1:T3" });
    expect(res.Data).toEqual([[2], [20], [6]]);
  });

  // Last: the array formula makes every later whole-sheet export fail.
  test("a range holding a legacy array formula is refused: writing it back cell by cell would break it", async () => {
    const at = { workbook: s.name, sheet: "Sheet1" };
    await xl(s, `$wb.Worksheets.Item('Sheet1').Range('V1:V3').FormulaArray = '=R1:R3*3'`);
    expect(await failure(callTool("excel_export_csv", { ...at, range: "U1:V3", path: csv }))).toMatch(/array formula/i);
  });
});
