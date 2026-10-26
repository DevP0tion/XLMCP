import { test, expect, beforeAll, afterAll } from "bun:test";
import { existsSync, unlinkSync, writeFileSync } from "fs";
import { tmpdir } from "os";
import { join, basename } from "path";
import { randomUUID } from "crypto";
import { dispose } from "../../src/services/powershell.js";
import { callTool } from "../helpers/tools.js";
import { describeExcel, openScratch, closeScratch, xl, type Scratch } from "../helpers/scratch.js";

// P0-2: DisplayAlerts belongs to the user's whole Excel instance. Tools must leave it as found and
// suppress prompts only around their own COM calls. The test sets it True first (the user's Excel
// already reads False, which would make an "unchanged" check pass trivially) and restores it at the end.
describeExcel("DisplayAlerts is left as the user set it (Excel)", () => {
  let s: Scratch;
  let original = "";
  const p2 = join(tmpdir(), `xlmcp_it_${randomUUID()}_b.xlsx`);
  const wb2 = basename(p2);
  const alerts = async () => (await xl(s, `[string]$app.DisplayAlerts`)).trim();
  const sheetNames = async () => (await xl(s, `@($wb.Worksheets | ForEach-Object { $_.Name }) -join ','`)).trim();

  beforeAll(async () => {
    s = await openScratch();
    original = (await xl(s, `[string]$app.DisplayAlerts`)).trim();
    await xl(s, `$app.DisplayAlerts = $true`);
  });
  afterAll(async () => {
    try {
      if (s) await xl(s, `$app.DisplayAlerts = $${original === "True" ? "true" : "false"}`);
    } finally {
      await closeScratch(s);
      if (existsSync(p2)) unlinkSync(p2);
      await dispose();
    }
  });

  test("fresh session binding does not switch it off", async () => {
    // Concurrent calls grow the pool; each new session binds to Excel on its first call.
    await Promise.all(
      Array.from({ length: 4 }, () => callTool("excel_read_range", { workbook: s.name, sheet: "Sheet1", range: "A1" }))
    );
    expect(await alerts()).toBe("True");
  });

  test("delete_sheet: no confirmation hang, setting restored", async () => {
    await xl(s, `$n = $wb.Worksheets.Add(); $n.Name = 'ToDelete'; $n.Range('A1').Value2 = 'x'`);
    await callTool("excel_delete_sheet", { workbook: s.name, name: "ToDelete" });
    expect(await sheetNames()).not.toContain("ToDelete");
    expect(await alerts()).toBe("True");
  });

  test("merge_cells over two values: no hang, setting restored", async () => {
    await callTool("excel_write_range", { workbook: s.name, sheet: "Sheet1", startCell: "A10", data: [["a", "b"]] });
    await callTool("excel_merge_cells", { workbook: s.name, sheet: "Sheet1", range: "A10:B10" });
    expect((await xl(s, `[string]$wb.Worksheets.Item('Sheet1').Range('A10').MergeCells`)).trim()).toBe("True");
    expect(await alerts()).toBe("True");
  });

  test("create_workbook SaveAs over an existing file: no hang, setting restored", async () => {
    writeFileSync(p2, "");
    const res = await callTool("excel_create_workbook", { savePath: p2 });
    expect(res.Name).toBe(wb2);
    expect(await alerts()).toBe("True");
  });

  test("save_workbook SaveAs over an existing file: no hang, setting restored", async () => {
    await callTool("excel_save_workbook", { workbook: wb2, savePath: p2 });
    expect(await alerts()).toBe("True");
  });

  test("close_workbook and open_workbook: setting restored", async () => {
    await callTool("excel_close_workbook", { workbook: wb2, save: false });
    expect(await alerts()).toBe("True");
    const res = await callTool("excel_open_workbook", { filePath: p2 });
    expect(res.Name).toBe(wb2);
    expect(await alerts()).toBe("True");
    await callTool("excel_close_workbook", { workbook: wb2, save: false });
  });

  test("copy_sheet: setting restored", async () => {
    await callTool("excel_copy_sheet", { workbook: s.name, source: "Sheet1", newName: "Copied" });
    expect(await sheetNames()).toContain("Copied");
    expect(await alerts()).toBe("True");
  });

  test("pivot over non-empty cells: no replace-prompt hang, setting restored", async () => {
    await callTool("excel_write_range", {
      workbook: s.name, sheet: "Sheet1", startCell: "D1",
      data: [["Cat", "Amt"], ["a", "1"], ["b", "2"], ["a", "3"]],
    });
    await callTool("excel_write_range", { workbook: s.name, sheet: "Sheet1", startCell: "H4", data: [["occupied"]] });
    await callTool("excel_create_pivot_table", {
      workbook: s.name, sheet: "Sheet1", sourceRange: "D1:E4", destCell: "G1", name: "PvtTest",
      rowFields: ["Cat"], dataFields: [{ field: "Amt" }],
    });
    expect(await alerts()).toBe("True");
    await callTool("excel_edit_pivot_layout", {
      workbook: s.name, sheet: "Sheet1", name: "PvtTest", addFields: [{ field: "Amt", orientation: "row" }],
    });
    expect(await alerts()).toBe("True");
  });
});
