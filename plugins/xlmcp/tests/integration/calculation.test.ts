import { test, expect, beforeAll, afterAll } from "bun:test";
import { dispose } from "../../src/services/powershell.js";
import { callTool } from "../helpers/tools.js";
import { describeExcel, openScratch, closeScratch, xl, type Scratch } from "../helpers/scratch.js";

const XL_MANUAL = "-4135";

// P0-4: bulk writes must restore the user's calculation mode / ScreenUpdating, not force Automatic.
describeExcel("calculation mode is restored after bulk writes (Excel)", () => {
  let s: Scratch;
  let original = { calc: "", su: "" };
  const appState = async () => {
    const [calc, su] = (await xl(s, `"$([int]$app.Calculation)|$($app.ScreenUpdating)"`)).trim().split("|");
    return { calc, su };
  };
  const rows40 = Array.from({ length: 40 }, (_, i) => [`r${i}`, String(i)]);

  beforeAll(async () => {
    s = await openScratch();
    original = await appState();
    await xl(s, `$app.Calculation = ${XL_MANUAL}`);
  });
  afterAll(async () => {
    try {
      if (s) await xl(s, `$app.Calculation = ${original.calc}; $app.ScreenUpdating = $${original.su === "True" ? "true" : "false"}`);
    } finally {
      await closeScratch(s);
      await dispose();
    }
  });

  test("write_range of 40 rows keeps manual calculation", async () => {
    await callTool("excel_write_range", { workbook: s.name, sheet: "Sheet1", startCell: "A1", data: rows40 });
    expect(await appState()).toEqual({ calc: XL_MANUAL, su: original.su });
  });

  test("copy_paste_range of 40 rows keeps manual calculation", async () => {
    await callTool("excel_copy_paste_range", { workbook: s.name, sheet: "Sheet1", sourceRange: "A1:B40", destCell: "D1" });
    expect(await appState()).toEqual({ calc: XL_MANUAL, su: original.su });
  });
});
