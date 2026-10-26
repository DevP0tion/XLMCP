import { test, expect } from "bun:test";
import { dispose } from "../../src/services/powershell.js";
import { callTool } from "../helpers/tools.js";
import { describeExcel, openScratch, closeScratch } from "../helpers/scratch.js";

function excelPids(): number[] {
  const p = Bun.spawnSync(["powershell.exe", "-NoProfile", "-Command", "@(Get-Process EXCEL -ErrorAction SilentlyContinue).Id -join ','"]);
  return p.stdout.toString().trim().split(",").filter(Boolean).map(Number);
}

describeExcel("pool dispose (Excel)", () => {
  // Owned (XLMCP-created) instances must actually exit on dispose. Tool scripts leave COM refs in
  // global scope; killing the PS process without releasing them kept Excel alive ~6 min.
  test("an XLMCP-created instance exits promptly after dispose", async () => {
    const before = new Set(excelPids());
    const s = await openScratch(); // openScratch refuses if Excel was already running, so a new PID is ours
    const ours = excelPids().filter((p) => !before.has(p));
    expect(ours.length).toBe(1);
    await callTool("excel_write_range", { workbook: s.name, sheet: "Sheet1", startCell: "A1", data: [["x"]] });
    await callTool("excel_read_range", { workbook: s.name, sheet: "Sheet1", range: "A1:A1" });
    await closeScratch(s);
    await dispose();
    const deadline = Date.now() + 15_000;
    while (Date.now() < deadline && excelPids().includes(ours[0])) await Bun.sleep(500);
    expect(excelPids()).not.toContain(ours[0]);
  });
});
