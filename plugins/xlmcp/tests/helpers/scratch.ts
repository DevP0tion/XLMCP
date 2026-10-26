import { describe } from "bun:test";
import { existsSync, unlinkSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";
import { randomUUID } from "crypto";
import { runPS } from "../../src/services/powershell.js";
import { callTool } from "./tools.js";

/** Excel integration tests run only with XLMCP_IT=1, so a plain `bun test` never touches Excel. */
export const describeExcel = process.env.XLMCP_IT ? describe : describe.skip;

export interface Scratch {
  name: string;
  path: string;
}

function excelRunning(): boolean {
  const p = Bun.spawnSync(["powershell.exe", "-NoProfile", "-Command", "@(Get-Process EXCEL -ErrorAction SilentlyContinue).Count"]);
  return parseInt(p.stdout.toString().trim() || "0", 10) > 0;
}

/**
 * Create a scratch workbook saved under %TEMP% with a unique name. Tests must address it by
 * this name only: an omitted workbook falls back to ActiveWorkbook, which may be the user's file.
 */
let checkedForUserExcel = false;

export async function openScratch(): Promise<Scratch> {
  // The scratch book (and any app-level setting a test touches) lives in whatever Excel instance
  // the resolver binds. Refuse to share the user's instance unless explicitly allowed.
  // Checked once per run: later Excel processes are the tests' own instance.
  if (!checkedForUserExcel) {
    if (excelRunning() && !process.env.XLMCP_IT_ALLOW_USER_EXCEL) {
      throw new Error("Excel is already running (possibly with user workbooks). Set XLMCP_IT_ALLOW_USER_EXCEL=1 to run anyway.");
    }
    checkedForUserExcel = true;
  }
  const path = join(tmpdir(), `xlmcp_it_${randomUUID()}.xlsx`);
  const res = await callTool<{ Name: string; Path: string }>("excel_create_workbook", { savePath: path });
  if (res.Path.toLowerCase() !== path.toLowerCase()) {
    throw new Error(`Scratch workbook saved to unexpected path: ${res.Path}`);
  }
  return { name: res.Name, path };
}

export async function closeScratch(s: Scratch | undefined): Promise<void> {
  if (!s) return;
  try {
    await callTool("excel_close_workbook", { workbook: s.name, save: false });
  } finally {
    if (existsSync(s.path)) unlinkSync(s.path);
  }
}

/** Run PS against the scratch workbook; `$wb` and `$app` (its own Application) are bound. */
export function xl(s: Scratch, script: string): Promise<string> {
  return runPS(`
    $wb = Resolve-Workbook '${s.name}'
    $app = $wb.Application
    ${script}
  `);
}
