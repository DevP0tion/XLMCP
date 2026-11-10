import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { writeFileSync } from "fs";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";
import { readRange } from "./read-range.tool.js";
import { fromField } from "./import-csv.tool.js";

/**
 * Cell -> CSV field text, chosen so that fromField gives the same cell back. Formula cells export
 * their formula; text that fromField would read as another type gets Excel's ' text prefix.
 */
export function toField(value: unknown, formula: unknown): string {
  // a text cell holding "=1+1" reports it as .Formula too; only a real formula differs from its value
  if (typeof formula === "string" && formula.startsWith("=") && formula !== value) return formula;
  // no formula text: an empty cell, or one a dynamic array spills into (a constant there would block the spill)
  // shortcut: a zero-length string constant also has no formula text and comes back as a blank cell
  if (formula === "" || value === null || value === undefined) return "";
  if (typeof value === "boolean") return value ? "TRUE" : "FALSE";
  const s = String(value);
  if (typeof value === "number") return typeof fromField(s) === "number" ? s : value.toExponential();
  return typeof fromField(s) === "string" && !/^['=]/.test(s) ? s : `'${s}`;
}

const quote = (f: string) => (/[",\r\n]/.test(f) ? `"${f.replace(/"/g, '""')}"` : f);

/** RFC 4180 text, CRLF after every record. */
export function toCsv(rows: string[][]): string {
  return rows.map((r) => r.map(quote).join(",") + "\r\n").join("");
}

export function register(server: McpServer) {
  server.registerTool(
    "excel_export_csv",
    {
      title: "Export CSV",
      description:
        "Save a range as a UTF-8 CSV file to edit on disk; excel_import_csv writes it back. Without range: A1 through " +
        "the last used cell, so CSV line n is sheet row n. Formula cells are saved as formulas (=...), other cells as " +
        "raw values (dates are serial numbers); text that would read back as a number, boolean or formula gets a ' prefix. " +
        "Cells a dynamic array spills into are saved empty (its formula refills them). Legacy array formulas are refused.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        range: z.string().optional().describe("Range address (e.g. A1:C10). A1 to the last used cell if omitted"),
        path: z.string().describe("Absolute path of the CSV file to write (overwritten if it exists)"),
      },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ workbook, sheet, range, path }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const rangeExpr = range
        ? `$ws.Range('${psEscape(range)}')`
        : "$ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($ws.UsedRange.Row + $ws.UsedRange.Rows.Count - 1, $ws.UsedRange.Column + $ws.UsedRange.Columns.Count - 1))";
      const meta = parseJSON<{ HasArray: string; Formula2: boolean }>(await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $r = ${rangeExpr}
        $f2 = $false
        try { $f2 = $null -ne $r.Cells.Item(1, 1).Formula2 } catch { }
        @{ HasArray = "$($r.HasArray)"; Formula2 = $f2 } | ConvertTo-Json -Compress
      `, { retry: true }));
      // HasArray is True (all cells), False (none) or null (some)
      if (meta.HasArray !== "False") {
        throw new Error(
          "Range contains a legacy (Ctrl+Shift+Enter) array formula, which a CSV round trip would break. Export a range without it."
        );
      }
      const values = await readRange(wbName, shName, rangeExpr);
      // Formula2 is the text Excel 365 shows and keeps dynamic arrays intact; older Excel only has Formula
      const formulas = await readRange(wbName, shName, rangeExpr, undefined, meta.Formula2 ? "Formula2" : "Formula");
      // misaligned formulas would export values in their place, and the import would then overwrite the formulas
      if (formulas.Range !== values.Range) throw new Error("The sheet changed during the export. Run it again.");
      const rows = values.Data.map((row, i) => row.map((v, j) => toField(v, formulas.Data[i][j])));
      // BOM: without it Excel opens UTF-8 as the ANSI code page
      writeFileSync(path, "﻿" + toCsv(rows));
      return textContent({ Path: path, Range: values.Range, Rows: values.Rows, Cols: values.Cols });
    }
  );
}
