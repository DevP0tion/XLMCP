import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { readFileSync } from "fs";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";
import { toCell, writeChunked, DEFAULT_CHUNK_SIZE, type CellInput } from "./write-range.tool.js";

// numbers outside toCell's plain-decimal rule (tiny, huge, > 15 digits) travel in exponent form
const EXPONENT = /^-?\d+(\.\d+)?e[+-]\d+$/i;

/** CSV field -> the value Value2 should receive. Inverse of toField (export-csv). */
export function fromField(s: string): CellInput {
  if (s === "") return null;
  if (s[0] === "'" || s[0] === "=") return s; // ': Excel drops it and stores the rest as text. =: formula
  if (/^(true|false)$/i.test(s)) return s.toLowerCase() === "true";
  if (EXPONENT.test(s)) return Number(s);
  return toCell(s);
}

/** RFC 4180 parse. Accepts CRLF, LF or CR record ends; rows may be ragged. */
export function parseCsv(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let field = "";
  let quoted = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (quoted) {
      if (ch !== '"') field += ch;
      else if (text[i + 1] === '"') { field += '"'; i++; }
      else quoted = false;
    } else if (ch === '"') {
      quoted = true;
    } else if (ch === ",") {
      row.push(field);
      field = "";
    } else if (ch === "\n" || ch === "\r") {
      if (ch === "\r" && text[i + 1] === "\n") i++;
      row.push(field);
      field = "";
      rows.push(row);
      row = [];
    } else {
      field += ch;
    }
  }
  // otherwise the rest of the file would land in one cell
  if (quoted) throw new Error("CSV has an unterminated quoted field");
  if (field !== "" || row.length > 0) rows.push([...row, field]);
  return rows;
}

export function register(server: McpServer) {
  server.registerTool(
    "excel_import_csv",
    {
      title: "Import CSV",
      description:
        "Write a UTF-8 CSV file (as saved by excel_export_csv) into a sheet from startCell. Only the CSV's rectangle is " +
        "overwritten: empty fields clear their cell, cells outside it are left alone. Fields: =... is a formula, '... is " +
        "literal text, plain decimals and TRUE/FALSE are typed, anything else is text.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        path: z.string().describe("Absolute path of the CSV file"),
        startCell: z.string().default("A1").describe("Top-left cell to write at. Default A1"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, path, startCell }) => {
      const bytes = readFileSync(path);
      let text: string;
      try {
        // fatal: an ANSI (CP949...) file is rejected instead of written as mojibake. Also drops the BOM.
        text = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
      } catch {
        throw new Error(`Not a UTF-8 file (save it as UTF-8): ${path}`);
      }
      const rows = parseCsv(text);
      // blank lines an editor left at the end are not rows to clear
      while (rows.length > 0 && rows[rows.length - 1].join("") === "") rows.pop();
      const cols = rows.reduce((m, r) => Math.max(m, r.length), 0);
      if (cols === 0) return textContent({ success: true, rows: 0, cols: 0 });
      // ragged rows are padded: the writer sizes every row by the first one
      const data = rows.map((r) => Array.from({ length: cols }, (_, j) => fromField(r[j] ?? "")));

      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      return textContent(
        await writeChunked(wbName, shName, startCell, data, data.length, cols, DEFAULT_CHUNK_SIZE, true, true)
      );
    }
  );
}
