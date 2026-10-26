import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";
import { readRange } from "../cell/read-range.tool.js";
import { colNumToLetter } from "./insert-delete-rows-cols.tool.js";

// XlLookAt. Every Replace argument is passed explicitly: omitted ones reuse the settings Excel saved
// from the last Find/Replace call or the user's Find dialog.
const LOOK_AT = { whole: 1, part: 2 } as const;
const MAX_CELLS = 200;

type MatchOptions = { lookAt: "whole" | "part"; matchCase: boolean };

/** Text a cell is matched on: raw Value2/Formula (numbers unformatted, booleans TRUE/FALSE). */
function cellText(v: unknown): string {
  if (v === null || v === undefined) return "";
  if (typeof v === "boolean") return v ? "TRUE" : "FALSE";
  return String(v);
}

/** [row, col] offsets of matching cells, row-major. Empty cells never match. */
export function matchCells(data: unknown[][], find: string, o: MatchOptions): [number, number][] {
  // shortcut: error cells arrive as CVErr ints and are not matched by their #N/A text
  const needle = o.matchCase ? find : find.toLowerCase();
  const hits: [number, number][] = [];
  data.forEach((row, i) =>
    row.forEach((v, j) => {
      const raw = cellText(v);
      if (raw === "") return;
      const text = o.matchCase ? raw : raw.toLowerCase();
      if (o.lookAt === "whole" ? text === needle : text.includes(needle)) hits.push([i, j]);
    })
  );
  return hits;
}

/** Compress hits into contiguous column runs (e.g. F3:F116) given the range's top-left cell. */
export function toRanges(hits: [number, number][], startRow: number, startCol: number): string[] {
  const byCol = new Map<number, number[]>();
  for (const [i, j] of hits) {
    if (!byCol.has(j)) byCol.set(j, []);
    byCol.get(j)!.push(i);
  }
  const out: string[] = [];
  for (const j of [...byCol.keys()].sort((a, b) => a - b)) {
    const col = colNumToLetter(startCol + j);
    const rows = byCol.get(j)!.sort((a, b) => a - b);
    for (let k = 0; k < rows.length; ) {
      let e = k;
      while (e + 1 < rows.length && rows[e + 1] === rows[e] + 1) e++;
      const first = `${col}${startRow + rows[k]}`;
      out.push(k === e ? first : `${first}:${col}${startRow + rows[e]}`);
      k = e + 1;
    }
  }
  return out;
}

export function register(server: McpServer) {
  server.registerTool(
    "excel_find_replace",
    {
      title: "Find/Replace",
      description:
        "Find or replace text; find-only if replace is omitted. Find matches raw cell values (unformatted: " +
        "1000 shown as 1,000 matches '1000'; dates are serial numbers) or formula text, and returns Count + " +
        "contiguous Ranges. Replace runs Excel Replace on formula text and returns the changed-cell count. " +
        "Never depends on Excel's Find dialog settings.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        find: z.string().min(1).describe("Text to find"),
        replace: z.string().optional().describe("Replacement text. Find-only if omitted"),
        matchCase: z.boolean().default(false).describe("Case-sensitive"),
        lookAt: z.enum(["whole", "part"]).default("whole").describe("whole: entire cell content equals find. part: substring match"),
        lookIn: z.enum(["values", "formulas"]).default("values").describe("Find-only: search raw values or formula text. Replace always works on formula text"),
        includeCells: z.boolean().default(false).describe(`Find-only: also return up to ${MAX_CELLS} {Cell, Value}`),
        range: z.string().optional().describe("Search range. Entire sheet if omitted"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, find, replace, matchCase, lookAt, lookIn, includeCells, range }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const rangeExpr = range ? `$ws.Range('${psEscape(range)}')` : `$ws.UsedRange`;
      const mc = matchCase ? "$true" : "$false";

      if (replace !== undefined) {
        // Replace(What, Replacement, LookAt, SearchOrder=xlByRows, MatchCase, MatchByte, SearchFormat, ReplaceFormat)
        // COM's Boolean result is True even when nothing changed (e.g. protected sheet), so count changed
        // cells by diffing the formula text (what Replace searches) before and after.
        const raw = await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          if ($ws.ProtectContents) { throw "Sheet '$($ws.Name)' is protected. Unprotect it before replacing." }
          $r = ${rangeExpr}
          $before = @($r.Formula)
          [void]$r.Replace('${psEscape(find)}', '${psEscape(replace)}', ${LOOK_AT[lookAt]}, 1, ${mc}, $false, $false, $false)
          $after = @($r.Formula)
          $n = 0
          for ($k = 0; $k -lt $before.Count; $k++) { if ($before[$k] -cne $after[$k]) { $n++ } }
          @{ Success = $true; Replaced = $n } | ConvertTo-Json -Compress
        `);
        return textContent(parseJSON(raw));
      } else {
        // One bulk read + TS matching (COM Find cost FindNext + Address() + Value2 per match)
        const r = await readRange(wbName, shName, rangeExpr, undefined, lookIn === "values" ? "Value2" : "Formula");
        const hits = matchCells(r.Data, find, { lookAt, matchCase });
        const result: Record<string, unknown> = { Count: hits.length, Ranges: toRanges(hits, r.StartRow, r.StartCol) };
        if (includeCells) {
          result.Cells = hits.slice(0, MAX_CELLS).map(([i, j]) => ({
            Cell: `${colNumToLetter(r.StartCol + j)}${r.StartRow + i}`,
            Value: r.Data[i][j],
          }));
          if (hits.length > MAX_CELLS) result.Truncated = true;
        }
        return textContent(result);
      }
    }
  );
}
