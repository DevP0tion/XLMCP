import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { tmpdir } from "os";
import { writeFileSync, unlinkSync } from "fs";
import { join } from "path";
import { randomUUID } from "crypto";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";
import { readRange } from "./read-range.tool.js";
import { colNumToLetter } from "../data/insert-delete-rows-cols.tool.js";

export const DEFAULT_CHUNK_SIZE = 500;

interface FormulaEntry {
  rowOffset: number;
  colOffset: number;
  formula: string;
}

export type CellInput = string | number | boolean | null;

const STRICT_NUMBER = /^-?(0|[1-9]\d*)(\.\d+)?$/;
// Excel parses strings written through Value2 like typed input ("0852"->852, "1/2"->date, "TRUE"->bool,
// "#N/A"->error, "'q"->"q"); such strings get a ' prefix so they are stored verbatim as text.
const EXCEL_WOULD_PARSE = /\d|^(true|false)$|^[#'=]/i;

/** Value written through Value2 for one non-formula input cell. */
export function toCell(v: CellInput): string | number | boolean | null {
  if (v === null || v === "") return null; // empty cell, never a zero-length string (Text-formatted columns keep "")
  if (typeof v !== "string") return v;
  // > 15 significant digits would be rounded by Excel, so such digit strings stay text
  if (STRICT_NUMBER.test(v) && v.replace(/[-.]/g, "").replace(/^0+/, "").length <= 15) return Number(v);
  return EXCEL_WOULD_PARSE.test(v) ? `'${v}` : v;
}

export function register(server: McpServer) {
  server.registerTool(
    "excel_write_range",
    {
      title: "Write Range",
      description:
        "Write 2D array from start cell. Cells: string | number | boolean | null. null and \"\" leave the cell empty. " +
        "Strings are stored verbatim as text unless they are plain decimals (\"12\", \"-3.5\"); \"=...\" is a formula " +
        "unless formulas=false. Large data is written in sequential chunks.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        startCell: z.string().describe("Start cell address (e.g. A1)"),
        data: z
          .array(z.array(z.union([z.string(), z.number(), z.boolean(), z.null()])))
          .describe("2D array data. Each inner array is one row"),
        formulas: z.boolean().default(true).describe("Treat strings starting with = as formulas. false writes them as text"),
        verify: z.boolean().default(false).describe("Read back and return { mismatches, first: up to 5 {cell, expected, actual} }"),
        chunkSize: z.number().int().optional().describe("Chunk size in rows. Default 500"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, startCell, data, formulas, verify, chunkSize: cs }) => {
      const chunkSize = cs ?? DEFAULT_CHUNK_SIZE;
      const rows = data.length;
      const cols = data[0]?.length ?? 0;
      if (rows === 0 || cols === 0) {
        return textContent({ success: true, rows: 0, cols: 0 });
      }

      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';

      // 모든 크기: 청크 분할 + JSON 임시 파일 (인라인 경로는 1행 데이터를 잘못 펼쳐 제거)
      const result = await writeChunked(wbName, shName, startCell, data, rows, cols, chunkSize, formulas);
      if (!verify) return textContent(result);
      const check = await verifyWrite(wbName, shName, startCell, data, rows, cols, chunkSize, formulas);
      return textContent({ ...result, verify: check });
    }
  );
}

/** Read the written range back and compare it with the normalized input. */
async function verifyWrite(
  wbName: string,
  shName: string,
  startCell: string,
  data: CellInput[][],
  rows: number,
  cols: number,
  chunkSize: number,
  formulasOn: boolean
) {
  const isFormula = (v: CellInput): v is string => formulasOn && typeof v === "string" && v.startsWith("=");
  const rangeExpr = `$ws.Range('${psEscape(startCell)}').Resize(${rows}, ${cols})`;
  const values = await readRange(wbName, shName, rangeExpr, chunkSize);
  const formulaText = data.some((r) => r.some(isFormula))
    ? await readRange(wbName, shName, rangeExpr, chunkSize, "Formula")
    : undefined;
  const mismatches: { cell: string; expected: unknown; actual: unknown }[] = [];
  data.forEach((row, i) =>
    row.forEach((v, j) => {
      let expected: unknown, actual: unknown, same: boolean;
      if (isFormula(v)) {
        // shortcut: Excel upper-cases function names and refs; case inside string literals is not compared
        expected = v;
        actual = formulaText!.Data[i]?.[j];
        same = String(actual).toUpperCase() === v.toUpperCase();
      } else {
        const c = toCell(v);
        expected = typeof c === "string" ? v : c; // text is stored verbatim; the ' prefix is not part of it
        actual = values.Data[i]?.[j];
        same = actual === expected;
      }
      if (!same) mismatches.push({ cell: `${colNumToLetter(values.StartCol + j)}${values.StartRow + i}`, expected, actual });
    })
  );
  const clip = (v: unknown) => (typeof v === "string" && v.length > 100 ? `${v.slice(0, 100)}… (${v.length} chars)` : v);
  const first = mismatches.slice(0, 5).map((m) => ({ ...m, expected: clip(m.expected), actual: clip(m.actual) }));
  return { mismatches: mismatches.length, first };
}

/**
 * Write one block of cells to a JSON temp file and return the PS fragment that loads it into $arr
 * (object[rows,cols]). Formulas are returned separately and left $null in $arr.
 * raw: cells are already what Value2 should receive (import_csv), so toCell is skipped.
 */
export function buildFill(
  data: CellInput[][],
  rowOffset = 0,
  formulasOn = true,
  raw = false
): { script: string; formulas: FormulaEntry[]; tmpFile: string } {
  const rows = data.length;
  const cols = data[0]?.length ?? 0;
  const formulas: FormulaEntry[] = [];
  const values = data.map((row, ri) =>
    row.map((v, ci) => {
      if (formulasOn && typeof v === "string" && v.startsWith("=")) {
        formulas.push({ rowOffset: rowOffset + ri, colOffset: ci, formula: v });
        return null;
      }
      return raw ? v : toCell(v);
    })
  );
  const tmpFile = join(tmpdir(), `xlmcp_chunk_${randomUUID()}.json`);
  writeFileSync(tmpFile, JSON.stringify(values));
  const script = `
        $json = Get-Content '${psEscape(tmpFile)}' -Raw -Encoding UTF8
        $data = $json | ConvertFrom-Json
        $arr = New-Object 'object[,]' ${rows},${cols}
        for ($i = 0; $i -lt ${rows}; $i++) {
          for ($j = 0; $j -lt ${cols}; $j++) {
            $v = $data[$i][$j]
            if ($v -ne $null) { $arr[$i,$j] = $v }
          }
        }`;
  return { script, formulas, tmpFile };
}

/** Group formula cells into vertical runs (same column, consecutive rows); each run is one Range.Formula assignment. */
export function formulaRuns(formulas: FormulaEntry[]): { r: number; c: number; f: string[] }[] {
  // shortcut: vertical runs only; a formula row spanning columns costs one assignment per cell. Merge into rectangles if that gets common.
  const sorted = [...formulas].sort((a, b) => a.colOffset - b.colOffset || a.rowOffset - b.rowOffset);
  const runs: { r: number; c: number; f: string[] }[] = [];
  for (const e of sorted) {
    const last = runs[runs.length - 1];
    if (last && last.c === e.colOffset && last.r + last.f.length === e.rowOffset) last.f.push(e.formula);
    else runs.push({ r: e.rowOffset, c: e.colOffset, f: [e.formula] });
  }
  return runs;
}

/**
 * Read the user's Calculation / ScreenUpdating (instance-wide settings) before a bulk write and return
 * the PS statement that restores them. Captured in TS so a fallback can restore after a dead session.
 */
export async function captureAppState(wbName: string): Promise<string> {
  const prev = parseJSON<{ Calc: number; SU: boolean }>(await runPS(`
    $wb = Resolve-Workbook ${wbName}
    @{ Calc = [int]$excel.Calculation; SU = [bool]$excel.ScreenUpdating } | ConvertTo-Json -Compress
  `));
  return `$excel.Calculation = ${Number(prev.Calc)}; $excel.ScreenUpdating = $${prev.SU ? "true" : "false"}`;
}

// ── 청크 분할 + 임시 파일 + 순차 ──
// asStored (import_csv): cells are Value2-ready and formulas are Formula2 text, so dynamic arrays spill again.
export async function writeChunked(
  wbName: string,
  shName: string,
  startCell: string,
  data: CellInput[][],
  rows: number,
  cols: number,
  chunkSize: number,
  formulasOn: boolean,
  asStored = false
) {
  // 청크 분할 (각 청크는 자체 JSON 임시 파일)
  const chunks: { rowOffset: number; chunkRows: number; fill: ReturnType<typeof buildFill> }[] = [];
  for (let offset = 0; offset < rows; offset += chunkSize) {
    const chunkData = data.slice(offset, offset + chunkSize);
    chunks.push({ rowOffset: offset, chunkRows: chunkData.length, fill: buildFill(chunkData, offset, formulasOn, asStored) });
  }
  const formulas = chunks.flatMap((c) => c.fill.formulas);
  // 수식은 연속 블록 단위 JSON → Range.Formula 대입 (상수는 .Formula로 보내지 않음: 입력처럼 해석됨)
  const formulaFile = formulas.length > 0 ? join(tmpdir(), `xlmcp_formulas_${randomUUID()}.json`) : "";
  if (formulaFile) writeFileSync(formulaFile, JSON.stringify(formulaRuns(formulas)));

  try {
    const chunkScripts = chunks
      .map((chunk, i) => `
        $chunkStart = $ws.Cells.Item($start.Row + ${chunk.rowOffset}, $start.Column)
        $chunkEnd = $ws.Cells.Item($start.Row + ${chunk.rowOffset} + ${chunk.chunkRows} - 1, $start.Column + ${cols} - 1)
        $targetRange = $ws.Range($chunkStart, $chunkEnd)
        ${chunk.fill.script}
        XlmcpMark 'json parse+fill c${i}'
        $targetRange.Value2 = $arr
        XlmcpMark 'value2 write c${i}'
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($targetRange)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($chunkEnd)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($chunkStart)
        $targetRange = $null; $chunkEnd = $null; $chunkStart = $null`)
      .join("\n");

    const formulaCmds = formulaFile
      ? `
          $runs = Get-Content '${psEscape(formulaFile)}' -Raw -Encoding UTF8 | ConvertFrom-Json
          # Excel without dynamic arrays (2019 and older) has no Formula2
          $useF2 = $false
          ${asStored ? "try { $useF2 = $null -ne $start.Formula2 } catch { }" : ""}
          foreach ($run in @($runs)) {
            $fs = @($run.f)
            $fa = New-Object 'object[,]' $fs.Count,1
            for ($k = 0; $k -lt $fs.Count; $k++) { $fa[$k,0] = $fs[$k] }
            $fr = $ws.Range($ws.Cells.Item($start.Row + $run.r, $start.Column + $run.c), $ws.Cells.Item($start.Row + $run.r + $fs.Count - 1, $start.Column + $run.c))
            if ($useF2) { $fr.Formula2 = $fa } else { $fr.Formula = $fa }
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($fr); $fr = $null
          }`
      : "";

    const restore = await captureAppState(wbName);
    try {
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $start = $ws.Range('${psEscape(startCell)}')
        try {
          $excel.ScreenUpdating = $false
          $excel.Calculation = -4135
          ${chunkScripts}
          ${formulaCmds}
          XlmcpMark 'formulas'
        } finally {
          ${restore}
        }
      `);
    } catch (err) {
      // PS finally never ran if the session died or timed out mid-script
      await runPS(`if ($excel) { ${restore} }`).catch(() => { /* best-effort */ });
      throw err;
    }
  } finally {
    for (const f of [...chunks.map((c) => c.fill.tmpFile), formulaFile].filter(Boolean)) {
      try {
        unlinkSync(f);
      } catch {
        // ignore
      }
    }
  }

  return {
    success: true,
    rows,
    cols,
    chunks: chunks.length,
    mode: "serial",
  };
}
