import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { tmpdir } from "os";
import { writeFileSync, unlinkSync } from "fs";
import { join } from "path";
import { randomUUID } from "crypto";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

const DEFAULT_CHUNK_SIZE = 30;

interface FormulaEntry {
  rowOffset: number;
  colOffset: number;
  formula: string;
}

export function register(server: McpServer) {
  server.registerTool(
    "excel_write_range",
    {
      title: "Write Range",
      description:
        "Write 2D array from start cell. Auto chunked parallel write for large data.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        startCell: z.string().describe("Start cell address (e.g. A1)"),
        data: z
          .array(z.array(z.string()))
          .describe("2D array data. Each inner array is one row"),
        chunkSize: z.number().int().optional().describe("Chunk size in rows for parallel write. Default 30"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, startCell, data, chunkSize: cs }) => {
      const chunkSize = cs ?? DEFAULT_CHUNK_SIZE;
      const rows = data.length;
      const cols = data[0]?.length ?? 0;
      if (rows === 0 || cols === 0) {
        return textContent({ success: true, rows: 0, cols: 0 });
      }

      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';

      // 모든 크기: 청크 분할 + JSON 임시 파일 (인라인 경로는 1행 데이터를 잘못 펼쳐 제거)
      return writeChunked(wbName, shName, startCell, data, rows, cols, chunkSize);
    }
  );
}

/**
 * Write one block of cells to a JSON temp file and return the PS fragment that loads it into $arr
 * (object[rows,cols]). Formulas are returned separately and left $null in $arr.
 */
export function buildFill(data: string[][], rowOffset = 0): { script: string; formulas: FormulaEntry[]; tmpFile: string } {
  const rows = data.length;
  const cols = data[0]?.length ?? 0;
  const formulas: FormulaEntry[] = [];
  const values = data.map((row, ri) =>
    row.map((v, ci) => {
      if (v.startsWith("=")) {
        formulas.push({ rowOffset: rowOffset + ri, colOffset: ci, formula: v });
        return null;
      }
      const num = Number(v);
      return v !== "" && !isNaN(num) ? num : v;
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

// ── 청크 분할 + 임시 파일 + 순차 ──
async function writeChunked(
  wbName: string,
  shName: string,
  startCell: string,
  data: string[][],
  rows: number,
  cols: number,
  chunkSize: number
) {
  // 청크 분할 (각 청크는 자체 JSON 임시 파일)
  const chunks: { rowOffset: number; chunkRows: number; fill: ReturnType<typeof buildFill> }[] = [];
  for (let offset = 0; offset < rows; offset += chunkSize) {
    const chunkData = data.slice(offset, offset + chunkSize);
    chunks.push({ rowOffset: offset, chunkRows: chunkData.length, fill: buildFill(chunkData, offset) });
  }
  const formulas = chunks.flatMap((c) => c.fill.formulas);

  try {
    const chunkScripts = chunks
      .map((chunk) => `
        $chunkStart = $ws.Cells.Item($start.Row + ${chunk.rowOffset}, $start.Column)
        $chunkEnd = $ws.Cells.Item($start.Row + ${chunk.rowOffset} + ${chunk.chunkRows} - 1, $start.Column + ${cols} - 1)
        $targetRange = $ws.Range($chunkStart, $chunkEnd)
        ${chunk.fill.script}
        $targetRange.Value2 = $arr
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($targetRange)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($chunkEnd)
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($chunkStart)
        $targetRange = $null; $chunkEnd = $null; $chunkStart = $null`)
      .join("\n");

    const formulaCmds =
      formulas.length > 0
        ? formulas
            .map(
              (f) =>
                `$ws.Cells.Item($start.Row + ${f.rowOffset}, $start.Column + ${f.colOffset}).Formula = '${psEscape(f.formula)}'`
            )
            .join("\n        ")
        : "";

    await runPS(`
      $wb = Resolve-Workbook ${wbName}
      $ws = Resolve-Sheet $wb ${shName}
      $start = $ws.Range('${psEscape(startCell)}')
      try {
        $excel.ScreenUpdating = $false
        $excel.Calculation = -4135
        ${chunkScripts}
        ${formulaCmds}
      } finally {
        $excel.Calculation = -4105
        $excel.ScreenUpdating = $true
      }
    `);
  } finally {
    for (const c of chunks) {
      try {
        unlinkSync(c.fill.tmpFile);
      } catch {
        // ignore
      }
    }
    await runPS(`if ($excel) { $excel.Calculation = -4105; $excel.ScreenUpdating = $true }`)
      .catch(() => { /* best-effort restoration if PS process died mid-script */ });
  }

  return textContent({
    success: true,
    rows,
    cols,
    chunks: chunks.length,
    mode: "serial",
  });
}
