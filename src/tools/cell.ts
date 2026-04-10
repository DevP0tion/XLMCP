import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../services/powershell.js";
import { psEscape, textContent } from "../services/utils.js";
import { workbookParam, sheetParam } from "../schemas/common.js";

export function registerCellTools(server: McpServer) {
  // ── 단일 셀 읽기 ──
  server.registerTool(
    "excel_read_cell",
    {
      title: "셀 읽기",
      description: "단일 셀의 값, 수식, 표시 텍스트를 반환합니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        cell: z.string().describe("셀 주소 (예: A1, B3)"),
      },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ workbook, sheet, cell }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const raw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $c = $ws.Range('${psEscape(cell)}')
        @{
          Value = if ($c.Value2 -ne $null) { $c.Value2.ToString() } else { $null }
          Formula = $c.Formula
          Text = $c.Text
          NumberFormat = $c.NumberFormat
        } | ConvertTo-Json -Compress
      `);
      return textContent(JSON.parse(raw.trim()));
    }
  );

  // ── 단일 셀 쓰기 ──
  server.registerTool(
    "excel_write_cell",
    {
      title: "셀 쓰기",
      description: "단일 셀에 값 또는 수식을 입력합니다. '='로 시작하면 수식으로 처리됩니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        cell: z.string().describe("셀 주소 (예: A1)"),
        value: z.string().describe("입력할 값 또는 수식 (예: '=SUM(A1:A10)')"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, cell, value }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const isFormula = value.startsWith("=");
      const cmd = isFormula
        ? `$c.Formula = '${psEscape(value)}'`
        : `$c.Value2 = '${psEscape(value)}'`;
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $c = $ws.Range('${psEscape(cell)}')
        ${cmd}
      `);
      return textContent({ success: true });
    }
  );

  // ── 범위 읽기 ──
  server.registerTool(
    "excel_read_range",
    {
      title: "범위 읽기",
      description: "셀 범위의 값을 2D 배열로 반환합니다. 범위를 생략하면 UsedRange 전체를 읽습니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        range: z.string().optional().describe("범위 주소 (예: A1:C10). 생략 시 UsedRange"),
      },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ workbook, sheet, range }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const rangeExpr = range
        ? `$ws.Range('${psEscape(range)}')`
        : `$ws.UsedRange`;
      const raw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $r = ${rangeExpr}
        $rows = $r.Rows.Count
        $cols = $r.Columns.Count
        $data = @()
        for ($i = 1; $i -le $rows; $i++) {
          $row = @()
          for ($j = 1; $j -le $cols; $j++) {
            $v = $r.Cells.Item($i, $j).Value2
            $row += if ($v -ne $null) { $v.ToString() } else { "" }
          }
          $data += ,@($row)
        }
        @{
          Range = $r.Address()
          Rows = $rows
          Cols = $cols
          Data = $data
        } | ConvertTo-Json -Depth 10 -Compress
      `);
      return textContent(JSON.parse(raw.trim()));
    }
  );

  // ── 범위 쓰기 ──
  server.registerTool(
    "excel_write_range",
    {
      title: "범위 쓰기",
      description: "시작 셀부터 2D 배열 데이터를 입력합니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        startCell: z.string().describe("시작 셀 주소 (예: A1)"),
        data: z
          .array(z.array(z.string()))
          .describe("2D 배열 데이터. 각 내부 배열이 한 행"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, startCell, data }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const rows = data.length;
      const cols = data[0]?.length ?? 0;
      // PS 배열 리터럴 생성
      const psRows = data
        .map((row) => {
          const cells = row.map((v) => `'${psEscape(v)}'`).join(",");
          return `@(${cells})`;
        })
        .join(",");

      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $start = $ws.Range('${psEscape(startCell)}')
        $endRow = $start.Row + ${rows} - 1
        $endCol = $start.Column + ${cols} - 1
        $endCell = $ws.Cells.Item($endRow, $endCol)
        $targetRange = $ws.Range($start, $endCell)
        $arr = New-Object 'object[,]' ${rows},${cols}
        $srcData = @(${psRows})
        for ($i = 0; $i -lt ${rows}; $i++) {
          $row = @($srcData[$i])
          for ($j = 0; $j -lt ${cols}; $j++) {
            $val = $row[$j]
            if ($val -match '^\=') {
              $ws.Cells.Item($start.Row + $i, $start.Column + $j).Formula = $val
            } else {
              $arr[$i,$j] = $val
            }
          }
        }
        $targetRange.Value2 = $arr
      `);
      return textContent({ success: true, rows, cols });
    }
  );
}
