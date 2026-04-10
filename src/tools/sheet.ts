import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../services/powershell.js";
import { psEscape, textContent } from "../services/utils.js";
import { workbookParam } from "../schemas/common.js";

export function registerSheetTools(server: McpServer) {
  // ── 시트 목록 ──
  server.registerTool(
    "excel_list_sheets",
    {
      title: "시트 목록",
      description: "워크북의 모든 시트 이름을 반환합니다.",
      inputSchema: { workbook: workbookParam },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ workbook }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const raw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $names = @()
        foreach ($ws in $wb.Worksheets) { $names += $ws.Name }
        ConvertTo-Json @($names) -Compress
      `);
      return textContent({ sheets: JSON.parse(raw.trim()) });
    }
  );

  // ── 시트 추가 ──
  server.registerTool(
    "excel_create_sheet",
    {
      title: "시트 추가",
      description: "새 시트를 추가합니다.",
      inputSchema: {
        workbook: workbookParam,
        name: z.string().describe("새 시트 이름"),
        after: z.string().optional().describe("이 시트 뒤에 추가. 생략 시 맨 뒤"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, name, after }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const afterCmd = after
        ? `$ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item('${psEscape(after)}'))`
        : `$ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))`;
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        ${afterCmd}
        $ws.Name = '${psEscape(name)}'
      `);
      return textContent({ success: true, name });
    }
  );

  // ── 시트 삭제 ──
  server.registerTool(
    "excel_delete_sheet",
    {
      title: "시트 삭제",
      description: "시트를 삭제합니다.",
      inputSchema: {
        workbook: workbookParam,
        name: z.string().describe("삭제할 시트 이름"),
      },
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ workbook, name }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $wb.Worksheets.Item('${psEscape(name)}').Delete()
      `);
      return textContent({ success: true });
    }
  );

  // ── 시트 복사 ──
  server.registerTool(
    "excel_copy_sheet",
    {
      title: "시트 복사",
      description: "시트를 복사합니다. 같은 워크북 내에서 복사됩니다.",
      inputSchema: {
        workbook: workbookParam,
        source: z.string().describe("원본 시트 이름"),
        newName: z.string().optional().describe("복사본 시트 이름. 생략 시 자동 이름"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, source, newName }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const renameCmd = newName
        ? `$wb.ActiveSheet.Name = '${psEscape(newName)}'`
        : "";
      const raw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $src = $wb.Worksheets.Item('${psEscape(source)}')
        $src.Copy([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
        ${renameCmd}
        $wb.ActiveSheet.Name
      `);
      return textContent({ success: true, name: raw.trim() });
    }
  );

  // ── 시트 이름 변경 ──
  server.registerTool(
    "excel_rename_sheet",
    {
      title: "시트 이름 변경",
      description: "시트 이름을 변경합니다.",
      inputSchema: {
        workbook: workbookParam,
        oldName: z.string().describe("현재 시트 이름"),
        newName: z.string().describe("새 시트 이름"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, oldName, newName }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $wb.Worksheets.Item('${psEscape(oldName)}').Name = '${psEscape(newName)}'
      `);
      return textContent({ success: true });
    }
  );
}
