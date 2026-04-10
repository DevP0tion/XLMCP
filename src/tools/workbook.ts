import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../services/powershell.js";
import { psEscape, textContent, errorContent } from "../services/utils.js";
import { workbookParam } from "../schemas/common.js";

export function registerWorkbookTools(server: McpServer) {
  // ── 열린 워크북 목록 ──
  server.registerTool(
    "excel_list_open_workbooks",
    {
      title: "열린 워크북 목록",
      description: "현재 Excel에 열려 있는 모든 워크북의 이름과 경로를 반환합니다.",
      inputSchema: {},
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async () => {
      const raw = await runPS(`
        $result = @()
        foreach ($wb in $excel.Workbooks) {
          $result += @{ Name = $wb.Name; Path = $wb.FullName; Sheets = $wb.Worksheets.Count } | ConvertTo-Json -Compress
        }
        "[" + ($result -join ",") + "]"
      `);
      return textContent(JSON.parse(raw.trim()));
    }
  );

  // ── 활성 워크북 정보 ──
  server.registerTool(
    "excel_get_active_workbook",
    {
      title: "활성 워크북 정보",
      description: "현재 활성화된 워크북의 이름, 경로, 시트 수, 활성 시트 이름을 반환합니다.",
      inputSchema: {},
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async () => {
      const raw = await runPS(`
        $wb = $excel.ActiveWorkbook
        if (-not $wb) { throw "열려 있는 워크북이 없습니다." }
        @{
          Name = $wb.Name
          Path = $wb.FullName
          SheetCount = $wb.Worksheets.Count
          ActiveSheet = $wb.ActiveSheet.Name
        } | ConvertTo-Json -Compress
      `);
      return textContent(JSON.parse(raw.trim()));
    }
  );

  // ── 새 워크북 생성 ──
  server.registerTool(
    "excel_create_workbook",
    {
      title: "새 워크북 생성",
      description: "새 빈 워크북을 생성합니다. savePath를 지정하면 즉시 저장합니다.",
      inputSchema: {
        savePath: z.string().optional().describe("저장할 절대 경로 (.xlsx). 생략 시 저장하지 않음"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ savePath }) => {
      const saveCmd = savePath
        ? `$wb.SaveAs('${psEscape(savePath)}')`
        : "";
      const raw = await runPS(`
        $wb = $excel.Workbooks.Add()
        ${saveCmd}
        @{ Name = $wb.Name; Path = $wb.FullName } | ConvertTo-Json -Compress
      `);
      return textContent(JSON.parse(raw.trim()));
    }
  );

  // ── 워크북 저장 ──
  server.registerTool(
    "excel_save_workbook",
    {
      title: "워크북 저장",
      description: "워크북을 저장합니다. savePath를 지정하면 다른 이름으로 저장합니다.",
      inputSchema: {
        workbook: workbookParam,
        savePath: z.string().optional().describe("다른 이름으로 저장할 절대 경로"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, savePath }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const saveCmd = savePath
        ? `$wb.SaveAs('${psEscape(savePath)}')`
        : `$wb.Save()`;
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        ${saveCmd}
      `);
      return textContent({ success: true });
    }
  );

  // ── 워크북 닫기 ──
  server.registerTool(
    "excel_close_workbook",
    {
      title: "워크북 닫기",
      description: "워크북을 닫습니다. save 옵션으로 저장 여부를 지정합니다.",
      inputSchema: {
        workbook: workbookParam,
        save: z.boolean().default(false).describe("닫기 전 저장 여부"),
      },
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ workbook, save }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $wb.Close(${save ? "$true" : "$false"})
      `);
      return textContent({ success: true });
    }
  );
}
