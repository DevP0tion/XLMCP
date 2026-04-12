import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_insert_image",
    {
      title: "이미지 삽입",
      description:
        "이미지 파일을 시트에 삽입합니다. 이미지는 엑셀 파일 내에 임베딩되므로 원본 파일을 삭제해도 유지됩니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        filePath: z.string().describe("이미지 절대 경로 (PNG, JPG, BMP, GIF)"),
        cell: z.string().describe("배치 기준 셀 (예: E1)"),
        width: z.number().optional().describe("너비 px. 생략 시 원본 크기"),
        height: z.number().optional().describe("높이 px. 생략 시 원본 크기"),
        name: z.string().optional().describe("Shape 이름. 생략 시 자동"),
        keepAspect: z
          .boolean()
          .default(true)
          .describe("width 또는 height 하나만 지정 시 비율 유지"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, filePath, cell, width, height, name, keepAspect }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';

      // 크기 결정 PS 스크립트
      let sizeScript: string;
      if (width && height) {
        // 둘 다 지정: 그대로 사용
        sizeScript = `$pic.Width = ${width}; $pic.Height = ${height}`;
      } else if (width) {
        sizeScript = keepAspect
          ? `$ratio = $pic.Height / $pic.Width; $pic.Width = ${width}; $pic.Height = ${width} * $ratio`
          : `$pic.Width = ${width}`;
      } else if (height) {
        sizeScript = keepAspect
          ? `$ratio = $pic.Width / $pic.Height; $pic.Height = ${height}; $pic.Width = ${height} * $ratio`
          : `$pic.Height = ${height}`;
      } else {
        sizeScript = ""; // 원본 크기 유지
      }

      const nameScript = name
        ? `$pic.Name = '${psEscape(name)}'`
        : "";

      const raw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $pos = $ws.Range('${psEscape(cell)}')
        $pic = $ws.Shapes.AddPicture(
          '${psEscape(filePath)}',
          0,
          -1,
          $pos.Left,
          $pos.Top,
          -1,
          -1
        )
        ${sizeScript}
        ${nameScript}
        @{
          Name = $pic.Name
          Width = [math]::Round($pic.Width, 1)
          Height = [math]::Round($pic.Height, 1)
          Cell = '${psEscape(cell)}'
        } | ConvertTo-Json -Compress
      `);
      return textContent(parseJSON(raw));
    }
  );
}
