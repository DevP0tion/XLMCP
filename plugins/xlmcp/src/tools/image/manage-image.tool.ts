import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_manage_image",
    {
      title: "이미지 관리",
      description: "삽입된 이미지를 삭제, 이동, 크기 변경합니다. 이미지 이름은 excel_list_images로 확인할 수 있습니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        name: z.string().describe("이미지(Shape) 이름"),
        action: z.enum(["delete", "move", "resize"]).describe("동작: delete(삭제), move(이동), resize(크기 변경)"),
        cell: z.string().optional().describe("move 시 대상 셀 (예: A10)"),
        width: z.number().optional().describe("resize 시 너비 px"),
        height: z.number().optional().describe("resize 시 높이 px"),
        keepAspect: z
          .boolean()
          .default(true)
          .describe("resize 시 width 또는 height 하나만 지정하면 비율 유지"),
      },
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ workbook, sheet, name, action, cell, width, height, keepAspect }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';

      if (action === "delete") {
        await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $ws.Shapes.Item('${psEscape(name)}').Delete()
        `);
        return textContent({ success: true, action: "deleted", name });
      }

      if (action === "move") {
        if (!cell) throw new Error("move 시 cell 파라미터가 필요합니다.");
        const raw = await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $s = $ws.Shapes.Item('${psEscape(name)}')
          $pos = $ws.Range('${psEscape(cell)}')
          $s.Left = $pos.Left
          $s.Top = $pos.Top
          @{
            Name = $s.Name
            Left = [math]::Round($s.Left, 1)
            Top = [math]::Round($s.Top, 1)
          } | ConvertTo-Json -Compress
        `);
        return textContent(parseJSON(raw));
      }

      // resize
      let sizeScript: string;
      if (width && height) {
        sizeScript = `
          $s.LockAspectRatio = 0
          $s.Width = ${width}
          $s.Height = ${height}`;
      } else if (width) {
        sizeScript = keepAspect
          ? `$ratio = $s.Height / $s.Width; $s.Width = ${width}; $s.Height = ${width} * $ratio`
          : `$s.LockAspectRatio = 0; $s.Width = ${width}`;
      } else if (height) {
        sizeScript = keepAspect
          ? `$ratio = $s.Width / $s.Height; $s.Height = ${height}; $s.Width = ${height} * $ratio`
          : `$s.LockAspectRatio = 0; $s.Height = ${height}`;
      } else {
        throw new Error("resize 시 width 또는 height가 필요합니다.");
      }

      const raw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $s = $ws.Shapes.Item('${psEscape(name)}')
        ${sizeScript}
        @{
          Name = $s.Name
          Width = [math]::Round($s.Width, 1)
          Height = [math]::Round($s.Height, 1)
        } | ConvertTo-Json -Compress
      `);
      return textContent(parseJSON(raw));
    }
  );
}
