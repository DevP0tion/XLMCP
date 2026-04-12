import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_export_range_image",
    {
      title: "범위 이미지 내보내기",
      description: `셀 범위를 이미지 파일(PNG/JPG/BMP/GIF)로 내보냅니다.
시스템 클립보드를 사용하므로 실행 중 다른 작업은 일시 차단됩니다.`,
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        range: z.string().describe("캡처할 범위 (예: A1:D10)"),
        savePath: z.string().describe("저장 절대 경로 (예: F:\\output\\capture.png)"),
        format: z
          .enum(["png", "jpg", "bmp", "gif"])
          .default("png")
          .describe("이미지 형식"),
      },
      annotations: { readOnlyHint: true, destructiveHint: false },
    },
    async ({ workbook, sheet, range, savePath, format }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';

      const raw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $r = $ws.Range('${psEscape(range)}')

        # 범위를 클립보드에 복사 (xlScreen=1, xlPicture=2 → EMF, xlBitmap=2)
        $r.CopyPicture(1, 2)

        # 임시 ChartObject 생성 (범위 크기에 맞춤)
        $chartObj = $ws.ChartObjects.Add($r.Left, $r.Top, $r.Width, $r.Height)
        $chartObj.Chart.Paste()

        # 이미지 내보내기
        $chartObj.Chart.Export('${psEscape(savePath)}', '${format.toUpperCase()}')

        # 임시 ChartObject 삭제
        $chartObj.Delete()
        $excel.CutCopyMode = $false

        @{
          Path = '${psEscape(savePath)}'
          Format = '${format}'
          Width = [math]::Round($r.Width, 1)
          Height = [math]::Round($r.Height, 1)
        } | ConvertTo-Json -Compress
      `, { exclusive: true });
      return textContent(parseJSON(raw));
    }
  );
}
