import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_copy_paste_format",
    {
      title: "범위 서식 복사/붙여넣기",
      description: `범위의 서식(폰트, 색상, 테두리, 표시 형식 등)을 복사하여 대상 위치에 붙여넣습니다.
시스템 클립보드를 사용하므로 실행 중 다른 작업은 일시 차단됩니다.

⚠️ 이 도구는 서식만 복사합니다.
값이나 수식을 복사하려면 excel_copy_paste_range를 사용하세요.
값과 서식을 모두 복사하려면 두 도구를 순차 호출하세요.`,
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        sourceRange: z.string().describe("원본 범위 (예: A1:C10)"),
        destCell: z.string().describe("붙여넣기 시작 셀 (예: E1)"),
        destSheet: z.string().optional().describe("대상 시트. 생략 시 같은 시트"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, sourceRange, destCell, destSheet }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const dstShName = destSheet ? `'${psEscape(destSheet)}'` : shName;
      // xlPasteFormats = -4122
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $srcWs = Resolve-Sheet $wb ${shName}
        $dstWs = Resolve-Sheet $wb ${dstShName}
        $srcWs.Range('${psEscape(sourceRange)}').Copy()
        $dstWs.Range('${psEscape(destCell)}').PasteSpecial(-4122)
        $excel.CutCopyMode = $false
      `, { exclusive: true });
      return textContent({ success: true });
    }
  );
}
