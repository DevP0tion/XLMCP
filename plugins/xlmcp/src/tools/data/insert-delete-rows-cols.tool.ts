import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_insert_delete_rows_cols",
    {
      title: "행/열 삽입·삭제",
      description: "행 또는 열을 삽입하거나 삭제합니다.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        target: z.enum(["row", "column"]).describe("대상: row 또는 column"),
        action: z.enum(["insert", "delete"]).describe("동작: insert 또는 delete"),
        index: z.union([z.number().int(), z.string()]).describe("행 번호(1부터) 또는 열 번호/문자 (1 또는 'A')"),
        count: z.number().int().default(1).describe("삽입/삭제할 개수"),
      },
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ workbook, sheet, target, action, index, count }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';

      let rangeExpr: string;
      if (target === "row") {
        const idx = Number(index);
        rangeExpr = `$ws.Rows("${idx}:${idx + count - 1}")`;
      } else {
        if (typeof index === "number") {
          rangeExpr = `$ws.Columns("${index}:${index + count - 1}")`;
        } else {
          // 문자 입력: "A" → Columns("A:A"), count > 1 시 열 오프셋 계산
          const startCode = index.toUpperCase().charCodeAt(0);
          const endLetter = String.fromCharCode(startCode + count - 1);
          rangeExpr = `$ws.Columns("${index.toUpperCase()}:${endLetter}")`;
        }
      }
      const cmd = action === "insert" ? "Insert()" : "Delete()";
      await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        ${rangeExpr}.${cmd}
      `, { exclusive: true });
      return textContent({ success: true });
    }
  );
}
