import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { cancelTask, cancelAllTasks } from "../../services/powershell.js";
import { textContent } from "../../services/utils.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_cancel_queued",
    {
      title: "대기 작업 취소",
      description: `큐에 대기 중인 작업을 취소합니다.
taskId를 지정하면 해당 작업만 취소하고, 생략하면 모든 대기 작업을 취소합니다.
이미 실행 중인 작업은 취소할 수 없습니다.
대기 중인 작업 ID는 excel_pool_status로 확인할 수 있습니다.`,
      inputSchema: {
        taskId: z.number().int().optional().describe("취소할 작업 ID. 생략 시 전체 취소"),
      },
      annotations: { readOnlyHint: false, destructiveHint: true },
    },
    async ({ taskId }) => {
      if (taskId !== undefined) {
        const found = cancelTask(taskId);
        return textContent({
          success: found,
          message: found ? `작업 #${taskId} 취소됨` : `작업 #${taskId}을 찾을 수 없음 (이미 실행 중이거나 완료됨)`,
        });
      } else {
        const count = cancelAllTasks();
        return textContent({
          success: true,
          cancelled: count,
          message: count > 0 ? `대기 중인 ${count}개 작업 취소됨` : "취소할 대기 작업 없음",
        });
      }
    }
  );
}
