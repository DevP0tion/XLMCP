import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { register as createPivotTable } from "./create-pivot-table.tool.js";
import { register as editPivotLayout } from "./edit-pivot-layout.tool.js";

export function registerPivotTools(server: McpServer) {
  createPivotTable(server);
  editPivotLayout(server);
}
