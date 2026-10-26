import { z } from "zod";
import { registerWorkbookTools } from "../../src/tools/workbook/index.js";
import { registerSheetTools } from "../../src/tools/sheet/index.js";
import { registerCellTools } from "../../src/tools/cell/index.js";
import { registerFormatTools } from "../../src/tools/format/index.js";
import { registerDataTools } from "../../src/tools/data/index.js";
import { registerTableTools } from "../../src/tools/table/index.js";
import { registerChartTools } from "../../src/tools/chart/index.js";
import { registerPivotTools } from "../../src/tools/pivot/index.js";
import { registerValidationTools } from "../../src/tools/validation/index.js";
import { registerViewTools } from "../../src/tools/view/index.js";
import { registerImageTools } from "../../src/tools/image/index.js";
import { registerVbaTools } from "../../src/tools/vba/index.js";

type Handler = (args: any) => Promise<{ content: { type: string; text: string }[] }>;

// Captures registerTool() calls so tests can invoke real tool handlers without an MCP server.
const tools = new Map<string, { schema: z.ZodTypeAny; handler: Handler }>();
const fakeServer = {
  registerTool(name: string, def: { inputSchema?: z.ZodRawShape }, handler: Handler) {
    tools.set(name, { schema: z.object(def.inputSchema ?? {}), handler });
  },
};
for (const reg of [
  registerWorkbookTools, registerSheetTools, registerCellTools, registerFormatTools,
  registerDataTools, registerTableTools, registerChartTools, registerPivotTools,
  registerValidationTools, registerViewTools, registerImageTools, registerVbaTools,
]) {
  reg(fakeServer as any);
}

/** Raw MCP text of a tool call. Args go through the tool's zod schema so defaults apply, as in the real server. */
export async function callToolText(name: string, args: Record<string, unknown> = {}): Promise<string> {
  const t = tools.get(name);
  if (!t) throw new Error(`Unknown tool: ${name}`);
  const res = await t.handler(t.schema.parse(args));
  return res.content[0].text;
}

export async function callTool<T = any>(name: string, args: Record<string, unknown> = {}): Promise<T> {
  return JSON.parse(await callToolText(name, args));
}
