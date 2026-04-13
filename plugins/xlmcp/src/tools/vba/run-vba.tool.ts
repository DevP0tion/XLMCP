import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam } from "../../schemas/common.js";

export function register(server: McpServer) {
  server.registerTool(
    "excel_run_vba",
    {
      title: "Run VBA Macro",
      description: "Run a VBA macro (Sub or Function). Returns the result if Function.",
      inputSchema: {
        workbook: workbookParam,
        macro: z.string().describe("Macro name (e.g. 'MyModule.Hello' or 'Hello')"),
        args: z
          .array(z.union([z.string(), z.number(), z.boolean()]))
          .optional()
          .describe("Arguments to pass to the macro (max 30)"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, macro, args }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';

      // Build Application.Run arguments
      const argList = (args ?? [])
        .map((a) => {
          if (typeof a === "string") return `'${psEscape(a)}'`;
          if (typeof a === "boolean") return a ? "$true" : "$false";
          return String(a);
        })
        .join(", ");

      const macroRef = workbook
        ? `'${psEscape(workbook)}!${psEscape(macro)}'`
        : `'${psEscape(macro)}'`;

      const argsPart = argList ? `, ${argList}` : "";

      const raw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $result = $excel.Run(${macroRef}${argsPart})
        @{
          Macro = '${psEscape(macro)}'
          Result = if ($null -eq $result) { $null } else { $result }
        } | ConvertTo-Json -Compress
      `);
      return textContent(parseJSON(raw));
    }
  );
}
