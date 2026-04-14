import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

const orientationSchema = z.enum(["row", "column", "data", "page", "hidden"]);
const functionSchema = z.enum(["sum", "count", "average", "max", "min"]);

export function register(server: McpServer) {
  server.registerTool(
    "excel_edit_pivot_layout",
    {
      title: "Edit Pivot Table Layout",
      description: "Add, remove, or move fields in an existing pivot table. Change aggregate functions.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        name: z.string().describe("Pivot table name"),
        addFields: z
          .array(
            z.object({
              field: z.string().describe("Field name"),
              orientation: orientationSchema.describe("Area: row, column, data, page, or hidden"),
              function: functionSchema.optional().describe("Aggregate function (data fields only)"),
            })
          )
          .optional()
          .describe("Fields to add or move to a new area"),
        removeFields: z
          .array(z.string())
          .optional()
          .describe("Field names to remove (set to hidden)"),
        refreshData: z.boolean().default(false).describe("Refresh pivot cache after changes"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, name, addFields, removeFields, refreshData }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';

      // xlRowField=1, xlColumnField=2, xlPageField=3, xlDataField=4, xlHidden=0
      const oriMap: Record<string, number> = { row: 1, column: 2, page: 3, data: 4, hidden: 0 };
      const fnMap: Record<string, number> = { sum: -4157, count: -4112, average: -4106, max: -4136, min: -4139 };

      const cmds: string[] = [];

      if (removeFields) {
        for (const f of removeFields) {
          cmds.push(`$pf = $pvt.PivotFields('${psEscape(f)}'); $pf.Orientation = 0`);
        }
      }

      if (addFields) {
        for (const af of addFields) {
          const ori = oriMap[af.orientation];
          if (af.orientation === "data" && af.function) {
            cmds.push(`$pf = $pvt.PivotFields('${psEscape(af.field)}'); $pf.Orientation = ${ori}; $pf.Function = ${fnMap[af.function]}`);
          } else {
            cmds.push(`$pf = $pvt.PivotFields('${psEscape(af.field)}'); $pf.Orientation = ${ori}`);
          }
        }
      }

      const refreshCmd = refreshData ? "$pvt.PivotCache().Refresh()" : "";

      const raw = await runPS(`
        $wb = Resolve-Workbook ${wbName}
        $ws = Resolve-Sheet $wb ${shName}
        $pvt = $ws.PivotTables('${psEscape(name)}')
        ${cmds.join("\n        ")}
        ${refreshCmd}
        function Get-FieldNames($fields) {
          $names = @()
          try { for ($i = 1; $i -le $fields.Count; $i++) { $names += $fields.Item($i).Name } } catch {}
          return $names
        }
        @{
          Name = $pvt.Name
          Location = $pvt.TableRange1.Address()
          RowFields = Get-FieldNames $pvt.RowFields
          ColumnFields = Get-FieldNames $pvt.ColumnFields
          DataFields = Get-FieldNames $pvt.DataFields
          PageFields = Get-FieldNames $pvt.PageFields
        } | ConvertTo-Json -Compress
      `);
      return textContent(parseJSON(raw));
    }
  );
}
