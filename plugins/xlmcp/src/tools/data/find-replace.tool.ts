import type { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { z } from "zod";
import { runPS } from "../../services/powershell.js";
import { psEscape, textContent, parseJSON } from "../../services/utils.js";
import { workbookParam, sheetParam } from "../../schemas/common.js";

// XlLookAt / XlFindLookIn. Every Find/Replace argument is passed explicitly: omitted ones reuse the
// settings Excel saved from the last Find/Replace call or the user's Find dialog.
const LOOK_AT = { whole: 1, part: 2 } as const;
const LOOK_IN = { values: -4163, formulas: -4123 } as const;

export function register(server: McpServer) {
  server.registerTool(
    "excel_find_replace",
    {
      title: "Find/Replace",
      description: "Find or replace text. Find-only if replace is omitted. Matching never depends on Excel's Find dialog settings.",
      inputSchema: {
        workbook: workbookParam,
        sheet: sheetParam,
        find: z.string().describe("Text to find"),
        replace: z.string().optional().describe("Replacement text. Find-only if omitted"),
        matchCase: z.boolean().default(false).describe("Case-sensitive"),
        lookAt: z.enum(["whole", "part"]).default("whole").describe("whole: entire cell content equals find. part: substring match"),
        lookIn: z.enum(["values", "formulas"]).default("values").describe("Find-only: search values or formula text. Replace always works on formula text"),
        range: z.string().optional().describe("Search range. Entire sheet if omitted"),
      },
      annotations: { readOnlyHint: false, destructiveHint: false },
    },
    async ({ workbook, sheet, find, replace, matchCase, lookAt, lookIn, range }) => {
      const wbName = workbook ? `'${psEscape(workbook)}'` : '""';
      const shName = sheet ? `'${psEscape(sheet)}'` : '""';
      const rangeExpr = range ? `$ws.Range('${psEscape(range)}')` : `$ws.UsedRange`;
      const mc = matchCase ? "$true" : "$false";

      if (replace !== undefined) {
        // Replace(What, Replacement, LookAt, SearchOrder=xlByRows, MatchCase, MatchByte, SearchFormat, ReplaceFormat)
        const raw = await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $r = ${rangeExpr}
          $replaced = $r.Replace('${psEscape(find)}', '${psEscape(replace)}', ${LOOK_AT[lookAt]}, 1, ${mc}, $false, $false, $false)
          @{ Success = $replaced } | ConvertTo-Json -Compress
        `);
        return textContent(parseJSON(raw));
      } else {
        // Find(What, After, LookIn, LookAt, SearchOrder=xlByRows, SearchDirection=xlNext, MatchCase, MatchByte, SearchFormat)
        const raw = await runPS(`
          $wb = Resolve-Workbook ${wbName}
          $ws = Resolve-Sheet $wb ${shName}
          $r = ${rangeExpr}
          $results = @()
          $first = $r.Find('${psEscape(find)}', [Type]::Missing, ${LOOK_IN[lookIn]}, ${LOOK_AT[lookAt]}, 1, 1, ${mc}, $false, $false)
          if ($first) {
            $current = $first
            do {
              $results += @{ Cell = $current.Address(); Value = $current.Value2 }
              $current = $r.FindNext($current)
            } while ($current -and $current.Address() -ne $first.Address())
          }
          @{ Count = $results.Count; Matches = $results } | ConvertTo-Json -Depth 5 -Compress
        `);
        return textContent(parseJSON(raw));
      }
    }
  );
}
