/**
 * P1-3 benchmark: where does the time go in a large range write + read?
 * Creates its own scratch workbook (refuses if Excel is already running), enables XLMCP_DEBUG timing,
 * and prints per-step totals from %TEMP%\xlmcp.log.
 *
 * Run:  bun plugins/xlmcp/scripts/bench-range.ts [rows=2500] [cols=29] [chunkSize=30]
 */
process.env.XLMCP_DEBUG = "1";
process.env.XLMCP_TIMEOUT_MS ??= "300000"; // measure the whole operation, not the 30 s default
process.env.XLMCP_IT_ALLOW_USER_EXCEL = ""; // never share the user's Excel

const { existsSync, statSync, readFileSync } = await import("fs");
const { tmpdir } = await import("os");
const { join } = await import("path");
const { dispose } = await import("../src/services/powershell.js");
const { callTool, callToolText } = await import("../tests/helpers/tools.js");
const { openScratch, closeScratch } = await import("../tests/helpers/scratch.js");

const [rows, cols, chunkSize] = [2500, 29, 30].map((d, i) => Number(process.argv[2 + i]) || d);
const LOG = join(tmpdir(), "xlmcp.log");
const logSize = () => (existsSync(LOG) ? statSync(LOG).size : 0);

function colName(n: number): string {
  let s = "";
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

/** Sum log lines written since `from`, grouped by step label (chunk index stripped). */
function summarize(from: number): string {
  const text = readFileSync(LOG).subarray(from).toString("utf-8");
  const steps = new Map<string, { ms: number; n: number }>();
  for (const line of text.split(/\r?\n/)) {
    let m = line.match(/ s\d+ ps (.+?)(?: c\d+)? (\d+)ms$/);
    let label: string, ms: number;
    if (m) { label = `ps  ${m[1]}`; ms = +m[2]; }
    else if ((m = line.match(/ s\d+ runPS wait=(\d+)ms invoke=(\d+)ms/))) {
      add("pool queue wait", +m[1]); label = "pool invoke (total)"; ms = +m[2];
    } else if ((m = line.match(/ ts (.+?) (\d+)ms$/))) { label = `ts  ${m[1]}`; ms = +m[2]; }
    else continue;
    add(label, ms);
  }
  function add(label: string, ms: number) {
    const s = steps.get(label) ?? { ms: 0, n: 0 };
    s.ms += ms; s.n++;
    steps.set(label, s);
  }
  return [...steps].map(([k, v]) => `  ${k.padEnd(34)} ${String(v.ms).padStart(7)} ms  (x${v.n})`).join("\n");
}

const data = Array.from({ length: rows }, (_, r) =>
  Array.from({ length: cols }, (_, c) => (c % 3 === 0 ? `item-${r}-${c}` : String(r * cols + c + 0.25)))
);
const range = `A1:${colName(cols)}${rows}`;
const s = await openScratch();
try {
  let from = logSize();
  let t = Date.now();
  await callTool("excel_write_range", { workbook: s.name, sheet: "Sheet1", startCell: "A1", data, chunkSize });
  console.log(`write_range ${range} (chunk ${chunkSize}): ${Date.now() - t} ms wall\n${summarize(from)}`);

  from = logSize();
  t = Date.now();
  const text = await callToolText("excel_read_range", { workbook: s.name, sheet: "Sheet1", range, chunkSize });
  console.log(`read_range ${range} (chunk ${chunkSize}): ${Date.now() - t} ms wall, ${text.length} chars\n${summarize(from)}`);
} finally {
  await closeScratch(s);
  await dispose();
}
