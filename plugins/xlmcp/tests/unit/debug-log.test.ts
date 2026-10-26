import { test, expect } from "bun:test";
import { join } from "path";
import { tmpdir } from "os";
import { randomUUID } from "crypto";
import { existsSync, readFileSync } from "fs";

const pluginRoot = join(import.meta.dir, "../..");

// P1-3: XLMCP_DEBUG=1 appends per-step timings (PS marks + pool wait/invoke) to %TEMP%\xlmcp.log.
test("XLMCP_DEBUG=1 logs PS step marks and pool timings", () => {
  const marker = `probe-${randomUUID()}`;
  const code = `
    import { runPS, dispose } from "./src/services/powershell.ts";
    runPS("Start-Sleep -Milliseconds 50; XlmcpMark '${marker}'")
      .then(() => console.log("OK"), (e) => console.log("ERR " + e.message))
      .finally(() => dispose().then(() => process.exit(0)));
  `;
  const p = Bun.spawnSync(["bun", "-e", code], { cwd: pluginRoot, env: { ...process.env, XLMCP_DEBUG: "1" } });
  expect(p.stdout.toString()).toContain("OK");
  const logPath = join(tmpdir(), "xlmcp.log");
  const log = existsSync(logPath) ? readFileSync(logPath, "utf-8") : "";
  const at = log.indexOf(marker);
  expect(at).toBeGreaterThan(-1);
  // the PS mark line carries the session id and elapsed ms; a pool line with wait/invoke follows
  expect(log.slice(log.lastIndexOf("\n", at) + 1, log.indexOf("\n", at))).toMatch(/ s\d+ ps probe-[\w-]+ \d+ms/);
  expect(log.slice(at)).toMatch(/ s\d+ runPS wait=\d+ms invoke=\d+ms ok/);
}, 30_000);
