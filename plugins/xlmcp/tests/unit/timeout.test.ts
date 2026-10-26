import { test, expect, afterAll } from "bun:test";
import { join } from "path";
import { runPS, dispose } from "../../src/services/powershell.js";

const pluginRoot = join(import.meta.dir, "../..");

afterAll(() => dispose());

// P1-1: the invoke timeout is configurable per call.
test("runPS honors a per-call timeoutMs", async () => {
  const err = await runPS("Start-Sleep -Milliseconds 2500; 'done'", { timeoutMs: 500 }).then(
    () => null,
    (e: Error) => e
  );
  expect(err?.message ?? "resolved").toContain("Timeout");
}, 30_000);

// P1-1: XLMCP_TIMEOUT_MS sets the default (read at import, so run in a child process).
test("XLMCP_TIMEOUT_MS sets the default timeout", () => {
  const code = `
    import { runPS, dispose } from "./src/services/powershell.ts";
    runPS("Start-Sleep -Milliseconds 2500; 'done'")
      .then((r) => console.log("RESOLVED " + r.trim()), (e) => console.log("REJECTED " + e.message))
      .finally(() => dispose().then(() => process.exit(0)));
  `;
  const p = Bun.spawnSync(["bun", "-e", code], {
    cwd: pluginRoot,
    env: { ...process.env, XLMCP_TIMEOUT_MS: "500" },
  });
  const out = p.stdout.toString();
  expect(out).toContain("REJECTED");
  expect(out).toContain("Timeout");
}, 30_000);
