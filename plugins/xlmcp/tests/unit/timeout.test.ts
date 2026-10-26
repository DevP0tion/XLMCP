import { test, expect, afterAll } from "bun:test";
import { join } from "path";
import { tmpdir } from "os";
import { randomUUID } from "crypto";
import { existsSync, readFileSync, unlinkSync } from "fs";
import { runPS, dispose } from "../../src/services/powershell.js";

const pluginRoot = join(import.meta.dir, "../..");

/** Run a slow script that appends one line to a counter file per execution; return run count and error. */
async function slowCountingCall(retry: boolean) {
  const counter = join(tmpdir(), `xlmcp_counter_${randomUUID()}.txt`);
  const t0 = Date.now();
  const err = await runPS(`Add-Content -Path '${counter}' -Value x; Start-Sleep -Milliseconds 2500`, {
    timeoutMs: 500,
    retry,
  }).then(() => null, (e: Error) => e);
  const elapsed = Date.now() - t0;
  await Bun.sleep(4000); // long enough for a hidden re-run to have appended its line
  const runs = existsSync(counter) ? readFileSync(counter, "utf-8").trim().split(/\r?\n/).length : 0;
  if (existsSync(counter)) unlinkSync(counter);
  return { err, elapsed, runs };
}

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

// P0-3: a timeout must not silently re-run a (possibly non-idempotent) script.
test("timed-out script runs once and fails with type Timeout", async () => {
  const { err, elapsed, runs } = await slowCountingCall(false);
  expect(runs).toBe(1);
  const e = JSON.parse(err!.message);
  expect(e.type).toBe("Timeout");
  expect(e.message).toContain("partially applied");
  expect(elapsed).toBeLessThan(4000); // no queued graceful dispose, no retry
}, 30_000);

test("retry: true re-runs a timed-out script once", async () => {
  const { err, runs } = await slowCountingCall(true);
  expect(runs).toBe(2);
  expect(JSON.parse(err!.message).type).toBe("Timeout");
}, 30_000);
