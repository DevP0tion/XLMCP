import { writeFileSync, unlinkSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";
import { randomUUID } from "crypto";

/**
 * Run a script in a fresh Windows PowerShell 5.1 process (no pool, no Excel) and return stdout.
 * Used by the PS-only unit layer to execute generated script fragments.
 */
export function runPowerShell(script: string): string {
  const file = join(tmpdir(), `xlmcp_unit_${randomUUID()}.ps1`);
  // UTF-8 BOM: PS 5.1 reads BOM-less .ps1 files in the ANSI code page.
  writeFileSync(file, "﻿" + script, "utf-8");
  try {
    const p = Bun.spawnSync([
      "powershell.exe", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", file,
    ]);
    const out = p.stdout.toString();
    const err = p.stderr.toString();
    if (p.exitCode !== 0 || err.trim()) throw new Error(`PowerShell failed (${p.exitCode}): ${err || out}`);
    return out;
  } finally {
    try { unlinkSync(file); } catch { /* ignore */ }
  }
}
