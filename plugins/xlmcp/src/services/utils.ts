import { appendFileSync } from "fs";
import { tmpdir } from "os";
import { join } from "path";
import { env } from "./env.js";

/** XLMCP_DEBUG=1: per-step timing log (%TEMP%\xlmcp.log). null when off. */
export const DEBUG_LOG = env("DEBUG") === "1" ? join(tmpdir(), "xlmcp.log") : null;

export function debugLog(msg: string): void {
  if (!DEBUG_LOG) return;
  try {
    appendFileSync(DEBUG_LOG, `${new Date().toISOString()} ${msg}\n`);
  } catch { /* logging must never break a tool */ }
}

/** PowerShell 싱글쿼트 문자열 내 특수문자 이스케이프 */
export function psEscape(str: string): string {
  // 싱글쿼트 문자열 내에서는 '' 만 이스케이프하면 됨
  // 단, 백틱·$·" 등은 싱글쿼트 안에서 리터럴 처리되므로 안전
  // 줄바꿈/탭/널은 제거하여 명령 주입 방지
  return str
    .replace(/\0/g, "")
    .replace(/\r?\n/g, " ")
    .replace(/'/g, "''");
}

/**
 * PowerShell 출력에서 JSON을 안전하게 추출·파싱.
 * 경고 메시지 등이 섞여 있어도 첫 번째 JSON 구조를 찾아 파싱한다.
 */
export function parseJSON<T = unknown>(raw: string): T {
  const trimmed = raw.trim();

  // 그대로 파싱 시도
  try {
    return JSON.parse(trimmed);
  } catch {
    // JSON 시작 지점 탐색 (객체 또는 배열)
    const objStart = trimmed.indexOf("{");
    const arrStart = trimmed.indexOf("[");
    let start = -1;
    if (objStart === -1) start = arrStart;
    else if (arrStart === -1) start = objStart;
    else start = Math.min(objStart, arrStart);

    if (start === -1) {
      throw new Error(`JSON parse failed: ${trimmed.slice(0, 200)}`);
    }

    // 해당 지점부터 파싱 시도 (끝에서부터 줄여가며)
    const sub = trimmed.slice(start);
    try {
      return JSON.parse(sub);
    } catch {
      throw new Error(`JSON parse failed: ${sub.slice(0, 200)}`);
    }
  }
}

/**
 * Run a PS fragment with Excel's DisplayAlerts off (for COM calls that would otherwise raise a modal
 * prompt), restoring the previous value: the setting belongs to the user's whole Excel instance.
 */
export function withAlertsOff(cmd: string): string {
  return `
        $__prevAlerts = $excel.DisplayAlerts
        $excel.DisplayAlerts = $false
        try {
          ${cmd}
        } finally {
          $excel.DisplayAlerts = $__prevAlerts
        }`;
}

/** MCP text content 래퍼: compact JSON, a top-level Data 2D array gets one row per line. */
export function textContent(data: unknown) {
  return { content: [{ type: "text" as const, text: toJSON(data) }] };
}

function toJSON(data: unknown): string {
  const d = data as { Data?: unknown };
  if (!d || typeof d !== "object" || !Array.isArray(d.Data) || d.Data.length === 0) return JSON.stringify(data);
  const { Data, ...rest } = d as { Data: unknown[] };
  const head = JSON.stringify(rest).slice(0, -1); // drop the closing brace
  const rows = Data.map((r) => JSON.stringify(r)).join(",\n");
  return `${head}${head.length > 1 ? "," : ""}"Data":[\n${rows}\n]}`;
}

/** hex RGB → [r, g, b] 변환 */
export function hexToRgb(hex: string): [number, number, number] {
  const h = hex.replace("#", "");
  return [parseInt(h.substring(0, 2), 16), parseInt(h.substring(2, 4), 16), parseInt(h.substring(4, 6), 16)];
}

/** [r, g, b] → Excel OLE 색상값 변환 */
export function rgbToOle([r, g, b]: [number, number, number]): number {
  return r + g * 256 + b * 65536;
}
