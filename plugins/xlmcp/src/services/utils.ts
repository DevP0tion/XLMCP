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

/** runPS 결과를 JSON으로 파싱. 실패 시 raw 텍스트 반환 */
export function parseJSON<T = unknown>(raw: string): T {
  const trimmed = raw.trim();
  return JSON.parse(trimmed);
}

/** MCP text content 래퍼 */
export function textContent(data: unknown) {
  return { content: [{ type: "text" as const, text: JSON.stringify(data, null, 2) }] };
}

/** MCP error content 래퍼 */
export function errorContent(message: string) {
  return { content: [{ type: "text" as const, text: JSON.stringify({ error: message }) }], isError: true };
}
