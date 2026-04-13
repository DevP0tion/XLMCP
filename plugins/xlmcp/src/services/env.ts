/**
 * 플러그인 모드 감지 + 환경변수 리졸버.
 *
 * - 플러그인 모드: CLAUDE_PLUGIN_OPTION_* 존재 시 → CLAUDE_PLUGIN_OPTION_* 만 사용
 * - MCP 모드: 그 외 → XLMCP_* 사용
 */

const isPluginMode = Object.keys(process.env).some((k) =>
  k.startsWith("CLAUDE_PLUGIN_OPTION_")
);

/**
 * 환경변수 읽기. 플러그인 모드에 따라 소스 자동 분기.
 *
 * @param key - 변수명 (예: "LANG", "POOL_SIZE")
 * @returns 환경변수 값 또는 undefined
 *
 * @example
 * env("LANG")       // 플러그인: CLAUDE_PLUGIN_OPTION_LANG / MCP: XLMCP_LANG
 * env("POOL_SIZE")  // 플러그인: CLAUDE_PLUGIN_OPTION_POOL_SIZE / MCP: XLMCP_POOL_SIZE
 */
export function env(key: string): string | undefined {
  if (isPluginMode) {
    return process.env[`CLAUDE_PLUGIN_OPTION_${key}`];
  }
  return process.env[`XLMCP_${key}`];
}

export { isPluginMode };
