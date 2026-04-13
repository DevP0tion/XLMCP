/**
 * Plugin mode detection + env resolver.
 *
 * - Plugin mode: CLAUDE_PLUGIN_OPTION_* exists → read CLAUDE_PLUGIN_OPTION_* only
 * - MCP mode: otherwise → read XLMCP_*
 */

const isPluginMode = Object.keys(process.env).some((k) =>
  k.startsWith("CLAUDE_PLUGIN_OPTION_")
);

/**
 * Read environment variable. Auto-switches source based on mode.
 *
 * @param key - Variable name (e.g. "POOL_SIZE")
 * @returns Value or undefined
 *
 * @example
 * env("POOL_SIZE")  // Plugin: CLAUDE_PLUGIN_OPTION_POOL_SIZE / MCP: XLMCP_POOL_SIZE
 */
export function env(key: string): string | undefined {
  if (isPluginMode) {
    return process.env[`CLAUDE_PLUGIN_OPTION_${key}`];
  }
  return process.env[`XLMCP_${key}`];
}

export { isPluginMode };
