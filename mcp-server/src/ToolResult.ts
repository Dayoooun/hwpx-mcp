/**
 * MCP tool-result envelope and argument checks, kept out of index.ts so they
 * can be tested without starting the stdio server.
 */
import type { CallToolResult } from '@modelcontextprotocol/sdk/types.js';

export function success(data: unknown): CallToolResult {
  return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
}

export function error(message: string): CallToolResult {
  // isError tells the MCP client the call failed. Without it, a missing
  // argument or a refused write came back as a normal result whose body merely
  // contained {"error": …}, and agents treated it as success (reported
  // 2026-09-24). The JSON body is kept for clients that read it.
  return { content: [{ type: 'text', text: JSON.stringify({ error: message }) }], isError: true };
}

/**
 * The declared-required arguments of `toolName` that `args` leaves undefined or
 * null. Reporting them by name replaced a generic handler failure: omitting
 * `section_index` used to surface as "Failed to insert paragraph", which reads
 * like document corruption.
 */
export function findMissingArgs(
  requiredByTool: ReadonlyMap<string, readonly string[]>,
  toolName: string,
  args: Record<string, unknown> | undefined,
): string[] {
  const required = requiredByTool.get(toolName);
  if (!required || required.length === 0) return [];
  return required.filter(key => args?.[key] === undefined || args?.[key] === null);
}
