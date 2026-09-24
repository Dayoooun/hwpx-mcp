/**
 * Minimal MCP stdio client for end-to-end tests.
 *
 * Spawns a real server process and talks JSON-RPC over stdin/stdout, exactly
 * as Claude Code or any MCP host does. The server under test is chosen by
 * HWPX_MCP_SERVER:
 *   - unset / "local"      → node dist/index.js of this checkout (built first)
 *   - "npm:<version>"      → npx -y @kimdayoun/hwpx-mcp@<version>
 * The version matrix in CI runs the same e2e files against several values.
 */
import { spawn, type ChildProcessWithoutNullStreams } from 'node:child_process';
import path from 'node:path';

export interface ToolResult {
  isError: boolean;
  /** Parsed JSON body when the text is JSON, otherwise the raw text. */
  body: any;
  raw: string;
}

export class McpClient {
  private proc: ChildProcessWithoutNullStreams;
  private buf = '';
  private nextId = 1;
  private waiters = new Map<number, (msg: any) => void>();
  stderr = '';

  private constructor(proc: ChildProcessWithoutNullStreams) {
    this.proc = proc;
    proc.stdout.on('data', chunk => this.onData(chunk.toString()));
    proc.stderr.on('data', chunk => { this.stderr += chunk.toString(); });
  }

  static serverLabel(): string {
    return process.env.HWPX_MCP_SERVER || 'local';
  }

  static async start(cwd: string): Promise<McpClient> {
    const target = McpClient.serverLabel();
    const proc = target.startsWith('npm:')
      ? spawn('npx', ['-y', `@kimdayoun/hwpx-mcp@${target.slice(4)}`], { cwd, stdio: ['pipe', 'pipe', 'pipe'] })
      : spawn(process.execPath, [path.resolve(__dirname, '../../dist/index.js')], { cwd, stdio: ['pipe', 'pipe', 'pipe'] });
    const client = new McpClient(proc as ChildProcessWithoutNullStreams);
    await client.request('initialize', {
      protocolVersion: '2024-11-05',
      capabilities: {},
      clientInfo: { name: 'hwpx-mcp-e2e', version: '1' },
    });
    client.notify('notifications/initialized');
    return client;
  }

  private onData(text: string) {
    this.buf += text;
    let nl: number;
    while ((nl = this.buf.indexOf('\n')) >= 0) {
      const line = this.buf.slice(0, nl).trim();
      this.buf = this.buf.slice(nl + 1);
      if (!line) continue;
      let msg: any;
      try { msg = JSON.parse(line); } catch { continue; }
      if (msg.id !== undefined && this.waiters.has(msg.id)) {
        this.waiters.get(msg.id)!(msg);
        this.waiters.delete(msg.id);
      }
    }
  }

  private notify(method: string, params?: unknown) {
    this.proc.stdin.write(JSON.stringify({ jsonrpc: '2.0', method, params }) + '\n');
  }

  request(method: string, params: unknown, timeoutMs = 120_000): Promise<any> {
    const id = this.nextId++;
    this.proc.stdin.write(JSON.stringify({ jsonrpc: '2.0', id, method, params }) + '\n');
    return new Promise((resolve, reject) => {
      const t = setTimeout(() => {
        this.waiters.delete(id);
        reject(new Error(`MCP timeout: ${method}\nstderr: ${this.stderr.slice(-400)}`));
      }, timeoutMs);
      this.waiters.set(id, msg => { clearTimeout(t); resolve(msg); });
    });
  }

  async listTools(): Promise<string[]> {
    const r = await this.request('tools/list', {});
    return r.result.tools.map((t: { name: string }) => t.name);
  }

  async call(name: string, args: Record<string, unknown>): Promise<ToolResult> {
    const r = await this.request('tools/call', { name, arguments: args });
    if (r.error) return { isError: true, body: r.error, raw: JSON.stringify(r.error) };
    const raw = (r.result?.content ?? []).map((c: { text?: string }) => c.text ?? '').join('');
    let body: any = raw;
    try { body = JSON.parse(raw); } catch { /* plain text */ }
    return { isError: r.result?.isError === true, body, raw };
  }

  /** call() that fails the test with the server's message when the tool reports an error. */
  async ok(name: string, args: Record<string, unknown>): Promise<any> {
    const r = await this.call(name, args);
    if (r.isError || (r.body && typeof r.body === 'object' && 'error' in r.body)) {
      throw new Error(`${name} failed: ${r.raw.slice(0, 400)}`);
    }
    return r.body;
  }

  close() {
    this.proc.kill();
  }
}
