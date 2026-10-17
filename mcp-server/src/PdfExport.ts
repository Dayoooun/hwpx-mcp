/**
 * HWP/HWPX → PDF through Hancom's own layout engine, on whichever OS the server runs.
 *
 *   macOS    native/hwp2pdf drives libSimpleEngineLib, the engine 한/글 for Mac's Quick Look
 *            preview uses (call order recovered with REA from HwpQLExtPreview). No window, no focus.
 *   Windows  PowerShell creates HWPFrame.HwpObject (한/글 COM) → Open → SaveAs PDF → Quit.
 *
 * Both write into a private directory next to the output and rename into place, so a failed
 * conversion never leaves a partial PDF and never touches an existing file at output_path.
 */
import { spawn } from 'child_process';
import * as fs from 'fs';
import * as path from 'path';

export type PdfEngine = 'hancom-mac' | 'hancom-windows';
export type PdfEngineRequest = PdfEngine | 'auto';

export interface PdfResult {
  output_path: string;
  pages: number;
  engine: PdfEngine;
}

const ENGINE_PLATFORM: Record<PdfEngine, NodeJS.Platform> = {
  'hancom-mac': 'darwin',
  'hancom-windows': 'win32',
};

/** Which engine runs `requested` on `platform`; throws with what the user must install or choose. */
export function pickEngine(requested: PdfEngineRequest | undefined, platform: NodeJS.Platform): PdfEngine {
  const want = requested ?? 'auto';
  if (want === 'auto') {
    if (platform === 'darwin') return 'hancom-mac';
    if (platform === 'win32') return 'hancom-windows';
    throw new Error(
      `PDF export needs Hancom's engine, which runs only on macOS (한/글 for Mac) or Windows (한/글). ` +
      `This server runs on ${platform}.`
    );
  }
  if (!(want in ENGINE_PLATFORM)) throw new Error(`Unknown engine: ${want}. Use auto, hancom-mac or hancom-windows.`);
  if (ENGINE_PLATFORM[want] !== platform) {
    throw new Error(`Engine ${want} runs only on ${ENGINE_PLATFORM[want]}; this server runs on ${platform}. Use engine "auto".`);
  }
  return want;
}

/**
 * PowerShell for the Windows engine. Paths arrive through environment variables, never inside the
 * script text, so quotes or `$` in a file name cannot change what runs.
 */
export const WINDOWS_SCRIPT = [
  '$ErrorActionPreference = "Stop"',
  '[Console]::OutputEncoding = [Text.Encoding]::UTF8',
  '$h = New-Object -ComObject HWPFrame.HwpObject',
  'try {',
  // Suppresses the "allow access" prompt when the FilePathChecker module is registered; harmless otherwise.
  '  try { $h.RegisterModule("FilePathCheckDLL", "FilePathCheckerModule") | Out-Null } catch {}',
  '  $h.SetMessageBoxMode(0x00020000) | Out-Null',
  '  if (-not $h.Open($env:HWPX_PDF_IN, "", "forceopen:true")) { throw "open failed" }',
  '  $pages = $h.PageCount',
  '  if (-not $h.SaveAs($env:HWPX_PDF_OUT, "PDF", "")) { throw "SaveAs PDF failed" }',
  '  Write-Output ("{""pages"":" + $pages + "}")',
  '} finally { try { $h.Clear(1) } catch {}; try { $h.Quit() } catch {} }',
].join('\r\n');

/** Command, arguments and environment for one Windows conversion. Paths are only in `env`. */
export function windowsInvocation(scriptPath: string, input: string, output: string) {
  return {
    command: 'powershell.exe',
    args: ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', scriptPath],
    env: { HWPX_PDF_IN: input, HWPX_PDF_OUT: output },
  };
}

function hancomMacApp(env: NodeJS.ProcessEnv): string {
  return env.HWP_APP || '/Applications/Hancom Office HWP.app';
}

/** The bundled converter; HWPX_HWP2PDF overrides it (a locally built copy). */
function hwp2pdfBinary(env: NodeJS.ProcessEnv): string {
  return env.HWPX_HWP2PDF || path.resolve(__dirname, '..', 'native', 'hwp2pdf');
}

/** Why `engine` cannot run here, or null when it can. */
export function engineUnavailable(engine: PdfEngine, env: NodeJS.ProcessEnv = process.env): string | null {
  if (engine === 'hancom-mac') {
    const app = hancomMacApp(env);
    if (!fs.existsSync(path.join(app, 'Contents', 'PlugIns', 'HwpQLExtPreview.appex'))) {
      return `한/글 for Mac (Hancom Office HWP) is not installed at ${app}. Install it from the Mac App Store, or set HWP_APP.`;
    }
    const bin = hwp2pdfBinary(env);
    if (!fs.existsSync(bin)) return `PDF converter missing: ${bin}`;
  }
  return null;
}

interface RunResult { code: number | null; signal: NodeJS.Signals | null; timedOut: boolean; stdout: string; stderr: string }

function run(command: string, args: string[], env: NodeJS.ProcessEnv, timeoutMs: number): Promise<RunResult> {
  return new Promise((resolve, reject) => {
    const child = spawn(command, args, { env, stdio: ['ignore', 'pipe', 'pipe'], windowsHide: true });
    let stdout = '';
    let stderr = '';
    let timedOut = false;
    child.stdout.on('data', d => { stdout += d; });
    child.stderr.on('data', d => { stderr += d; });
    const timer = setTimeout(() => { timedOut = true; child.kill('SIGKILL'); }, timeoutMs);
    child.on('error', e => { clearTimeout(timer); reject(e); });
    child.on('close', (code, signal) => { clearTimeout(timer); resolve({ code, signal, timedOut, stdout, stderr }); });
  });
}

/** Our own kill on timeout must not read as an engine crash or a conversion error. */
function throwIfTimedOut(r: RunResult, timeoutMs: number): void {
  if (r.timedOut) throw new Error(`PDF conversion timed out after ${timeoutMs / 1000}s and was stopped`);
}

/** hwp2pdf exit codes (native/hwp2pdf.m). */
const MAC_EXIT: Record<number, string> = {
  3: 'engine could not be loaded',
  4: 'engine failed to start',
  5: 'the engine could not open the document (corrupt, password-protected or unsupported)',
  6: 'engine symbols missing (unsupported 한/글 version)',
  7: 'the document has no pages',
  8: 'a page failed to draw',
  9: 'the PDF could not be written',
};

function isPdf(file: string): boolean {
  try {
    const fd = fs.openSync(file, 'r');
    try {
      const head = Buffer.alloc(5);
      return fs.readSync(fd, head, 0, 5, 0) === 5 && head.toString('latin1') === '%PDF-';
    } finally { fs.closeSync(fd); }
  } catch { return false; }
}

export interface ExportOptions {
  engine?: PdfEngineRequest;
  platform?: NodeJS.Platform;
  env?: NodeJS.ProcessEnv;
  timeoutMs?: number;
}

/** Convert `input` (HWP/HWPX on disk) to `output`. Throws on any failure; `output` is untouched then. */
export async function exportPdf(input: string, output: string, opts: ExportOptions = {}): Promise<PdfResult> {
  const env = opts.env ?? process.env;
  const engine = pickEngine(opts.engine, opts.platform ?? process.platform);
  const missing = engineUnavailable(engine, env);
  if (missing) throw new Error(missing);

  const src = path.resolve(input);
  const dest = path.resolve(output);
  if (!fs.existsSync(src)) throw new Error(`Input not found: ${src}`);
  const destDir = path.dirname(dest);
  if (!fs.existsSync(destDir)) throw new Error(`Directory does not exist: ${destDir}`);
  const existing = fs.lstatSync(dest, { throwIfNoEntry: false });
  if (existing && !existing.isFile()) throw new Error('output_path must be a regular file');

  const work = fs.mkdtempSync(path.join(destDir, '.hwpx-pdf-'));
  const staged = path.join(work, 'out.pdf');
  const timeoutMs = opts.timeoutMs ?? 180_000;
  try {
    let pages: number;
    if (engine === 'hancom-mac') {
      const r = await run(hwp2pdfBinary(env), [src, staged], { ...env, TMPDIR: work }, timeoutMs);
      throwIfTimedOut(r, timeoutMs);
      if (r.code !== 0) {
        const why = r.signal ? `the Hancom engine crashed on this document (${r.signal})`
          : MAC_EXIT[r.code ?? -1] ?? `converter exited with ${r.code}`;
        throw new Error(`PDF conversion failed: ${why}`);
      }
      pages = JSON.parse(r.stdout.trim().split('\n').pop() || '{}').pages;
    } else {
      const script = path.join(work, 'convert.ps1');
      // UTF-8 with BOM: Windows PowerShell 5.1 reads a BOM-less script as the ANSI code page.
      fs.writeFileSync(script, '\ufeff' + WINDOWS_SCRIPT, 'utf8');
      const inv = windowsInvocation(script, src, staged);
      const r = await run(inv.command, inv.args, { ...env, ...inv.env }, timeoutMs);
      throwIfTimedOut(r, timeoutMs);
      if (r.code !== 0) {
        const why = (r.stderr || r.stdout).trim().split(/\r?\n/).slice(-3).join(' ');
        throw new Error(`PDF conversion failed in 한/글 (Windows): ${why || `exit ${r.code}`}`);
      }
      pages = JSON.parse(r.stdout.trim().split(/\r?\n/).pop() || '{}').pages;
    }
    if (!isPdf(staged)) throw new Error('PDF conversion failed: the engine produced no PDF');
    fs.renameSync(staged, dest);
    return { output_path: dest, pages: Number(pages) || 0, engine };
  } finally {
    fs.rmSync(work, { recursive: true, force: true });
  }
}
