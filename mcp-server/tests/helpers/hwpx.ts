/**
 * Shared helpers for module / regression tests.
 *
 * Every assertion about "did the edit land" is made on the SAVED file, and
 * where possible by reopening it with the parser — never on the in-memory
 * model alone. The memory model reporting success while the saved XML was
 * wrong is the defect class these tests exist for.
 */
import JSZip from 'jszip';
import { HwpxDocument } from '../../src/HwpxDocument';

export async function sectionXml(buf: Buffer, index = 0): Promise<string> {
  const file = (await JSZip.loadAsync(buf)).file(`Contents/section${index}.xml`);
  if (!file) throw new Error(`section${index}.xml missing`);
  return file.async('string');
}

export async function reopen(buf: Buffer): Promise<HwpxDocument> {
  return HwpxDocument.createFromBuffer('reopened', 'reopened.hwpx', buf);
}

/** Save → reopen: what a user gets back after closing and reopening the file. */
export async function roundTrip(doc: HwpxDocument): Promise<{ buf: Buffer; doc: HwpxDocument }> {
  const buf = await doc.save();
  return { buf, doc: await reopen(buf) };
}

/** Rewrite one section's XML inside a saved document and reopen it. */
export async function withSectionXml(
  doc: HwpxDocument,
  edit: (xml: string) => string,
  index = 0,
): Promise<HwpxDocument> {
  const zip = await JSZip.loadAsync(await doc.save());
  const path = `Contents/section${index}.xml`;
  zip.file(path, edit(await zip.file(path)!.async('string')));
  return reopen(await zip.generateAsync({ type: 'nodebuffer' }));
}

/**
 * Open / close balance of an element. Self-closing tags are neither.
 * 0 means every <tag> has its </tag>.
 */
export function tagBalance(xml: string, tag: string): number {
  const open = (xml.match(new RegExp(`<${tag}(?=[\\s>])`, 'g')) || []).length;
  const selfClosing = (xml.match(new RegExp(`<${tag}\\b[^>]*/>`, 'g')) || []).length;
  const close = (xml.match(new RegExp(`</${tag}>`, 'g')) || []).length;
  return open - selfClosing - close;
}

/** Every structural tag Hancom needs balanced to open the file. */
export function assertBalanced(xml: string): Record<string, number> {
  const out: Record<string, number> = {};
  for (const t of ['hp:p', 'hp:run', 'hp:tbl', 'hp:tr', 'hp:tc', 'hp:subList']) {
    const b = tagBalance(xml, t);
    if (b !== 0) out[t] = b;
  }
  return out;
}

/** Rows of the first table as their <hp:cellAddr rowAddr> values (unique per row). */
export function rowAddrsOfFirstTable(xml: string): string[] {
  const t = xml.slice(xml.indexOf('<hp:tbl'), xml.indexOf('</hp:tbl>') + 9);
  return t
    .split('</hp:tr>')
    .filter(r => r.includes('<hp:tr'))
    .map(r => [...new Set([...r.matchAll(/rowAddr="(\d+)"/g)].map(m => m[1]))].join('/'));
}

/** Plain text of a memory paragraph element. */
export function paragraphText(doc: HwpxDocument, section: number, element: number): string {
  const el = doc.content.sections[section].elements[element];
  if (!el || el.type !== 'paragraph') return '';
  return el.data.runs.map((r: { text: string }) => r.text).join('');
}

/**
 * Element index of the first paragraph whose text equals `text`.
 *
 * getParagraphs() lists paragraphs only, but each entry's `.index` is the
 * ELEMENT index that update_paragraph_text / copy_paragraph expect. The array
 * position (findIndex on that list) is off by one per table above the
 * paragraph — it only looks right in documents without tables.
 */
export function paragraphIndexOf(doc: HwpxDocument, section: number, text: string): number {
  const hit = doc.getParagraphs(section).find(p => p.text === text);
  if (!hit) throw new Error(`No paragraph with text ${JSON.stringify(text)} in section ${section}`);
  return hit.index;
}

/** Plain text of one table cell (all its paragraphs joined by \n). */
export function cellText(doc: HwpxDocument, section: number, tableIndex: number, row: number, col: number): string {
  const tables = doc.content.sections[section].elements.filter(e => e.type === 'table');
  const cell = tables[tableIndex]?.data.rows[row]?.cells[col];
  if (!cell) return '';
  return (cell.paragraphs || [])
    .map((p: { runs: Array<{ text: string }> }) => p.runs.map(r => r.text).join(''))
    .join('\n');
}
