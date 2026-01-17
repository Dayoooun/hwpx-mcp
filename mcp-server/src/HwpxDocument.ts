import JSZip from 'jszip';
import { HwpxParser } from './HwpxParser';
import {
  HwpxContent,
  HwpxParagraph,
  TextRun,
  CharacterStyle,
  ParagraphStyle,
  HwpxTable,
  TableCell,
  TableRow,
  SectionElement,
  HwpxSection,
  HwpxImage,
  PageSettings,
  Footnote,
  Endnote,
  Memo,
  ColumnDef,
  CharShape,
  ParaShape,
  StyleDef,
  HwpxLine,
  HwpxRect,
  HwpxEllipse,
  HwpxEquation,
  HeaderFooter,
} from './types';

type DocumentFormat = 'hwpx' | 'hwp';

const MAX_UNDO_STACK_SIZE = 50;

export class HwpxDocument {
  private _id: string;
  private _path: string;
  private _zip: JSZip | null;
  private _content: HwpxContent;
  private _isDirty = false;
  private _format: DocumentFormat;

  private _undoStack: string[] = [];
  private _redoStack: string[] = [];
  private _pendingTextReplacements: Array<{ oldText: string; newText: string; options: { caseSensitive?: boolean; regex?: boolean; replaceAll?: boolean } }> = [];
  private _pendingDirectTextUpdates: Array<{ oldText: string; newText: string }> = [];
  private _pendingTableCellUpdates: Array<{ sectionIndex: number; tableIndex: number; tableId: string; row: number; col: number; text: string; charShapeId?: number }> = [];

  private constructor(id: string, path: string, zip: JSZip | null, content: HwpxContent, format: DocumentFormat) {
    this._id = id;
    this._path = path;
    this._zip = zip;
    this._content = content;
    this._format = format;
  }

  public static async createFromBuffer(id: string, path: string, data: Buffer): Promise<HwpxDocument> {
    const extension = path.toLowerCase();

    if (extension.endsWith('.hwp')) {
      // HWP parsing would go here - for now return empty content
      const content: HwpxContent = {
        metadata: {},
        sections: [],
        images: new Map(),
        binItems: new Map(),
        binData: new Map(),
        footnotes: [],
        endnotes: [],
      };
      return new HwpxDocument(id, path, null, content, 'hwp');
    } else {
      const zip = await JSZip.loadAsync(data);
      const content = await HwpxParser.parse(zip);
      return new HwpxDocument(id, path, zip, content, 'hwpx');
    }
  }

  public static createNew(id: string, title?: string, creator?: string): HwpxDocument {
    const now = new Date().toISOString();
    const content: HwpxContent = {
      metadata: {
        title: title || 'Untitled',
        creator: creator || 'Unknown',
        createdDate: now,
        modifiedDate: now,
      },
      sections: [{
        id: Math.random().toString(36).substring(2, 11),
        elements: [{
          type: 'paragraph',
          data: {
            id: Math.random().toString(36).substring(2, 11),
            runs: [{ text: '' }],
          },
        }],
        pageSettings: {
          width: 59528,
          height: 84188,
          marginTop: 4252,
          marginBottom: 4252,
          marginLeft: 4252,
          marginRight: 4252,
        },
      }],
      images: new Map(),
      binItems: new Map(),
      binData: new Map(),
      footnotes: [],
      endnotes: [],
    };

    // Create a new zip with basic HWPX structure
    const zip = new JSZip();

    // Add minimal required files for a valid HWPX document
    zip.file('mimetype', 'application/hwp+zip');
    zip.file('Contents/header.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:title>${title || 'Untitled'}</hh:title>
  <hh:creator>${creator || 'Unknown'}</hh:creator>
  <hh:createdDate>${now}</hh:createdDate>
  <hh:modifiedDate>${now}</hh:modifiedDate>
</hh:head>`);
    zip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hp:sec xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p>
    <hp:run>
      <hp:t></hp:t>
    </hp:run>
  </hp:p>
</hp:sec>`);

    return new HwpxDocument(id, 'new-document.hwpx', zip, content, 'hwpx');
  }

  get id(): string { return this._id; }
  get path(): string { return this._path; }
  get format(): DocumentFormat { return this._format; }
  get isDirty(): boolean { return this._isDirty; }
  get zip(): JSZip | null { return this._zip; }
  get content(): HwpxContent { return this._content; }

  // ============================================================
  // Undo/Redo
  // ============================================================

  private saveState(): void {
    const state = this.serializeContent();
    this._undoStack.push(state);
    if (this._undoStack.length > MAX_UNDO_STACK_SIZE) {
      this._undoStack.shift();
    }
    this._redoStack = [];
  }

  private serializeContent(): string {
    return JSON.stringify({
      sections: this._content.sections,
      metadata: this._content.metadata,
    });
  }

  private deserializeContent(state: string): void {
    const parsed = JSON.parse(state);
    this._content.sections = parsed.sections;
    this._content.metadata = parsed.metadata;
  }

  canUndo(): boolean { return this._undoStack.length > 0; }
  canRedo(): boolean { return this._redoStack.length > 0; }

  undo(): boolean {
    if (!this.canUndo()) return false;
    const currentState = this.serializeContent();
    this._redoStack.push(currentState);
    const previousState = this._undoStack.pop()!;
    this.deserializeContent(previousState);
    this._isDirty = true;
    return true;
  }

  redo(): boolean {
    if (!this.canRedo()) return false;
    const currentState = this.serializeContent();
    this._undoStack.push(currentState);
    const nextState = this._redoStack.pop()!;
    this.deserializeContent(nextState);
    this._isDirty = true;
    return true;
  }

  // ============================================================
  // Content Access
  // ============================================================

  getSerializableContent(): object {
    return {
      metadata: this._content.metadata,
      sections: this._content.sections,
      images: Array.from(this._content.images.entries()),
      footnotes: this._content.footnotes,
      endnotes: this._content.endnotes,
    };
  }

  getAllText(): string {
    let text = '';
    for (const section of this._content.sections) {
      for (const element of section.elements) {
        if (element.type === 'paragraph') {
          text += element.data.runs.map(r => r.text).join('') + '\n';
        }
      }
    }
    return text;
  }

  getStructure(): object {
    return {
      format: this._format,
      sections: this._content.sections.map((s, i) => {
        let paragraphs = 0, tables = 0, images = 0;
        for (const el of s.elements) {
          if (el.type === 'paragraph') paragraphs++;
          if (el.type === 'table') tables++;
          if (el.type === 'image') images++;
        }
        return { section: i, paragraphs, tables, images };
      }),
    };
  }

  // ============================================================
  // Paragraph Operations
  // ============================================================

  private findParagraphByPath(sectionIndex: number, elementIndex: number): HwpxParagraph | null {
    const section = this._content.sections[sectionIndex];
    if (!section) return null;
    const element = section.elements[elementIndex];
    if (!element || element.type !== 'paragraph') return null;
    return element.data;
  }

  getParagraphs(sectionIndex?: number): Array<{ section: number; index: number; text: string; style?: ParagraphStyle }> {
    const paragraphs: Array<{ section: number; index: number; text: string; style?: ParagraphStyle }> = [];
    const sections = sectionIndex !== undefined
      ? [{ section: this._content.sections[sectionIndex], idx: sectionIndex }]
      : this._content.sections.map((s, i) => ({ section: s, idx: i }));

    for (const { section, idx } of sections) {
      if (!section) continue;
      section.elements.forEach((el, ei) => {
        if (el.type === 'paragraph') {
          paragraphs.push({
            section: idx,
            index: ei,
            text: el.data.runs.map(r => r.text).join(''),
            style: el.data.paraStyle,
          });
        }
      });
    }
    return paragraphs;
  }

  getParagraph(sectionIndex: number, paragraphIndex: number): { text: string; runs: TextRun[]; style?: ParagraphStyle } | null {
    const para = this.findParagraphByPath(sectionIndex, paragraphIndex);
    if (!para) return null;
    return {
      text: para.runs.map(r => r.text).join(''),
      runs: para.runs,
      style: para.paraStyle,
    };
  }

  updateParagraphText(sectionIndex: number, elementIndex: number, runIndex: number, text: string): void {
    const paragraph = this.findParagraphByPath(sectionIndex, elementIndex);
    if (!paragraph || !paragraph.runs[runIndex]) return;

    // Track the old text for XML update
    const oldText = paragraph.runs[runIndex].text;
    if (oldText && oldText !== text && this._zip) {
      this._pendingDirectTextUpdates.push({ oldText, newText: text });
    }

    this.saveState();
    paragraph.runs[runIndex].text = text;
    this._isDirty = true;
  }

  updateParagraphRuns(sectionIndex: number, elementIndex: number, runs: TextRun[]): void {
    const paragraph = this.findParagraphByPath(sectionIndex, elementIndex);
    if (!paragraph) return;
    this.saveState();
    paragraph.runs = runs;
    this._isDirty = true;
  }

  insertParagraph(sectionIndex: number, afterElementIndex: number, text: string = ''): number {
    const section = this._content.sections[sectionIndex];
    if (!section) return -1;

    this.saveState();
    const newParagraph: HwpxParagraph = {
      id: Math.random().toString(36).substring(2, 11),
      runs: [{ text }],
    };

    const newElement: SectionElement = { type: 'paragraph', data: newParagraph };
    section.elements.splice(afterElementIndex + 1, 0, newElement);
    this._isDirty = true;
    return afterElementIndex + 1;
  }

  deleteParagraph(sectionIndex: number, elementIndex: number): boolean {
    const section = this._content.sections[sectionIndex];
    if (!section || elementIndex < 0 || elementIndex >= section.elements.length) return false;

    this.saveState();
    section.elements.splice(elementIndex, 1);
    this._isDirty = true;
    return true;
  }

  appendTextToParagraph(sectionIndex: number, elementIndex: number, text: string): void {
    const paragraph = this.findParagraphByPath(sectionIndex, elementIndex);
    if (!paragraph) return;

    this.saveState();
    paragraph.runs.push({ text });
    this._isDirty = true;
  }

  // ============================================================
  // Character Style Operations
  // ============================================================

  applyCharacterStyle(sectionIndex: number, elementIndex: number, runIndex: number, style: Partial<CharacterStyle>): void {
    const paragraph = this.findParagraphByPath(sectionIndex, elementIndex);
    if (!paragraph || !paragraph.runs[runIndex]) return;

    this.saveState();
    const run = paragraph.runs[runIndex];
    run.charStyle = { ...run.charStyle, ...style };
    this._isDirty = true;
  }

  getCharacterStyle(sectionIndex: number, elementIndex: number, runIndex?: number): CharacterStyle | CharacterStyle[] | null {
    const paragraph = this.findParagraphByPath(sectionIndex, elementIndex);
    if (!paragraph) return null;

    if (runIndex !== undefined) {
      return paragraph.runs[runIndex]?.charStyle || null;
    }
    return paragraph.runs.map(r => r.charStyle || {});
  }

  // ============================================================
  // Paragraph Style Operations
  // ============================================================

  applyParagraphStyle(sectionIndex: number, elementIndex: number, style: Partial<ParagraphStyle>): void {
    const paragraph = this.findParagraphByPath(sectionIndex, elementIndex);
    if (!paragraph) return;

    this.saveState();
    paragraph.paraStyle = { ...paragraph.paraStyle, ...style };
    this._isDirty = true;
  }

  getParagraphStyle(sectionIndex: number, elementIndex: number): ParagraphStyle | null {
    const paragraph = this.findParagraphByPath(sectionIndex, elementIndex);
    return paragraph?.paraStyle || null;
  }

  // ============================================================
  // Table Operations
  // ============================================================

  private findTable(sectionIndex: number, tableIndex: number): HwpxTable | null {
    const section = this._content.sections[sectionIndex];
    if (!section) return null;
    const tables = section.elements.filter(el => el.type === 'table');
    return tables[tableIndex]?.data as HwpxTable || null;
  }

  getTables(): Array<{ section: number; index: number; rows: number; cols: number }> {
    const tables: Array<{ section: number; index: number; rows: number; cols: number }> = [];
    this._content.sections.forEach((section, si) => {
      let tableIndex = 0;
      section.elements.forEach(el => {
        if (el.type === 'table') {
          const table = el.data as HwpxTable;
          tables.push({
            section: si,
            index: tableIndex++,
            rows: table.rows.length,
            cols: table.rows[0]?.cells.length || 0,
          });
        }
      });
    });
    return tables;
  }

  getTable(sectionIndex: number, tableIndex: number): { rows: number; cols: number; data: any[][] } | null {
    const table = this.findTable(sectionIndex, tableIndex);
    if (!table) return null;

    return {
      rows: table.rows.length,
      cols: table.rows[0]?.cells.length || 0,
      data: table.rows.map(row => row.cells.map(cell => ({
        text: cell.paragraphs.map(p => p.runs.map(r => r.text).join('')).join('\n'),
        style: cell,
      }))),
    };
  }

  getTableCell(sectionIndex: number, tableIndex: number, row: number, col: number): { text: string; cell: TableCell } | null {
    const table = this.findTable(sectionIndex, tableIndex);
    if (!table) return null;
    const cell = table.rows[row]?.cells[col];
    if (!cell) return null;
    return {
      text: cell.paragraphs.map(p => p.runs.map(r => r.text).join('')).join('\n'),
      cell,
    };
  }

  updateTableCell(sectionIndex: number, tableIndex: number, row: number, col: number, text: string, charShapeId?: number): boolean {
    const table = this.findTable(sectionIndex, tableIndex);
    if (!table) return false;
    const cell = table.rows[row]?.cells[col];
    if (!cell) return false;

    // Track cell update for XML sync (works for both empty and non-empty cells)
    // Store table ID for reliable XML matching
    // charShapeId is optional - if provided, it will override the existing charPrIDRef
    this._pendingTableCellUpdates.push({ sectionIndex, tableIndex, tableId: table.id, row, col, text, charShapeId });

    this.saveState();
    if (cell.paragraphs.length > 0 && cell.paragraphs[0].runs.length > 0) {
      cell.paragraphs[0].runs[0].text = text;
    } else {
      cell.paragraphs = [{ id: Math.random().toString(36).substring(2, 11), runs: [{ text }] }];
    }
    this._isDirty = true;
    return true;
  }

  setCellProperties(sectionIndex: number, tableIndex: number, row: number, col: number, props: Partial<TableCell>): boolean {
    const table = this.findTable(sectionIndex, tableIndex);
    if (!table) return false;
    const cell = table.rows[row]?.cells[col];
    if (!cell) return false;

    this.saveState();
    Object.assign(cell, props);
    this._isDirty = true;
    return true;
  }

  insertTableRow(sectionIndex: number, tableIndex: number, afterRowIndex: number, cellTexts?: string[]): boolean {
    const table = this.findTable(sectionIndex, tableIndex);
    if (!table || !table.rows[afterRowIndex]) return false;

    this.saveState();
    const templateRow = table.rows[afterRowIndex];
    const colCount = templateRow.cells.length;

    const newRow = {
      cells: Array.from({ length: colCount }, (_, i) => ({
        paragraphs: [{
          id: Math.random().toString(36).substring(2, 11),
          runs: [{ text: cellTexts?.[i] || '' }],
        }],
      })),
    };

    table.rows.splice(afterRowIndex + 1, 0, newRow as any);
    this._isDirty = true;
    return true;
  }

  deleteTableRow(sectionIndex: number, tableIndex: number, rowIndex: number): boolean {
    const table = this.findTable(sectionIndex, tableIndex);
    if (!table || table.rows.length <= 1) return false;

    this.saveState();
    table.rows.splice(rowIndex, 1);
    this._isDirty = true;
    return true;
  }

  insertTableColumn(sectionIndex: number, tableIndex: number, afterColIndex: number): boolean {
    const table = this.findTable(sectionIndex, tableIndex);
    if (!table) return false;

    this.saveState();
    for (const row of table.rows) {
      row.cells.splice(afterColIndex + 1, 0, {
        paragraphs: [{
          id: Math.random().toString(36).substring(2, 11),
          runs: [{ text: '' }],
        }],
      } as any);
    }
    this._isDirty = true;
    return true;
  }

  deleteTableColumn(sectionIndex: number, tableIndex: number, colIndex: number): boolean {
    const table = this.findTable(sectionIndex, tableIndex);
    if (!table || (table.rows[0]?.cells.length || 0) <= 1) return false;

    this.saveState();
    for (const row of table.rows) {
      row.cells.splice(colIndex, 1);
    }
    this._isDirty = true;
    return true;
  }

  getTableAsCsv(sectionIndex: number, tableIndex: number, delimiter: string = ','): string | null {
    const table = this.findTable(sectionIndex, tableIndex);
    if (!table) return null;

    return table.rows.map(row =>
      row.cells.map(cell => {
        const text = cell.paragraphs.map(p => p.runs.map(r => r.text).join('')).join(' ');
        if (text.includes(delimiter) || text.includes('"') || text.includes('\n')) {
          return `"${text.replace(/"/g, '""')}"`;
        }
        return text;
      }).join(delimiter)
    ).join('\n');
  }

  // ============================================================
  // Search & Replace
  // ============================================================

  searchText(query: string, options: { caseSensitive?: boolean; regex?: boolean } = {}): Array<{ section: number; element: number; text: string; matches: string[]; count: number }> {
    const { caseSensitive = false, regex = false } = options;
    let pattern: RegExp;

    if (regex) {
      pattern = new RegExp(query, caseSensitive ? 'g' : 'gi');
    } else {
      const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      pattern = new RegExp(escaped, caseSensitive ? 'g' : 'gi');
    }

    const results: Array<{ section: number; element: number; text: string; matches: string[]; count: number }> = [];

    this._content.sections.forEach((section, si) => {
      section.elements.forEach((el, ei) => {
        if (el.type === 'paragraph') {
          const text = el.data.runs.map(r => r.text).join('');
          const found = text.match(pattern);
          if (found) {
            results.push({
              section: si,
              element: ei,
              text,
              matches: found,
              count: found.length,
            });
          }
        }
      });
    });

    return results;
  }

  replaceText(oldText: string, newText: string, options: { caseSensitive?: boolean; regex?: boolean; replaceAll?: boolean } = {}): number {
    const { caseSensitive = false, regex = false, replaceAll = true } = options;
    let pattern: RegExp;

    if (regex) {
      pattern = new RegExp(oldText, caseSensitive ? (replaceAll ? 'g' : '') : (replaceAll ? 'gi' : 'i'));
    } else {
      const escaped = oldText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      pattern = new RegExp(escaped, caseSensitive ? (replaceAll ? 'g' : '') : (replaceAll ? 'gi' : 'i'));
    }

    this.saveState();
    let count = 0;

    // Update in-memory content
    for (const section of this._content.sections) {
      for (const element of section.elements) {
        if (element.type === 'paragraph') {
          for (const run of element.data.runs) {
            const matches = run.text.match(pattern);
            if (matches) {
              count += matches.length;
              run.text = run.text.replace(pattern, newText);
            }
          }
        }
        // Also handle table cells
        if (element.type === 'table') {
          const table = element.data as HwpxTable;
          for (const row of table.rows) {
            for (const cell of row.cells) {
              for (const para of cell.paragraphs) {
                for (const run of para.runs) {
                  const matches = run.text.match(pattern);
                  if (matches) {
                    count += matches.length;
                    run.text = run.text.replace(pattern, newText);
                  }
                }
              }
            }
          }
        }
      }
    }

    // Also update directly in the ZIP XML files for safe saving
    if (count > 0 && this._zip) {
      this._pendingTextReplacements = this._pendingTextReplacements || [];
      this._pendingTextReplacements.push({ oldText, newText, options });
      this._isDirty = true;
    }

    return count;
  }

  // ============================================================
  // Metadata
  // ============================================================

  getMetadata(): HwpxContent['metadata'] {
    return this._content.metadata;
  }

  setMetadata(metadata: Partial<HwpxContent['metadata']>): void {
    this.saveState();
    this._content.metadata = { ...this._content.metadata, ...metadata };
    this._isDirty = true;
  }

  // ============================================================
  // Page Settings
  // ============================================================

  getPageSettings(sectionIndex: number = 0): PageSettings | null {
    const section = this._content.sections[sectionIndex];
    return section?.pageSettings || null;
  }

  setPageSettings(sectionIndex: number, settings: Partial<PageSettings>): boolean {
    const section = this._content.sections[sectionIndex];
    if (!section) return false;

    this.saveState();
    section.pageSettings = { ...section.pageSettings, ...settings } as PageSettings;
    this._isDirty = true;
    return true;
  }

  // ============================================================
  // Statistics
  // ============================================================

  getWordCount(): { characters: number; charactersNoSpaces: number; words: number; paragraphs: number } {
    let characters = 0;
    let charactersNoSpaces = 0;
    let words = 0;
    let paragraphs = 0;

    for (const section of this._content.sections) {
      for (const element of section.elements) {
        if (element.type === 'paragraph') {
          paragraphs++;
          const text = element.data.runs.map(r => r.text).join('');
          characters += text.length;
          charactersNoSpaces += text.replace(/\s/g, '').length;
          words += text.trim().split(/\s+/).filter(w => w.length > 0).length;
        }
      }
    }

    return { characters, charactersNoSpaces, words, paragraphs };
  }

  // ============================================================
  // Copy/Move Operations
  // ============================================================

  copyParagraph(sourceSection: number, sourceParagraph: number, targetSection: number, targetAfter: number): boolean {
    const srcSection = this._content.sections[sourceSection];
    const tgtSection = this._content.sections[targetSection];
    if (!srcSection || !tgtSection) return false;

    const srcElement = srcSection.elements[sourceParagraph];
    if (!srcElement || srcElement.type !== 'paragraph') return false;

    this.saveState();
    const copy = JSON.parse(JSON.stringify(srcElement));
    copy.data.id = Math.random().toString(36).substring(2, 11);
    tgtSection.elements.splice(targetAfter + 1, 0, copy);
    this._isDirty = true;
    return true;
  }

  moveParagraph(sourceSection: number, sourceParagraph: number, targetSection: number, targetAfter: number): boolean {
    const srcSection = this._content.sections[sourceSection];
    const tgtSection = this._content.sections[targetSection];
    if (!srcSection || !tgtSection) return false;

    const srcElement = srcSection.elements[sourceParagraph];
    if (!srcElement || srcElement.type !== 'paragraph') return false;

    this.saveState();
    srcSection.elements.splice(sourceParagraph, 1);
    tgtSection.elements.splice(targetAfter + 1, 0, srcElement);
    this._isDirty = true;
    return true;
  }

  // ============================================================
  // Images
  // ============================================================

  getImages(): Array<{ id: string; width: number; height: number }> {
    return Array.from(this._content.images.values()).map(img => ({
      id: img.id,
      width: img.width,
      height: img.height,
    }));
  }

  // ============================================================
  // Table Creation
  // ============================================================

  insertTable(sectionIndex: number, afterElementIndex: number, rows: number, cols: number, options?: { width?: number; cellWidth?: number }): { tableIndex: number } | null {
    const section = this._content.sections[sectionIndex];
    if (!section) return null;
    if (rows <= 0 || cols <= 0) return null;

    this.saveState();

    const tableId = Math.random().toString(36).substring(2, 11);
    const defaultWidth = options?.width || 42520; // Default table width in hwpunit
    const cellWidth = options?.cellWidth || Math.floor(defaultWidth / cols);

    const tableRows: TableRow[] = [];
    for (let r = 0; r < rows; r++) {
      const cells: TableCell[] = [];
      for (let c = 0; c < cols; c++) {
        cells.push({
          colAddr: c,
          rowAddr: r,
          colSpan: 1,
          rowSpan: 1,
          width: cellWidth,
          paragraphs: [{
            id: Math.random().toString(36).substring(2, 11),
            runs: [{ text: '' }],
          }],
        });
      }
      tableRows.push({ cells });
    }

    const newTable: HwpxTable = {
      id: tableId,
      rowCount: rows,
      colCount: cols,
      rows: tableRows,
      width: defaultWidth,
    };

    const newElement: SectionElement = { type: 'table', data: newTable };
    section.elements.splice(afterElementIndex + 1, 0, newElement);

    // Calculate table index
    let tableIndex = 0;
    for (let i = 0; i <= afterElementIndex + 1; i++) {
      if (section.elements[i]?.type === 'table') {
        if (i === afterElementIndex + 1) break;
        tableIndex++;
      }
    }

    this._isDirty = true;
    return { tableIndex };
  }

  // ============================================================
  // Header/Footer Operations
  // ============================================================

  getHeader(sectionIndex: number): { paragraphs: any[] } | null {
    const section = this._content.sections[sectionIndex];
    if (!section || !section.header) return null;
    return {
      paragraphs: section.header.paragraphs.map(p => ({
        id: p.id,
        text: p.runs.map(r => r.text).join(''),
        runs: p.runs,
      })),
    };
  }

  setHeader(sectionIndex: number, text: string): boolean {
    const section = this._content.sections[sectionIndex];
    if (!section) return false;

    this.saveState();

    const headerParagraph: HwpxParagraph = {
      id: Math.random().toString(36).substring(2, 11),
      runs: [{ text }],
    };

    if (!section.header) {
      section.header = {
        paragraphs: [headerParagraph],
      };
    } else {
      section.header.paragraphs = [headerParagraph];
    }

    this._isDirty = true;
    return true;
  }

  getFooter(sectionIndex: number): { paragraphs: any[] } | null {
    const section = this._content.sections[sectionIndex];
    if (!section || !section.footer) return null;
    return {
      paragraphs: section.footer.paragraphs.map(p => ({
        id: p.id,
        text: p.runs.map(r => r.text).join(''),
        runs: p.runs,
      })),
    };
  }

  setFooter(sectionIndex: number, text: string): boolean {
    const section = this._content.sections[sectionIndex];
    if (!section) return false;

    this.saveState();

    const footerParagraph: HwpxParagraph = {
      id: Math.random().toString(36).substring(2, 11),
      runs: [{ text }],
    };

    if (!section.footer) {
      section.footer = {
        paragraphs: [footerParagraph],
      };
    } else {
      section.footer.paragraphs = [footerParagraph];
    }

    this._isDirty = true;
    return true;
  }

  // ============================================================
  // Footnote/Endnote Operations
  // ============================================================

  getFootnotes(): Footnote[] {
    return this._content.footnotes || [];
  }

  insertFootnote(sectionIndex: number, paragraphIndex: number, text: string): { id: string } | null {
    const paragraph = this.findParagraphByPath(sectionIndex, paragraphIndex);
    if (!paragraph) return null;

    this.saveState();

    const footnoteId = Math.random().toString(36).substring(2, 11);
    const footnoteNumber = (this._content.footnotes?.length || 0) + 1;

    const footnote: Footnote = {
      id: footnoteId,
      number: footnoteNumber,
      type: 'footnote',
      paragraphs: [{
        id: Math.random().toString(36).substring(2, 11),
        runs: [{ text }],
      }],
    };

    if (!this._content.footnotes) {
      this._content.footnotes = [];
    }
    this._content.footnotes.push(footnote);

    // Add footnote reference to the paragraph
    paragraph.runs.push({
      text: '',
      footnoteRef: footnoteNumber,
    });

    this._isDirty = true;
    return { id: footnoteId };
  }

  getEndnotes(): Endnote[] {
    return this._content.endnotes || [];
  }

  insertEndnote(sectionIndex: number, paragraphIndex: number, text: string): { id: string } | null {
    const paragraph = this.findParagraphByPath(sectionIndex, paragraphIndex);
    if (!paragraph) return null;

    this.saveState();

    const endnoteId = Math.random().toString(36).substring(2, 11);
    const endnoteNumber = (this._content.endnotes?.length || 0) + 1;

    const endnote: Endnote = {
      id: endnoteId,
      number: endnoteNumber,
      paragraphs: [{
        id: Math.random().toString(36).substring(2, 11),
        runs: [{ text }],
      }],
    };

    if (!this._content.endnotes) {
      this._content.endnotes = [];
    }
    this._content.endnotes.push(endnote);

    // Add endnote reference to the paragraph
    paragraph.runs.push({
      text: '',
      endnoteRef: endnoteNumber,
    });

    this._isDirty = true;
    return { id: endnoteId };
  }

  // ============================================================
  // Bookmark/Hyperlink Operations
  // ============================================================

  getBookmarks(): { name: string; section: number; paragraph: number }[] {
    const bookmarks: { name: string; section: number; paragraph: number }[] = [];

    this._content.sections.forEach((section, si) => {
      section.elements.forEach((el, ei) => {
        if (el.type === 'paragraph') {
          for (const run of el.data.runs) {
            if (run.field?.fieldType === 'Bookmark' || run.field?.fieldType === 'bookmark') {
              bookmarks.push({
                name: run.field.name || '',
                section: si,
                paragraph: ei,
              });
            }
          }
        }
      });
    });

    return bookmarks;
  }

  insertBookmark(sectionIndex: number, paragraphIndex: number, name: string): boolean {
    const paragraph = this.findParagraphByPath(sectionIndex, paragraphIndex);
    if (!paragraph) return false;

    this.saveState();

    paragraph.runs.push({
      text: '',
      field: {
        fieldType: 'Bookmark',
        name,
      },
    });

    this._isDirty = true;
    return true;
  }

  getHyperlinks(): { url: string; text: string; section: number; paragraph: number }[] {
    const hyperlinks: { url: string; text: string; section: number; paragraph: number }[] = [];

    this._content.sections.forEach((section, si) => {
      section.elements.forEach((el, ei) => {
        if (el.type === 'paragraph') {
          for (const run of el.data.runs) {
            if (run.hyperlink) {
              hyperlinks.push({
                url: run.hyperlink.url,
                text: run.text || run.hyperlink.name || '',
                section: si,
                paragraph: ei,
              });
            }
          }
        }
      });
    });

    return hyperlinks;
  }

  insertHyperlink(sectionIndex: number, paragraphIndex: number, url: string, text: string): boolean {
    const paragraph = this.findParagraphByPath(sectionIndex, paragraphIndex);
    if (!paragraph) return false;

    this.saveState();

    paragraph.runs.push({
      text,
      hyperlink: {
        fieldType: 'Hyperlink',
        url,
        name: text,
      },
    });

    this._isDirty = true;
    return true;
  }

  // ============================================================
  // Image Operations
  // ============================================================

  insertImage(sectionIndex: number, afterElementIndex: number, imageData: { data: string; mimeType: string; width: number; height: number }): { id: string } | null {
    const section = this._content.sections[sectionIndex];
    if (!section) return null;

    this.saveState();

    const imageId = Math.random().toString(36).substring(2, 11);
    const binaryId = Math.random().toString(36).substring(2, 11);

    const newImage: HwpxImage = {
      id: imageId,
      binaryId,
      width: imageData.width,
      height: imageData.height,
      data: imageData.data,
      mimeType: imageData.mimeType,
    };

    // Store image in the images map
    this._content.images.set(imageId, newImage);

    // Store binary data
    this._content.binData.set(binaryId, {
      id: binaryId,
      encoding: 'Base64',
      data: imageData.data,
    });

    // Add image element to section
    const newElement: SectionElement = { type: 'image', data: newImage };
    section.elements.splice(afterElementIndex + 1, 0, newElement);

    this._isDirty = true;
    return { id: imageId };
  }

  updateImageSize(imageId: string, width: number, height: number): boolean {
    const image = this._content.images.get(imageId);
    if (!image) return false;

    this.saveState();

    image.width = width;
    image.height = height;

    this._isDirty = true;
    return true;
  }

  deleteImage(imageId: string): boolean {
    const image = this._content.images.get(imageId);
    if (!image) return false;

    this.saveState();

    // Remove from images map
    this._content.images.delete(imageId);

    // Remove binary data if exists
    if (image.binaryId) {
      this._content.binData.delete(image.binaryId);
    }

    // Remove from sections
    for (const section of this._content.sections) {
      const index = section.elements.findIndex(el => el.type === 'image' && el.data.id === imageId);
      if (index !== -1) {
        section.elements.splice(index, 1);
        break;
      }
    }

    this._isDirty = true;
    return true;
  }

  // ============================================================
  // Drawing Objects (Line, Rect, Ellipse)
  // ============================================================

  insertLine(sectionIndex: number, x1: number, y1: number, x2: number, y2: number, options?: { color?: string; width?: number }): { id: string } | null {
    const section = this._content.sections[sectionIndex];
    if (!section) return null;

    this.saveState();

    const lineId = Math.random().toString(36).substring(2, 11);

    const newLine: HwpxLine = {
      id: lineId,
      x1,
      y1,
      x2,
      y2,
      strokeColor: options?.color || '#000000',
      strokeWidth: options?.width || 1,
    };

    const newElement: SectionElement = { type: 'line', data: newLine };
    section.elements.push(newElement);

    this._isDirty = true;
    return { id: lineId };
  }

  insertRect(sectionIndex: number, x: number, y: number, width: number, height: number, options?: { fillColor?: string; strokeColor?: string }): { id: string } | null {
    const section = this._content.sections[sectionIndex];
    if (!section) return null;

    this.saveState();

    const rectId = Math.random().toString(36).substring(2, 11);

    const newRect: HwpxRect = {
      id: rectId,
      x,
      y,
      width,
      height,
      fillColor: options?.fillColor,
      strokeColor: options?.strokeColor || '#000000',
    };

    const newElement: SectionElement = { type: 'rect', data: newRect };
    section.elements.push(newElement);

    this._isDirty = true;
    return { id: rectId };
  }

  insertEllipse(sectionIndex: number, cx: number, cy: number, rx: number, ry: number, options?: { fillColor?: string; strokeColor?: string }): { id: string } | null {
    const section = this._content.sections[sectionIndex];
    if (!section) return null;

    this.saveState();

    const ellipseId = Math.random().toString(36).substring(2, 11);

    const newEllipse: HwpxEllipse = {
      id: ellipseId,
      cx,
      cy,
      rx,
      ry,
      fillColor: options?.fillColor,
      strokeColor: options?.strokeColor || '#000000',
    };

    const newElement: SectionElement = { type: 'ellipse', data: newEllipse };
    section.elements.push(newElement);

    this._isDirty = true;
    return { id: ellipseId };
  }

  // ============================================================
  // Equation Operations
  // ============================================================

  insertEquation(sectionIndex: number, afterElementIndex: number, script: string): { id: string } | null {
    const section = this._content.sections[sectionIndex];
    if (!section) return null;

    this.saveState();

    const equationId = Math.random().toString(36).substring(2, 11);

    const newEquation: HwpxEquation = {
      id: equationId,
      script,
      lineMode: false,
      baseUnit: 1000,
    };

    const newElement: SectionElement = { type: 'equation', data: newEquation };
    section.elements.splice(afterElementIndex + 1, 0, newElement);

    this._isDirty = true;
    return { id: equationId };
  }

  getEquations(): { id: string; script: string }[] {
    const equations: { id: string; script: string }[] = [];

    for (const section of this._content.sections) {
      for (const element of section.elements) {
        if (element.type === 'equation') {
          equations.push({
            id: element.data.id,
            script: element.data.script || '',
          });
        }
      }
    }

    return equations;
  }

  // ============================================================
  // Memo Operations
  // ============================================================

  getMemos(): Memo[] {
    const memos: Memo[] = [];

    for (const section of this._content.sections) {
      if (section.memos) {
        memos.push(...section.memos);
      }
    }

    return memos;
  }

  insertMemo(sectionIndex: number, paragraphIndex: number, content: string, author?: string): { id: string } | null {
    const section = this._content.sections[sectionIndex];
    if (!section) return null;

    const paragraph = this.findParagraphByPath(sectionIndex, paragraphIndex);
    if (!paragraph) return null;

    this.saveState();

    const memoId = Math.random().toString(36).substring(2, 11);

    const memo: Memo = {
      id: memoId,
      author: author || 'Unknown',
      date: new Date().toISOString(),
      content: [content],
    };

    if (!section.memos) {
      section.memos = [];
    }
    section.memos.push(memo);

    // Mark the paragraph as having a memo
    if (paragraph.runs.length > 0) {
      paragraph.runs[paragraph.runs.length - 1].hasMemo = true;
      paragraph.runs[paragraph.runs.length - 1].memoId = memoId;
    }

    this._isDirty = true;
    return { id: memoId };
  }

  deleteMemo(memoId: string): boolean {
    let found = false;

    for (const section of this._content.sections) {
      if (section.memos) {
        const index = section.memos.findIndex(m => m.id === memoId);
        if (index !== -1) {
          this.saveState();
          section.memos.splice(index, 1);
          found = true;
          break;
        }
      }
    }

    if (found) {
      // Remove memo reference from paragraphs
      for (const section of this._content.sections) {
        for (const element of section.elements) {
          if (element.type === 'paragraph') {
            for (const run of element.data.runs) {
              if (run.memoId === memoId) {
                run.hasMemo = false;
                run.memoId = undefined;
              }
            }
          }
        }
      }

      this._isDirty = true;
    }

    return found;
  }

  // ============================================================
  // Section Operations
  // ============================================================

  getSections(): { index: number; pageSettings: PageSettings }[] {
    return this._content.sections.map((section, index) => ({
      index,
      pageSettings: section.pageSettings || {
        width: 59528,
        height: 84188,
        marginTop: 4252,
        marginBottom: 4252,
        marginLeft: 4252,
        marginRight: 4252,
      },
    }));
  }

  insertSection(afterSectionIndex: number): number {
    this.saveState();

    const newSection: HwpxSection = {
      id: Math.random().toString(36).substring(2, 11),
      elements: [{
        type: 'paragraph',
        data: {
          id: Math.random().toString(36).substring(2, 11),
          runs: [{ text: '' }],
        },
      }],
      pageSettings: {
        width: 59528,
        height: 84188,
        marginTop: 4252,
        marginBottom: 4252,
        marginLeft: 4252,
        marginRight: 4252,
      },
    };

    const insertIndex = afterSectionIndex + 1;
    this._content.sections.splice(insertIndex, 0, newSection);

    this._isDirty = true;
    return insertIndex;
  }

  deleteSection(sectionIndex: number): boolean {
    if (sectionIndex < 0 || sectionIndex >= this._content.sections.length) return false;
    if (this._content.sections.length <= 1) return false; // Cannot delete the last section

    this.saveState();
    this._content.sections.splice(sectionIndex, 1);
    this._isDirty = true;
    return true;
  }

  // ============================================================
  // Style Operations
  // ============================================================

  getStyles(): { id: number; name: string; type: string }[] {
    if (!this._content.styles?.styles) return [];

    return Array.from(this._content.styles.styles.values()).map(style => ({
      id: style.id,
      name: style.name || '',
      type: style.type || 'Para',
    }));
  }

  getCharShapes(): CharShape[] {
    if (!this._content.styles?.charShapes) return [];
    return Array.from(this._content.styles.charShapes.values());
  }

  getParaShapes(): ParaShape[] {
    if (!this._content.styles?.paraShapes) return [];
    return Array.from(this._content.styles.paraShapes.values());
  }

  applyStyle(sectionIndex: number, paragraphIndex: number, styleId: number): boolean {
    const paragraph = this.findParagraphByPath(sectionIndex, paragraphIndex);
    if (!paragraph) return false;

    if (!this._content.styles?.styles) return false;
    const style = this._content.styles.styles.get(styleId);
    if (!style) return false;

    this.saveState();

    paragraph.style = styleId;

    // Apply paragraph shape if defined
    if (style.paraPrIdRef !== undefined && this._content.styles.paraShapes) {
      const paraShape = this._content.styles.paraShapes.get(style.paraPrIdRef);
      if (paraShape) {
        paragraph.paraStyle = {
          align: paraShape.align?.toLowerCase() as ParagraphStyle['align'],
          lineSpacing: paraShape.lineSpacing,
          marginTop: paraShape.marginTop,
          marginBottom: paraShape.marginBottom,
          marginLeft: paraShape.marginLeft,
          marginRight: paraShape.marginRight,
          firstLineIndent: paraShape.firstLineIndent,
        };
      }
    }

    // Apply character shape if defined
    if (style.charPrIdRef !== undefined && this._content.styles.charShapes) {
      const charShape = this._content.styles.charShapes.get(style.charPrIdRef);
      if (charShape) {
        for (const run of paragraph.runs) {
          run.charStyle = {
            bold: charShape.bold,
            italic: charShape.italic,
            underline: charShape.underline,
            fontSize: charShape.height ? charShape.height / 100 : undefined,
            fontColor: charShape.textColor,
          };
        }
      }
    }

    this._isDirty = true;
    return true;
  }

  // ============================================================
  // Column Definition Operations
  // ============================================================

  getColumnDef(sectionIndex: number): ColumnDef | null {
    const section = this._content.sections[sectionIndex];
    if (!section) return null;
    return section.columnDef || null;
  }

  setColumnDef(sectionIndex: number, columns: number, gap?: number): boolean {
    const section = this._content.sections[sectionIndex];
    if (!section) return false;
    if (columns < 1) return false;

    this.saveState();

    const defaultGap = gap || 850; // Default gap in hwpunit (approx 8.5mm)

    section.columnDef = {
      type: 'Newspaper',
      count: columns,
      sameSize: true,
      sameGap: defaultGap,
      columns: Array.from({ length: columns }, () => ({
        width: 0, // Will be calculated based on page width
        gap: defaultGap,
      })),
    };

    this._isDirty = true;
    return true;
  }

  // ============================================================
  // Save
  // ============================================================

  async save(): Promise<Buffer> {
    if (!this._zip) throw new Error('Cannot save HWP files');
    await this.syncContentToZip();
    return await this._zip.generateAsync({ type: 'nodebuffer' });
  }

  private async syncContentToZip(): Promise<void> {
    if (!this._zip) return;

    // Apply table cell updates first (preserves original XML structure)
    if (this._pendingTableCellUpdates && this._pendingTableCellUpdates.length > 0) {
      await this.applyTableCellUpdatesToXml();
      this._pendingTableCellUpdates = [];
    }

    // Apply direct text updates (from updateParagraphText)
    if (this._pendingDirectTextUpdates && this._pendingDirectTextUpdates.length > 0) {
      await this.applyDirectTextUpdatesToXml();
      this._pendingDirectTextUpdates = [];
    }

    // Apply text replacements (from replaceText)
    if (this._pendingTextReplacements && this._pendingTextReplacements.length > 0) {
      await this.applyTextReplacementsToXml();
      this._pendingTextReplacements = [];
    }

    // NOTE: Do NOT call syncCharShapesToZip() here.
    // The current serialization is incomplete and loses critical attributes
    // (textColor, shadeColor, symMark, underline, strikeout, outline, shadow).
    // Original header.xml charPr/charShape elements should be preserved as-is.
    // Only sync charShapes when they are explicitly modified.

    // Sync metadata
    await this.syncMetadataToZip();

    this._isDirty = false;
  }

  /**
   * Apply table cell updates to XML while preserving original structure.
   * This function modifies only the text content of specific cells,
   * keeping all other XML elements, attributes, and structure intact.
   *
   * Safety features:
   * - Backs up original XML before modification
   * - Validates XML structure after changes
   * - Reverts to original if validation fails
   */
  private async applyTableCellUpdatesToXml(): Promise<void> {
    if (!this._zip) return;

    // Group updates by section for efficiency
    const updatesBySection = new Map<number, Array<{ tableId: string; row: number; col: number; text: string; charShapeId?: number }>>();
    for (const update of this._pendingTableCellUpdates) {
      const sectionUpdates = updatesBySection.get(update.sectionIndex) || [];
      sectionUpdates.push({ tableId: update.tableId, row: update.row, col: update.col, text: update.text, charShapeId: update.charShapeId });
      updatesBySection.set(update.sectionIndex, sectionUpdates);
    }

    // Process each section that has updates
    for (const [sectionIndex, updates] of updatesBySection) {
      const sectionPath = `Contents/section${sectionIndex}.xml`;
      const file = this._zip.file(sectionPath);
      if (!file) continue;

      // Backup original XML for safety
      const originalXml = await file.async('string');
      let xml = originalXml;

      // Group updates by table ID
      const updatesByTableId = new Map<string, Array<{ row: number; col: number; text: string; charShapeId?: number }>>();
      for (const update of updates) {
        const tableUpdates = updatesByTableId.get(update.tableId) || [];
        tableUpdates.push({ row: update.row, col: update.col, text: update.text, charShapeId: update.charShapeId });
        updatesByTableId.set(update.tableId, tableUpdates);
      }

      // Process each table that has updates (by ID)
      for (const [tableId, tableUpdates] of updatesByTableId) {
        // Find the table by ID in XML
        const tableMatch = this.findTableById(xml, tableId);
        if (!tableMatch) continue;

        // Safety check: validate indices before substring operations
        if (tableMatch.startIndex < 0 || tableMatch.endIndex > xml.length || tableMatch.startIndex > tableMatch.endIndex) {
          console.warn(`[HwpxDocument] Invalid table indices for table ${tableId}, skipping update`);
          continue;
        }

        // Update the table XML with cell changes
        const updatedTableXml = this.updateTableCellsInXml(tableMatch.xml, tableUpdates);

        // Safety check: ensure updated XML is not empty or drastically smaller
        if (!updatedTableXml || updatedTableXml.length < tableMatch.xml.length * 0.5) {
          console.warn(`[HwpxDocument] Suspicious table update result for table ${tableId}, skipping`);
          continue;
        }

        // Replace the old table XML with the updated one
        xml = xml.substring(0, tableMatch.startIndex) + updatedTableXml + xml.substring(tableMatch.endIndex);
      }

      // Validate modified XML before saving
      if (!this.validateXmlStructure(xml)) {
        console.error(`[HwpxDocument] XML validation failed for ${sectionPath}, reverting to original`);
        // Revert to original - don't save corrupted XML
        continue;
      }

      this._zip.file(sectionPath, xml);
    }
  }

  /**
   * Basic XML structure validation.
   * Checks for common corruption indicators.
   * Note: This is intentionally lenient to avoid false positives.
   */
  private validateXmlStructure(xml: string): boolean {
    // Must start with XML declaration or root element
    if (!xml.trim().startsWith('<?xml') && !xml.trim().startsWith('<')) {
      console.warn(`[HwpxDocument] XML does not start with declaration or element`);
      return false;
    }

    // Check for truncated XML (ends with incomplete tag)
    if (xml.match(/<[^>]*$/)) {
      console.warn(`[HwpxDocument] XML appears truncated (incomplete tag at end)`);
      return false;
    }

    // Check for broken opening tags (< followed by another < without >)
    // This catches cases like: <tag<broken>
    if (xml.match(/<[^>]*</)) {
      console.warn(`[HwpxDocument] Broken opening tag detected`);
      return false;
    }

    // Check for empty or near-empty content (likely corruption)
    if (xml.trim().length < 100) {
      console.warn(`[HwpxDocument] Suspiciously short XML content`);
      return false;
    }

    return true;
  }

  /**
   * Find a table by its ID in XML.
   */
  private findTableById(xml: string, tableId: string): { xml: string; startIndex: number; endIndex: number } | null {
    // Match table with specific ID
    const tableStartRegex = new RegExp(`<(?:hp|hs|hc):tbl[^>]*\\bid="${tableId}"[^>]*>`, 'g');
    const match = tableStartRegex.exec(xml);

    if (!match) {
      // Try alternate ID format (id='...' instead of id="...")
      const altRegex = new RegExp(`<(?:hp|hs|hc):tbl[^>]*\\bid='${tableId}'[^>]*>`, 'g');
      const altMatch = altRegex.exec(xml);
      if (!altMatch) return null;
      return this.extractTableFromMatch(xml, altMatch);
    }

    return this.extractTableFromMatch(xml, match);
  }

  /**
   * Extract complete table XML from a regex match.
   */
  private extractTableFromMatch(xml: string, match: RegExpExecArray): { xml: string; startIndex: number; endIndex: number } | null {
    const startIndex = match.index;
    const prefix = match[0].match(/<(hp|hs|hc):tbl/)?.[1] || 'hp';

    // Find the matching closing tag
    const endTag = `</${prefix}:tbl>`;
    let depth = 1;
    let pos = match.index + match[0].length;

    while (depth > 0 && pos < xml.length) {
      const nextOpen = xml.indexOf(`<${prefix}:tbl`, pos);
      const nextClose = xml.indexOf(endTag, pos);

      if (nextClose === -1) return null;

      if (nextOpen !== -1 && nextOpen < nextClose) {
        depth++;
        pos = nextOpen + 1;
      } else {
        depth--;
        if (depth === 0) {
          const endIndex = nextClose + endTag.length;
          return {
            xml: xml.substring(startIndex, endIndex),
            startIndex,
            endIndex
          };
        }
        pos = nextClose + 1;
      }
    }

    return null;
  }

  /**
   * Find all tables in XML and return their positions and content.
   */
  private findAllTables(xml: string): Array<{ xml: string; startIndex: number; endIndex: number }> {
    const tables: Array<{ xml: string; startIndex: number; endIndex: number }> = [];

    // Match both hp:tbl and hs:tbl (different namespace prefixes)
    const tableStartRegex = /<(?:hp|hs|hc):tbl[^>]*>/g;
    let match;

    while ((match = tableStartRegex.exec(xml)) !== null) {
      const startIndex = match.index;
      const prefix = match[0].match(/<(hp|hs|hc):tbl/)?.[1] || 'hp';

      // Find the matching closing tag
      const endTag = `</${prefix}:tbl>`;
      let depth = 1;
      let pos = match.index + match[0].length;

      while (depth > 0 && pos < xml.length) {
        const nextOpen = xml.indexOf(`<${prefix}:tbl`, pos);
        const nextClose = xml.indexOf(endTag, pos);

        if (nextClose === -1) break;

        if (nextOpen !== -1 && nextOpen < nextClose) {
          depth++;
          pos = nextOpen + 1;
        } else {
          depth--;
          if (depth === 0) {
            const endIndex = nextClose + endTag.length;
            tables.push({
              xml: xml.substring(startIndex, endIndex),
              startIndex,
              endIndex
            });
          }
          pos = nextClose + 1;
        }
      }
    }

    return tables;
  }

  /**
   * Find all elements of a given type using depth tracking.
   * This correctly handles nested elements (e.g., nested tables).
   * @param xml The XML string to search in
   * @param elementName The element name without namespace prefix (e.g., 'tr', 'tc')
   * @returns Array of elements with their positions
   */
  private findAllElementsWithDepth(xml: string, elementName: string): Array<{ xml: string; startIndex: number; endIndex: number }> {
    const elements: Array<{ xml: string; startIndex: number; endIndex: number }> = [];

    // Match all namespace prefixes (hp, hs, hc)
    const startPattern = new RegExp(`<(hp|hs|hc):${elementName}[^>]*>`, 'g');
    let match;

    while ((match = startPattern.exec(xml)) !== null) {
      const startIndex = match.index;
      const prefix = match[1];
      const openTag = `<${prefix}:${elementName}`;
      const closeTag = `</${prefix}:${elementName}>`;

      let depth = 1;
      let pos = match.index + match[0].length;

      while (depth > 0 && pos < xml.length) {
        const nextOpen = xml.indexOf(openTag, pos);
        const nextClose = xml.indexOf(closeTag, pos);

        if (nextClose === -1) break;

        if (nextOpen !== -1 && nextOpen < nextClose) {
          // Found another opening tag - go deeper
          depth++;
          pos = nextOpen + 1;
        } else {
          // Found closing tag
          depth--;
          if (depth === 0) {
            const endIndex = nextClose + closeTag.length;
            elements.push({
              xml: xml.substring(startIndex, endIndex),
              startIndex,
              endIndex
            });
            // Update regex lastIndex to continue after this element
            startPattern.lastIndex = endIndex;
          }
          pos = nextClose + 1;
        }
      }
    }

    return elements;
  }

  /**
   * Update specific cells in a table XML string.
   */
  private updateTableCellsInXml(tableXml: string, updates: Array<{ row: number; col: number; text: string; charShapeId?: number }>): string {
    let result = tableXml;

    // Find all rows using depth tracking to handle nested tables correctly
    const rows = this.findAllElementsWithDepth(tableXml, 'tr');

    // Sort updates by row (descending) to process from end to start (avoid index shifting)
    const sortedUpdates = [...updates].sort((a, b) => b.row - a.row || b.col - a.col);

    for (const update of sortedUpdates) {
      if (update.row >= rows.length) continue;

      const rowData = rows[update.row];
      const updatedRowXml = this.updateCellInRow(rowData.xml, update.col, update.text, update.charShapeId);

      result = result.substring(0, rowData.startIndex) + updatedRowXml + result.substring(rowData.endIndex);

      // Update row positions for subsequent rows (those with lower indices)
      const lengthDiff = updatedRowXml.length - rowData.xml.length;
      for (let i = update.row - 1; i >= 0; i--) {
        // Earlier rows (lower indices) are before in the string, so they don't need adjustment
        // But we need to update the row data for current processing
      }
    }

    return result;
  }

  /**
   * Update a specific cell in a row XML string.
   */
  private updateCellInRow(rowXml: string, colIndex: number, newText: string, charShapeId?: number): string {
    // Find all cells in this row using depth tracking to handle nested tables correctly
    const cells = this.findAllElementsWithDepth(rowXml, 'tc');

    if (colIndex >= cells.length) return rowXml;

    const cellData = cells[colIndex];
    const updatedCellXml = this.updateTextInCell(cellData.xml, newText, charShapeId);

    return rowXml.substring(0, cellData.startIndex) + updatedCellXml + rowXml.substring(cellData.endIndex);
  }

  /**
   * Reset lineseg values to default so Hancom Word recalculates line layout.
   * When text content changes, the old lineseg values (horzsize, textpos, etc.)
   * no longer match the new text, causing rendering issues like overlapping text.
   * By resetting to default values, Hancom Word will recalculate proper line breaks.
   */
  private resetLinesegInXml(xml: string): string {
    // Find all linesegarray elements and reset their lineseg children to default values
    // Default lineseg: single line with minimal values - Hancom Word will recalculate
    const defaultLineseg = '<hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850" spacing="600" horzpos="0" horzsize="0" flags="0"/>';

    // Replace linesegarray content with default single lineseg
    // This pattern matches <hp:linesegarray>...</hp:linesegarray> and replaces content
    const linesegArrayPattern = /(<(?:hp|hs|hc):linesegarray[^>]*>)[\s\S]*?(<\/(?:hp|hs|hc):linesegarray>)/g;

    return xml.replace(linesegArrayPattern, (match, openTag, closeTag) => {
      // Determine the prefix (hp, hs, or hc) from the opening tag
      const prefixMatch = openTag.match(/<(hp|hs|hc):linesegarray/);
      const prefix = prefixMatch ? prefixMatch[1] : 'hp';
      const defaultSeg = `<${prefix}:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850" spacing="600" horzpos="0" horzsize="0" flags="0"/>`;
      return openTag + defaultSeg + closeTag;
    });
  }

  /**
   * Update text content in a cell XML string.
   * Handles both existing text replacement and empty cell population.
   * If charShapeId is provided, overrides the charPrIDRef attribute.
   */
  private updateTextInCell(cellXml: string, newText: string, charShapeId?: number): string {
    const escapedText = this.escapeXml(newText);
    let xml = cellXml;

    // If charShapeId is provided, update charPrIDRef in the first run tag
    if (charShapeId !== undefined) {
      xml = xml.replace(
        /(<(?:hp|hs|hc):run\s+)charPrIDRef="[^"]*"/,
        `$1charPrIDRef="${charShapeId}"`
      );
    }

    // Pattern 1: Cell has existing <hp:t> or <hs:t> or <hc:t> tags with content
    const tTagPattern = /(<(?:hp|hs|hc):t[^>]*>)([^<]*)(<\/(?:hp|hs|hc):t>)/g;
    let foundText = false;
    let result = xml.replace(tTagPattern, (match, openTag, _oldText, closeTag, offset) => {
      // Only replace the first text occurrence
      if (!foundText) {
        foundText = true;
        return openTag + escapedText + closeTag;
      }
      return match;
    });

    if (foundText) return this.resetLinesegInXml(result);

    // Pattern 2: Cell has empty <hp:t/> or <hp:t></hp:t> tags
    const emptyTTagPattern = /<((?:hp|hs|hc):t)([^>]*)\s*\/>/;
    const emptyTMatch = xml.match(emptyTTagPattern);
    if (emptyTMatch) {
      const updated = xml.replace(emptyTTagPattern, `<${emptyTMatch[1]}${emptyTMatch[2]}>${escapedText}</${emptyTMatch[1]}>`);
      return this.resetLinesegInXml(updated);
    }

    // Pattern 3a: Self-closing <hp:run .../> - expand to full run with text
    const selfClosingRunPattern = /<((?:hp|hs|hc):run)([^>]*)\s*\/>/;
    const selfClosingRunMatch = xml.match(selfClosingRunPattern);
    if (selfClosingRunMatch) {
      const tagName = selfClosingRunMatch[1]; // e.g., "hp:run"
      let attrs = selfClosingRunMatch[2];
      // If charShapeId is provided, update or add charPrIDRef
      if (charShapeId !== undefined) {
        if (attrs.includes('charPrIDRef=')) {
          attrs = attrs.replace(/charPrIDRef="[^"]*"/, `charPrIDRef="${charShapeId}"`);
        } else {
          attrs = ` charPrIDRef="${charShapeId}"` + attrs;
        }
      }
      const prefix = tagName.split(':')[0]; // e.g., "hp"
      const updated = xml.replace(selfClosingRunPattern, `<${tagName}${attrs}><${prefix}:t>${escapedText}</${prefix}:t></${tagName}>`);
      return this.resetLinesegInXml(updated);
    }

    // Pattern 3b: Cell has <hp:run> but no <hp:t> - add text inside run
    const runPattern = /(<(?:hp|hs|hc):run[^>]*>)([\s\S]*?)(<\/(?:hp|hs|hc):run>)/;
    const runMatch = xml.match(runPattern);
    if (runMatch) {
      const prefix = runMatch[1].match(/<(hp|hs|hc):run/)?.[1] || 'hp';
      const newRunContent = runMatch[2] + `<${prefix}:t>${escapedText}</${prefix}:t>`;
      const updated = xml.replace(runPattern, runMatch[1] + newRunContent + runMatch[3]);
      return this.resetLinesegInXml(updated);
    }

    // Pattern 4: Cell has <hp:subList><hp:p> structure - find the paragraph and add text
    const subListPattern = /(<(?:hp|hs|hc):subList[^>]*>[\s\S]*?<(?:hp|hs|hc):p[^>]*>)([\s\S]*?)(<\/(?:hp|hs|hc):p>)/;
    const subListMatch = xml.match(subListPattern);
    if (subListMatch) {
      const prefix = subListMatch[1].match(/<(hp|hs|hc):subList/)?.[1] || 'hp';
      // Check if there's already a run
      if (!subListMatch[2].includes(':run')) {
        const charAttr = charShapeId !== undefined ? ` charPrIDRef="${charShapeId}"` : '';
        const newContent = subListMatch[2] + `<${prefix}:run${charAttr}><${prefix}:t>${escapedText}</${prefix}:t></${prefix}:run>`;
        const updated = xml.replace(subListPattern, subListMatch[1] + newContent + subListMatch[3]);
        return this.resetLinesegInXml(updated);
      }
    }

    // Pattern 5: Cell has only <hp:p> without subList
    const pPattern = /(<(?:hp|hs|hc):p[^>]*>)([\s\S]*?)(<\/(?:hp|hs|hc):p>)/;
    const pMatch = xml.match(pPattern);
    if (pMatch) {
      const prefix = pMatch[1].match(/<(hp|hs|hc):p/)?.[1] || 'hp';
      if (!pMatch[2].includes(':run') && !pMatch[2].includes(':t>')) {
        const charAttr = charShapeId !== undefined ? ` charPrIDRef="${charShapeId}"` : '';
        const newContent = pMatch[2] + `<${prefix}:run${charAttr}><${prefix}:t>${escapedText}</${prefix}:t></${prefix}:run>`;
        const updated = xml.replace(pPattern, pMatch[1] + newContent + pMatch[3]);
        return this.resetLinesegInXml(updated);
      }
    }

    // Fallback: return unchanged (shouldn't happen in well-formed HWPX)
    return xml;
  }

  /**
   * Apply direct text updates (exact match replacement)
   */
  private async applyDirectTextUpdatesToXml(): Promise<void> {
    if (!this._zip) return;

    let sectionIndex = 0;
    while (true) {
      const sectionPath = `Contents/section${sectionIndex}.xml`;
      const file = this._zip.file(sectionPath);
      if (!file) break;

      let xml = await file.async('string');

      for (const update of this._pendingDirectTextUpdates) {
        const escapedOld = this.escapeXml(update.oldText);
        const escapedNew = this.escapeXml(update.newText);

        // Replace text anywhere within <hp:t> tags (may contain other tags like <hp:tab/>)
        // First try exact match at the start of <hp:t> content
        const pattern1 = new RegExp(`(<hp:t[^>]*>)${this.escapeRegex(escapedOld)}`, 'g');
        xml = xml.replace(pattern1, `$1${escapedNew}`);

        // Also try simple text replacement for cases where text is standalone
        xml = xml.replace(new RegExp(`>${this.escapeRegex(escapedOld)}<`, 'g'), `>${escapedNew}<`);
      }

      this._zip.file(sectionPath, xml);
      sectionIndex++;
    }
  }

  private escapeRegex(str: string): string {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  /**
   * Apply text replacements directly to XML files.
   * This is the safest approach as it preserves the original XML structure.
   */
  private async applyTextReplacementsToXml(): Promise<void> {
    if (!this._zip) return;

    // Get all section files
    let sectionIndex = 0;
    while (true) {
      const sectionPath = `Contents/section${sectionIndex}.xml`;
      const file = this._zip.file(sectionPath);
      if (!file) break;

      let xml = await file.async('string');

      // Apply each pending replacement to the XML
      for (const replacement of this._pendingTextReplacements) {
        const { oldText, newText, options } = replacement;
        const { caseSensitive = false, regex = false, replaceAll = true } = options;

        // Create pattern for matching text inside <hp:t> tags
        let searchPattern: RegExp;
        if (regex) {
          searchPattern = new RegExp(oldText, caseSensitive ? (replaceAll ? 'g' : '') : (replaceAll ? 'gi' : 'i'));
        } else {
          const escaped = oldText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
          searchPattern = new RegExp(escaped, caseSensitive ? (replaceAll ? 'g' : '') : (replaceAll ? 'gi' : 'i'));
        }

        // Replace text within <hp:t> tags while preserving XML structure
        xml = xml.replace(/<hp:t([^>]*)>([^<]*)<\/hp:t>/g, (match, attrs, textContent) => {
          const newTextContent = textContent.replace(searchPattern, this.escapeXml(newText));
          return `<hp:t${attrs}>${newTextContent}</hp:t>`;
        });
      }

      this._zip.file(sectionPath, xml);
      sectionIndex++;
    }

    // Update metadata in header.xml if needed
    await this.syncMetadataToZip();
  }

  /**
   * Sync structural changes (paragraph text, table cells, etc.)
   * Regenerates section XML from _content to handle new elements.
   */
  private async syncStructuralChangesToZip(): Promise<void> {
    if (!this._zip) return;

    // Regenerate each section XML from content
    for (let sectionIndex = 0; sectionIndex < this._content.sections.length; sectionIndex++) {
      const sectionPath = `Contents/section${sectionIndex}.xml`;
      const section = this._content.sections[sectionIndex];
      const newXml = this.generateSectionXml(section);
      this._zip.file(sectionPath, newXml);
    }

    // Sync metadata
    await this.syncMetadataToZip();
  }

  /**
   * Generate complete section XML from HwpxSection content.
   */
  private generateSectionXml(section: HwpxSection): string {
    let xml = `<?xml version="1.0" encoding="UTF-8"?>\n`;
    xml += `<hp:sec xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">\n`;

    for (const element of section.elements) {
      if (element.type === 'paragraph') {
        xml += this.generateParagraphXml(element.data as HwpxParagraph);
      } else if (element.type === 'table') {
        xml += this.generateTableXml(element.data as HwpxTable);
      }
    }

    xml += `</hp:sec>`;
    return xml;
  }

  /**
   * Generate paragraph XML from HwpxParagraph.
   */
  private generateParagraphXml(paragraph: HwpxParagraph): string {
    let xml = `  <hp:p>\n`;
    for (const run of paragraph.runs) {
      xml += `    <hp:run>\n`;
      xml += `      <hp:t>${this.escapeXml(run.text)}</hp:t>\n`;
      xml += `    </hp:run>\n`;
    }
    xml += `  </hp:p>\n`;
    return xml;
  }

  /**
   * Generate table XML from HwpxTable.
   */
  private generateTableXml(table: HwpxTable): string {
    let xml = `  <hp:tbl rowCount="${table.rowCount}" colCount="${table.colCount}">\n`;

    for (const row of table.rows) {
      xml += `    <hp:tr>\n`;
      for (const cell of row.cells) {
        xml += `      <hp:tc colAddr="${cell.colAddr}" rowAddr="${cell.rowAddr}" colSpan="${cell.colSpan}" rowSpan="${cell.rowSpan}">\n`;
        for (const para of cell.paragraphs) {
          xml += `        <hp:p>\n`;
          for (const run of para.runs) {
            xml += `          <hp:run>\n`;
            xml += `            <hp:t>${this.escapeXml(run.text)}</hp:t>\n`;
            xml += `          </hp:run>\n`;
          }
          xml += `        </hp:p>\n`;
        }
        xml += `      </hp:tc>\n`;
      }
      xml += `    </hp:tr>\n`;
    }

    xml += `  </hp:tbl>\n`;
    return xml;
  }

  /**
   * Update section XML with current content.
   * Handles paragraphs and table cells.
   */
  private updateSectionXml(xml: string, section: HwpxSection): string {
    let updatedXml = xml;

    // Build a map of element index to paragraph data for quick lookup
    const paragraphMap = new Map<number, HwpxParagraph>();
    const tableMap = new Map<number, HwpxTable>();

    let paragraphCount = 0;
    let tableCount = 0;

    for (const element of section.elements) {
      if (element.type === 'paragraph') {
        paragraphMap.set(paragraphCount, element.data as HwpxParagraph);
        paragraphCount++;
      } else if (element.type === 'table') {
        tableMap.set(tableCount, element.data as HwpxTable);
        tableCount++;
      }
    }

    // Track table positions to skip paragraphs inside tables
    const tablePositions: Array<{ start: number; end: number }> = [];
    const tableRegex = /<hp:tbl\b[^>]*>[\s\S]*?<\/hp:tbl>/g;
    let tableMatch;
    while ((tableMatch = tableRegex.exec(xml)) !== null) {
      tablePositions.push({ start: tableMatch.index, end: tableMatch.index + tableMatch[0].length });
    }

    // Update paragraphs outside of tables
    let paragraphIndex = 0;
    const paragraphRegex = /<hp:p\b[^>]*>([\s\S]*?)<\/hp:p>/g;
    updatedXml = xml.replace(paragraphRegex, (match, _inner, offset) => {
      // Check if this paragraph is inside a table
      const isInTable = tablePositions.some(pos => offset >= pos.start && offset < pos.end);

      if (isInTable) {
        return match; // Don't modify paragraphs inside tables here
      }

      const paragraph = paragraphMap.get(paragraphIndex);
      paragraphIndex++;

      if (paragraph) {
        return this.updateParagraphXml(match, paragraph);
      }
      return match;
    });

    // Update table cells
    let tableIndex = 0;
    updatedXml = updatedXml.replace(/<hp:tbl\b[^>]*>([\s\S]*?)<\/hp:tbl>/g, (tblMatch) => {
      const table = tableMap.get(tableIndex);
      tableIndex++;

      if (!table) {
        return tblMatch;
      }

      let rowIndex = 0;
      return tblMatch.replace(/<hp:tr[^>]*>([\s\S]*?)<\/hp:tr>/g, (rowMatch) => {
        if (rowIndex >= table.rows.length) {
          rowIndex++;
          return rowMatch;
        }

        const row = table.rows[rowIndex];
        rowIndex++;

        let cellIndex = 0;
        return rowMatch.replace(/<hp:tc\b([^>]*)>([\s\S]*?)<\/hp:tc>/g, (cellMatch, cellAttrs, cellContent) => {
          if (cellIndex >= row.cells.length) {
            cellIndex++;
            return cellMatch;
          }

          const cell = row.cells[cellIndex];
          cellIndex++;

          // Update cell content - replace text in paragraphs
          let updatedCellContent = cellContent;
          if (cell.paragraphs && cell.paragraphs.length > 0) {
            let cellParaIndex = 0;
            updatedCellContent = cellContent.replace(/<hp:p\b[^>]*>([\s\S]*?)<\/hp:p>/g, (paraMatch: string) => {
              if (cellParaIndex < cell.paragraphs.length) {
                const para = cell.paragraphs[cellParaIndex];
                cellParaIndex++;
                return this.updateParagraphXml(paraMatch, para);
              }
              cellParaIndex++;
              return paraMatch;
            });
          }

          return `<hp:tc${cellAttrs}>${updatedCellContent}</hp:tc>`;
        });
      });
    });

    return updatedXml;
  }

  /**
   * Update paragraph XML with new text content.
   */
  private updateParagraphXml(xml: string, paragraph: HwpxParagraph): string {
    const fullText = paragraph.runs.map(r => r.text).join('');

    // Update all <hp:t> tags with the combined text
    // For simplicity, put all text in the first <hp:t> tag and empty the rest
    let firstTextTag = true;
    return xml.replace(/<hp:t([^>]*)>([^<]*)<\/hp:t>/g, (_match, attrs, _oldText) => {
      if (firstTextTag) {
        firstTextTag = false;
        return `<hp:t${attrs}>${this.escapeXml(fullText)}</hp:t>`;
      }
      // Empty subsequent text tags
      return `<hp:t${attrs}></hp:t>`;
    });
  }

  /**
   * Serialize a CharShape object to XML string.
   * This preserves all character style properties including spacing (자간).
   */
  private serializeCharShape(charShape: CharShape): string {
    // Use original tag name (charPr or charShape), default to charPr for compatibility
    const tagName = charShape.tagName || 'charPr';
    let xml = `<hh:${tagName} id="${charShape.id}"`;

    // Basic attributes - height takes precedence over fontSize (height is the raw value)
    if (charShape.height !== undefined) {
      xml += ` height="${charShape.height}"`;
    } else if (charShape.fontSize !== undefined) {
      xml += ` height="${Math.round(charShape.fontSize * 100)}"`;
    }
    if (charShape.color) xml += ` textColor="${charShape.color}"`;
    if (charShape.backgroundColor) xml += ` shadeColor="${charShape.backgroundColor}"`;
    if (charShape.useFontSpace !== undefined) xml += ` useFontSpace="${charShape.useFontSpace ? '1' : '0'}"`;
    if (charShape.useKerning !== undefined) xml += ` useKerning="${charShape.useKerning ? '1' : '0'}"`;
    if (charShape.borderFillId !== undefined) xml += ` borderFillIDRef="${charShape.borderFillId}"`;

    xml += `>`;

    // Font references
    if (charShape.fontRefs) {
      xml += `<hh:fontRef`;
      if (charShape.fontRefs.hangul !== undefined) xml += ` hangul="${charShape.fontRefs.hangul}"`;
      if (charShape.fontRefs.latin !== undefined) xml += ` latin="${charShape.fontRefs.latin}"`;
      if (charShape.fontRefs.hanja !== undefined) xml += ` hanja="${charShape.fontRefs.hanja}"`;
      if (charShape.fontRefs.japanese !== undefined) xml += ` japanese="${charShape.fontRefs.japanese}"`;
      if (charShape.fontRefs.other !== undefined) xml += ` other="${charShape.fontRefs.other}"`;
      if (charShape.fontRefs.symbol !== undefined) xml += ` symbol="${charShape.fontRefs.symbol}"`;
      if (charShape.fontRefs.user !== undefined) xml += ` user="${charShape.fontRefs.user}"`;
      xml += `/>`;
    }

    // Ratio (장평)
    if (charShape.ratio) {
      xml += `<hh:ratio`;
      if (charShape.ratio.hangul !== undefined) xml += ` hangul="${charShape.ratio.hangul}"`;
      if (charShape.ratio.latin !== undefined) xml += ` latin="${charShape.ratio.latin}"`;
      if (charShape.ratio.hanja !== undefined) xml += ` hanja="${charShape.ratio.hanja}"`;
      if (charShape.ratio.japanese !== undefined) xml += ` japanese="${charShape.ratio.japanese}"`;
      if (charShape.ratio.other !== undefined) xml += ` other="${charShape.ratio.other}"`;
      if (charShape.ratio.symbol !== undefined) xml += ` symbol="${charShape.ratio.symbol}"`;
      if (charShape.ratio.user !== undefined) xml += ` user="${charShape.ratio.user}"`;
      xml += `/>`;
    }

    // Spacing (자간) - CRITICAL for character spacing!
    if (charShape.charSpacing) {
      xml += `<hh:spacing`;
      if (charShape.charSpacing.hangul !== undefined) xml += ` hangul="${charShape.charSpacing.hangul}"`;
      if (charShape.charSpacing.latin !== undefined) xml += ` latin="${charShape.charSpacing.latin}"`;
      if (charShape.charSpacing.hanja !== undefined) xml += ` hanja="${charShape.charSpacing.hanja}"`;
      if (charShape.charSpacing.japanese !== undefined) xml += ` japanese="${charShape.charSpacing.japanese}"`;
      if (charShape.charSpacing.other !== undefined) xml += ` other="${charShape.charSpacing.other}"`;
      if (charShape.charSpacing.symbol !== undefined) xml += ` symbol="${charShape.charSpacing.symbol}"`;
      if (charShape.charSpacing.user !== undefined) xml += ` user="${charShape.charSpacing.user}"`;
      xml += `/>`;
    }

    // Relative size (상대크기)
    if (charShape.relSize) {
      xml += `<hh:relSz`;
      if (charShape.relSize.hangul !== undefined) xml += ` hangul="${charShape.relSize.hangul}"`;
      if (charShape.relSize.latin !== undefined) xml += ` latin="${charShape.relSize.latin}"`;
      if (charShape.relSize.hanja !== undefined) xml += ` hanja="${charShape.relSize.hanja}"`;
      if (charShape.relSize.japanese !== undefined) xml += ` japanese="${charShape.relSize.japanese}"`;
      if (charShape.relSize.other !== undefined) xml += ` other="${charShape.relSize.other}"`;
      if (charShape.relSize.symbol !== undefined) xml += ` symbol="${charShape.relSize.symbol}"`;
      if (charShape.relSize.user !== undefined) xml += ` user="${charShape.relSize.user}"`;
      xml += `/>`;
    }

    // Char offset (글자위치)
    if (charShape.charOffset) {
      xml += `<hh:offset`;
      if (charShape.charOffset.hangul !== undefined) xml += ` hangul="${charShape.charOffset.hangul}"`;
      if (charShape.charOffset.latin !== undefined) xml += ` latin="${charShape.charOffset.latin}"`;
      if (charShape.charOffset.hanja !== undefined) xml += ` hanja="${charShape.charOffset.hanja}"`;
      if (charShape.charOffset.japanese !== undefined) xml += ` japanese="${charShape.charOffset.japanese}"`;
      if (charShape.charOffset.other !== undefined) xml += ` other="${charShape.charOffset.other}"`;
      if (charShape.charOffset.symbol !== undefined) xml += ` symbol="${charShape.charOffset.symbol}"`;
      if (charShape.charOffset.user !== undefined) xml += ` user="${charShape.charOffset.user}"`;
      xml += `/>`;
    }

    // Bold/Italic
    if (charShape.bold) xml += `<hh:bold/>`;
    if (charShape.italic) xml += `<hh:italic/>`;

    // Underline
    if (charShape.underline && typeof charShape.underline === 'object') {
      xml += `<hh:underline type="${charShape.underline.type.toUpperCase()}" shape="${charShape.underline.shape.toUpperCase()}" color="${charShape.underline.color}"/>`;
    }

    // Strikeout
    if (charShape.strikeout && typeof charShape.strikeout === 'object') {
      xml += `<hh:strikeout type="${charShape.strikeout.type.toUpperCase()}" shape="${charShape.strikeout.shape.toUpperCase()}" color="${charShape.strikeout.color}"/>`;
    }

    // Outline
    if (charShape.outline) {
      const outlineType = typeof charShape.outline === 'object' ? charShape.outline.type : charShape.outline;
      xml += `<hh:outline type="${outlineType.toUpperCase()}"/>`;
    }

    // Shadow
    if (charShape.shadow && typeof charShape.shadow === 'object' && charShape.shadow.type !== 'None') {
      xml += `<hh:shadow type="${charShape.shadow.type.toUpperCase()}"`;
      if (charShape.shadow.offsetX !== undefined) xml += ` offsetX="${Math.round(charShape.shadow.offsetX * 100)}"`;
      if (charShape.shadow.offsetY !== undefined) xml += ` offsetY="${Math.round(charShape.shadow.offsetY * 100)}"`;
      if (charShape.shadow.color) xml += ` color="${charShape.shadow.color}"`;
      xml += `/>`;
    }

    // Emboss/Engrave
    if (charShape.emboss) xml += `<hh:emboss/>`;
    if (charShape.engrave) xml += `<hh:engrave/>`;

    // SymMark (강조점)
    if (charShape.symMark && charShape.symMark !== 'None') {
      xml += `<hh:symMark symMarkType="${charShape.symMark.toUpperCase()}"/>`;
    }

    xml += `</hh:${tagName}>`;
    return xml;
  }

  /**
   * Sync charShapes from memory to header.xml.
   * This ensures character styles (including spacing) are preserved after save.
   */
  private async syncCharShapesToZip(): Promise<void> {
    if (!this._zip || !this._content.styles?.charShapes) return;

    const headerPath = 'Contents/header.xml';
    let headerXml = await this._zip.file(headerPath)?.async('string');
    if (!headerXml) return;

    // Debug: count charShapes before modification
    const originalCharShapeCount = (headerXml.match(/<hh:charShape/gi) || []).length;
    const originalCharPrCount = (headerXml.match(/<hh:charPr/gi) || []).length;
    console.log(`[HwpxDocument] syncCharShapesToZip: original charShape=${originalCharShapeCount}, charPr=${originalCharPrCount}`);

    // For each charShape in memory, update or preserve in XML
    for (const [id, charShape] of this._content.styles.charShapes) {
      const newXml = this.serializeCharShape(charShape);

      // Try to match existing charShape with this ID (supports both hh:charShape and hh:charPr)
      // Use non-greedy match with [\s\S]*? for content between tags
      const charShapePattern = new RegExp(
        `<hh:charShape[^>]*\\bid="${id}"[^>]*>[\\s\\S]*?</hh:charShape>`,
        'i'
      );
      const charPrPattern = new RegExp(
        `<hh:charPr[^>]*\\bid="${id}"[^>]*>[\\s\\S]*?</hh:charPr>`,
        'i'
      );

      if (headerXml.match(charShapePattern)) {
        headerXml = headerXml.replace(charShapePattern, newXml);
      } else if (headerXml.match(charPrPattern)) {
        headerXml = headerXml.replace(charPrPattern, newXml);
      }
      // If no match found, the charShape might be new - but we don't add new ones
      // to avoid corrupting the structure. Original charShapes are preserved.
    }

    // Debug: count charShapes/charPr after modification
    const newCharShapeCount = (headerXml.match(/<hh:charShape/gi) || []).length;
    const newCharPrCount = (headerXml.match(/<hh:charPr/gi) || []).length;
    console.log(`[HwpxDocument] syncCharShapesToZip: after update charShape=${newCharShapeCount}, charPr=${newCharPrCount}`);

    const totalBefore = originalCharShapeCount + originalCharPrCount;
    const totalAfter = newCharShapeCount + newCharPrCount;
    if (totalAfter < totalBefore) {
      console.warn(`[HwpxDocument] WARNING: charShape/charPr count decreased from ${totalBefore} to ${totalAfter}`);
    }

    this._zip.file(headerPath, headerXml);
  }

  /**
   * Sync metadata to header.xml
   */
  private async syncMetadataToZip(): Promise<void> {
    if (!this._zip) return;

    const headerPath = 'Contents/header.xml';
    let headerXml = await this._zip.file(headerPath)?.async('string');
    if (headerXml && this._content.metadata) {
      const meta = this._content.metadata;
      if (meta.title) {
        headerXml = headerXml.replace(/<hh:title[^>]*>[^<]*<\/hh:title>/,
          `<hh:title>${this.escapeXml(meta.title)}</hh:title>`);
      }
      if (meta.creator) {
        headerXml = headerXml.replace(/<hh:creator[^>]*>[^<]*<\/hh:creator>/,
          `<hh:creator>${this.escapeXml(meta.creator)}</hh:creator>`);
      }
      if (meta.subject) {
        headerXml = headerXml.replace(/<hh:subject[^>]*>[^<]*<\/hh:subject>/,
          `<hh:subject>${this.escapeXml(meta.subject)}</hh:subject>`);
      }
      if (meta.description) {
        headerXml = headerXml.replace(/<hh:description[^>]*>[^<]*<\/hh:description>/,
          `<hh:description>${this.escapeXml(meta.description)}</hh:description>`);
      }
      this._zip.file(headerPath, headerXml);
    }
  }

  private escapeXml(text: string): string {
    return text
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&apos;');
  }
}
