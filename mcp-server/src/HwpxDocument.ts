import JSZip from 'jszip';
import pako from 'pako';
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
  private _pendingNestedTableInserts: Array<{ sectionIndex: number; parentTableIndex: number; row: number; col: number; nestedRows: number; nestedCols: number; data: string[][] }> = [];
  private _pendingImageInserts: Array<{ sectionIndex: number; afterElementIndex: number; imageId: string; binaryId: string; data: string; mimeType: string; width: number; height: number }> = [];

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

    // version.xml - required for HWPX format (한컴 형식)
    zip.file('version.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hv:HCFVersion xmlns:hv="http://www.hancom.co.kr/hwpml/2011/version" tagetApplication="WORDPROCESSOR" major="5" minor="1" micro="1" buildNumber="0" os="1" xmlVersion="1.5" application="Hancom Office Hangul" appVersion="14, 0, 0, 0 WIN32LEWindows_10"/>`);

    // META-INF/container.xml - required for package structure
    zip.file('META-INF/container.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf"><ocf:rootfiles><ocf:rootfile full-path="Contents/content.hpf" media-type="application/hwpml-package+xml"/></ocf:rootfiles></ocf:container>`);

    // META-INF/manifest.xml - empty but required
    zip.file('META-INF/manifest.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><odf:manifest xmlns:odf="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"/>`);

    // settings.xml - application settings
    zip.file('settings.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><ha:HWPApplicationSetting xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"><ha:CaretPosition listIDRef="0" paraIDRef="0" pos="0"/></ha:HWPApplicationSetting>`);

    // Contents/content.hpf - package manifest with all namespaces
    zip.file('Contents/content.hpf', `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><opf:package xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:opf="http://www.idpf.org/2007/opf/" xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" xmlns:epub="http://www.idpf.org/2007/ops" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" version="" unique-identifier="" id=""><opf:metadata><opf:title>${title || 'Untitled'}</opf:title><opf:language>ko</opf:language><opf:meta name="creator" content="text">${creator || 'Unknown'}</opf:meta><opf:meta name="CreatedDate" content="text">${now}</opf:meta><opf:meta name="ModifiedDate" content="text">${now}</opf:meta></opf:metadata><opf:manifest><opf:item id="header" href="Contents/header.xml" media-type="application/xml"/><opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/><opf:item id="settings" href="settings.xml" media-type="application/xml"/></opf:manifest><opf:spine><opf:itemref idref="header" linear="yes"/><opf:itemref idref="section0" linear="yes"/></opf:spine></opf:package>`);

    // Contents/header.xml with all namespaces
    zip.file('Contents/header.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hh:head xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:opf="http://www.idpf.org/2007/opf/" xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" xmlns:epub="http://www.idpf.org/2007/ops" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0" version="1.1" secCnt="1">
  <hh:beginNum page="1" footnote="1" endnote="1" pic="1" tbl="1" equation="1"/>
  <hh:refList>
    <hh:fontfaces itemCnt="1">
      <hh:fontface lang="HANGUL" fontCnt="1">
        <hh:font id="0" face="함초롬바탕" type="TTF"/>
      </hh:fontface>
    </hh:fontfaces>
    <hh:borderFills itemCnt="1">
      <hh:borderFill id="1" threeD="0" shadow="0" centerLine="NONE" breakCellSeparateLine="0">
        <hh:slash type="NONE"/>
        <hh:backSlash type="NONE"/>
        <hh:leftBorder type="NONE" width="0.1mm" color="#000000"/>
        <hh:rightBorder type="NONE" width="0.1mm" color="#000000"/>
        <hh:topBorder type="NONE" width="0.1mm" color="#000000"/>
        <hh:bottomBorder type="NONE" width="0.1mm" color="#000000"/>
        <hh:diagonal type="NONE" width="0.1mm" color="#000000"/>
      </hh:borderFill>
    </hh:borderFills>
    <hh:charProperties itemCnt="1">
      <hh:charPr id="0" height="1000" textColor="#000000" shadeColor="none" useFontSpace="0" useKerning="0" symMark="NONE" borderFillIDRef="1">
        <hh:fontRef hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
        <hh:ratio hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>
        <hh:spacing hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
        <hh:relSz hangul="100" latin="100" hanja="100" japanese="100" other="100" symbol="100" user="100"/>
        <hh:offset hangul="0" latin="0" hanja="0" japanese="0" other="0" symbol="0" user="0"/>
        <hh:italic/>
        <hh:bold/>
        <hh:underline type="NONE" shape="SOLID" color="#000000"/>
        <hh:strikeout type="NONE" shape="SOLID" color="#000000"/>
        <hh:outline type="NONE"/>
        <hh:shadow type="NONE" color="#B2B2B2" offsetX="10" offsetY="10"/>
        <hh:emboss/>
        <hh:engrave/>
        <hh:supscript/>
        <hh:subscript/>
      </hh:charPr>
    </hh:charProperties>
    <hh:tabProperties itemCnt="1">
      <hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/>
    </hh:tabProperties>
    <hh:numberings itemCnt="0"/>
    <hh:bullets itemCnt="0"/>
    <hh:paraProperties itemCnt="1">
      <hh:paraPr id="0" tabPrIDRef="0" condense="0" fontLineHeight="0" snapToGrid="1" suppressLineNumbers="0" checked="0">
        <hh:align horizontal="JUSTIFY" vertical="BASELINE"/>
        <hh:heading type="NONE" idRef="0" level="0"/>
        <hh:breakSetting breakLatinWord="KEEP_WORD" breakNonLatinWord="KEEP_WORD" widowOrphan="0" keepWithNext="0" keepLines="0" pageBreakBefore="0" lineWrap="BREAK"/>
        <hh:autoSpacing eAsianEng="0" eAsianNum="0"/>
      </hh:paraPr>
    </hh:paraProperties>
    <hh:styles itemCnt="1">
      <hh:style id="0" type="PARA" name="바탕글" engName="Body" paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="0" langId="1042" lockForm="0"/>
    </hh:styles>
    <hh:memoProperties itemCnt="0"/>
  </hh:refList>
  <hh:compatibleDocument targetProgram="HWP201X"/>
  <hh:docOption>
    <hh:linkinfo path="" pageInherit="0" footnoteInherit="0"/>
  </hh:docOption>
  <hh:trackChangeConfig flags="0"/>
</hh:head>`);

    // Contents/section0.xml with all namespaces
    zip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><hs:sec xmlns:ha="http://www.hancom.co.kr/hwpml/2011/app" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph" xmlns:hp10="http://www.hancom.co.kr/hwpml/2016/paragraph" xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hhs="http://www.hancom.co.kr/hwpml/2011/history" xmlns:hm="http://www.hancom.co.kr/hwpml/2011/master-page" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:opf="http://www.idpf.org/2007/opf/" xmlns:ooxmlchart="http://www.hancom.co.kr/hwpml/2016/ooxmlchart" xmlns:hwpunitchar="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar" xmlns:epub="http://www.idpf.org/2007/ops" xmlns:config="urn:oasis:names:tc:opendocument:xmlns:config:1.0"><hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0"><hp:secPr id="" textDirection="HORIZONTAL" spaceColumns="1134" tabStop="8000" tabStopVal="4000" tabStopUnit="HWPUNIT" outlineShapeIDRef="1" memoShapeIDRef="0" textVerticalWidthHead="0" masterPageCnt="0"><hp:grid lineGrid="0" charGrid="0" wongoji="0"/><hp:startNum pageStartsOn="BOTH" page="0" pic="0" tbl="0" equation="0"/><hp:visibility hideFirstHeader="0" hideFirstFooter="0" hideFirstMasterPage="0" border="SHOW_ALL" fill="SHOW_ALL" hideFirstPageNum="0" hideFirstEmptyLine="0" showLineNumber="0"/><hp:pagePr landscape="NARROWER" width="59528" height="84188" gutterType="LEFT_ONLY"><hp:pageMar header="4252" footer="4252" left="8504" right="8504" top="5668" bottom="4252" gutter="0"/></hp:pagePr><hp:footNotePr><hp:autoNumFormat type="DIGIT"/><hp:noteLine length="-1" type="SOLID" width="0.12mm" color="#000000"/><hp:noteSpacing aboveLine="850" belowLine="567" betweenNotes="283"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="EACH_COLUMN" beneathText="0"/></hp:footNotePr><hp:endNotePr><hp:autoNumFormat type="DIGIT"/><hp:noteLine length="14692" type="SOLID" width="0.12mm" color="#000000"/><hp:noteSpacing aboveLine="850" belowLine="567" betweenNotes="0"/><hp:numbering type="CONTINUOUS" newNum="1"/><hp:placement place="END_OF_DOCUMENT" beneathText="0"/></hp:endNotePr></hp:secPr><hp:t></hp:t></hp:run></hp:p></hs:sec>`);

    // Create empty BinData folder
    zip.folder('BinData');

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

  searchText(query: string, options: { caseSensitive?: boolean; regex?: boolean; includeTables?: boolean } = {}): Array<{ section: number; element: number; text: string; matches: string[]; count: number; location?: { type: 'paragraph' | 'table'; tableIndex?: number; row?: number; col?: number } }> {
    const { caseSensitive = false, regex = false, includeTables = true } = options;
    let pattern: RegExp;

    if (regex) {
      pattern = new RegExp(query, caseSensitive ? 'g' : 'gi');
    } else {
      const escaped = query.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      pattern = new RegExp(escaped, caseSensitive ? 'g' : 'gi');
    }

    const results: Array<{ section: number; element: number; text: string; matches: string[]; count: number; location?: { type: 'paragraph' | 'table'; tableIndex?: number; row?: number; col?: number } }> = [];

    this._content.sections.forEach((section, si) => {
      let tableIndex = 0;
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
              location: { type: 'paragraph' },
            });
          }
        }
        // Search in table cells
        if (el.type === 'table' && includeTables) {
          const table = el.data as HwpxTable;
          table.rows.forEach((row, ri) => {
            row.cells.forEach((cell, ci) => {
              const cellText = cell.paragraphs.map(p => p.runs.map(r => r.text).join('')).join('\n');
              const found = cellText.match(pattern);
              if (found) {
                results.push({
                  section: si,
                  element: ei,
                  text: cellText,
                  matches: found,
                  count: found.length,
                  location: { type: 'table', tableIndex, row: ri, col: ci },
                });
              }
            });
          });
          tableIndex++;
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

  /**
   * Replace text within a specific table cell.
   * This is more targeted than replaceText and works directly on cell content.
   */
  replaceTextInCell(
    sectionIndex: number,
    tableIndex: number,
    row: number,
    col: number,
    oldText: string,
    newText: string,
    options: { caseSensitive?: boolean; regex?: boolean; replaceAll?: boolean } = {}
  ): { success: boolean; count: number; error?: string } {
    const { caseSensitive = false, regex = false, replaceAll = true } = options;

    const section = this._content.sections[sectionIndex];
    if (!section) {
      return { success: false, count: 0, error: `Section ${sectionIndex} not found` };
    }

    // Find the table
    let tableCount = 0;
    let targetTable: HwpxTable | null = null;
    for (const el of section.elements) {
      if (el.type === 'table') {
        if (tableCount === tableIndex) {
          targetTable = el.data as HwpxTable;
          break;
        }
        tableCount++;
      }
    }

    if (!targetTable) {
      return { success: false, count: 0, error: `Table ${tableIndex} not found in section ${sectionIndex}` };
    }

    if (row >= targetTable.rows.length) {
      return { success: false, count: 0, error: `Row ${row} out of range (max: ${targetTable.rows.length - 1})` };
    }

    if (col >= targetTable.rows[row].cells.length) {
      return { success: false, count: 0, error: `Column ${col} out of range (max: ${targetTable.rows[row].cells.length - 1})` };
    }

    const cell = targetTable.rows[row].cells[col];
    this.saveState();

    let pattern: RegExp;
    if (regex) {
      pattern = new RegExp(oldText, caseSensitive ? (replaceAll ? 'g' : '') : (replaceAll ? 'gi' : 'i'));
    } else {
      const escaped = oldText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
      pattern = new RegExp(escaped, caseSensitive ? (replaceAll ? 'g' : '') : (replaceAll ? 'gi' : 'i'));
    }

    let count = 0;

    // Replace in memory
    for (const para of cell.paragraphs) {
      for (const run of para.runs) {
        const matches = run.text.match(pattern);
        if (matches) {
          count += matches.length;
          run.text = run.text.replace(pattern, newText);
        }
      }
    }

    // Add to pending cell updates for XML sync
    if (count > 0 && this._zip) {
      const tableId = targetTable.id || '';
      const cellText = cell.paragraphs.map(p => p.runs.map(r => r.text).join('')).join('\n');

      // Use existing pending table cell update mechanism
      this._pendingTableCellUpdates = this._pendingTableCellUpdates || [];
      this._pendingTableCellUpdates.push({
        sectionIndex,
        tableIndex,
        tableId,
        row,
        col,
        text: cellText,
      });
      this._isDirty = true;
    }

    return { success: true, count };
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

  /**
   * Insert a nested table inside a table cell.
   * @param sectionIndex Section index
   * @param parentTableIndex Parent table index
   * @param row Row index in parent table
   * @param col Column index in parent table
   * @param nestedRows Number of rows in nested table
   * @param nestedCols Number of columns in nested table
   * @param options Optional data for cells
   */
  insertNestedTable(
    sectionIndex: number,
    parentTableIndex: number,
    row: number,
    col: number,
    nestedRows: number,
    nestedCols: number,
    options?: { data?: string[][] }
  ): boolean {
    const section = this._content.sections[sectionIndex];
    if (!section) return false;
    if (nestedRows <= 0 || nestedCols <= 0) return false;

    // Find the parent table
    let tableCount = 0;
    let parentTable: HwpxTable | null = null;
    for (const element of section.elements) {
      if (element.type === 'table') {
        if (tableCount === parentTableIndex) {
          parentTable = element.data as HwpxTable;
          break;
        }
        tableCount++;
      }
    }

    if (!parentTable) return false;
    if (row >= parentTable.rows.length) return false;
    if (col >= parentTable.rows[row].cells.length) return false;

    this.saveState();

    // Store the nested table insertion request for XML processing
    if (!this._pendingNestedTableInserts) {
      this._pendingNestedTableInserts = [];
    }

    this._pendingNestedTableInserts.push({
      sectionIndex,
      parentTableIndex,
      row,
      col,
      nestedRows,
      nestedCols,
      data: options?.data || []
    });

    this._isDirty = true;
    return true;
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

  /**
   * Insert an image into the document.
   *
   * @param sectionIndex Section to insert into
   * @param afterElementIndex Insert after this element index (-1 for beginning)
   * @param imageData Image data including base64 data and MIME type
   *   - width: Target width in points (optional if preserveAspectRatio is true)
   *   - height: Target height in points (optional if preserveAspectRatio is true)
   *   - preserveAspectRatio: If true, maintains original image aspect ratio.
   *     When only width is specified, height is auto-calculated.
   *     When only height is specified, width is auto-calculated.
   *     When neither is specified, uses original dimensions (scaled to fit if too large).
   * @returns Object with image ID or null on failure
   */
  insertImage(
    sectionIndex: number,
    afterElementIndex: number,
    imageData: {
      data: string;
      mimeType: string;
      width?: number;
      height?: number;
      preserveAspectRatio?: boolean;
    }
  ): { id: string; actualWidth: number; actualHeight: number } | null {
    const section = this._content.sections[sectionIndex];
    if (!section) return null;

    this.saveState();

    // Calculate final dimensions
    let finalWidth = imageData.width ?? 400;  // Default 400pt
    let finalHeight = imageData.height ?? 300; // Default 300pt

    if (imageData.preserveAspectRatio) {
      // Get original image dimensions
      const originalDims = this.getImageDimensionsFromBase64(imageData.data, imageData.mimeType);

      if (originalDims) {
        const aspectRatio = originalDims.width / originalDims.height;

        if (imageData.width !== undefined && imageData.height === undefined) {
          // Width specified, calculate height
          finalWidth = imageData.width;
          finalHeight = Math.round(imageData.width / aspectRatio);
        } else if (imageData.height !== undefined && imageData.width === undefined) {
          // Height specified, calculate width
          finalHeight = imageData.height;
          finalWidth = Math.round(imageData.height * aspectRatio);
        } else if (imageData.width === undefined && imageData.height === undefined) {
          // Neither specified, use original dimensions (convert pixels to points, 1pt = 1.333px)
          // Cap at reasonable max size (e.g., 500pt)
          const maxSize = 500;
          const originalWidthPt = originalDims.width * 0.75; // px to pt
          const originalHeightPt = originalDims.height * 0.75;

          if (originalWidthPt > maxSize || originalHeightPt > maxSize) {
            // Scale down to fit within maxSize while preserving aspect ratio
            const scale = Math.min(maxSize / originalWidthPt, maxSize / originalHeightPt);
            finalWidth = Math.round(originalWidthPt * scale);
            finalHeight = Math.round(originalHeightPt * scale);
          } else {
            finalWidth = Math.round(originalWidthPt);
            finalHeight = Math.round(originalHeightPt);
          }
        }
        // If both width and height are specified with preserveAspectRatio, use width and recalculate height
        else {
          finalWidth = imageData.width!;
          finalHeight = Math.round(imageData.width! / aspectRatio);
        }
      }
    }

    // Generate sequential image ID (image1, image2, ...)
    const existingImageIds = this.getExistingImageIds();
    let nextNum = 1;
    while (existingImageIds.has(`image${nextNum}`)) {
      nextNum++;
    }
    const imageId = `image${nextNum}`;
    const binaryId = imageId; // Use same ID for binary reference

    const newImage: HwpxImage = {
      id: imageId,
      binaryId,
      width: finalWidth,
      height: finalHeight,
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

    // Add to pending inserts for XML sync
    this._pendingImageInserts.push({
      sectionIndex,
      afterElementIndex,
      imageId,
      binaryId,
      data: imageData.data,
      mimeType: imageData.mimeType,
      width: finalWidth,
      height: finalHeight,
    });

    this._isDirty = true;
    return { id: imageId, actualWidth: finalWidth, actualHeight: finalHeight };
  }

  /**
   * Get existing image IDs from ZIP file
   */
  private getExistingImageIds(): Set<string> {
    const ids = new Set<string>();
    if (this._zip) {
      this._zip.forEach((relativePath) => {
        if (relativePath.startsWith('BinData/image')) {
          const match = relativePath.match(/BinData\/(image\d+)\./);
          if (match) {
            ids.add(match[1]);
          }
        }
      });
    }
    // Also include pending inserts
    for (const insert of this._pendingImageInserts) {
      ids.add(insert.imageId);
    }
    return ids;
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

    // Apply nested table inserts
    if (this._pendingNestedTableInserts && this._pendingNestedTableInserts.length > 0) {
      await this.applyNestedTableInsertsToXml();
      this._pendingNestedTableInserts = [];
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

    // Apply image inserts
    if (this._pendingImageInserts && this._pendingImageInserts.length > 0) {
      await this.applyImageInsertsToZip();
      this._pendingImageInserts = [];
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
   * Apply nested table inserts to XML.
   * Inserts a new table inside a cell of an existing table.
   */
  private async applyNestedTableInsertsToXml(): Promise<void> {
    if (!this._zip) return;

    // Group inserts by section
    const insertsBySection = new Map<number, Array<{ parentTableIndex: number; row: number; col: number; nestedRows: number; nestedCols: number; data: string[][] }>>();
    for (const insert of this._pendingNestedTableInserts) {
      const sectionInserts = insertsBySection.get(insert.sectionIndex) || [];
      sectionInserts.push({
        parentTableIndex: insert.parentTableIndex,
        row: insert.row,
        col: insert.col,
        nestedRows: insert.nestedRows,
        nestedCols: insert.nestedCols,
        data: insert.data
      });
      insertsBySection.set(insert.sectionIndex, sectionInserts);
    }

    // Process each section
    for (const [sectionIndex, inserts] of insertsBySection) {
      const sectionPath = `Contents/section${sectionIndex}.xml`;
      const file = this._zip.file(sectionPath);
      if (!file) continue;

      let xml = await file.async('string');

      // Group inserts by parentTableIndex to handle multiple inserts to same table correctly
      const insertsByTable = new Map<number, Array<{ row: number; col: number; nestedRows: number; nestedCols: number; data: string[][] }>>();
      for (const insert of inserts) {
        const tableInserts = insertsByTable.get(insert.parentTableIndex) || [];
        tableInserts.push({
          row: insert.row,
          col: insert.col,
          nestedRows: insert.nestedRows,
          nestedCols: insert.nestedCols,
          data: insert.data
        });
        insertsByTable.set(insert.parentTableIndex, tableInserts);
      }

      // Process tables in descending order of index to avoid position shifting between tables
      const sortedTableIndices = [...insertsByTable.keys()].sort((a, b) => b - a);

      for (const tableIndex of sortedTableIndices) {
        const tableInserts = insertsByTable.get(tableIndex)!;

        // Re-find tables for each table we process (positions may have shifted)
        const currentTables = this.findAllTables(xml);
        if (tableIndex >= currentTables.length) continue;

        const tableData = currentTables[tableIndex];
        let tableXml = tableData.xml;

        // Count tags before all insertions for this table
        const beforeTblOpen = (xml.match(/<(?:hp|hs|hc):tbl/g) || []).length;
        const beforeTblClose = (xml.match(/<\/(?:hp|hs|hc):tbl>/g) || []).length;

        // Sort inserts by row (desc), then col (desc) to avoid position shifting within table
        tableInserts.sort((a, b) => {
          if (a.row !== b.row) return b.row - a.row;
          return b.col - a.col;
        });

        // Process all inserts for this table, re-finding elements after each modification
        for (const insert of tableInserts) {
          // Re-find rows in current tableXml (positions change after each insert)
          const rows = this.findAllElementsWithDepth(tableXml, 'tr');
          if (insert.row >= rows.length) continue;

          const rowXml = rows[insert.row].xml;
          const cells = this.findAllElementsWithDepth(rowXml, 'tc');
          if (insert.col >= cells.length) continue;

          const cellXml = cells[insert.col].xml;

          // Generate nested table XML
          const nestedTableXml = this.generateNestedTableXml(insert.nestedRows, insert.nestedCols, insert.data);

          // Insert nested table into cell
          const updatedCellXml = this.insertNestedTableIntoCell(cellXml, nestedTableXml);

          // Update the row with the new cell
          const updatedRowXml = rowXml.substring(0, cells[insert.col].startIndex) +
            updatedCellXml +
            rowXml.substring(cells[insert.col].endIndex);

          // Update the table with the new row
          tableXml = tableXml.substring(0, rows[insert.row].startIndex) +
            updatedRowXml +
            tableXml.substring(rows[insert.row].endIndex);
        }

        // Apply all inserts for this table to the main XML at once
        xml = xml.substring(0, tableData.startIndex) +
          tableXml +
          xml.substring(tableData.endIndex);

        // Validate XML integrity: should have +N for both open and close tags
        const expectedIncrease = tableInserts.length;
        const afterTblOpen = (xml.match(/<(?:hp|hs|hc):tbl/g) || []).length;
        const afterTblClose = (xml.match(/<\/(?:hp|hs|hc):tbl>/g) || []).length;

        if (afterTblOpen !== beforeTblOpen + expectedIncrease || afterTblClose !== beforeTblClose + expectedIncrease) {
          console.error(`[HwpxDocument] XML corruption detected in nested table insertion! tbl tags: ${beforeTblOpen}→${afterTblOpen} open (expected +${expectedIncrease}), ${beforeTblClose}→${afterTblClose} close (expected +${expectedIncrease})`);
          throw new Error(`Nested table insertion failed: XML tag imbalance detected (expected +${expectedIncrease} for open and close tags)`);
        }
      }

      this._zip.file(sectionPath, xml);
    }
  }

  /**
   * Generate XML for a nested table.
   */
  private generateNestedTableXml(rows: number, cols: number, data: string[][]): string {
    // Generate unique ID
    const id = Math.floor(Math.random() * 2000000000) + 100000000;
    const zOrder = Math.floor(Math.random() * 100);

    // Calculate sizes (in hwpunit, 1 hwpunit = 0.1mm)
    const cellWidth = 8000; // ~80mm per cell
    const cellHeight = 1400; // ~14mm per cell
    const tableWidth = cellWidth * cols;
    const tableHeight = cellHeight * rows;

    let xml = `<hp:tbl id="${id}" zOrder="${zOrder}" numberingType="TABLE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" pageBreak="NONE" repeatHeader="0" rowCnt="${rows}" colCnt="${cols}" cellSpacing="0" borderFillIDRef="7" noAdjust="0">`;

    // Size element
    xml += `<hp:sz width="${tableWidth}" widthRelTo="ABSOLUTE" height="${tableHeight}" heightRelTo="ABSOLUTE" protect="0"/>`;

    // Position element (treat as character for inline table)
    xml += `<hp:pos treatAsChar="1" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>`;

    // Original position
    xml += `<hp:outMargin left="0" right="0" top="0" bottom="0"/>`;

    // Inside margin
    xml += `<hp:inMargin left="0" right="0" top="0" bottom="0"/>`;

    // Cell zone list (column widths)
    xml += `<hp:cellzoneList>`;
    for (let c = 0; c < cols; c++) {
      xml += `<hp:cellzone startRowAddr="0" startColAddr="${c}" endRowAddr="${rows - 1}" endColAddr="${c}" borderFillIDRef="7"/>`;
    }
    xml += `</hp:cellzoneList>`;

    // Table rows
    for (let r = 0; r < rows; r++) {
      xml += `<hp:tr>`;
      for (let c = 0; c < cols; c++) {
        const cellText = (data[r] && data[r][c]) ? this.escapeXml(data[r][c]) : '';

        xml += `<hp:tc name="" header="0" hasMargin="0" protect="0" editable="0" dirty="0" borderFillIDRef="7">`;
        xml += `<hp:subList id="" textDirection="HORIZONTAL" lineWrap="BREAK" vertAlign="CENTER" linkListIDRef="0" linkListNextIDRef="0" textWidth="0" textHeight="0" hasTextRef="0" hasNumRef="0">`;
        xml += `<hp:p id="0" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0">`;
        xml += `<hp:run charPrIDRef="0">`;
        xml += `<hp:t>${cellText}</hp:t>`;
        xml += `</hp:run>`;
        xml += `<hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="1000" textheight="1000" baseline="850" spacing="600" horzpos="0" horzsize="0" flags="0"/></hp:linesegarray>`;
        xml += `</hp:p>`;
        xml += `</hp:subList>`;
        xml += `<hp:cellAddr colAddr="${c}" rowAddr="${r}"/>`;
        xml += `<hp:cellSpan colSpan="1" rowSpan="1"/>`;
        xml += `<hp:cellSz width="${cellWidth}" height="${cellHeight}"/>`;
        xml += `<hp:cellMargin left="141" right="141" top="141" bottom="141"/>`;
        xml += `</hp:tc>`;
      }
      xml += `</hp:tr>`;
    }

    xml += `</hp:tbl>`;
    return xml;
  }

  /**
   * Insert a nested table XML into a cell XML.
   * Finds the last <hp:p> in the cell and inserts the table inside a run.
   */
  private insertNestedTableIntoCell(cellXml: string, nestedTableXml: string): string {
    // Find the subList in the cell
    const subListMatch = cellXml.match(/<hp:subList[^>]*>/);
    if (!subListMatch) {
      // No subList, try to add to paragraph directly
      const pMatch = cellXml.match(/<hp:p[^>]*>/);
      if (pMatch) {
        const insertPos = cellXml.indexOf(pMatch[0]) + pMatch[0].length;
        const runXml = `<hp:run charPrIDRef="0"><hp:t> </hp:t>${nestedTableXml}<hp:t/></hp:run>`;
        return cellXml.substring(0, insertPos) + runXml + cellXml.substring(insertPos);
      }
      return cellXml;
    }

    // Find the last </hp:p> before </hp:subList>
    const subListEnd = cellXml.lastIndexOf('</hp:subList>');
    if (subListEnd === -1) return cellXml;

    // Find the last </hp:p> before subList end
    const lastPEnd = cellXml.lastIndexOf('</hp:p>', subListEnd);
    if (lastPEnd === -1) return cellXml;

    // Find the corresponding <hp:p> tag
    let pStart = cellXml.lastIndexOf('<hp:p', lastPEnd);
    if (pStart === -1) return cellXml;

    // Find the end of the opening <hp:p ...> tag
    const pTagEnd = cellXml.indexOf('>', pStart) + 1;

    // Create new run with nested table
    const runXml = `<hp:run charPrIDRef="0"><hp:t> </hp:t>${nestedTableXml}<hp:t/></hp:run>`;

    // Insert after the opening <hp:p> tag
    return cellXml.substring(0, pTagEnd) + runXml + cellXml.substring(pTagEnd);
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
            // Skip nested tables by moving regex lastIndex to end of current table
            tableStartRegex.lastIndex = endIndex;
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
   * Groups updates by row to avoid index corruption when multiple cells in the same row are updated.
   */
  private updateTableCellsInXml(tableXml: string, updates: Array<{ row: number; col: number; text: string; charShapeId?: number }>): string {
    let result = tableXml;

    // Find all rows using depth tracking to handle nested tables correctly
    const rows = this.findAllElementsWithDepth(tableXml, 'tr');

    // Group updates by row to process each row only once
    const updatesByRow = new Map<number, Array<{ col: number; text: string; charShapeId?: number }>>();
    for (const update of updates) {
      if (update.row >= rows.length) continue;
      if (!updatesByRow.has(update.row)) {
        updatesByRow.set(update.row, []);
      }
      updatesByRow.get(update.row)!.push({ col: update.col, text: update.text, charShapeId: update.charShapeId });
    }

    // Sort row indices descending to process from end to start (avoid index shifting)
    const sortedRowIndices = Array.from(updatesByRow.keys()).sort((a, b) => b - a);

    for (const rowIndex of sortedRowIndices) {
      const rowData = rows[rowIndex];
      const cellUpdates = updatesByRow.get(rowIndex)!;

      // Apply all cell updates to this row at once
      const updatedRowXml = this.updateMultipleCellsInRow(rowData.xml, cellUpdates);

      result = result.substring(0, rowData.startIndex) + updatedRowXml + result.substring(rowData.endIndex);
    }

    return result;
  }

  /**
   * Update multiple cells in a single row XML string.
   * Processes cells from right to left (descending col order) to avoid index shifting.
   */
  private updateMultipleCellsInRow(rowXml: string, updates: Array<{ col: number; text: string; charShapeId?: number }>): string {
    let result = rowXml;

    // Find all cells in this row using depth tracking to handle nested tables correctly
    const cells = this.findAllElementsWithDepth(rowXml, 'tc');

    // Sort updates by col descending to process from right to left (avoid index shifting)
    const sortedUpdates = [...updates].sort((a, b) => b.col - a.col);

    for (const update of sortedUpdates) {
      if (update.col >= cells.length) continue;

      const cellData = cells[update.col];
      const updatedCellXml = this.updateTextInCell(cellData.xml, update.text, update.charShapeId);

      result = result.substring(0, cellData.startIndex) + updatedCellXml + result.substring(cellData.endIndex);
    }

    return result;
  }

  /**
   * Update a specific cell in a row XML string.
   * @deprecated Use updateMultipleCellsInRow for better index handling
   */
  private updateCellInRow(rowXml: string, colIndex: number, newText: string, charShapeId?: number): string {
    return this.updateMultipleCellsInRow(rowXml, [{ col: colIndex, text: newText, charShapeId }]);
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

  /**
   * Get raw XML content of a section.
   * Useful for AI-based document manipulation.
   */
  public async getSectionXml(sectionIndex: number): Promise<string | null> {
    if (!this._zip) return null;

    const sectionPath = `Contents/section${sectionIndex}.xml`;
    const file = this._zip.file(sectionPath);
    if (!file) return null;

    return await file.async('string');
  }

  /**
   * Set (replace) raw XML content of a section.
   * WARNING: This completely replaces the section XML. Use with caution.
   * The XML must be valid HWPML format.
   *
   * @param sectionIndex The section index to replace
   * @param xml The new XML content (must be valid HWPML)
   * @param validate If true, performs basic XML validation before replacing
   * @returns Object with success status and any validation errors
   */
  public async setSectionXml(
    sectionIndex: number,
    xml: string,
    validate: boolean = true
  ): Promise<{ success: boolean; error?: string }> {
    if (!this._zip) {
      return { success: false, error: 'Document not loaded or is HWP format (read-only)' };
    }

    const sectionPath = `Contents/section${sectionIndex}.xml`;
    const existingFile = this._zip.file(sectionPath);
    if (!existingFile) {
      return { success: false, error: `Section ${sectionIndex} does not exist` };
    }

    if (validate) {
      // Basic XML structure validation
      const validation = this.validateSectionXml(xml);
      if (!validation.valid) {
        return { success: false, error: validation.error };
      }
    }

    // Save undo state
    const originalXml = await existingFile.async('string');
    this.saveState();

    // Replace section XML
    this._zip.file(sectionPath, xml);
    this._isDirty = true;

    // Re-parse the section to update in-memory content
    try {
      const updatedContent = await HwpxParser.parse(this._zip);
      this._content = updatedContent;
      return { success: true };
    } catch (parseError) {
      // Rollback on parse error
      this._zip.file(sectionPath, originalXml);
      return {
        success: false,
        error: `XML parsing failed: ${parseError instanceof Error ? parseError.message : String(parseError)}`
      };
    }
  }

  /**
   * Validate section XML structure.
   */
  private validateSectionXml(xml: string): { valid: boolean; error?: string } {
    // Check for required root element
    if (!xml.includes('<hs:sec') && !xml.includes('<hp:sec')) {
      return { valid: false, error: 'Missing section root element (<hs:sec> or <hp:sec>)' };
    }

    // Check for basic tag balance
    const openTags = (xml.match(/<(?:hp|hs|hc):[a-zA-Z]+[^/>]*>/g) || []).length;
    const closeTags = (xml.match(/<\/(?:hp|hs|hc):[a-zA-Z]+>/g) || []).length;
    const selfCloseTags = (xml.match(/<(?:hp|hs|hc):[a-zA-Z]+[^>]*\/>/g) || []).length;

    // Note: This is a rough check. openTags - selfCloseTags should equal closeTags
    // But some tags might be counted in both patterns, so we do a simpler check
    const pOpen = (xml.match(/<(?:hp|hs):p[^>]*>/g) || []).length;
    const pClose = (xml.match(/<\/(?:hp|hs):p>/g) || []).length;
    if (pOpen !== pClose) {
      return { valid: false, error: `Paragraph tag imbalance: ${pOpen} open, ${pClose} close` };
    }

    const tblOpen = (xml.match(/<(?:hp|hs):tbl[^>]*>/g) || []).length;
    const tblClose = (xml.match(/<\/(?:hp|hs):tbl>/g) || []).length;
    if (tblOpen !== tblClose) {
      return { valid: false, error: `Table tag imbalance: ${tblOpen} open, ${tblClose} close` };
    }

    return { valid: true };
  }

  /**
   * Render Mermaid diagram and insert as image into the document.
   * Uses mermaid.ink API for rendering.
   *
   * @param mermaidCode The Mermaid diagram code
   * @param sectionIndex Section to insert into
   * @param afterElementIndex Insert after this element index (-1 for beginning)
   * @param options Optional rendering options
   *   - width: Target width in points (optional)
   *   - height: Target height in points (optional)
   *   - preserveAspectRatio: If true, maintains original image aspect ratio (default: true)
   *     When only width is specified, height is auto-calculated.
   *     When only height is specified, width is auto-calculated.
   * @returns Object with image ID and actual dimensions, or error
   */
  public async renderMermaidToImage(
    mermaidCode: string,
    sectionIndex: number,
    afterElementIndex: number,
    options?: {
      width?: number;
      height?: number;
      theme?: 'default' | 'dark' | 'forest' | 'neutral';
      backgroundColor?: string;
      preserveAspectRatio?: boolean;
    }
  ): Promise<{ success: boolean; imageId?: string; actualWidth?: number; actualHeight?: number; error?: string }> {
    if (!this._zip) {
      return { success: false, error: 'Document not loaded or is HWP format' };
    }

    const section = this._content.sections[sectionIndex];
    if (!section) {
      return { success: false, error: `Section ${sectionIndex} not found` };
    }

    try {
      // Create state object for mermaid.ink (same format as mermaid.live)
      const stateObject = {
        code: mermaidCode,
        mermaid: { theme: options?.theme || 'default' },
        autoSync: true,
        updateDiagram: true
      };

      // Encode using pako deflate + base64 URL-safe (mermaid.live format)
      const jsonString = JSON.stringify(stateObject);
      const compressed = pako.deflate(jsonString, { level: 9 });
      const base64Code = Buffer.from(compressed)
        .toString('base64')
        .replace(/\+/g, '-')
        .replace(/\//g, '_');

      // Build URL with optional background color
      let url = `https://mermaid.ink/img/pako:${base64Code}?type=png`;
      if (options?.backgroundColor) {
        // Remove # from color if present
        const bgColor = options.backgroundColor.replace(/^#/, '');
        url += `&bgColor=${bgColor}`;
      }
      const response = await fetch(url);

      if (!response.ok) {
        return {
          success: false,
          error: `Mermaid rendering failed: ${response.status} ${response.statusText}`
        };
      }

      // Get image data as base64
      const arrayBuffer = await response.arrayBuffer();
      const imageBase64 = Buffer.from(arrayBuffer).toString('base64');

      // preserveAspectRatio defaults to true for Mermaid diagrams
      const preserveAspectRatio = options?.preserveAspectRatio !== false;

      // Insert image using existing method with preserveAspectRatio support
      const result = this.insertImage(sectionIndex, afterElementIndex, {
        data: imageBase64,
        mimeType: 'image/png',
        width: options?.width,
        height: options?.height,
        preserveAspectRatio,
      });

      if (result) {
        return {
          success: true,
          imageId: result.id,
          actualWidth: result.actualWidth,
          actualHeight: result.actualHeight
        };
      } else {
        return { success: false, error: 'Failed to insert image into document' };
      }
    } catch (err) {
      return {
        success: false,
        error: `Mermaid rendering error: ${err instanceof Error ? err.message : String(err)}`
      };
    }
  }

  /**
   * Get list of available sections.
   */
  public async getAvailableSections(): Promise<number[]> {
    if (!this._zip) return [];

    const sections: number[] = [];
    const files = this._zip.file(/Contents\/section\d+\.xml/);

    for (const file of files) {
      const match = file.name.match(/section(\d+)\.xml/);
      if (match) {
        sections.push(parseInt(match[1], 10));
      }
    }

    return sections.sort((a, b) => a - b);
  }

  // ============================================================
  // Image Insert to ZIP (BinData, content.hpf, section XML)
  // ============================================================

  /**
   * Apply pending image inserts to ZIP file.
   * 1. Add image file to BinData/ folder
   * 2. Update content.hpf manifest
   * 3. Add hp:pic tag to section XML
   */
  private async applyImageInsertsToZip(): Promise<void> {
    if (!this._zip || this._pendingImageInserts.length === 0) return;

    for (const insert of this._pendingImageInserts) {
      // 1. Add image file to BinData/ folder
      const extension = this.getExtensionFromMimeType(insert.mimeType);
      const binDataPath = `BinData/${insert.imageId}.${extension}`;
      const imageBuffer = Buffer.from(insert.data, 'base64');
      this._zip.file(binDataPath, imageBuffer);

      // 2. Update content.hpf manifest
      await this.addImageToContentHpf(insert.imageId, binDataPath, insert.mimeType);

      // 3. Add hp:pic tag to section XML
      await this.addImageToSectionXml(
        insert.sectionIndex,
        insert.afterElementIndex,
        insert.imageId,
        insert.width,
        insert.height
      );
    }
  }

  /**
   * Get file extension from MIME type
   */
  private getExtensionFromMimeType(mimeType: string): string {
    const mimeToExt: Record<string, string> = {
      'image/png': 'png',
      'image/jpeg': 'jpg',
      'image/jpg': 'jpg',
      'image/gif': 'gif',
      'image/bmp': 'bmp',
      'image/webp': 'webp',
    };
    return mimeToExt[mimeType] || 'png';
  }

  /**
   * Extract original image dimensions from base64 encoded image data.
   * Supports PNG and JPEG formats.
   * @param base64Data Base64 encoded image data
   * @param mimeType MIME type of the image
   * @returns { width, height } or null if unable to parse
   */
  private getImageDimensionsFromBase64(base64Data: string, mimeType: string): { width: number; height: number } | null {
    try {
      const buffer = Buffer.from(base64Data, 'base64');

      if (mimeType === 'image/png') {
        // PNG format:
        // Bytes 0-7: PNG signature (89 50 4E 47 0D 0A 1A 0A)
        // Bytes 8-11: IHDR chunk length
        // Bytes 12-15: "IHDR" chunk type
        // Bytes 16-19: Width (big-endian)
        // Bytes 20-23: Height (big-endian)
        if (buffer.length < 24) return null;

        // Verify PNG signature
        const pngSignature = [0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A];
        for (let i = 0; i < 8; i++) {
          if (buffer[i] !== pngSignature[i]) return null;
        }

        const width = buffer.readUInt32BE(16);
        const height = buffer.readUInt32BE(20);
        return { width, height };
      }

      if (mimeType === 'image/jpeg' || mimeType === 'image/jpg') {
        // JPEG format: Find SOF0 (0xFF 0xC0) or SOF2 (0xFF 0xC2) marker
        // After marker: 2 bytes length, 1 byte precision, 2 bytes height, 2 bytes width
        if (buffer.length < 2) return null;

        // Verify JPEG signature (starts with 0xFF 0xD8)
        if (buffer[0] !== 0xFF || buffer[1] !== 0xD8) return null;

        let offset = 2;
        while (offset < buffer.length - 8) {
          if (buffer[offset] !== 0xFF) {
            offset++;
            continue;
          }

          const marker = buffer[offset + 1];

          // SOF markers (Start of Frame)
          if (marker === 0xC0 || marker === 0xC1 || marker === 0xC2) {
            // Skip marker (2 bytes) and length (2 bytes) and precision (1 byte)
            const height = buffer.readUInt16BE(offset + 5);
            const width = buffer.readUInt16BE(offset + 7);
            return { width, height };
          }

          // Skip to next marker
          if (marker === 0xD8 || marker === 0xD9 || (marker >= 0xD0 && marker <= 0xD7)) {
            // These markers have no length field
            offset += 2;
          } else {
            // Read length and skip
            const length = buffer.readUInt16BE(offset + 2);
            offset += 2 + length;
          }
        }
        return null;
      }

      // GIF format
      if (mimeType === 'image/gif') {
        if (buffer.length < 10) return null;
        // GIF87a or GIF89a signature
        const sig = buffer.slice(0, 6).toString('ascii');
        if (sig !== 'GIF87a' && sig !== 'GIF89a') return null;

        const width = buffer.readUInt16LE(6);
        const height = buffer.readUInt16LE(8);
        return { width, height };
      }

      return null;
    } catch {
      return null;
    }
  }

  /**
   * Add image entry to content.hpf manifest
   */
  private async addImageToContentHpf(imageId: string, binDataPath: string, mimeType: string): Promise<void> {
    if (!this._zip) return;

    const contentHpfFile = this._zip.file('Contents/content.hpf');
    if (!contentHpfFile) return;

    let contentHpf = await contentHpfFile.async('string');

    // Find </opf:manifest> and insert new item before it
    const manifestEndTag = '</opf:manifest>';
    const insertPos = contentHpf.indexOf(manifestEndTag);
    if (insertPos === -1) return;

    // Generate hash key (simple base64 of image id for uniqueness)
    const hashKey = Buffer.from(imageId + Date.now().toString()).toString('base64').substring(0, 22) + '==';

    // Create new item entry
    const newItem = `<opf:item id="${imageId}" href="${binDataPath}" media-type="${mimeType}" isEmbeded="1" hashkey="${hashKey}"/>`;

    // Insert before </opf:manifest>
    contentHpf = contentHpf.substring(0, insertPos) + newItem + contentHpf.substring(insertPos);

    this._zip.file('Contents/content.hpf', contentHpf);
  }

  /**
   * Add hp:pic tag to section XML
   */
  private async addImageToSectionXml(
    sectionIndex: number,
    afterElementIndex: number,
    imageId: string,
    width: number,
    height: number
  ): Promise<void> {
    if (!this._zip) return;

    const sectionPath = `Contents/section${sectionIndex}.xml`;
    const sectionFile = this._zip.file(sectionPath);
    if (!sectionFile) return;

    let sectionXml = await sectionFile.async('string');

    // Convert width/height from points to hwpunit (1pt ≈ 100 hwpunit)
    const hwpWidth = Math.round(width * 100);
    const hwpHeight = Math.round(height * 100);

    // Generate unique IDs
    const picId = Math.floor(Math.random() * 2000000000) + 100000000;
    const instId = Math.floor(Math.random() * 2000000000) + 100000000;
    const zOrder = Math.floor(Math.random() * 100);

    // Generate hp:pic XML tag
    const picXml = this.generateImagePicXml(picId, instId, zOrder, imageId, hwpWidth, hwpHeight);

    // Wrap pic in a paragraph structure (required by HWPML)
    // Image must be inside <hp:p><hp:run>...</hp:run></hp:p>
    const paraId = Math.floor(Math.random() * 2000000000);
    const fullParagraphXml = `<hp:p id="${paraId}" paraPrIDRef="0" styleIDRef="0" pageBreak="0" columnBreak="0" merged="0"><hp:run charPrIDRef="0">${picXml}<hp:t/></hp:run><hp:linesegarray><hp:lineseg textpos="0" vertpos="0" vertsize="1600" textheight="1600" baseline="1360" spacing="960" horzpos="0" horzsize="0" flags="393216"/></hp:linesegarray></hp:p>`;

    // Find insertion point - after the specified paragraph
    // Find all paragraphs
    const paraRegex = /<hp:p[^>]*>[\s\S]*?<\/hp:p>/g;
    const paragraphs: Array<{ start: number; end: number }> = [];
    let match;
    while ((match = paraRegex.exec(sectionXml)) !== null) {
      paragraphs.push({ start: match.index, end: match.index + match[0].length });
    }

    // Insert after the specified paragraph (or at the beginning if -1)
    let insertPos: number;
    if (afterElementIndex < 0 || paragraphs.length === 0) {
      // Insert at the beginning of section (after <hs:sec ...> or <hp:sec ...>)
      const secStartMatch = sectionXml.match(/<(?:hs|hp):sec[^>]*>/);
      insertPos = secStartMatch ? secStartMatch.index! + secStartMatch[0].length : 0;
    } else if (afterElementIndex >= paragraphs.length) {
      // Insert at the end (before </hs:sec> or </hp:sec>)
      let secEndMatch = sectionXml.lastIndexOf('</hs:sec>');
      if (secEndMatch === -1) {
        secEndMatch = sectionXml.lastIndexOf('</hp:sec>');
      }
      insertPos = secEndMatch !== -1 ? secEndMatch : sectionXml.length;
    } else {
      // Insert after the specified paragraph
      insertPos = paragraphs[afterElementIndex].end;
    }

    // Insert the image paragraph XML
    sectionXml = sectionXml.substring(0, insertPos) + fullParagraphXml + sectionXml.substring(insertPos);

    this._zip.file(sectionPath, sectionXml);
  }

  /**
   * Generate hp:pic XML tag for image
   */
  private generateImagePicXml(
    picId: number,
    instId: number,
    zOrder: number,
    binaryItemId: string,
    width: number,
    height: number
  ): string {
    return `<hp:pic id="${picId}" zOrder="${zOrder}" numberingType="PICTURE" textWrap="TOP_AND_BOTTOM" textFlow="BOTH_SIDES" lock="0" dropcapstyle="None" href="" groupLevel="0" instid="${instId}" reverse="0">
  <hp:offset x="0" y="0"/>
  <hp:orgSz width="${width}" height="${height}"/>
  <hp:curSz width="${width}" height="${height}"/>
  <hp:flip horizontal="0" vertical="0"/>
  <hp:rotationInfo angle="0" centerX="${Math.round(width / 2)}" centerY="${Math.round(height / 2)}" rotateimage="1"/>
  <hp:renderingInfo>
    <hc:transMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>
    <hc:scaMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>
    <hc:rotMatrix e1="1" e2="0" e3="0" e4="0" e5="1" e6="0"/>
  </hp:renderingInfo>
  <hc:img binaryItemIDRef="${binaryItemId}" bright="0" contrast="0" effect="REAL_PIC" alpha="0"/>
  <hp:imgRect>
    <hc:pt0 x="0" y="0"/>
    <hc:pt1 x="${width}" y="0"/>
    <hc:pt2 x="${width}" y="${height}"/>
    <hc:pt3 x="0" y="${height}"/>
  </hp:imgRect>
  <hp:imgClip left="0" right="${width}" top="0" bottom="${height}"/>
  <hp:inMargin left="0" right="0" top="0" bottom="0"/>
  <hp:imgDim dimwidth="${width}" dimheight="${height}"/>
  <hp:effects/>
  <hp:sz width="${width}" widthRelTo="ABSOLUTE" height="${height}" heightRelTo="ABSOLUTE" protect="0"/>
  <hp:pos treatAsChar="0" affectLSpacing="0" flowWithText="1" allowOverlap="0" holdAnchorAndSO="0" vertRelTo="PARA" horzRelTo="COLUMN" vertAlign="TOP" horzAlign="LEFT" vertOffset="0" horzOffset="0"/>
  <hp:outMargin left="0" right="0" top="0" bottom="0"/>
  <hp:shapeComment>Inserted by HWPX MCP</hp:shapeComment>
</hp:pic>`;
  }

  // ============================================================
  // XML Analysis and Repair Tools
  // ============================================================

  /**
   * Analyze XML for issues like tag imbalance, malformed elements, etc.
   * @param sectionIndex Section to analyze (optional, all sections if not specified)
   * @returns Detailed analysis report
   */
  public async analyzeXml(sectionIndex?: number): Promise<{
    hasIssues: boolean;
    sections: Array<{
      sectionIndex: number;
      issues: Array<{
        type: 'tag_imbalance' | 'malformed_tag' | 'unclosed_tag' | 'orphan_close_tag' | 'nesting_error';
        severity: 'error' | 'warning';
        message: string;
        position?: number;
        context?: string;
        suggestedFix?: string;
      }>;
      tagCounts: Record<string, { open: number; close: number; balance: number }>;
    }>;
    summary: string;
  }> {
    if (!this._zip) {
      return { hasIssues: true, sections: [], summary: 'Document not loaded' };
    }

    const result: Awaited<ReturnType<typeof this.analyzeXml>> = {
      hasIssues: false,
      sections: [],
      summary: ''
    };

    // Get sections to analyze
    const sectionsToAnalyze = sectionIndex !== undefined
      ? [sectionIndex]
      : await this.getAvailableSections();

    for (const secIdx of sectionsToAnalyze) {
      const sectionPath = `Contents/section${secIdx}.xml`;
      const file = this._zip.file(sectionPath);
      if (!file) continue;

      const xml = await file.async('string');
      const sectionResult = this.analyzeXmlContent(xml, secIdx);

      if (sectionResult.issues.length > 0) {
        result.hasIssues = true;
      }
      result.sections.push(sectionResult);
    }

    // Generate summary
    const totalIssues = result.sections.reduce((sum, s) => sum + s.issues.length, 0);
    const errors = result.sections.reduce((sum, s) => sum + s.issues.filter(i => i.severity === 'error').length, 0);
    const warnings = result.sections.reduce((sum, s) => sum + s.issues.filter(i => i.severity === 'warning').length, 0);

    result.summary = totalIssues === 0
      ? 'No XML issues detected'
      : `Found ${totalIssues} issue(s): ${errors} error(s), ${warnings} warning(s)`;

    return result;
  }

  /**
   * Analyze XML content for issues
   */
  private analyzeXmlContent(xml: string, sectionIndex: number): {
    sectionIndex: number;
    issues: Array<{
      type: 'tag_imbalance' | 'malformed_tag' | 'unclosed_tag' | 'orphan_close_tag' | 'nesting_error';
      severity: 'error' | 'warning';
      message: string;
      position?: number;
      context?: string;
      suggestedFix?: string;
    }>;
    tagCounts: Record<string, { open: number; close: number; balance: number }>;
  } {
    const issues: Array<{
      type: 'tag_imbalance' | 'malformed_tag' | 'unclosed_tag' | 'orphan_close_tag' | 'nesting_error';
      severity: 'error' | 'warning';
      message: string;
      position?: number;
      context?: string;
      suggestedFix?: string;
    }> = [];

    const tagCounts: Record<string, { open: number; close: number; balance: number }> = {};

    // List of important HWPX tags to check
    const tagsToCheck = [
      'hp:p', 'hp:run', 'hp:t', 'hp:tbl', 'hp:tr', 'hp:tc', 'hp:subList',
      'hp:pic', 'hp:container', 'hp:sec', 'hp:colPr', 'hp:paraPr',
      'hs:tbl', 'hs:tr', 'hs:tc', 'hc:tbl', 'hc:tr', 'hc:tc'
    ];

    // Count open and close tags
    for (const tag of tagsToCheck) {
      const openRegex = new RegExp(`<${tag.replace(/([.?*+^$[\]\\(){}|-])/g, '\\$1')}(?:\\s|>|\\/)`, 'g');
      const closeRegex = new RegExp(`</${tag.replace(/([.?*+^$[\]\\(){}|-])/g, '\\$1')}>`, 'g');
      const selfCloseRegex = new RegExp(`<${tag.replace(/([.?*+^$[\]\\(){}|-])/g, '\\$1')}[^>]*/\\s*>`, 'g');

      const openMatches = xml.match(openRegex) || [];
      const closeMatches = xml.match(closeRegex) || [];
      const selfCloseMatches = xml.match(selfCloseRegex) || [];

      // Self-closing tags count as both open and close
      const openCount = openMatches.length;
      const closeCount = closeMatches.length + selfCloseMatches.length;
      const balance = openCount - closeCount;

      tagCounts[tag] = { open: openCount, close: closeCount, balance };

      if (balance !== 0) {
        issues.push({
          type: (balance > 0 ? 'unclosed_tag' : 'orphan_close_tag') as 'unclosed_tag' | 'orphan_close_tag',
          severity: 'error' as const,
          message: `Tag imbalance for <${tag}>: ${openCount} open, ${closeCount} close (balance: ${balance > 0 ? '+' : ''}${balance})`,
          suggestedFix: balance > 0
            ? `Add ${balance} closing </${tag}> tag(s)`
            : `Remove ${Math.abs(balance)} orphan </${tag}> tag(s)`
        });
      }
    }

    // Find specific problem locations for tbl tags (most common issue)
    const tblIssues = this.findTblTagIssues(xml);
    for (const tblIssue of tblIssues) {
      issues.push({
        ...tblIssue,
        severity: 'error'
      });
    }

    // Check for common nesting errors
    const nestingIssues = this.checkNestingErrors(xml);
    issues.push(...nestingIssues);

    return { sectionIndex, issues, tagCounts };
  }

  /**
   * Find specific issues with tbl (table) tags
   */
  private findTblTagIssues(xml: string): Array<{
    type: 'tag_imbalance' | 'orphan_close_tag' | 'unclosed_tag';
    message: string;
    position: number;
    context: string;
    suggestedFix?: string;
  }> {
    const issues: Array<{
      type: 'tag_imbalance' | 'orphan_close_tag' | 'unclosed_tag';
      message: string;
      position: number;
      context: string;
      suggestedFix?: string;
    }> = [];

    // Track table tag positions
    const tblOpenRegex = /<(?:hp|hs|hc):tbl[^>]*>/g;
    const tblCloseRegex = /<\/(?:hp|hs|hc):tbl>/g;

    interface TagPosition {
      type: 'open' | 'close';
      position: number;
      tag: string;
    }

    const allPositions: TagPosition[] = [];

    let match;
    while ((match = tblOpenRegex.exec(xml)) !== null) {
      allPositions.push({ type: 'open', position: match.index, tag: match[0] });
    }
    while ((match = tblCloseRegex.exec(xml)) !== null) {
      allPositions.push({ type: 'close', position: match.index, tag: match[0] });
    }

    // Sort by position
    allPositions.sort((a, b) => a.position - b.position);

    // Track nesting depth
    let depth = 0;
    for (const pos of allPositions) {
      if (pos.type === 'open') {
        depth++;
      } else {
        depth--;
        if (depth < 0) {
          // Found orphan closing tag
          const contextStart = Math.max(0, pos.position - 50);
          const contextEnd = Math.min(xml.length, pos.position + 50);
          issues.push({
            type: 'orphan_close_tag',
            message: `Orphan closing table tag at position ${pos.position}`,
            position: pos.position,
            context: xml.substring(contextStart, contextEnd).replace(/\n/g, ' '),
            suggestedFix: `Remove the orphan ${pos.tag} tag or add matching opening tag before it`
          });
          depth = 0; // Reset to continue checking
        }
      }
    }

    if (depth > 0) {
      // Unclosed tables
      issues.push({
        type: 'unclosed_tag',
        message: `${depth} unclosed table tag(s) found`,
        position: xml.length,
        context: 'End of document',
        suggestedFix: `Add ${depth} closing </hp:tbl> tag(s)`
      });
    }

    return issues;
  }

  /**
   * Check for common nesting errors
   */
  private checkNestingErrors(xml: string): Array<{
    type: 'nesting_error';
    severity: 'warning';
    message: string;
    position?: number;
    context?: string;
  }> {
    const issues: Array<{
      type: 'nesting_error';
      severity: 'warning';
      message: string;
      position?: number;
      context?: string;
    }> = [];

    // Check for tc outside of tr
    const tcOutsideTr = /<(?:hp|hs|hc):tc[^>]*>(?:(?!<(?:hp|hs|hc):tr[^>]*>).)*?<\/(?:hp|hs|hc):tc>/gs;
    // This is simplified - a full check would need proper nesting validation

    // Check for tr outside of tbl
    const trPattern = /<(?:hp|hs|hc):tr[^>]*>/g;
    const tblPattern = /<(?:hp|hs|hc):tbl[^>]*>/g;

    // Simple check: count if tr appears without preceding tbl
    let match;
    let lastTblPos = -1;
    let lastTblClosePos = -1;

    const tblClosePattern = /<\/(?:hp|hs|hc):tbl>/g;

    while ((match = tblPattern.exec(xml)) !== null) {
      lastTblPos = match.index;
    }

    tblPattern.lastIndex = 0;

    // More sophisticated nesting check would go here
    // For now, we rely on tag counting

    return issues;
  }

  /**
   * Attempt to repair XML issues in a section
   * @param sectionIndex Section to repair
   * @param options Repair options
   * @returns Repair result
   */
  public async repairXml(
    sectionIndex: number,
    options: {
      removeOrphanCloseTags?: boolean;
      addMissingCloseTags?: boolean;
      fixTableStructure?: boolean;
      backup?: boolean;
    } = {}
  ): Promise<{
    success: boolean;
    message: string;
    repairsApplied: string[];
    originalXml?: string;
  }> {
    if (!this._zip) {
      return { success: false, message: 'Document not loaded', repairsApplied: [] };
    }

    const sectionPath = `Contents/section${sectionIndex}.xml`;
    const file = this._zip.file(sectionPath);
    if (!file) {
      return { success: false, message: `Section ${sectionIndex} not found`, repairsApplied: [] };
    }

    const originalXml = await file.async('string');
    let xml = originalXml;
    const repairsApplied: string[] = [];

    // Analyze current issues
    const analysis = this.analyzeXmlContent(xml, sectionIndex);
    if (analysis.issues.length === 0) {
      return { success: true, message: 'No issues to repair', repairsApplied: [] };
    }

    const opts = {
      removeOrphanCloseTags: options.removeOrphanCloseTags ?? true,
      addMissingCloseTags: options.addMissingCloseTags ?? true,
      fixTableStructure: options.fixTableStructure ?? true,
      backup: options.backup ?? true
    };

    // Repair orphan close tags
    if (opts.removeOrphanCloseTags) {
      const result = this.removeOrphanCloseTags(xml);
      if (result.modified) {
        xml = result.xml;
        repairsApplied.push(...result.repairs);
      }
    }

    // Fix table structure
    if (opts.fixTableStructure) {
      const result = this.fixTableStructure(xml);
      if (result.modified) {
        xml = result.xml;
        repairsApplied.push(...result.repairs);
      }
    }

    // Validate after repairs
    const afterAnalysis = this.analyzeXmlContent(xml, sectionIndex);
    const remainingErrors = afterAnalysis.issues.filter(i => i.severity === 'error');

    if (remainingErrors.length > 0) {
      // Don't save if errors remain
      return {
        success: false,
        message: `Repair incomplete: ${remainingErrors.length} error(s) remain`,
        repairsApplied,
        originalXml: opts.backup ? originalXml : undefined
      };
    }

    // Save repaired XML
    this._zip.file(sectionPath, xml);
    this._isDirty = true;

    return {
      success: true,
      message: `Repaired ${repairsApplied.length} issue(s)`,
      repairsApplied,
      originalXml: opts.backup ? originalXml : undefined
    };
  }

  /**
   * Remove orphan closing tags
   */
  private removeOrphanCloseTags(xml: string): { xml: string; modified: boolean; repairs: string[] } {
    const repairs: string[] = [];
    let modified = false;

    // Find and track all table tags
    const allTags: Array<{ type: 'open' | 'close'; pos: number; tag: string; prefix: string }> = [];

    const openRegex = /<(hp|hs|hc):tbl[^>]*>/g;
    const closeRegex = /<\/(hp|hs|hc):tbl>/g;

    let match;
    while ((match = openRegex.exec(xml)) !== null) {
      allTags.push({ type: 'open', pos: match.index, tag: match[0], prefix: match[1] });
    }
    while ((match = closeRegex.exec(xml)) !== null) {
      allTags.push({ type: 'close', pos: match.index, tag: match[0], prefix: match[1] });
    }

    allTags.sort((a, b) => a.pos - b.pos);

    // Find orphan close tags
    const orphanPositions: number[] = [];
    let depth = 0;

    for (const tag of allTags) {
      if (tag.type === 'open') {
        depth++;
      } else {
        depth--;
        if (depth < 0) {
          orphanPositions.push(tag.pos);
          depth = 0;
        }
      }
    }

    // Remove orphan tags in reverse order to maintain positions
    if (orphanPositions.length > 0) {
      orphanPositions.sort((a, b) => b - a);
      for (const pos of orphanPositions) {
        const closeTagMatch = xml.substring(pos).match(/^<\/(?:hp|hs|hc):tbl>/);
        if (closeTagMatch) {
          xml = xml.substring(0, pos) + xml.substring(pos + closeTagMatch[0].length);
          repairs.push(`Removed orphan closing tag at position ${pos}`);
          modified = true;
        }
      }
    }

    return { xml, modified, repairs };
  }

  /**
   * Fix table structure issues
   */
  private fixTableStructure(xml: string): { xml: string; modified: boolean; repairs: string[] } {
    const repairs: string[] = [];
    let modified = false;

    // Find tables and check their structure
    const tables = this.findAllTables(xml);

    for (let i = tables.length - 1; i >= 0; i--) {
      const table = tables[i];
      const tableXml = table.xml;

      // Check for incomplete table (missing rows/cells)
      const hasRows = /<(?:hp|hs|hc):tr/.test(tableXml);
      const hasCells = /<(?:hp|hs|hc):tc/.test(tableXml);

      if (!hasRows && !hasCells) {
        // Empty table structure - this might be intentional, skip
        continue;
      }

      // Check row/cell balance
      const trOpen = (tableXml.match(/<(?:hp|hs|hc):tr/g) || []).length;
      const trClose = (tableXml.match(/<\/(?:hp|hs|hc):tr>/g) || []).length;

      if (trOpen !== trClose) {
        // Row imbalance - complex repair needed
        repairs.push(`Table at position ${table.startIndex} has row imbalance: ${trOpen} open, ${trClose} close`);
      }

      const tcOpen = (tableXml.match(/<(?:hp|hs|hc):tc/g) || []).length;
      const tcClose = (tableXml.match(/<\/(?:hp|hs|hc):tc>/g) || []).length;

      if (tcOpen !== tcClose) {
        repairs.push(`Table at position ${table.startIndex} has cell imbalance: ${tcOpen} open, ${tcClose} close`);
      }
    }

    return { xml, modified, repairs };
  }

  /**
   * Get raw XML of a section for manual inspection/editing
   */
  public async getRawSectionXml(sectionIndex: number): Promise<string | null> {
    if (!this._zip) return null;

    const sectionPath = `Contents/section${sectionIndex}.xml`;
    const file = this._zip.file(sectionPath);
    if (!file) return null;

    return file.async('string');
  }

  /**
   * Set raw XML of a section (use with caution)
   */
  public async setRawSectionXml(sectionIndex: number, xml: string, validate: boolean = true): Promise<{
    success: boolean;
    message: string;
    issues?: Array<{ type: string; message: string }>;
  }> {
    if (!this._zip) {
      return { success: false, message: 'Document not loaded' };
    }

    // Validate XML if requested
    if (validate) {
      const analysis = this.analyzeXmlContent(xml, sectionIndex);
      const errors = analysis.issues.filter(i => i.severity === 'error');

      if (errors.length > 0) {
        return {
          success: false,
          message: `XML validation failed: ${errors.length} error(s)`,
          issues: errors.map(e => ({ type: e.type, message: e.message }))
        };
      }
    }

    const sectionPath = `Contents/section${sectionIndex}.xml`;
    this._zip.file(sectionPath, xml);
    this._isDirty = true;

    // Note: Internal state is not automatically updated.
    // Save and reopen the document to see changes in other tools.

    return { success: true, message: 'Section XML updated successfully. Save and reopen to refresh internal state.' };
  }
}
