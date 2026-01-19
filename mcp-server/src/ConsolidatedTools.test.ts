/**
 * Tests for Consolidated Tools (통합 도구)
 *
 * 이 테스트들은 새로운 통합 도구들이 기존 114개 도구의 기능을
 * 올바르게 래핑하는지 검증합니다.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import JSZip from 'jszip';
import { HwpxDocument } from './HwpxDocument';

describe('Consolidated Tools (통합 도구)', () => {
  let tempDir: string;

  beforeEach(async () => {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'hwpx-consolidated-test-'));
  });

  afterEach(() => {
    if (fs.existsSync(tempDir)) {
      fs.rmSync(tempDir, { recursive: true });
    }
  });

  /**
   * 테스트용 HWPX 파일 생성 (단락 + 테이블 + 이미지 포함)
   */
  async function createTestDocument(): Promise<HwpxDocument> {
    const zip = new JSZip();
    zip.file('mimetype', 'application/hwp+zip');

    zip.file('META-INF/manifest.xml', `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:media-type="application/hwp+zip" manifest:full-path="/"/>
  <manifest:file-entry manifest:media-type="application/xml" manifest:full-path="Contents/section0.xml"/>
  <manifest:file-entry manifest:media-type="application/xml" manifest:full-path="Contents/header.xml"/>
</manifest:manifest>`);

    zip.file('Contents/content.hpf', `<?xml version="1.0" encoding="UTF-8"?>
<hpf:package xmlns:hpf="http://www.hancom.co.kr/hwpml/2011/packageList">
  <hpf:fileItem>Contents/header.xml</hpf:fileItem>
  <hpf:fileItem>Contents/section0.xml</hpf:fileItem>
</hpf:package>`);

    zip.file('Contents/header.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"
         xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core">
  <hh:refList>
    <hh:fontfaces itemCnt="1">
      <hh:fontface lang="HANGUL" fontCnt="1">
        <hh:font id="0" face="함초롬바탕" type="TTF"/>
      </hh:fontface>
    </hh:fontfaces>
    <hh:charProperties itemCnt="1">
      <hh:charPr id="0" height="1000" textColor="#000000"/>
    </hh:charProperties>
    <hh:tabProperties itemCnt="1">
      <hh:tabPr id="0" autoTabLeft="0" autoTabRight="0"/>
    </hh:tabProperties>
    <hh:paraProperties itemCnt="1">
      <hh:paraPr id="0" tabPrIDRef="0">
        <hh:align horizontal="JUSTIFY" vertical="BASELINE"/>
        <hh:margin>
          <hc:intent value="0" unit="HWPUNIT"/>
          <hc:left value="0" unit="HWPUNIT"/>
          <hc:right value="0" unit="HWPUNIT"/>
          <hc:prev value="0" unit="HWPUNIT"/>
          <hc:next value="0" unit="HWPUNIT"/>
        </hh:margin>
        <hh:lineSpacing type="PERCENT" value="160" unit="HWPUNIT"/>
      </hh:paraPr>
    </hh:paraProperties>
    <hh:styles itemCnt="1">
      <hh:style id="0" type="PARA" name="바탕글" paraPrIDRef="0" charPrIDRef="0"/>
    </hh:styles>
  </hh:refList>
</hh:head>`);

    zip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"
        xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core">
  <hp:p id="p1" paraPrIDRef="0" styleIDRef="0">
    <hp:run charPrIDRef="0"><hp:t>1. 연구개발 개요</hp:t></hp:run>
  </hp:p>
  <hp:p id="p2" paraPrIDRef="0" styleIDRef="0">
    <hp:run charPrIDRef="0"><hp:t>본 연구는 AI 기술을 활용한 문서 처리에 관한 것입니다.</hp:t></hp:run>
  </hp:p>
  <hp:tbl id="100" rowCnt="2" colCnt="2">
    <hp:tr>
      <hp:tc>
        <hp:cellAddr colAddr="0" rowAddr="0"/>
        <hp:cellSpan colSpan="1" rowSpan="1"/>
        <hp:p id="p3" paraPrIDRef="0" styleIDRef="0">
          <hp:run charPrIDRef="0"><hp:t>항목</hp:t></hp:run>
        </hp:p>
      </hp:tc>
      <hp:tc>
        <hp:cellAddr colAddr="1" rowAddr="0"/>
        <hp:cellSpan colSpan="1" rowSpan="1"/>
        <hp:p id="p4" paraPrIDRef="0" styleIDRef="0">
          <hp:run charPrIDRef="0"><hp:t>내용</hp:t></hp:run>
        </hp:p>
      </hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc>
        <hp:cellAddr colAddr="0" rowAddr="1"/>
        <hp:cellSpan colSpan="1" rowSpan="1"/>
        <hp:p id="p5" paraPrIDRef="0" styleIDRef="0">
          <hp:run charPrIDRef="0"><hp:t>연구 목표</hp:t></hp:run>
        </hp:p>
      </hp:tc>
      <hp:tc>
        <hp:cellAddr colAddr="1" rowAddr="1"/>
        <hp:cellSpan colSpan="1" rowSpan="1"/>
        <hp:p id="p6" paraPrIDRef="0" styleIDRef="0">
          <hp:run charPrIDRef="0"><hp:t>문서 자동화 시스템 개발</hp:t></hp:run>
        </hp:p>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
  <hp:p id="p7" paraPrIDRef="0" styleIDRef="0">
    <hp:run charPrIDRef="0"><hp:t>2. 기술개발 목표</hp:t></hp:run>
  </hp:p>
  <hp:tbl id="200" rowCnt="2" colCnt="2">
    <hp:tr>
      <hp:tc>
        <hp:cellAddr colAddr="0" rowAddr="0"/>
        <hp:cellSpan colSpan="1" rowSpan="1"/>
        <hp:p id="p8" paraPrIDRef="0" styleIDRef="0">
          <hp:run charPrIDRef="0"><hp:t>목표</hp:t></hp:run>
        </hp:p>
      </hp:tc>
      <hp:tc>
        <hp:cellAddr colAddr="1" rowAddr="0"/>
        <hp:cellSpan colSpan="1" rowSpan="1"/>
        <hp:p id="p9" paraPrIDRef="0" styleIDRef="0">
          <hp:run charPrIDRef="0"><hp:t>성과지표</hp:t></hp:run>
        </hp:p>
      </hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc>
        <hp:cellAddr colAddr="0" rowAddr="1"/>
        <hp:cellSpan colSpan="1" rowSpan="1"/>
        <hp:p id="p10" paraPrIDRef="0" styleIDRef="0">
          <hp:run charPrIDRef="0"><hp:t>정확도 향상</hp:t></hp:run>
        </hp:p>
      </hp:tc>
      <hp:tc>
        <hp:cellAddr colAddr="1" rowAddr="1"/>
        <hp:cellSpan colSpan="1" rowSpan="1"/>
        <hp:p id="p11" paraPrIDRef="0" styleIDRef="0">
          <hp:run charPrIDRef="0"><hp:t>95% 이상</hp:t></hp:run>
        </hp:p>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
</hs:sec>`);

    const testFilePath = path.join(tempDir, 'test-consolidated.hwpx');
    const buffer = await zip.generateAsync({ type: 'nodebuffer' });
    fs.writeFileSync(testFilePath, buffer);

    return HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);
  }

  // ============================================================
  // find_position 통합 도구 테스트
  // ============================================================
  describe('findPosition (위치 찾기 통합)', () => {
    it('should find table by header text', async () => {
      const doc = await createTestDocument();

      // getTableMap()은 테이블 직전 단락을 header로 저장
      // 첫 번째 테이블 직전 단락: "본 연구는 AI 기술을..."
      const result = doc.findPosition('table', 'AI 기술');

      expect(result).not.toBeNull();
      expect(result?.type).toBe('table');
      expect(result?.tableIndex).toBe(0);
    });

    it('should find paragraph by text', async () => {
      const doc = await createTestDocument();

      const result = doc.findPosition('paragraph', 'AI 기술');

      expect(result).not.toBeNull();
      expect(result?.type).toBe('paragraph');
      expect(result?.sectionIndex).toBe(0);
    });

    it('should find insert position after header', async () => {
      const doc = await createTestDocument();

      const result = doc.findPosition('insert_point', '기술개발 목표');

      expect(result).not.toBeNull();
      expect(result?.type).toBe('insert_point');
      expect(result?.sectionIndex).toBe(0);
    });

    it('should return null for not found', async () => {
      const doc = await createTestDocument();

      const result = doc.findPosition('table', '존재하지 않는 텍스트');

      expect(result).toBeNull();
    });

    it('should find table in cell content', async () => {
      const doc = await createTestDocument();

      // first_row_preview로 테이블 찾기 (첫 번째 행 셀 내용)
      const result = doc.findPosition('table', '항목');

      expect(result).not.toBeNull();
      expect(result?.tableIndex).toBe(0);
    });
  });

  // ============================================================
  // query_table 통합 도구 테스트
  // ============================================================
  describe('queryTable (테이블 조회 통합)', () => {
    it('should get table list (summary mode)', async () => {
      const doc = await createTestDocument();

      const result = doc.queryTable({ mode: 'list' });

      expect(result.tables).toHaveLength(2);
      expect(result.tables[0].rowCount).toBe(2);
      expect(result.tables[0].colCount).toBe(2);
    });

    it('should get full table data', async () => {
      const doc = await createTestDocument();

      const result = doc.queryTable({ tableIndex: 0, mode: 'full' });

      expect(result.table).not.toBeNull();
      expect(result.table?.rows).toHaveLength(2);
    });

    it('should get specific cell', async () => {
      const doc = await createTestDocument();

      const result = doc.queryTable({ tableIndex: 0, row: 1, col: 1, mode: 'cell' });

      expect(result.cell).not.toBeNull();
      expect(result.cell?.text).toContain('문서 자동화');
    });

    it('should get table map with headers', async () => {
      const doc = await createTestDocument();

      const result = doc.queryTable({ mode: 'map' });

      expect(result.map).toHaveLength(2);
      // getTableMap()은 테이블 직전 단락을 header로 저장
      // 첫 번째 테이블 직전: "본 연구는 AI 기술을 활용한 문서 처리에 관한 것입니다."
      expect(result.map[0].header).toContain('AI 기술');
    });

    it('should return null for invalid table index', async () => {
      const doc = await createTestDocument();

      const result = doc.queryTable({ tableIndex: 99, mode: 'full' });

      expect(result.table).toBeNull();
    });
  });

  // ============================================================
  // modifyContent 통합 도구 테스트
  // ============================================================
  describe('modifyContent (내용 수정 통합)', () => {
    it('should update table cell', async () => {
      const doc = await createTestDocument();

      const result = doc.modifyContent({
        type: 'cell',
        tableIndex: 0,
        row: 1,
        col: 1,
        text: '새로운 내용'
      });

      expect(result).toBe(true);

      // 변경 확인
      const cell = doc.getTableCell(0, 0, 1, 1);
      expect(cell?.text).toBe('새로운 내용');
    });

    it('should replace text globally', async () => {
      const doc = await createTestDocument();

      const result = doc.modifyContent({
        type: 'replace',
        oldText: '연구',
        newText: 'R&D',
        replaceAll: true
      });

      expect(result).toBe(true);
    });

    it('should update paragraph text', async () => {
      const doc = await createTestDocument();

      const result = doc.modifyContent({
        type: 'paragraph',
        sectionIndex: 0,
        paragraphIndex: 1,
        text: '수정된 단락 내용'
      });

      expect(result).toBe(true);
    });

    it('should reject invalid parameters', async () => {
      const doc = await createTestDocument();

      const result = doc.modifyContent({
        type: 'cell',
        tableIndex: 99,
        row: 0,
        col: 0,
        text: '테스트'
      });

      expect(result).toBe(false);
    });
  });

  // ============================================================
  // applyStyle 통합 도구 테스트
  // ============================================================
  describe('applyStyle (스타일 적용 통합)', () => {
    it('should apply hanging indent to paragraph', async () => {
      const doc = await createTestDocument();

      const result = doc.applyConsolidatedStyle({
        target: 'paragraph',
        sectionIndex: 0,
        paragraphIndex: 1,
        style: { hangingIndent: 15 }
      });

      expect(result).toBe(true);
      expect(doc.getHangingIndent(0, 1)).toBe(15);
    });

    it('should apply hanging indent to table cell', async () => {
      const doc = await createTestDocument();

      const result = doc.applyConsolidatedStyle({
        target: 'table_cell',
        tableIndex: 0,
        row: 1,
        col: 0,
        paragraphIndex: 0,
        style: { hangingIndent: 20 }
      });

      expect(result).toBe(true);
      expect(doc.getTableCellHangingIndent(0, 0, 1, 0, 0)).toBe(20);
    });

    it('should remove hanging indent', async () => {
      const doc = await createTestDocument();

      // 먼저 설정
      doc.applyConsolidatedStyle({
        target: 'paragraph',
        sectionIndex: 0,
        paragraphIndex: 1,
        style: { hangingIndent: 15 }
      });

      // 제거
      const result = doc.applyConsolidatedStyle({
        target: 'paragraph',
        sectionIndex: 0,
        paragraphIndex: 1,
        style: { hangingIndent: 0 }  // 0 = 제거
      });

      expect(result).toBe(true);
      expect(doc.getHangingIndent(0, 1)).toBe(0);
    });

    it('should apply paragraph alignment', async () => {
      const doc = await createTestDocument();

      const result = doc.applyConsolidatedStyle({
        target: 'paragraph',
        sectionIndex: 0,
        paragraphIndex: 1,
        style: { align: 'center' }
      });

      expect(result).toBe(true);
    });
  });

  // ============================================================
  // 저장 후 검증 테스트
  // ============================================================
  describe('persistence after save (저장 후 유지)', () => {
    it('should persist all modifications after save', async () => {
      const doc = await createTestDocument();

      // 여러 수정 작업
      doc.modifyContent({
        type: 'cell',
        tableIndex: 0,
        row: 1,
        col: 1,
        text: '수정된 셀'
      });

      doc.applyConsolidatedStyle({
        target: 'table_cell',
        tableIndex: 0,
        row: 1,
        col: 0,
        paragraphIndex: 0,
        style: { hangingIndent: 15 }
      });

      // 저장
      const savedBuffer = await doc.save();
      const savedPath = path.join(tempDir, 'saved-consolidated.hwpx');
      fs.writeFileSync(savedPath, savedBuffer);

      // 다시 열어서 확인
      const savedData = fs.readFileSync(savedPath);
      const reloadedDoc = await HwpxDocument.createFromBuffer('reload-id', savedPath, savedData);

      // 셀 내용 확인
      const cell = reloadedDoc.getTableCell(0, 0, 1, 1);
      expect(cell?.text).toBe('수정된 셀');
    });
  });

  // ============================================================
  // 기존 도구와의 호환성 테스트
  // ============================================================
  describe('backward compatibility (기존 도구 호환성)', () => {
    it('should work alongside existing tools', async () => {
      const doc = await createTestDocument();

      // 새 통합 도구 사용
      doc.modifyContent({
        type: 'cell',
        tableIndex: 0,
        row: 0,
        col: 0,
        text: '통합 도구로 수정'
      });

      // 기존 도구로 확인
      const cell = doc.getTableCell(0, 0, 0, 0);
      expect(cell?.text).toBe('통합 도구로 수정');

      // 기존 도구로 수정
      doc.updateTableCell(0, 0, 0, 1, '기존 도구로 수정');

      // 새 통합 도구로 확인
      const result = doc.queryTable({ tableIndex: 0, row: 0, col: 1, mode: 'cell' });
      expect(result.cell?.text).toBe('기존 도구로 수정');
    });
  });
});
