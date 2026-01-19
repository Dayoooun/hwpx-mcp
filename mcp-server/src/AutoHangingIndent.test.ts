/**
 * Tests for Auto Hanging Indent Integration (자동 내어쓰기 통합 테스트)
 *
 * HwpxDocument.setAutoHangingIndent() 및 setTableCellAutoHangingIndent() 테스트
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import JSZip from 'jszip';
import { HwpxDocument } from './HwpxDocument';

describe('Auto Hanging Indent Integration (자동 내어쓰기 통합)', () => {
  let tempDir: string;

  beforeEach(async () => {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'hwpx-auto-indent-test-'));
  });

  afterEach(() => {
    if (fs.existsSync(tempDir)) {
      fs.rmSync(tempDir, { recursive: true });
    }
  });

  /**
   * 테스트용 HWPX 파일 생성 (다양한 마커가 있는 단락들)
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

    // 다양한 마커가 있는 단락들 포함
    zip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"
        xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core">
  <hp:p id="p1" paraPrIDRef="0" styleIDRef="0">
    <hp:run charPrIDRef="0"><hp:t>○ 불릿 마커 항목</hp:t></hp:run>
  </hp:p>
  <hp:p id="p2" paraPrIDRef="0" styleIDRef="0">
    <hp:run charPrIDRef="0"><hp:t>1. 숫자 마커 항목</hp:t></hp:run>
  </hp:p>
  <hp:p id="p3" paraPrIDRef="0" styleIDRef="0">
    <hp:run charPrIDRef="0"><hp:t>가. 한글 마커 항목</hp:t></hp:run>
  </hp:p>
  <hp:p id="p4" paraPrIDRef="0" styleIDRef="0">
    <hp:run charPrIDRef="0"><hp:t>일반 텍스트 (마커 없음)</hp:t></hp:run>
  </hp:p>
  <hp:p id="p5" paraPrIDRef="0" styleIDRef="0">
    <hp:run charPrIDRef="0"><hp:t>- 대시 마커 항목</hp:t></hp:run>
  </hp:p>
  <hp:tbl id="100" rowCnt="2" colCnt="2">
    <hp:tr>
      <hp:tc>
        <hp:cellAddr colAddr="0" rowAddr="0"/>
        <hp:cellSpan colSpan="1" rowSpan="1"/>
        <hp:p id="p6" paraPrIDRef="0" styleIDRef="0">
          <hp:run charPrIDRef="0"><hp:t>① 원문자 셀</hp:t></hp:run>
        </hp:p>
      </hp:tc>
      <hp:tc>
        <hp:cellAddr colAddr="1" rowAddr="0"/>
        <hp:cellSpan colSpan="1" rowSpan="1"/>
        <hp:p id="p7" paraPrIDRef="0" styleIDRef="0">
          <hp:run charPrIDRef="0"><hp:t>(1) 괄호 숫자 셀</hp:t></hp:run>
        </hp:p>
      </hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc>
        <hp:cellAddr colAddr="0" rowAddr="1"/>
        <hp:cellSpan colSpan="1" rowSpan="1"/>
        <hp:p id="p8" paraPrIDRef="0" styleIDRef="0">
          <hp:run charPrIDRef="0"><hp:t>일반 셀 내용</hp:t></hp:run>
        </hp:p>
      </hp:tc>
      <hp:tc>
        <hp:cellAddr colAddr="1" rowAddr="1"/>
        <hp:cellSpan colSpan="1" rowSpan="1"/>
        <hp:p id="p9" paraPrIDRef="0" styleIDRef="0">
          <hp:run charPrIDRef="0"><hp:t>A. 알파벳 셀</hp:t></hp:run>
        </hp:p>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
</hs:sec>`);

    const testFilePath = path.join(tempDir, 'test-auto-indent.hwpx');
    const buffer = await zip.generateAsync({ type: 'nodebuffer' });
    fs.writeFileSync(testFilePath, buffer);

    return HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);
  }

  // ============================================================
  // 단락 자동 내어쓰기 테스트
  // ============================================================
  describe('setAutoHangingIndent (단락 자동 내어쓰기)', () => {
    it('should auto-detect and apply hanging indent for bullet marker', async () => {
      const doc = await createTestDocument();

      // "○ 불릿 마커 항목" - 첫 번째 단락
      const indent = doc.setAutoHangingIndent(0, 0, 10);

      // v2: ○ (1em) + 공백 (0.5em) = 1.5em × 10pt × 1.3 = 19.5pt
      expect(indent).toBeGreaterThan(15);
      expect(indent).toBeLessThan(25);

      // 실제로 설정되었는지 확인
      const appliedIndent = doc.getHangingIndent(0, 0);
      expect(appliedIndent).toBeCloseTo(indent, 1);
    });

    it('should auto-detect and apply hanging indent for number marker', async () => {
      const doc = await createTestDocument();

      // "1. 숫자 마커 항목" - 두 번째 단락
      const indent = doc.setAutoHangingIndent(0, 1, 10);

      // v2: 1 (0.6em) + . (0.35em) + 공백 (0.5em) = 1.45em × 10pt × 1.3 = 18.85pt
      expect(indent).toBeGreaterThan(15);
      expect(indent).toBeLessThan(25);
    });

    it('should auto-detect and apply hanging indent for Korean marker', async () => {
      const doc = await createTestDocument();

      // "가. 한글 마커 항목" - 세 번째 단락
      const indent = doc.setAutoHangingIndent(0, 2, 10);

      // v2: 가 (1em) + . (0.35em) + 공백 (0.5em) = 1.85em × 10pt × 1.3 = 24.05pt
      expect(indent).toBeGreaterThan(20);
      expect(indent).toBeLessThan(30);
    });

    it('should return 0 for text without marker', async () => {
      const doc = await createTestDocument();

      // "일반 텍스트 (마커 없음)" - 네 번째 단락
      const indent = doc.setAutoHangingIndent(0, 3, 10);

      expect(indent).toBe(0);
    });

    it('should auto-detect dash marker', async () => {
      const doc = await createTestDocument();

      // "- 대시 마커 항목" - 다섯 번째 단락
      const indent = doc.setAutoHangingIndent(0, 4, 10);

      // v2: - (0.5em) + 공백 (0.5em) = 1em × 10pt × 1.3 = 13pt
      expect(indent).toBeGreaterThan(10);
      expect(indent).toBeLessThan(18);
    });

    it('should scale with font size', async () => {
      const doc = await createTestDocument();

      const indent10pt = doc.setAutoHangingIndent(0, 0, 10);

      // 다시 생성해서 20pt로 테스트
      const doc2 = await createTestDocument();
      const indent20pt = doc2.setAutoHangingIndent(0, 0, 20);

      // 20pt는 10pt의 약 2배
      expect(indent20pt).toBeCloseTo(indent10pt * 2, 1);
    });
  });

  // ============================================================
  // 테이블 셀 자동 내어쓰기 테스트
  // ============================================================
  describe('setTableCellAutoHangingIndent (테이블 셀 자동 내어쓰기)', () => {
    it('should auto-detect circled number in table cell', async () => {
      const doc = await createTestDocument();

      // "① 원문자 셀" - 테이블 (0,0) 셀
      const indent = doc.setTableCellAutoHangingIndent(0, 0, 0, 0, 0, 10);

      // v2: ① (1em) + 공백 (0.5em) = 1.5em × 10pt × 1.3 = 19.5pt
      expect(indent).toBeGreaterThan(15);
      expect(indent).toBeLessThan(25);

      // 실제로 설정되었는지 확인
      const appliedIndent = doc.getTableCellHangingIndent(0, 0, 0, 0, 0);
      expect(appliedIndent).toBeCloseTo(indent, 1);
    });

    it('should auto-detect parenthesized number in table cell', async () => {
      const doc = await createTestDocument();

      // "(1) 괄호 숫자 셀" - 테이블 (0,1) 셀
      const indent = doc.setTableCellAutoHangingIndent(0, 0, 0, 1, 0, 10);

      // v2: ( (0.4em) + 1 (0.6em) + ) (0.4em) + 공백 (0.5em) = 1.9em × 10pt × 1.3 = 24.7pt
      expect(indent).toBeGreaterThan(20);
      expect(indent).toBeLessThan(30);
    });

    it('should return 0 for cell without marker', async () => {
      const doc = await createTestDocument();

      // "일반 셀 내용" - 테이블 (1,0) 셀
      const indent = doc.setTableCellAutoHangingIndent(0, 0, 1, 0, 0, 10);

      expect(indent).toBe(0);
    });

    it('should auto-detect alpha marker in table cell', async () => {
      const doc = await createTestDocument();

      // "A. 알파벳 셀" - 테이블 (1,1) 셀
      const indent = doc.setTableCellAutoHangingIndent(0, 0, 1, 1, 0, 10);

      // v2: A (0.7em) + . (0.35em) + 공백 (0.5em) = 1.55em × 10pt × 1.3 = 20.15pt
      expect(indent).toBeGreaterThan(16);
      expect(indent).toBeLessThan(25);
    });
  });

  // ============================================================
  // 저장 후 유지 테스트
  // ============================================================
  describe('persistence after save (저장 후 유지)', () => {
    it('should persist auto hanging indent after save', async () => {
      const doc = await createTestDocument();

      // 자동 내어쓰기 적용
      const indent = doc.setAutoHangingIndent(0, 0, 10);
      expect(indent).toBeGreaterThan(0);

      // 저장
      const savedBuffer = await doc.save();
      const savedPath = path.join(tempDir, 'saved-auto-indent.hwpx');
      fs.writeFileSync(savedPath, savedBuffer);

      // 다시 열어서 확인
      const savedData = fs.readFileSync(savedPath);
      const reloadedDoc = await HwpxDocument.createFromBuffer('reload-id', savedPath, savedData);

      const appliedIndent = reloadedDoc.getHangingIndent(0, 0);
      expect(appliedIndent).toBeCloseTo(indent, 1);
    });

    it('should persist table cell auto hanging indent after save', async () => {
      const doc = await createTestDocument();

      // 테이블 셀 자동 내어쓰기 적용
      const indent = doc.setTableCellAutoHangingIndent(0, 0, 0, 0, 0, 10);
      expect(indent).toBeGreaterThan(0);

      // 저장
      const savedBuffer = await doc.save();
      const savedPath = path.join(tempDir, 'saved-table-auto-indent.hwpx');
      fs.writeFileSync(savedPath, savedBuffer);

      // 다시 열어서 확인
      const savedData = fs.readFileSync(savedPath);
      const reloadedDoc = await HwpxDocument.createFromBuffer('reload-id', savedPath, savedData);

      const appliedIndent = reloadedDoc.getTableCellHangingIndent(0, 0, 0, 0, 0);
      expect(appliedIndent).toBeCloseTo(indent, 1);
    });
  });

  // ============================================================
  // 엣지 케이스 테스트
  // ============================================================
  describe('edge cases', () => {
    it('should return 0 for invalid section index', async () => {
      const doc = await createTestDocument();

      const indent = doc.setAutoHangingIndent(99, 0, 10);

      expect(indent).toBe(0);
    });

    it('should return 0 for invalid element index', async () => {
      const doc = await createTestDocument();

      const indent = doc.setAutoHangingIndent(0, 99, 10);

      expect(indent).toBe(0);
    });

    it('should return 0 for invalid table parameters', async () => {
      const doc = await createTestDocument();

      const indent = doc.setTableCellAutoHangingIndent(0, 99, 0, 0, 0, 10);

      expect(indent).toBe(0);
    });

    it('should use default font size when not provided', async () => {
      const doc = await createTestDocument();

      // fontSize를 제공하지 않음
      const indent = doc.setAutoHangingIndent(0, 0);

      // v2: 기본 폰트 크기(12pt)로 계산됨
      // ○ (1em) + 공백 (0.5em) = 1.5em × 12pt × 1.3 = 23.4pt
      expect(indent).toBeGreaterThan(18);
      expect(indent).toBeLessThan(30);
    });
  });
});
