/**
 * Tests for Table Cell Hanging Indent (테이블 셀 내 내어쓰기) functionality
 *
 * 테이블 셀 내부 단락에 내어쓰기를 적용하는 기능 테스트
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';
import JSZip from 'jszip';
import { HwpxDocument } from './HwpxDocument';

describe('HwpxDocument - Table Cell Hanging Indent (테이블 셀 내어쓰기)', () => {
  let testFilePath: string;
  let tempDir: string;

  beforeEach(async () => {
    tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'hwpx-cell-indent-test-'));
  });

  afterEach(() => {
    if (fs.existsSync(tempDir)) {
      fs.rmSync(tempDir, { recursive: true });
    }
  });

  /**
   * 테스트용 HWPX 파일 생성 (테이블 포함)
   */
  async function createTestFileWithTable(): Promise<HwpxDocument> {
    const zip = new JSZip();
    zip.file('mimetype', 'application/hwp+zip');

    // META-INF/manifest.xml
    zip.file('META-INF/manifest.xml', `<?xml version="1.0" encoding="UTF-8"?>
<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">
  <manifest:file-entry manifest:media-type="application/hwp+zip" manifest:full-path="/"/>
  <manifest:file-entry manifest:media-type="application/xml" manifest:full-path="Contents/section0.xml"/>
  <manifest:file-entry manifest:media-type="application/xml" manifest:full-path="Contents/header.xml"/>
</manifest:manifest>`);

    // Contents/content.hpf
    zip.file('Contents/content.hpf', `<?xml version="1.0" encoding="UTF-8"?>
<hpf:package xmlns:hpf="http://www.hancom.co.kr/hwpml/2011/packageList">
  <hpf:fileItem>Contents/header.xml</hpf:fileItem>
  <hpf:fileItem>Contents/section0.xml</hpf:fileItem>
</hpf:package>`);

    // Contents/header.xml - 기본 paraPr 정의 포함
    zip.file('Contents/header.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"
         xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core"
         xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
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

    // Contents/section0.xml - 테이블 포함
    zip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph"
        xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core">
  <hp:p id="1" paraPrIDRef="0" styleIDRef="0">
    <hp:run charPrIDRef="0"><hp:t>문서 제목</hp:t></hp:run>
  </hp:p>
  <hp:tbl id="tbl1" rowCnt="2" colCnt="2">
    <hp:tr>
      <hp:tc>
        <hp:cellAddr colAddr="0" rowAddr="0"/>
        <hp:cellSpan colSpan="1" rowSpan="1"/>
        <hp:p id="p1" paraPrIDRef="0" styleIDRef="0">
          <hp:run charPrIDRef="0"><hp:t>1. 첫 번째 셀 내용입니다.</hp:t></hp:run>
        </hp:p>
        <hp:p id="p2" paraPrIDRef="0" styleIDRef="0">
          <hp:run charPrIDRef="0"><hp:t>추가 단락입니다.</hp:t></hp:run>
        </hp:p>
      </hp:tc>
      <hp:tc>
        <hp:cellAddr colAddr="1" rowAddr="0"/>
        <hp:cellSpan colSpan="1" rowSpan="1"/>
        <hp:p id="p3" paraPrIDRef="0" styleIDRef="0">
          <hp:run charPrIDRef="0"><hp:t>2. 두 번째 셀 내용입니다.</hp:t></hp:run>
        </hp:p>
      </hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc>
        <hp:cellAddr colAddr="0" rowAddr="1"/>
        <hp:cellSpan colSpan="1" rowSpan="1"/>
        <hp:p id="p4" paraPrIDRef="0" styleIDRef="0">
          <hp:run charPrIDRef="0"><hp:t>3. 세 번째 셀 내용입니다.</hp:t></hp:run>
        </hp:p>
      </hp:tc>
      <hp:tc>
        <hp:cellAddr colAddr="1" rowAddr="1"/>
        <hp:cellSpan colSpan="1" rowSpan="1"/>
        <hp:p id="p5" paraPrIDRef="0" styleIDRef="0">
          <hp:run charPrIDRef="0"><hp:t>4. 네 번째 셀 내용입니다.</hp:t></hp:run>
        </hp:p>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
</hs:sec>`);

    testFilePath = path.join(tempDir, 'test-cell-hanging-indent.hwpx');
    const buffer = await zip.generateAsync({ type: 'nodebuffer' });
    fs.writeFileSync(testFilePath, buffer);

    return HwpxDocument.createFromBuffer('test-id', testFilePath, buffer);
  }

  describe('setTableCellHangingIndent', () => {
    it('should set hanging indent on a table cell paragraph', async () => {
      const doc = await createTestFileWithTable();

      // 테이블 셀 (0,0)의 첫 번째 단락에 내어쓰기 설정: 15pt
      const result = doc.setTableCellHangingIndent(0, 0, 0, 0, 0, 15);

      expect(result).toBe(true);
    });

    it('should set hanging indent on second paragraph in a cell', async () => {
      const doc = await createTestFileWithTable();

      // 테이블 셀 (0,0)의 두 번째 단락에 내어쓰기 설정
      const result = doc.setTableCellHangingIndent(0, 0, 0, 0, 1, 20);

      expect(result).toBe(true);
    });

    it('should persist hanging indent after save', async () => {
      const doc = await createTestFileWithTable();

      // 내어쓰기 설정
      doc.setTableCellHangingIndent(0, 0, 0, 0, 0, 20);

      // 저장
      const savedPath = path.join(tempDir, 'saved-cell-indent.hwpx');
      const savedBuffer = await doc.save();
      fs.writeFileSync(savedPath, savedBuffer);

      // 저장된 파일의 header.xml 확인
      const savedData = fs.readFileSync(savedPath);
      const savedZip = await JSZip.loadAsync(savedData);
      const headerXml = await savedZip.file('Contents/header.xml')?.async('string');

      // 새 paraPr이 추가되었는지 확인 (음수 intent)
      expect(headerXml).toContain('intent value="-2000"'); // -20pt = -2000 HWPUNIT
      expect(headerXml).toContain('left value="2000"');    // 20pt = 2000 HWPUNIT
    });

    it('should update paragraph paraPrIDRef in table cell after save', async () => {
      const doc = await createTestFileWithTable();

      // 내어쓰기 설정
      doc.setTableCellHangingIndent(0, 0, 0, 0, 0, 10);

      // 저장
      const savedPath = path.join(tempDir, 'saved-cell-ref.hwpx');
      const savedBuffer = await doc.save();
      fs.writeFileSync(savedPath, savedBuffer);

      // section XML에서 테이블 셀 내 단락의 paraPrIDRef가 새 ID를 참조하는지 확인
      const savedData = fs.readFileSync(savedPath);
      const savedZip = await JSZip.loadAsync(savedData);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      // 테이블 내 첫 번째 셀의 첫 번째 단락(id="p1")의 paraPrIDRef가 0이 아닌 새 ID로 변경되어야 함
      expect(sectionXml).toMatch(/<hp:p[^>]*id="p1"[^>]*paraPrIDRef="[1-9]\d*"/);
    });

    it('should reject invalid table index', async () => {
      const doc = await createTestFileWithTable();

      const result = doc.setTableCellHangingIndent(0, 99, 0, 0, 0, 15);

      expect(result).toBe(false);
    });

    it('should reject invalid row/col index', async () => {
      const doc = await createTestFileWithTable();

      expect(doc.setTableCellHangingIndent(0, 0, 99, 0, 0, 15)).toBe(false);
      expect(doc.setTableCellHangingIndent(0, 0, 0, 99, 0, 15)).toBe(false);
    });

    it('should accept paragraph index that does not exist in memory model (for multi-line text support)', async () => {
      const doc = await createTestFileWithTable();

      // 메모리 모델에는 paragraph가 없어도 pending에 추가됩니다.
      // 이는 updateTableCell()로 여러 줄 텍스트를 입력할 때,
      // 메모리 모델에는 1개의 paragraph만 유지되지만
      // XML에는 여러 paragraph가 생성되기 때문입니다.
      // 실제 paragraph 존재 검증은 save() 시 XML 처리에서 수행됩니다.
      const result = doc.setTableCellHangingIndent(0, 0, 0, 0, 99, 15);

      // pending에 추가되었으므로 true 반환
      expect(result).toBe(true);
    });

    it('should reject zero or negative indent', async () => {
      const doc = await createTestFileWithTable();

      expect(doc.setTableCellHangingIndent(0, 0, 0, 0, 0, 0)).toBe(false);
      expect(doc.setTableCellHangingIndent(0, 0, 0, 0, 0, -5)).toBe(false);
    });
  });

  describe('getTableCellHangingIndent', () => {
    it('should return hanging indent value for a table cell paragraph', async () => {
      const doc = await createTestFileWithTable();

      // 내어쓰기 설정
      doc.setTableCellHangingIndent(0, 0, 0, 0, 0, 25);

      // 내어쓰기 값 조회
      const indent = doc.getTableCellHangingIndent(0, 0, 0, 0, 0);

      expect(indent).toBe(25);
    });

    it('should return 0 for cell paragraph without hanging indent', async () => {
      const doc = await createTestFileWithTable();

      const indent = doc.getTableCellHangingIndent(0, 0, 0, 0, 0);

      expect(indent).toBe(0);
    });

    it('should return null for invalid indices', async () => {
      const doc = await createTestFileWithTable();

      expect(doc.getTableCellHangingIndent(0, 99, 0, 0, 0)).toBeNull();
      expect(doc.getTableCellHangingIndent(0, 0, 99, 0, 0)).toBeNull();
      expect(doc.getTableCellHangingIndent(0, 0, 0, 99, 0)).toBeNull();
      expect(doc.getTableCellHangingIndent(0, 0, 0, 0, 99)).toBeNull();
    });
  });

  describe('removeTableCellHangingIndent', () => {
    it('should remove hanging indent from a table cell paragraph', async () => {
      const doc = await createTestFileWithTable();

      // 먼저 내어쓰기 설정
      doc.setTableCellHangingIndent(0, 0, 0, 0, 0, 15);
      expect(doc.getTableCellHangingIndent(0, 0, 0, 0, 0)).toBe(15);

      // 내어쓰기 제거
      const result = doc.removeTableCellHangingIndent(0, 0, 0, 0, 0);

      expect(result).toBe(true);
      expect(doc.getTableCellHangingIndent(0, 0, 0, 0, 0)).toBe(0);
    });
  });

  describe('multiple cells (다중 셀)', () => {
    it('should apply same indent to multiple cells and reuse paraPr', async () => {
      const doc = await createTestFileWithTable();

      // 여러 셀에 같은 내어쓰기 적용
      doc.setTableCellHangingIndent(0, 0, 0, 0, 0, 15); // 셀 (0,0)
      doc.setTableCellHangingIndent(0, 0, 0, 1, 0, 15); // 셀 (0,1)
      doc.setTableCellHangingIndent(0, 0, 1, 0, 0, 15); // 셀 (1,0)

      // 저장
      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const headerXml = await savedZip.file('Contents/header.xml')?.async('string');

      // paraPr ID가 하나만 추가되었는지 확인 (같은 indent 값은 재사용)
      const intentMatches = headerXml?.match(/intent value="-1500"/g);
      expect(intentMatches?.length).toBe(1);
    });

    it('should create separate paraPr for different indent values', async () => {
      const doc = await createTestFileWithTable();

      // 다른 내어쓰기 값 적용
      doc.setTableCellHangingIndent(0, 0, 0, 0, 0, 10); // 셀 (0,0)
      doc.setTableCellHangingIndent(0, 0, 0, 1, 0, 20); // 셀 (0,1)

      // 저장
      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const headerXml = await savedZip.file('Contents/header.xml')?.async('string');

      // 두 개의 다른 intent 값이 있어야 함
      expect(headerXml).toContain('intent value="-1000"'); // 10pt
      expect(headerXml).toContain('intent value="-2000"'); // 20pt
    });

    it('should use last value when same cell paragraph is changed multiple times', async () => {
      const doc = await createTestFileWithTable();

      // 같은 셀 단락에 여러 번 내어쓰기 변경
      doc.setTableCellHangingIndent(0, 0, 0, 0, 0, 10);
      doc.setTableCellHangingIndent(0, 0, 0, 0, 0, 15);
      doc.setTableCellHangingIndent(0, 0, 0, 0, 0, 20); // 최종 값

      // 메모리 상태 확인
      expect(doc.getTableCellHangingIndent(0, 0, 0, 0, 0)).toBe(20);

      // 저장
      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const headerXml = await savedZip.file('Contents/header.xml')?.async('string');

      // 최종 값인 20pt만 저장되어야 함
      expect(headerXml).toContain('intent value="-2000"');
      // 중간 값들은 저장되지 않아야 함 (최적화)
      expect(headerXml).not.toContain('intent value="-1000"');
      expect(headerXml).not.toContain('intent value="-1500"');
    });
  });

  describe('edge cases (경계 케이스)', () => {
    it('should reject invalid section index', async () => {
      const doc = await createTestFileWithTable();

      const result = doc.setTableCellHangingIndent(99, 0, 0, 0, 0, 15);

      expect(result).toBe(false);
    });

    it('should persist removal after save', async () => {
      const doc = await createTestFileWithTable();

      // 내어쓰기 설정 후 제거
      doc.setTableCellHangingIndent(0, 0, 0, 0, 0, 15);
      doc.removeTableCellHangingIndent(0, 0, 0, 0, 0);

      // 저장
      const savedPath = path.join(tempDir, 'saved-removed-indent.hwpx');
      const savedBuffer = await doc.save();
      fs.writeFileSync(savedPath, savedBuffer);

      // 저장된 파일의 section XML 확인 - paraPrIDRef가 0이어야 함
      const savedData = fs.readFileSync(savedPath);
      const savedZip = await JSZip.loadAsync(savedData);
      const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

      // 테이블 내 첫 번째 셀의 첫 번째 단락(id="p1")의 paraPrIDRef가 0이어야 함
      expect(sectionXml).toMatch(/<hp:p[^>]*id="p1"[^>]*paraPrIDRef="0"/);
    });

    it('should handle multiple paragraphs in same cell independently', async () => {
      const doc = await createTestFileWithTable();

      // 같은 셀의 다른 단락에 다른 내어쓰기
      doc.setTableCellHangingIndent(0, 0, 0, 0, 0, 10); // 첫 번째 단락
      doc.setTableCellHangingIndent(0, 0, 0, 0, 1, 20); // 두 번째 단락

      expect(doc.getTableCellHangingIndent(0, 0, 0, 0, 0)).toBe(10);
      expect(doc.getTableCellHangingIndent(0, 0, 0, 0, 1)).toBe(20);

      // 저장
      const savedBuffer = await doc.save();
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const headerXml = await savedZip.file('Contents/header.xml')?.async('string');

      // 두 개의 다른 paraPr이 생성되어야 함
      expect(headerXml).toContain('intent value="-1000"');
      expect(headerXml).toContain('intent value="-2000"');
    });
  });
});
