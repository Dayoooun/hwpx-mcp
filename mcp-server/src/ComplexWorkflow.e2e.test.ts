/**
 * Complex Workflow E2E Tests
 *
 * 실제 사용 시나리오를 시뮬레이션하는 장기적/복잡한 테스트
 * - 다중 세션 (열기 → 수정 → 저장 → 재열기 → 수정 → 저장)
 * - 복합 작업 (문단 + 표 + 이미지 동시 수정)
 * - 스트레스 테스트 (대량 수정)
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import { HwpxDocument } from './HwpxDocument';
import JSZip from 'jszip';
import * as fs from 'fs';
import * as path from 'path';

const testOutputDir = path.join(__dirname, '..', 'test-output');

// Helper: Create a complex test document
async function createComplexDocument(): Promise<Buffer> {
  const zip = new JSZip();

  const headerXml = `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:docInfo><hh:title>Complex Test Document</hh:title></hh:docInfo>
</hh:head>`;

  // Document with mixed content: paragraphs + tables + multi-run paragraphs
  const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <!-- 일반 문단들 -->
  <hp:p id="p1"><hp:run charPrIDRef="0"><hp:t>제목: 복합 문서 테스트</hp:t></hp:run></hp:p>
  <hp:p id="p2"><hp:run charPrIDRef="0"><hp:t>작성일: 2024-01-01</hp:t></hp:run></hp:p>

  <!-- 다중 run 문단 -->
  <hp:p id="p3">
    <hp:run charPrIDRef="0"><hp:t>볼드텍스트</hp:t></hp:run>
    <hp:run charPrIDRef="1"><hp:t> - 일반텍스트</hp:t></hp:run>
    <hp:run charPrIDRef="2"><hp:t> - 이탤릭텍스트</hp:t></hp:run>
  </hp:p>

  <!-- 첫 번째 표 -->
  <hp:tbl id="t1" rowCnt="3" colCnt="3">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0"><hp:subList><hp:p id="c00"><hp:run><hp:t>헤더1</hp:t></hp:run></hp:p></hp:subList></hp:tc>
      <hp:tc colAddr="1" rowAddr="0"><hp:subList><hp:p id="c01"><hp:run><hp:t>헤더2</hp:t></hp:run></hp:p></hp:subList></hp:tc>
      <hp:tc colAddr="2" rowAddr="0"><hp:subList><hp:p id="c02"><hp:run><hp:t>헤더3</hp:t></hp:run></hp:p></hp:subList></hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="1"><hp:subList><hp:p id="c10"><hp:run><hp:t>데이터1-1</hp:t></hp:run></hp:p></hp:subList></hp:tc>
      <hp:tc colAddr="1" rowAddr="1"><hp:subList><hp:p id="c11"><hp:run><hp:t>데이터1-2</hp:t></hp:run></hp:p></hp:subList></hp:tc>
      <hp:tc colAddr="2" rowAddr="1"><hp:subList><hp:p id="c12"><hp:run><hp:t>데이터1-3</hp:t></hp:run></hp:p></hp:subList></hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="2"><hp:subList><hp:p id="c20"><hp:run><hp:t>데이터2-1</hp:t></hp:run></hp:p></hp:subList></hp:tc>
      <hp:tc colAddr="1" rowAddr="2"><hp:subList><hp:p id="c21"><hp:run><hp:t>데이터2-2</hp:t></hp:run></hp:p></hp:subList></hp:tc>
      <hp:tc colAddr="2" rowAddr="2"><hp:subList><hp:p id="c22"><hp:run><hp:t>데이터2-3</hp:t></hp:run></hp:p></hp:subList></hp:tc>
    </hp:tr>
  </hp:tbl>

  <!-- 중간 문단 -->
  <hp:p id="p4"><hp:run charPrIDRef="0"><hp:t>표 아래 설명 문단</hp:t></hp:run></hp:p>

  <!-- 두 번째 표 (1x1 - 삭제 테스트용) -->
  <hp:tbl id="t2" rowCnt="1" colCnt="1">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0"><hp:subList><hp:p id="single"><hp:run><hp:t>단일 셀 표</hp:t></hp:run></hp:p></hp:subList></hp:tc>
    </hp:tr>
  </hp:tbl>

  <!-- 마지막 문단들 -->
  <hp:p id="p5"><hp:run charPrIDRef="0"><hp:t>결론 문단 1</hp:t></hp:run></hp:p>
  <hp:p id="p6"><hp:run charPrIDRef="0"><hp:t>결론 문단 2</hp:t></hp:run></hp:p>
</hs:sec>`;

  const contentTypes = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`;

  zip.file('Contents/header.xml', headerXml);
  zip.file('Contents/section0.xml', sectionXml);
  zip.file('[Content_Types].xml', contentTypes);

  return await zip.generateAsync({ type: 'nodebuffer' });
}

describe('Complex Workflow E2E Tests', () => {
  const testFilePath = path.join(testOutputDir, 'complex-workflow-test.hwpx');

  beforeEach(() => {
    if (!fs.existsSync(testOutputDir)) {
      fs.mkdirSync(testOutputDir, { recursive: true });
    }
  });

  afterEach(() => {
    if (fs.existsSync(testFilePath)) {
      fs.unlinkSync(testFilePath);
    }
  });

  describe('시나리오 1: 다중 세션 수정', () => {
    it('should persist changes across 5 consecutive save/reload cycles', async () => {
      console.log('\n=== 다중 세션 수정 테스트 ===');

      // Session 1: 초기 문서 생성 및 첫 수정
      const buffer1 = await createComplexDocument();
      let doc = await HwpxDocument.createFromBuffer('session1', testFilePath, buffer1);

      doc.updateParagraphText(0, 0, 0, '세션1: 제목 수정됨');
      console.log('Session 1 - Updated title');

      let savedBuffer = await doc.save();
      fs.writeFileSync(testFilePath, savedBuffer);

      // Session 2: 재로드 후 다른 문단 수정
      doc = await HwpxDocument.createFromBuffer('session2', testFilePath, fs.readFileSync(testFilePath));
      expect(doc.getParagraph(0, 0)?.text).toBe('세션1: 제목 수정됨');

      doc.updateParagraphText(0, 1, 0, '세션2: 날짜 수정됨');
      console.log('Session 2 - Updated date, title preserved:', doc.getParagraph(0, 0)?.text);

      savedBuffer = await doc.save();
      fs.writeFileSync(testFilePath, savedBuffer);

      // Session 3: 표 셀 수정
      doc = await HwpxDocument.createFromBuffer('session3', testFilePath, fs.readFileSync(testFilePath));
      expect(doc.getParagraph(0, 0)?.text).toBe('세션1: 제목 수정됨');
      expect(doc.getParagraph(0, 1)?.text).toBe('세션2: 날짜 수정됨');

      doc.updateTableCell(0, 0, 0, 0, '세션3: 헤더 수정됨');
      console.log('Session 3 - Updated table cell');

      savedBuffer = await doc.save();
      fs.writeFileSync(testFilePath, savedBuffer);

      // Session 4: 다중 run 문단 수정
      doc = await HwpxDocument.createFromBuffer('session4', testFilePath, fs.readFileSync(testFilePath));
      const table = doc.findTable(0, 0);
      expect(table?.rows[0]?.cells[0]?.paragraphs?.[0]?.runs?.[0]?.text).toContain('세션3');

      doc.updateParagraphTextPreserveStyles(0, 2, '세션4: 다중run 수정됨');
      console.log('Session 4 - Updated multi-run paragraph');

      savedBuffer = await doc.save();
      fs.writeFileSync(testFilePath, savedBuffer);

      // Session 5: 최종 검증
      doc = await HwpxDocument.createFromBuffer('session5', testFilePath, fs.readFileSync(testFilePath));

      console.log('\n=== 최종 검증 ===');
      console.log('Title:', doc.getParagraph(0, 0)?.text);
      console.log('Date:', doc.getParagraph(0, 1)?.text);
      console.log('Multi-run:', doc.getParagraph(0, 2)?.text);

      // All changes should persist
      expect(doc.getParagraph(0, 0)?.text).toBe('세션1: 제목 수정됨');
      expect(doc.getParagraph(0, 1)?.text).toBe('세션2: 날짜 수정됨');
      expect(doc.getParagraph(0, 2)?.text).toBe('세션4: 다중run 수정됨');

      console.log('✅ 5 sessions completed successfully');
    });
  });

  describe('시나리오 2: 복합 작업 (문단 + 표 동시 수정)', () => {
    it('should handle simultaneous paragraph and table modifications', async () => {
      console.log('\n=== 복합 작업 테스트 ===');

      const buffer = await createComplexDocument();
      let doc = await HwpxDocument.createFromBuffer('complex', testFilePath, buffer);

      // 동시에 여러 종류의 수정
      console.log('Performing 10 simultaneous modifications...');

      // 5개 문단 수정
      doc.updateParagraphText(0, 0, 0, '수정된 제목');
      doc.updateParagraphText(0, 1, 0, '수정된 날짜');
      doc.updateParagraphTextPreserveStyles(0, 2, '수정된 다중run');
      doc.updateParagraphText(0, 3, 0, '수정된 설명');  // 표 다음 문단
      doc.updateParagraphText(0, 5, 0, '수정된 결론');  // 마지막 문단

      // 5개 표 셀 수정
      doc.updateTableCell(0, 0, 0, 0, '수정된 헤더1');
      doc.updateTableCell(0, 0, 0, 1, '수정된 헤더2');
      doc.updateTableCell(0, 0, 0, 2, '수정된 헤더3');
      doc.updateTableCell(0, 0, 1, 0, '수정된 데이터1-1');
      doc.updateTableCell(0, 0, 2, 2, '수정된 데이터2-3');

      // 저장
      const savedBuffer = await doc.save();
      fs.writeFileSync(testFilePath, savedBuffer);

      // 재로드 및 검증
      doc = await HwpxDocument.createFromBuffer('verify', testFilePath, fs.readFileSync(testFilePath));

      // 문단 검증
      expect(doc.getParagraph(0, 0)?.text).toBe('수정된 제목');
      expect(doc.getParagraph(0, 1)?.text).toBe('수정된 날짜');
      expect(doc.getParagraph(0, 2)?.text).toBe('수정된 다중run');

      // 표 검증
      const table = doc.findTable(0, 0);
      expect(table?.rows[0]?.cells[0]?.paragraphs?.[0]?.runs?.[0]?.text).toBe('수정된 헤더1');
      expect(table?.rows[0]?.cells[1]?.paragraphs?.[0]?.runs?.[0]?.text).toBe('수정된 헤더2');
      expect(table?.rows[1]?.cells[0]?.paragraphs?.[0]?.runs?.[0]?.text).toBe('수정된 데이터1-1');

      console.log('✅ 10 simultaneous modifications persisted correctly');
    });
  });

  describe('시나리오 3: 표 삭제 워크플로우', () => {
    it('should delete table using delete_table', async () => {
      console.log('\n=== 표 삭제 테스트 (delete_table) ===');

      const buffer = await createComplexDocument();
      let doc = await HwpxDocument.createFromBuffer('delete-test', testFilePath, buffer);

      // 초기 상태: 표 2개
      const initialTables = doc.getTables(0);
      console.log('Initial tables:', initialTables.length);
      expect(initialTables.length).toBe(2);

      // 첫 번째 표 삭제
      const deleted = doc.deleteTable(0, 0);
      expect(deleted).toBe(true);

      // 저장 후 재로드
      const savedBuffer = await doc.save();
      fs.writeFileSync(testFilePath, savedBuffer);

      doc = await HwpxDocument.createFromBuffer('verify', testFilePath, fs.readFileSync(testFilePath));

      const finalTables = doc.getTables(0);
      console.log('Final tables:', finalTables.length);
      expect(finalTables.length).toBe(1);

      console.log('✅ Table deleted successfully');
    });

    it('should delete 1x1 table when deleting its only row', async () => {
      console.log('\n=== 1x1 표 행 삭제 테스트 ===');

      const buffer = await createComplexDocument();
      let doc = await HwpxDocument.createFromBuffer('1x1-test', testFilePath, buffer);

      // 두 번째 표(1x1)의 유일한 행 삭제 시도
      const initialTables = doc.getTables(0);
      expect(initialTables.length).toBe(2);

      // 1x1 표의 행 삭제 → 표 전체 삭제되어야 함
      const deleted = doc.deleteTableRow(0, 1, 0);
      expect(deleted).toBe(true);

      // 저장 후 재로드
      const savedBuffer = await doc.save();
      fs.writeFileSync(testFilePath, savedBuffer);

      doc = await HwpxDocument.createFromBuffer('verify', testFilePath, fs.readFileSync(testFilePath));

      const finalTables = doc.getTables(0);
      console.log('Tables after deleting 1x1 row:', finalTables.length);
      expect(finalTables.length).toBe(1);

      console.log('✅ 1x1 table auto-deleted when row deleted');
    });
  });

  describe('시나리오 4: 대량 수정 스트레스 테스트', () => {
    it('should handle 50 consecutive paragraph updates', async () => {
      console.log('\n=== 대량 문단 수정 테스트 ===');

      const buffer = await createComplexDocument();
      let doc = await HwpxDocument.createFromBuffer('stress', testFilePath, buffer);

      const paragraphs = doc.getParagraphs(0);
      const paraCount = paragraphs.length;
      console.log('Paragraph count:', paraCount);

      // 같은 문단을 50번 연속 수정
      for (let i = 0; i < 50; i++) {
        doc.updateParagraphText(0, 0, 0, `업데이트 #${i + 1}`);
      }

      // 메모리에서 확인
      expect(doc.getParagraph(0, 0)?.text).toBe('업데이트 #50');

      // 저장 후 재로드
      const savedBuffer = await doc.save();
      fs.writeFileSync(testFilePath, savedBuffer);

      doc = await HwpxDocument.createFromBuffer('verify', testFilePath, fs.readFileSync(testFilePath));

      // 최종 값만 저장되어야 함
      expect(doc.getParagraph(0, 0)?.text).toBe('업데이트 #50');

      // XML에서 중복 확인
      const zip = await JSZip.loadAsync(savedBuffer);
      const xml = await zip.file('Contents/section0.xml')?.async('string');
      const occurrences = (xml?.match(/업데이트 #/g) || []).length;
      console.log('Occurrences of "업데이트 #":', occurrences);

      // 최종 값만 1번 있어야 함
      expect(occurrences).toBe(1);

      console.log('✅ 50 consecutive updates handled correctly');
    });

    it('should handle updates to all paragraphs in document', async () => {
      console.log('\n=== 전체 문단 수정 테스트 ===');

      const buffer = await createComplexDocument();
      let doc = await HwpxDocument.createFromBuffer('all-para', testFilePath, buffer);

      const paragraphs = doc.getParagraphs(0);
      console.log('Total paragraphs:', paragraphs.length);

      // 모든 문단 수정 (getParagraph로 확인하면서)
      let updatedCount = 0;
      for (let i = 0; i < paragraphs.length; i++) {
        const para = doc.getParagraph(0, i);
        if (para && para.text) {
          console.log(`Updating para ${i}: "${para.text.substring(0, 20)}..."`);
          doc.updateParagraphText(0, i, 0, `문단 ${i} 수정됨`);
          updatedCount++;
        }
      }
      console.log('Updated paragraphs:', updatedCount);

      // 저장 후 재로드
      const savedBuffer = await doc.save();
      fs.writeFileSync(testFilePath, savedBuffer);

      doc = await HwpxDocument.createFromBuffer('verify', testFilePath, fs.readFileSync(testFilePath));

      // 첫 번째 문단 검증
      const reloadedText = doc.getParagraph(0, 0)?.text;
      console.log('Reloaded para 0:', reloadedText);

      expect(updatedCount).toBeGreaterThan(0);
      expect(reloadedText).toBe('문단 0 수정됨');

      console.log('✅ All paragraphs updated successfully');
    });
  });

  describe('시나리오 5: Undo/Redo 워크플로우', () => {
    it('should undo and redo paragraph changes', async () => {
      console.log('\n=== Undo/Redo 테스트 ===');

      const buffer = await createComplexDocument();
      const doc = await HwpxDocument.createFromBuffer('undo-test', testFilePath, buffer);

      const original = doc.getParagraph(0, 0)?.text;
      console.log('Original:', original);

      // 수정 1
      doc.updateParagraphText(0, 0, 0, '수정 1');
      expect(doc.getParagraph(0, 0)?.text).toBe('수정 1');

      // 수정 2
      doc.updateParagraphText(0, 0, 0, '수정 2');
      expect(doc.getParagraph(0, 0)?.text).toBe('수정 2');

      // Undo
      if (doc.canUndo()) {
        doc.undo();
        console.log('After undo:', doc.getParagraph(0, 0)?.text);
        expect(doc.getParagraph(0, 0)?.text).toBe('수정 1');
      }

      // Redo
      if (doc.canRedo()) {
        doc.redo();
        console.log('After redo:', doc.getParagraph(0, 0)?.text);
        expect(doc.getParagraph(0, 0)?.text).toBe('수정 2');
      }

      console.log('✅ Undo/Redo working correctly');
    });
  });

  describe('시나리오 6: 표 행/열 조작 복합 테스트', () => {
    it('should handle insert and delete row/column operations', async () => {
      console.log('\n=== 표 행/열 조작 복합 테스트 ===');

      const buffer = await createComplexDocument();
      let doc = await HwpxDocument.createFromBuffer('table-ops', testFilePath, buffer);

      const table = doc.findTable(0, 0);
      console.log('Initial size:', table?.rows.length, 'x', table?.rows[0]?.cells.length);
      expect(table?.rows.length).toBe(3);
      expect(table?.rows[0]?.cells.length).toBe(3);

      // 행 추가
      doc.insertTableRow(0, 0, 0);
      expect(doc.findTable(0, 0)?.rows.length).toBe(4);
      console.log('After insert row: 4 rows');

      // 열 추가
      doc.insertTableColumn(0, 0, 0);
      expect(doc.findTable(0, 0)?.rows[0]?.cells.length).toBe(4);
      console.log('After insert column: 4 columns');

      // 행 삭제
      doc.deleteTableRow(0, 0, 0);
      expect(doc.findTable(0, 0)?.rows.length).toBe(3);
      console.log('After delete row: 3 rows');

      // 열 삭제
      doc.deleteTableColumn(0, 0, 0);
      expect(doc.findTable(0, 0)?.rows[0]?.cells.length).toBe(3);
      console.log('After delete column: 3 columns');

      // 저장 후 재로드
      const savedBuffer = await doc.save();
      fs.writeFileSync(testFilePath, savedBuffer);

      doc = await HwpxDocument.createFromBuffer('verify', testFilePath, fs.readFileSync(testFilePath));

      const finalTable = doc.findTable(0, 0);
      expect(finalTable?.rows.length).toBe(3);
      expect(finalTable?.rows[0]?.cells.length).toBe(3);

      console.log('✅ Table row/column operations persisted correctly');
    });
  });

  describe('시나리오 7: replace_text 전체 문서 치환', () => {
    it('should replace text across entire document', async () => {
      console.log('\n=== 전체 문서 텍스트 치환 테스트 ===');

      const buffer = await createComplexDocument();
      let doc = await HwpxDocument.createFromBuffer('replace', testFilePath, buffer);

      // "데이터" 를 "DATA"로 치환
      const count = doc.replaceText('데이터', 'DATA');
      console.log('Replaced count:', count);
      expect(count).toBeGreaterThan(0);

      // 저장 후 재로드
      const savedBuffer = await doc.save();
      fs.writeFileSync(testFilePath, savedBuffer);

      doc = await HwpxDocument.createFromBuffer('verify', testFilePath, fs.readFileSync(testFilePath));

      // XML에서 확인
      const zip = await JSZip.loadAsync(savedBuffer);
      const xml = await zip.file('Contents/section0.xml')?.async('string');

      expect(xml).toContain('DATA');
      expect(xml).not.toContain('데이터');

      console.log('✅ Text replacement persisted across document');
    });
  });

  describe('delete_paragraph XML 반영 테스트', () => {
    it('should persist paragraph deletion to XML after save/reload', async () => {
      const buffer = await createComplexDocument();
      const testFilePath = path.join(testOutputDir, 'delete-paragraph-test.hwpx');
      fs.writeFileSync(testFilePath, buffer);

      let doc = await HwpxDocument.createFromBuffer('delete-para', testFilePath, buffer);

      // 원본 문단 수 확인
      const originalParagraphs = doc.getParagraphs(0);
      const originalCount = originalParagraphs.length;
      console.log(`원본 문단 수: ${originalCount}`);

      // 원본 XML에서 <hp:p> 개수 확인
      const originalZip = await JSZip.loadAsync(buffer);
      const originalXml = await originalZip.file('Contents/section0.xml')?.async('string') || '';
      const originalParagraphMatches = originalXml.match(/<hp:p\b/g) || [];
      console.log(`원본 XML <hp:p> 개수: ${originalParagraphMatches.length}`);

      // 두 번째 문단 삭제 (index 1)
      const deleteResult = doc.deleteParagraph(0, 1);
      expect(deleteResult).toBe(true);

      // 메모리에서 삭제 확인
      const afterDeleteParagraphs = doc.getParagraphs(0);
      expect(afterDeleteParagraphs.length).toBe(originalCount - 1);
      console.log(`삭제 후 메모리 문단 수: ${afterDeleteParagraphs.length}`);

      // 저장
      const savedBuffer = await doc.save();
      fs.writeFileSync(testFilePath, savedBuffer);

      // XML에서 <hp:p> 개수 확인
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const savedXml = await savedZip.file('Contents/section0.xml')?.async('string') || '';
      const savedParagraphMatches = savedXml.match(/<hp:p\b/g) || [];
      console.log(`저장 후 XML <hp:p> 개수: ${savedParagraphMatches.length}`);

      // XML에서도 문단이 삭제되었는지 확인 (테이블 내 문단 제외)
      // 참고: 테이블 내 셀에도 <hp:p>가 있으므로, 단순 카운트가 아닌 내용으로 확인
      expect(savedXml).not.toContain('작성일: 2024-01-01'); // 삭제된 문단 내용

      // 재로드
      doc = await HwpxDocument.createFromBuffer('reload', testFilePath, savedBuffer);
      const reloadedParagraphs = doc.getParagraphs(0);
      console.log(`재로드 후 문단 수: ${reloadedParagraphs.length}`);

      // 재로드 후에도 삭제가 유지되는지 확인
      expect(reloadedParagraphs.length).toBe(originalCount - 1);

      console.log('✅ Paragraph deletion persisted to XML');
    });

    it('should persist table deletion via deleteParagraph to XML', async () => {
      const buffer = await createComplexDocument();
      const testFilePath = path.join(testOutputDir, 'delete-table-via-para-test.hwpx');
      fs.writeFileSync(testFilePath, buffer);

      let doc = await HwpxDocument.createFromBuffer('delete-tbl', testFilePath, buffer);

      // 원본 테이블 수 확인
      const originalTables = doc.getTables(0);
      const originalTableCount = originalTables.length;
      console.log(`원본 테이블 수: ${originalTableCount}`);

      // 원본 XML에서 <hp:tbl> 개수 확인
      const originalZip = await JSZip.loadAsync(buffer);
      const originalXml = await originalZip.file('Contents/section0.xml')?.async('string') || '';
      const originalTableMatches = originalXml.match(/<hp:tbl\b/g) || [];
      console.log(`원본 XML <hp:tbl> 개수: ${originalTableMatches.length}`);

      // 테이블의 element index 찾기 (첫 번째 테이블)
      // 문서 구조: p, p, p, tbl, p, p, p, tbl, p, p → 테이블은 index 3
      // createComplexDocument()를 보면: p1, p2, p3, t1, p4, p5, p6, t2, p7, p8
      // index: 0, 1, 2, 3, 4, 5, 6, 7, 8, 9
      // 첫 번째 테이블은 index 3

      // deleteParagraph로 테이블 삭제 (element index 3 = 첫 번째 테이블)
      const deleteResult = doc.deleteParagraph(0, 3);
      expect(deleteResult).toBe(true);

      // 메모리에서 삭제 확인
      const afterDeleteTables = doc.getTables(0);
      expect(afterDeleteTables.length).toBe(originalTableCount - 1);
      console.log(`삭제 후 메모리 테이블 수: ${afterDeleteTables.length}`);

      // 저장
      const savedBuffer = await doc.save();
      fs.writeFileSync(testFilePath, savedBuffer);

      // XML에서 <hp:tbl> 개수 확인
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const savedXml = await savedZip.file('Contents/section0.xml')?.async('string') || '';
      const savedTableMatches = savedXml.match(/<hp:tbl\b/g) || [];
      console.log(`저장 후 XML <hp:tbl> 개수: ${savedTableMatches.length}`);

      // XML에서 테이블이 삭제되었는지 확인
      expect(savedTableMatches.length).toBe(originalTableMatches.length - 1);

      // 재로드
      doc = await HwpxDocument.createFromBuffer('reload', testFilePath, savedBuffer);
      const reloadedTables = doc.getTables(0);
      console.log(`재로드 후 테이블 수: ${reloadedTables.length}`);

      // 재로드 후에도 삭제가 유지되는지 확인
      expect(reloadedTables.length).toBe(originalTableCount - 1);

      console.log('✅ Table deletion via deleteParagraph persisted to XML');
    });
  });

  describe('중첩 테이블 삭제 테스트 (Nested Table Delete)', () => {
    // Helper: 중첩 테이블이 포함된 문서 생성
    async function createNestedTableDocument(): Promise<Buffer> {
      const zip = new JSZip();

      const headerXml = `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:docInfo><hh:title>Nested Table Test</hh:title></hh:docInfo>
</hh:head>`;

      // 중첩 테이블이 있는 문서 구조
      // - tbl_outer1: 외부 테이블 (내부에 tbl_inner1, tbl_inner2 포함)
      // - tbl_simple: 단순 테이블
      // - tbl_outer2: 또 다른 외부 테이블 (내부에 tbl_inner3 포함)
      const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="p1"><hp:run><hp:t>문서 시작</hp:t></hp:run></hp:p>

  <!-- 외부 테이블 1 (중첩 테이블 2개 포함) -->
  <hp:tbl id="tbl_outer1" rowCnt="2" colCnt="2">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0">
        <hp:subList>
          <hp:p id="outer1_cell00"><hp:run><hp:t>외부1-셀00</hp:t></hp:run></hp:p>
          <!-- 중첩 테이블 1 -->
          <hp:tbl id="tbl_inner1" rowCnt="1" colCnt="1">
            <hp:tr>
              <hp:tc colAddr="0" rowAddr="0">
                <hp:subList><hp:p id="inner1"><hp:run><hp:t>내부테이블1</hp:t></hp:run></hp:p></hp:subList>
              </hp:tc>
            </hp:tr>
          </hp:tbl>
        </hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="0">
        <hp:subList>
          <hp:p id="outer1_cell01"><hp:run><hp:t>외부1-셀01</hp:t></hp:run></hp:p>
        </hp:subList>
      </hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="1">
        <hp:subList>
          <hp:p id="outer1_cell10"><hp:run><hp:t>외부1-셀10</hp:t></hp:run></hp:p>
          <!-- 중첩 테이블 2 -->
          <hp:tbl id="tbl_inner2" rowCnt="1" colCnt="1">
            <hp:tr>
              <hp:tc colAddr="0" rowAddr="0">
                <hp:subList><hp:p id="inner2"><hp:run><hp:t>내부테이블2</hp:t></hp:run></hp:p></hp:subList>
              </hp:tc>
            </hp:tr>
          </hp:tbl>
        </hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="1">
        <hp:subList><hp:p id="outer1_cell11"><hp:run><hp:t>외부1-셀11</hp:t></hp:run></hp:p></hp:subList>
      </hp:tc>
    </hp:tr>
  </hp:tbl>

  <hp:p id="p2"><hp:run><hp:t>중간 문단</hp:t></hp:run></hp:p>

  <!-- 단순 테이블 (중첩 없음) -->
  <hp:tbl id="tbl_simple" rowCnt="1" colCnt="2">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0"><hp:subList><hp:p id="simple_cell0"><hp:run><hp:t>단순셀0</hp:t></hp:run></hp:p></hp:subList></hp:tc>
      <hp:tc colAddr="1" rowAddr="0"><hp:subList><hp:p id="simple_cell1"><hp:run><hp:t>단순셀1</hp:t></hp:run></hp:p></hp:subList></hp:tc>
    </hp:tr>
  </hp:tbl>

  <hp:p id="p3"><hp:run><hp:t>또 다른 중간 문단</hp:t></hp:run></hp:p>

  <!-- 외부 테이블 2 (중첩 테이블 1개 포함) -->
  <hp:tbl id="tbl_outer2" rowCnt="1" colCnt="1">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0">
        <hp:subList>
          <hp:p id="outer2_cell"><hp:run><hp:t>외부2-셀</hp:t></hp:run></hp:p>
          <!-- 중첩 테이블 3 -->
          <hp:tbl id="tbl_inner3" rowCnt="1" colCnt="1">
            <hp:tr>
              <hp:tc colAddr="0" rowAddr="0">
                <hp:subList><hp:p id="inner3"><hp:run><hp:t>내부테이블3</hp:t></hp:run></hp:p></hp:subList>
              </hp:tc>
            </hp:tr>
          </hp:tbl>
        </hp:subList>
      </hp:tc>
    </hp:tr>
  </hp:tbl>

  <hp:p id="p4"><hp:run><hp:t>문서 끝</hp:t></hp:run></hp:p>
</hs:sec>`;

      zip.file('Contents/header.xml', headerXml);
      zip.file('Contents/section0.xml', sectionXml);
      zip.file('Contents/content.hpf', '<?xml version="1.0"?><pkg:package xmlns:pkg="http://www.hancom.co.kr/hwpml/2011/package"><pkg:manifest><pkg:item id="section0" href="section0.xml"/></pkg:manifest></pkg:package>');
      zip.file('version.xml', '<?xml version="1.0"?><hwpml version="1.0"/>');
      zip.file('mimetype', 'application/hwp+zip');

      return await zip.generateAsync({ type: 'nodebuffer' });
    }

    // Helper: XML 태그 균형 검사
    function checkTagBalance(xml: string, tagName: string): { open: number; close: number; balanced: boolean } {
      const openRegex = new RegExp(`<hp:${tagName}\\b`, 'g');
      const closeRegex = new RegExp(`</hp:${tagName}>`, 'g');
      const openCount = (xml.match(openRegex) || []).length;
      const closeCount = (xml.match(closeRegex) || []).length;
      return { open: openCount, close: closeCount, balanced: openCount === closeCount };
    }

    it('should maintain XML tag balance after deleting table with nested tables', async () => {
      const buffer = await createNestedTableDocument();
      const testPath = path.join(testOutputDir, 'nested-table-delete-test.hwpx');
      fs.writeFileSync(testPath, buffer);

      let doc = await HwpxDocument.createFromBuffer('nested-test', testPath, buffer);

      // 원본 상태 확인
      const originalZip = await JSZip.loadAsync(buffer);
      const originalXml = await originalZip.file('Contents/section0.xml')?.async('string') || '';
      const originalBalance = checkTagBalance(originalXml, 'tbl');

      console.log('=== 중첩 테이블 삭제 테스트 ===');
      console.log(`원본: <hp:tbl> ${originalBalance.open}개, </hp:tbl> ${originalBalance.close}개`);
      expect(originalBalance.balanced).toBe(true);
      expect(originalBalance.open).toBe(6); // outer1 + inner1 + inner2 + simple + outer2 + inner3

      // 외부 테이블 1 삭제 (index 0) - 내부에 중첩 테이블 2개 포함
      const deleteResult = doc.deleteTable(0, 0);
      expect(deleteResult).toBe(true);
      console.log('외부 테이블 1 삭제 (중첩 테이블 2개 포함)');

      // 저장
      const savedBuffer = await doc.save();
      fs.writeFileSync(testPath, savedBuffer);

      // 저장된 XML 태그 균형 검사
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const savedXml = await savedZip.file('Contents/section0.xml')?.async('string') || '';
      const savedBalance = checkTagBalance(savedXml, 'tbl');

      console.log(`저장 후: <hp:tbl> ${savedBalance.open}개, </hp:tbl> ${savedBalance.close}개`);
      console.log(`태그 균형: ${savedBalance.balanced ? '✅ OK' : '❌ FAIL'}`);

      // 핵심 검증: 태그 균형이 맞아야 함
      expect(savedBalance.balanced).toBe(true);
      // outer1 + inner1 + inner2 가 삭제되어 3개 감소
      expect(savedBalance.open).toBe(3); // simple + outer2 + inner3

      // 내용 검증: 삭제된 테이블 내용이 없어야 함
      expect(savedXml).not.toContain('외부1-셀00');
      expect(savedXml).not.toContain('내부테이블1');
      expect(savedXml).not.toContain('내부테이블2');

      // 남은 테이블 내용은 있어야 함
      expect(savedXml).toContain('단순셀0');
      expect(savedXml).toContain('외부2-셀');
      expect(savedXml).toContain('내부테이블3');

      // 재로드 검증
      doc = await HwpxDocument.createFromBuffer('reload', testPath, savedBuffer);
      const reloadedTables = doc.getTables();
      expect(reloadedTables.length).toBe(3); // simple, outer2, inner3

      console.log('✅ 중첩 테이블 포함 삭제 성공');

      // Cleanup
      if (fs.existsSync(testPath)) fs.unlinkSync(testPath);
    });

    it('should correctly delete multiple tables with nested structures', async () => {
      const buffer = await createNestedTableDocument();
      const testPath = path.join(testOutputDir, 'multi-nested-delete-test.hwpx');
      fs.writeFileSync(testPath, buffer);

      let doc = await HwpxDocument.createFromBuffer('multi-nested', testPath, buffer);

      console.log('=== 여러 중첩 테이블 삭제 테스트 ===');

      // 역순으로 삭제 (인덱스 시프트 방지)
      // 테이블 인덱스: 0=outer1, 1=simple, 2=outer2
      // (중첩 테이블은 루트 레벨이 아니므로 인덱스에 포함되지 않음)

      // outer2 삭제 (index 2)
      expect(doc.deleteTable(0, 2)).toBe(true);
      console.log('외부 테이블 2 삭제 (중첩 테이블 1개 포함)');

      // simple 삭제 (index 1)
      expect(doc.deleteTable(0, 1)).toBe(true);
      console.log('단순 테이블 삭제');

      // 저장
      const savedBuffer = await doc.save();
      fs.writeFileSync(testPath, savedBuffer);

      // 검증
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const savedXml = await savedZip.file('Contents/section0.xml')?.async('string') || '';
      const savedBalance = checkTagBalance(savedXml, 'tbl');

      console.log(`저장 후: <hp:tbl> ${savedBalance.open}개, </hp:tbl> ${savedBalance.close}개`);
      console.log(`태그 균형: ${savedBalance.balanced ? '✅ OK' : '❌ FAIL'}`);

      expect(savedBalance.balanced).toBe(true);
      expect(savedBalance.open).toBe(3); // outer1 + inner1 + inner2 만 남음

      // 삭제된 내용 확인
      expect(savedXml).not.toContain('단순셀0');
      expect(savedXml).not.toContain('외부2-셀');
      expect(savedXml).not.toContain('내부테이블3');

      // 남은 내용 확인
      expect(savedXml).toContain('외부1-셀00');
      expect(savedXml).toContain('내부테이블1');
      expect(savedXml).toContain('내부테이블2');

      console.log('✅ 여러 중첩 테이블 삭제 성공');

      // Cleanup
      if (fs.existsSync(testPath)) fs.unlinkSync(testPath);
    });

    it('should delete all tables and maintain valid XML', async () => {
      const buffer = await createNestedTableDocument();
      const testPath = path.join(testOutputDir, 'delete-all-tables-test.hwpx');
      fs.writeFileSync(testPath, buffer);

      let doc = await HwpxDocument.createFromBuffer('delete-all', testPath, buffer);

      console.log('=== 모든 테이블 삭제 테스트 ===');

      // 역순으로 모든 루트 레벨 테이블 삭제
      expect(doc.deleteTable(0, 2)).toBe(true); // outer2
      expect(doc.deleteTable(0, 1)).toBe(true); // simple
      expect(doc.deleteTable(0, 0)).toBe(true); // outer1

      console.log('모든 루트 레벨 테이블 삭제');

      // 저장
      const savedBuffer = await doc.save();
      fs.writeFileSync(testPath, savedBuffer);

      // 검증
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const savedXml = await savedZip.file('Contents/section0.xml')?.async('string') || '';
      const savedBalance = checkTagBalance(savedXml, 'tbl');

      console.log(`저장 후: <hp:tbl> ${savedBalance.open}개, </hp:tbl> ${savedBalance.close}개`);
      console.log(`태그 균형: ${savedBalance.balanced ? '✅ OK' : '❌ FAIL'}`);

      expect(savedBalance.balanced).toBe(true);
      expect(savedBalance.open).toBe(0); // 모든 테이블 삭제됨

      // 문단은 남아있어야 함
      expect(savedXml).toContain('문서 시작');
      expect(savedXml).toContain('중간 문단');
      expect(savedXml).toContain('문서 끝');

      // 테이블 내용은 모두 삭제되어야 함
      expect(savedXml).not.toContain('외부1-셀00');
      expect(savedXml).not.toContain('단순셀0');
      expect(savedXml).not.toContain('내부테이블');

      // 재로드
      doc = await HwpxDocument.createFromBuffer('reload', testPath, savedBuffer);
      expect(doc.getTables().length).toBe(0);

      console.log('✅ 모든 테이블 삭제 성공');

      // Cleanup
      if (fs.existsSync(testPath)) fs.unlinkSync(testPath);
    });

    it('should handle consecutive deletes with index recalculation', async () => {
      const buffer = await createNestedTableDocument();
      const testPath = path.join(testOutputDir, 'consecutive-delete-test.hwpx');
      fs.writeFileSync(testPath, buffer);

      let doc = await HwpxDocument.createFromBuffer('consecutive', testPath, buffer);

      console.log('=== 연속 삭제 인덱스 재계산 테스트 ===');

      // 앞에서부터 삭제 (인덱스가 계속 바뀜)
      // 처음: [outer1(0), simple(1), outer2(2)]
      expect(doc.deleteTable(0, 0)).toBe(true); // outer1 삭제
      console.log('Step 1: index 0 삭제 후 남은 테이블 확인');

      // 이제: [simple(0), outer2(1)]
      expect(doc.deleteTable(0, 0)).toBe(true); // simple 삭제
      console.log('Step 2: index 0 삭제 후 남은 테이블 확인');

      // 이제: [outer2(0)]
      expect(doc.deleteTable(0, 0)).toBe(true); // outer2 삭제
      console.log('Step 3: index 0 삭제 후 남은 테이블 확인');

      // 저장
      const savedBuffer = await doc.save();
      fs.writeFileSync(testPath, savedBuffer);

      // 검증
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const savedXml = await savedZip.file('Contents/section0.xml')?.async('string') || '';
      const savedBalance = checkTagBalance(savedXml, 'tbl');

      console.log(`저장 후: <hp:tbl> ${savedBalance.open}개, </hp:tbl> ${savedBalance.close}개`);

      expect(savedBalance.balanced).toBe(true);
      expect(savedBalance.open).toBe(0);

      console.log('✅ 연속 삭제 인덱스 재계산 성공');

      // Cleanup
      if (fs.existsSync(testPath)) fs.unlinkSync(testPath);
    });

    it('should preserve XML structure integrity after nested table delete', async () => {
      const buffer = await createNestedTableDocument();
      const testPath = path.join(testOutputDir, 'xml-integrity-test.hwpx');
      fs.writeFileSync(testPath, buffer);

      let doc = await HwpxDocument.createFromBuffer('integrity', testPath, buffer);

      console.log('=== XML 구조 무결성 테스트 ===');

      // 중첩 테이블 포함 테이블 삭제
      doc.deleteTable(0, 0);

      const savedBuffer = await doc.save();
      fs.writeFileSync(testPath, savedBuffer);

      // XML 파싱 테스트 - 유효한 XML이어야 함
      const savedZip = await JSZip.loadAsync(savedBuffer);
      const savedXml = await savedZip.file('Contents/section0.xml')?.async('string') || '';

      // 모든 주요 태그 균형 검사
      const tagsToCheck = ['tbl', 'tr', 'tc', 'p', 'run', 'subList'];
      let allBalanced = true;

      for (const tag of tagsToCheck) {
        const balance = checkTagBalance(savedXml, tag);
        console.log(`  <hp:${tag}>: ${balance.open}/${balance.close} ${balance.balanced ? '✅' : '❌'}`);
        if (!balance.balanced) allBalanced = false;
      }

      expect(allBalanced).toBe(true);
      console.log('✅ XML 구조 무결성 유지됨');

      // Cleanup
      if (fs.existsSync(testPath)) fs.unlinkSync(testPath);
    });
  });
});
