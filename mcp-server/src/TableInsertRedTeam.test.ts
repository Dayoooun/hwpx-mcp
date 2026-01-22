/**
 * RED TEAM: 복잡한 테이블 삽입 테스트
 *
 * 테이블이 첫 번째 테이블 안에 삽입되는 버그를 재현하고 검증
 */

import { describe, it, expect, beforeEach } from 'vitest';
import { HwpxDocument } from './HwpxDocument';
import * as fs from 'fs';
import * as path from 'path';

// XML 태그 밸런스 검증 함수
function verifyTagBalance(xml: string): { balanced: boolean; details: Record<string, { open: number; close: number }> } {
  const tags = ['hp:tbl', 'hp:tr', 'hp:tc', 'hp:p', 'hp:run', 'hp:subList'];
  const details: Record<string, { open: number; close: number }> = {};

  for (const tag of tags) {
    const openMatches = xml.match(new RegExp(`<${tag}[\\s>]`, 'g')) || [];
    const closeMatches = xml.match(new RegExp(`</${tag}>`, 'g')) || [];
    details[tag] = { open: openMatches.length, close: closeMatches.length };
  }

  const balanced = Object.values(details).every(d => d.open === d.close);
  return { balanced, details };
}

// 테이블이 독립적인지 검증 (다른 테이블 안에 있지 않은지)
function verifyIndependentTables(xml: string, expectedCount: number): {
  independent: boolean;
  topLevelCount: number;
  nestedCount: number;
  message: string;
} {
  // 모든 테이블 시작 위치 찾기
  const tableStarts: number[] = [];
  let pos = 0;
  while ((pos = xml.indexOf('<hp:tbl', pos)) !== -1) {
    tableStarts.push(pos);
    pos++;
  }

  // 각 테이블이 top-level인지 확인
  let topLevelCount = 0;
  let nestedCount = 0;

  for (const startPos of tableStarts) {
    // 이 위치 이전의 텍스트에서 열린 subList와 닫힌 subList 개수 비교
    const beforeText = xml.substring(0, startPos);
    const subListOpens = (beforeText.match(/<hp:subList/g) || []).length;
    const subListCloses = (beforeText.match(/<\/hp:subList>/g) || []).length;

    if (subListOpens === subListCloses) {
      topLevelCount++;
    } else {
      nestedCount++;
    }
  }

  return {
    independent: topLevelCount === expectedCount,
    topLevelCount,
    nestedCount,
    message: `Top-level: ${topLevelCount}, Nested: ${nestedCount}, Expected top-level: ${expectedCount}`
  };
}

describe('RED TEAM - 복잡한 테이블 삽입', () => {
  let doc: HwpxDocument;
  const testFilePath = path.join(__dirname, '../test-fixtures/test-document.hwpx');

  beforeEach(async () => {
    // HwpxDocument.createNew()를 사용하여 새 문서 생성
    // 테스트 fixture 파일이 있으면 로드, 없으면 새 문서 생성
    if (fs.existsSync(testFilePath)) {
      const buffer = fs.readFileSync(testFilePath);
      doc = await HwpxDocument.createFromBuffer('test-doc', testFilePath, buffer);
    } else {
      console.log('테스트 파일 없음, 새 문서 생성');
      doc = HwpxDocument.createNew('test-doc', 'Table Insert Test');
    }
  });

  describe('연속 테이블 삽입', () => {
    it('2개의 독립 테이블을 연속으로 삽입', async () => {
      // 첫 번째 테이블 삽입
      const result1 = doc.insertTable(0, 0, 2, 2);
      expect(result1).not.toBeNull();
      console.log('첫 번째 테이블 삽입:', result1);

      // 두 번째 테이블 삽입 (첫 번째 테이블 뒤에)
      const result2 = doc.insertTable(0, 1, 3, 3);
      expect(result2).not.toBeNull();
      console.log('두 번째 테이블 삽입:', result2);

      // 저장
      const outputPath = path.join(__dirname, '../test-output/two-tables.hwpx');
      const savedBuffer = await doc.save();
      fs.writeFileSync(outputPath, savedBuffer);

      // XML 검증
      const reloadedBuffer = fs.readFileSync(outputPath);
      const savedDoc = await HwpxDocument.createFromBuffer('saved-doc', outputPath, reloadedBuffer);
      const xml = await savedDoc.getSectionXml(0);

      const balance = verifyTagBalance(xml);
      console.log('태그 밸런스:', balance.details);
      expect(balance.balanced).toBe(true);

      const independence = verifyIndependentTables(xml, 2);
      console.log('테이블 독립성:', independence.message);
      expect(independence.independent).toBe(true);
    });

    it('3개의 독립 테이블을 연속으로 삽입', async () => {
      doc.insertTable(0, 0, 2, 2);
      doc.insertTable(0, 1, 2, 3);
      doc.insertTable(0, 2, 3, 2);

      const outputPath = path.join(__dirname, '../test-output/three-tables.hwpx');
      const savedBuffer = await doc.save();
      fs.writeFileSync(outputPath, savedBuffer);

      const reloadedBuffer = fs.readFileSync(outputPath);
      const savedDoc = await HwpxDocument.createFromBuffer('saved-doc', outputPath, reloadedBuffer);
      const xml = await savedDoc.getSectionXml(0);

      const balance = verifyTagBalance(xml);
      console.log('3개 테이블 밸런스:', balance.details);
      expect(balance.balanced).toBe(true);

      const independence = verifyIndependentTables(xml, 3);
      console.log('3개 테이블 독립성:', independence.message);
      expect(independence.independent).toBe(true);
    });

    it('5개의 독립 테이블을 연속으로 삽입', async () => {
      for (let i = 0; i < 5; i++) {
        const result = doc.insertTable(0, i, 2 + i, 2 + i);
        expect(result).not.toBeNull();
      }

      const outputPath = path.join(__dirname, '../test-output/five-tables.hwpx');
      const savedBuffer = await doc.save();
      fs.writeFileSync(outputPath, savedBuffer);

      const reloadedBuffer = fs.readFileSync(outputPath);
      const savedDoc = await HwpxDocument.createFromBuffer('saved-doc', outputPath, reloadedBuffer);
      const xml = await savedDoc.getSectionXml(0);

      const balance = verifyTagBalance(xml);
      console.log('5개 테이블 밸런스:', balance.details);
      expect(balance.balanced).toBe(true);

      const independence = verifyIndependentTables(xml, 5);
      console.log('5개 테이블 독립성:', independence.message);
      expect(independence.independent).toBe(true);
    });
  });

  describe('테이블 + 내용 수정 후 추가 테이블', () => {
    it('테이블 생성 → 셀 수정 → 새 테이블 생성', async () => {
      // 첫 번째 테이블
      const result1 = doc.insertTable(0, 0, 2, 2);
      expect(result1).not.toBeNull();

      // 첫 번째 테이블 셀 수정
      doc.updateTableCell(0, result1!.tableIndex, 0, 0, '표1 셀1');
      doc.updateTableCell(0, result1!.tableIndex, 0, 1, '표1 셀2');

      // 두 번째 테이블
      const result2 = doc.insertTable(0, 1, 3, 3);
      expect(result2).not.toBeNull();

      // 두 번째 테이블 셀 수정
      doc.updateTableCell(0, result2!.tableIndex, 0, 0, '표2 셀1');

      const outputPath = path.join(__dirname, '../test-output/table-edit-table.hwpx');
      const savedBuffer = await doc.save();
      fs.writeFileSync(outputPath, savedBuffer);

      const reloadedBuffer = fs.readFileSync(outputPath);
      const savedDoc = await HwpxDocument.createFromBuffer('saved-doc', outputPath, reloadedBuffer);
      const xml = await savedDoc.getSectionXml(0);

      const balance = verifyTagBalance(xml);
      expect(balance.balanced).toBe(true);

      const independence = verifyIndependentTables(xml, 2);
      console.log('테이블+수정+테이블:', independence.message);
      expect(independence.independent).toBe(true);

      // 내용 검증
      expect(xml).toContain('표1 셀1');
      expect(xml).toContain('표2 셀1');
    });
  });

  describe('중첩 테이블이 있는 상황', () => {
    it('중첩 테이블 포함 문서에 새 독립 테이블 추가', async () => {
      // 첫 번째 테이블 (나중에 중첩 테이블 추가됨)
      const result1 = doc.insertTable(0, 0, 3, 3);
      expect(result1).not.toBeNull();

      // 중첩 테이블 추가
      const nestedResult = doc.insertNestedTable(0, result1!.tableIndex, 1, 1, 2, 2);
      expect(nestedResult.success).toBe(true);

      // 독립적인 두 번째 테이블 추가
      const result2 = doc.insertTable(0, 1, 2, 2);
      expect(result2).not.toBeNull();

      const outputPath = path.join(__dirname, '../test-output/nested-plus-independent.hwpx');
      const savedBuffer = await doc.save();
      fs.writeFileSync(outputPath, savedBuffer);

      const reloadedBuffer = fs.readFileSync(outputPath);
      const savedDoc = await HwpxDocument.createFromBuffer('saved-doc', outputPath, reloadedBuffer);
      const xml = await savedDoc.getSectionXml(0);

      const balance = verifyTagBalance(xml);
      console.log('중첩+독립 테이블 밸런스:', balance.details);
      expect(balance.balanced).toBe(true);

      // 총 3개 테이블 (메인1 + 중첩1 + 독립1)
      // top-level은 2개여야 함
      const independence = verifyIndependentTables(xml, 2);
      console.log('중첩+독립 테이블 독립성:', independence.message);
      expect(independence.topLevelCount).toBe(2);
      expect(independence.nestedCount).toBe(1);
    });
  });

  describe('복잡한 시나리오', () => {
    it('10개 테이블을 번갈아가며 생성하고 수정', async () => {
      const tableResults: Array<{ tableIndex: number }> = [];

      // 10개 테이블 생성
      for (let i = 0; i < 10; i++) {
        const result = doc.insertTable(0, i, 2, 3);
        expect(result).not.toBeNull();
        tableResults.push(result!);
      }

      // 각 테이블 셀 수정
      for (let i = 0; i < 10; i++) {
        doc.updateTableCell(0, tableResults[i].tableIndex, 0, 0, `테이블${i + 1} 헤더`);
        doc.updateTableCell(0, tableResults[i].tableIndex, 1, 0, `테이블${i + 1} 데이터`);
      }

      const outputPath = path.join(__dirname, '../test-output/ten-tables-complex.hwpx');
      const savedBuffer = await doc.save();
      fs.writeFileSync(outputPath, savedBuffer);

      const reloadedBuffer = fs.readFileSync(outputPath);
      const savedDoc = await HwpxDocument.createFromBuffer('saved-doc', outputPath, reloadedBuffer);
      const xml = await savedDoc.getSectionXml(0);

      const balance = verifyTagBalance(xml);
      console.log('10개 테이블 밸런스:', balance.details);
      expect(balance.balanced).toBe(true);

      const independence = verifyIndependentTables(xml, 10);
      console.log('10개 테이블 독립성:', independence.message);
      expect(independence.independent).toBe(true);

      // 모든 테이블 내용 검증
      for (let i = 1; i <= 10; i++) {
        expect(xml).toContain(`테이블${i} 헤더`);
        expect(xml).toContain(`테이블${i} 데이터`);
      }
    });

    it('테이블 크기가 다양한 경우 (1x1 ~ 10x10)', async () => {
      const sizes = [[1, 1], [2, 3], [3, 2], [4, 4], [5, 3], [2, 6], [6, 2], [3, 5], [5, 5], [10, 10]];

      for (let i = 0; i < sizes.length; i++) {
        const [rows, cols] = sizes[i];
        const result = doc.insertTable(0, i, rows, cols);
        expect(result).not.toBeNull();

        // 첫 셀에 크기 정보 기록
        doc.updateTableCell(0, result!.tableIndex, 0, 0, `${rows}x${cols}`);
      }

      const outputPath = path.join(__dirname, '../test-output/various-sizes.hwpx');
      const savedBuffer = await doc.save();
      fs.writeFileSync(outputPath, savedBuffer);

      const reloadedBuffer = fs.readFileSync(outputPath);
      const savedDoc = await HwpxDocument.createFromBuffer('saved-doc', outputPath, reloadedBuffer);
      const xml = await savedDoc.getSectionXml(0);

      const balance = verifyTagBalance(xml);
      console.log('다양한 크기 테이블 밸런스:', balance.details);
      expect(balance.balanced).toBe(true);

      const independence = verifyIndependentTables(xml, 10);
      console.log('다양한 크기 테이블 독립성:', independence.message);
      expect(independence.independent).toBe(true);
    });

    it('중첩 테이블이 여러 개 있는 복잡한 구조', async () => {
      // 3개의 메인 테이블
      const table1 = doc.insertTable(0, 0, 4, 4);
      const table2 = doc.insertTable(0, 1, 3, 3);
      const table3 = doc.insertTable(0, 2, 5, 5);

      expect(table1).not.toBeNull();
      expect(table2).not.toBeNull();
      expect(table3).not.toBeNull();

      // 각 테이블에 중첩 테이블 추가
      doc.insertNestedTable(0, table1!.tableIndex, 1, 1, 2, 2);
      doc.insertNestedTable(0, table2!.tableIndex, 1, 1, 2, 2);
      doc.insertNestedTable(0, table3!.tableIndex, 2, 2, 3, 3);

      // 4번째 독립 테이블 추가
      const table4 = doc.insertTable(0, 3, 2, 2);
      expect(table4).not.toBeNull();

      const outputPath = path.join(__dirname, '../test-output/multi-nested-complex.hwpx');
      const savedBuffer = await doc.save();
      fs.writeFileSync(outputPath, savedBuffer);

      const reloadedBuffer = fs.readFileSync(outputPath);
      const savedDoc = await HwpxDocument.createFromBuffer('saved-doc', outputPath, reloadedBuffer);
      const xml = await savedDoc.getSectionXml(0);

      const balance = verifyTagBalance(xml);
      console.log('복잡한 중첩 구조 밸런스:', balance.details);
      expect(balance.balanced).toBe(true);

      // 총 7개 테이블: 4개 메인 + 3개 중첩
      // top-level은 4개
      const independence = verifyIndependentTables(xml, 4);
      console.log('복잡한 중첩 구조 독립성:', independence.message);
      expect(independence.topLevelCount).toBe(4);
      expect(independence.nestedCount).toBe(3);
    });
  });

  describe('엣지 케이스', () => {
    it('첫 요소 위치에 테이블 삽입', async () => {
      const result = doc.insertTable(0, -1, 2, 2);
      // -1은 맨 처음을 의미할 수 있음
      console.log('첫 위치 삽입 결과:', result);

      const outputPath = path.join(__dirname, '../test-output/first-position.hwpx');
      const savedBuffer = await doc.save();
      fs.writeFileSync(outputPath, savedBuffer);

      const reloadedBuffer = fs.readFileSync(outputPath);
      const savedDoc = await HwpxDocument.createFromBuffer('saved-doc', outputPath, reloadedBuffer);
      const xml = await savedDoc.getSectionXml(0);

      const balance = verifyTagBalance(xml);
      expect(balance.balanced).toBe(true);
    });

    it('매우 큰 테이블 (20x20) 생성', async () => {
      const result = doc.insertTable(0, 0, 20, 20);
      expect(result).not.toBeNull();

      // 일부 셀 수정
      for (let i = 0; i < 20; i++) {
        doc.updateTableCell(0, result!.tableIndex, i, 0, `Row ${i}`);
      }

      const outputPath = path.join(__dirname, '../test-output/large-table.hwpx');
      const savedBuffer = await doc.save();
      fs.writeFileSync(outputPath, savedBuffer);

      const reloadedBuffer = fs.readFileSync(outputPath);
      const savedDoc = await HwpxDocument.createFromBuffer('saved-doc', outputPath, reloadedBuffer);
      const xml = await savedDoc.getSectionXml(0);

      const balance = verifyTagBalance(xml);
      console.log('대형 테이블 밸런스:', balance.details);
      expect(balance.balanced).toBe(true);

      // 행/열 개수 검증
      const trCount = (xml.match(/<hp:tr>/g) || []).length;
      const tcCount = (xml.match(/<hp:tc /g) || []).length;
      console.log(`20x20 테이블: ${trCount} rows, ${tcCount} cells`);

      // 최소 20행, 400셀 이상
      expect(trCount).toBeGreaterThanOrEqual(20);
      expect(tcCount).toBeGreaterThanOrEqual(400);
    });

    it('빠른 연속 삽입 (20개 테이블)', async () => {
      const startTime = Date.now();

      for (let i = 0; i < 20; i++) {
        const result = doc.insertTable(0, i, 2, 2);
        expect(result).not.toBeNull();
      }

      const insertTime = Date.now() - startTime;
      console.log(`20개 테이블 삽입 시간: ${insertTime}ms`);

      const outputPath = path.join(__dirname, '../test-output/rapid-insert.hwpx');
      const savedBuffer = await doc.save();
      fs.writeFileSync(outputPath, savedBuffer);

      const reloadedBuffer = fs.readFileSync(outputPath);
      const savedDoc = await HwpxDocument.createFromBuffer('saved-doc', outputPath, reloadedBuffer);
      const xml = await savedDoc.getSectionXml(0);

      const balance = verifyTagBalance(xml);
      expect(balance.balanced).toBe(true);

      const independence = verifyIndependentTables(xml, 20);
      console.log('20개 빠른 삽입 독립성:', independence.message);
      expect(independence.independent).toBe(true);
    });
  });

  describe('RED TEAM - Save/Load 텍스트 손실 버그', () => {
    it('새 테이블 삽입 후 셀 업데이트, 저장, 리로드 시 텍스트 유지', async () => {
      // Step 1: 새 테이블 삽입
      const insertResult = doc.insertTable(0, 0, 3, 3);
      expect(insertResult).not.toBeNull();
      const tableIndex = insertResult!.tableIndex;
      console.log('삽입된 테이블 인덱스:', tableIndex);

      // Step 2: 셀에 텍스트 입력
      const testTexts = [
        { row: 0, col: 0, text: '제1장 서론' },
        { row: 0, col: 1, text: '테스트 데이터' },
        { row: 1, col: 0, text: '가. 배경' },
        { row: 1, col: 1, text: '상세 내용입니다' },
        { row: 2, col: 2, text: '결론' },
      ];

      for (const { row, col, text } of testTexts) {
        const result = doc.updateTableCell(0, tableIndex, row, col, text);
        expect(result).toBe(true);
        console.log(`셀 (${row},${col}) 업데이트: "${text}"`);
      }

      // Step 3: 메모리에서 텍스트 확인 (getTableCell 사용)
      console.log('저장 전 테이블 내용:');
      for (const { row, col, text } of testTexts) {
        const cell = doc.getTableCell(0, tableIndex, row, col);
        const cellText = cell?.text || '';
        console.log(`  (${row},${col}): "${cellText}" (기대: "${text}")`);
        expect(cellText).toBe(text);
      }

      // Step 4: 저장
      const outputPath = path.join(__dirname, '../test-output/saveload-text-bug.hwpx');
      const savedBuffer = await doc.save();
      fs.writeFileSync(outputPath, savedBuffer);
      console.log('파일 저장 완료:', outputPath);

      // Step 5: XML에서 텍스트 확인
      const xml = await doc.getSectionXml(0);
      console.log('저장된 XML에서 텍스트 검색:');
      for (const { text } of testTexts) {
        const found = xml.includes(text);
        console.log(`  "${text}": ${found ? '✓ 발견' : '✗ 없음'}`);
        expect(found).toBe(true); // 이 단계에서 실패하면 XML 저장 버그
      }

      // Step 6: 새 문서로 리로드
      const reloadedBuffer = fs.readFileSync(outputPath);
      const reloadedDoc = await HwpxDocument.createFromBuffer('reloaded-doc', outputPath, reloadedBuffer);
      console.log('문서 리로드 완료');

      // Step 7: 리로드된 문서에서 테이블 찾기
      const tables = reloadedDoc.getTables();
      console.log('리로드된 문서의 테이블 수:', tables.length);
      expect(tables.length).toBeGreaterThan(0);

      // Step 8: 리로드된 테이블에서 텍스트 확인 (getTableCell 사용)
      console.log('리로드된 테이블 내용:');
      for (const { row, col, text } of testTexts) {
        const cell = reloadedDoc.getTableCell(0, 0, row, col);
        const cellText = cell?.text || '';
        console.log(`  (${row},${col}): "${cellText}" (기대: "${text}")`);
        expect(cellText).toBe(text); // 이 단계에서 실패하면 파싱 버그
      }

      // Step 9: 테이블 구조 확인
      const reloadedTable = reloadedDoc.getTable(0, 0);
      expect(reloadedTable?.rows).toBe(3);
      expect(reloadedTable?.cols).toBe(3);
    });

    it('기존 테이블의 셀 업데이트 후 저장/리로드 시 텍스트 유지', async () => {
      // 이 테스트는 기존 테이블(파서로 로드된)의 셀을 업데이트하는 경우를 테스트
      // 먼저 테이블을 삽입하고 저장한 후, 다시 로드하여 업데이트 테스트

      // Step 1: 새 테이블 삽입 및 저장
      const insertResult = doc.insertTable(0, 0, 2, 2);
      expect(insertResult).not.toBeNull();

      const outputPath = path.join(__dirname, '../test-output/existing-table-test.hwpx');
      let savedBuffer = await doc.save();
      fs.writeFileSync(outputPath, savedBuffer);

      // Step 2: 새 문서로 로드 (이제 "기존" 테이블이 됨)
      const reloadedBuffer = fs.readFileSync(outputPath);
      const existingDoc = await HwpxDocument.createFromBuffer('existing-doc', outputPath, reloadedBuffer);

      // Step 3: 기존 테이블의 셀 업데이트
      const updateResult = existingDoc.updateTableCell(0, 0, 0, 0, '기존 테이블 업데이트 테스트');
      expect(updateResult).toBe(true);
      console.log('기존 테이블 셀 업데이트 완료');

      // Step 4: 다시 저장
      const outputPath2 = path.join(__dirname, '../test-output/existing-table-updated.hwpx');
      savedBuffer = await existingDoc.save();
      fs.writeFileSync(outputPath2, savedBuffer);

      // Step 5: XML에서 텍스트 확인
      const xml = await existingDoc.getSectionXml(0);
      const found = xml.includes('기존 테이블 업데이트 테스트');
      console.log('기존 테이블 업데이트 텍스트 XML 저장:', found ? '✓ 성공' : '✗ 실패');
      expect(found).toBe(true);

      // Step 6: 세 번째 로드
      const finalBuffer = fs.readFileSync(outputPath2);
      const finalDoc = await HwpxDocument.createFromBuffer('final-doc', outputPath2, finalBuffer);

      const finalCell = finalDoc.getTableCell(0, 0, 0, 0);
      const finalText = finalCell?.text || '';
      console.log('최종 리로드된 텍스트:', finalText);
      expect(finalText).toBe('기존 테이블 업데이트 테스트');
    });

    it('테이블 ID 동기화 검증', async () => {
      // 새 테이블 삽입 후 ID가 XML에 올바르게 저장되는지 검증
      const insertResult = doc.insertTable(0, 0, 2, 2);
      expect(insertResult).not.toBeNull();

      // _pendingTableInserts에서 테이블 ID 가져오기 (테스트 목적으로 private 접근)
      const pendingInserts = (doc as any)._pendingTableInserts;
      expect(pendingInserts.length).toBeGreaterThan(0);
      const memoryTableId = pendingInserts[pendingInserts.length - 1]?.tableId;
      console.log('메모리 테이블 ID (from _pendingTableInserts):', memoryTableId);
      expect(memoryTableId).toBeDefined();

      // 저장
      const savedBuffer = await doc.save();

      // XML에서 테이블 ID 확인
      const xml = await doc.getSectionXml(0);
      const idMatch = xml.match(/<hp:tbl[^>]*id="([^"]+)"/);
      const xmlTableId = idMatch?.[1];
      console.log('XML 테이블 ID:', xmlTableId);

      // ID가 일치해야 함
      expect(xmlTableId).toBe(memoryTableId);
    });
  });
});
