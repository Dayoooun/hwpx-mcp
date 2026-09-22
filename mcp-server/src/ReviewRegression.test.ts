/**
 * 2026-09-22 외부 리뷰 회귀 테스트
 *
 * 출처: hwpx-mcp-review.md (개선점 1~8)
 * 각 테스트는 리뷰에서 보고된 증상을 그대로 재현한다.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import JSZip from 'jszip';
import { HwpxDocument } from './HwpxDocument';

let workDir: string;

beforeEach(() => {
  workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'hwpx-review-'));
});

afterEach(() => {
  fs.rmSync(workDir, { recursive: true, force: true });
});

async function sectionXml(filePath: string, index = 0): Promise<string> {
  const zip = await JSZip.loadAsync(fs.readFileSync(filePath));
  const file = zip.file(`Contents/section${index}.xml`);
  if (!file) throw new Error(`section${index}.xml missing`);
  return file.async('string');
}

function countTag(xml: string, tag: string): { open: number; close: number } {
  const open = (xml.match(new RegExp(`<${tag}(?=[\\s/>])`, 'g')) || []).length;
  const close = (xml.match(new RegExp(`</${tag}>`, 'g')) || []).length;
  return { open, close };
}

async function saveTo(doc: HwpxDocument, filePath: string): Promise<void> {
  fs.writeFileSync(filePath, await doc.save());
}

async function reopen(id: string, filePath: string): Promise<HwpxDocument> {
  return HwpxDocument.createFromBuffer(id, filePath, fs.readFileSync(filePath));
}

describe('개선점 3: insert_table_row가 저장본 XML을 깨뜨린다', () => {
  it('행 삽입 후 저장해도 hp:tc 여닫기 짝이 맞는다', async () => {
    const doc = HwpxDocument.createNew('d1', 'row-insert');
    doc.insertParagraph(0, -1, '본문');
    expect(doc.insertTable(0, 0, 3, 3)).not.toBeNull();
    expect(doc.insertTableRow(0, 0, 0)).toBe(true);

    const out = path.join(workDir, 'row.hwpx');
    await saveTo(doc, out);

    const xml = await sectionXml(out);
    const tc = countTag(xml, 'hp:tc');
    const tr = countTag(xml, 'hp:tr');

    // 리뷰 증상: 닫는 태그가 남아 짝이 어긋났다
    expect(tc.open).toBe(tc.close);
    expect(tr.open).toBe(tr.close);
  });

  it('삽입된 행이 subList/p/run 구조를 그대로 보존한다', async () => {
    const doc = HwpxDocument.createNew('d2', 'row-structure');
    doc.insertParagraph(0, -1, '본문');
    doc.insertTable(0, 0, 2, 3);
    doc.insertTableRow(0, 0, 0);

    const out = path.join(workDir, 'row2.hwpx');
    await saveTo(doc, out);

    const xml = await sectionXml(out);
    const tbl = xml.slice(xml.indexOf('<hp:tbl'), xml.indexOf('</hp:tbl>') + 9);

    // 각 행의 셀 개수가 동일해야 한다 (구조 손실이 없다)
    const rows = tbl.split('</hp:tr>').filter(r => r.includes('<hp:tr'));
    const cellCounts = rows.map(r => (r.match(/<hp:tc(?=[\s/>])/g) || []).length);
    expect(new Set(cellCounts).size).toBe(1);
    expect(cellCounts[0]).toBe(3);

    // subList 여닫기도 어긋나면 안 된다
    const sub = countTag(tbl, 'hp:subList');
    expect(sub.open).toBe(sub.close);
  });

  it('cell_texts로 넘긴 글자가 삽입된 행에 들어간다', async () => {
    const doc = HwpxDocument.createNew('d3', 'row-texts');
    doc.insertParagraph(0, -1, '본문');
    doc.insertTable(0, 0, 2, 3);
    doc.insertTableRow(0, 0, 0, ['가', '나', '다']);

    const out = path.join(workDir, 'row3.hwpx');
    await saveTo(doc, out);

    const xml = await sectionXml(out);
    expect(xml).toContain('>가<');
    expect(xml).toContain('>나<');
    expect(xml).toContain('>다<');
  });
});

describe('개선점 3: insert_table_column도 같은 결함을 공유한다', () => {
  it('열 삽입 후 저장해도 hp:tc 여닫기 짝이 맞는다', async () => {
    const doc = HwpxDocument.createNew('d4', 'col-insert');
    doc.insertParagraph(0, -1, '본문');
    doc.insertTable(0, 0, 2, 2);
    expect(doc.insertTableColumn(0, 0, 0)).toBe(true);

    const out = path.join(workDir, 'col.hwpx');
    await saveTo(doc, out);

    const xml = await sectionXml(out);
    const tc = countTag(xml, 'hp:tc');
    expect(tc.open).toBe(tc.close);

    const sub = countTag(xml, 'hp:subList');
    expect(sub.open).toBe(sub.close);
  });
});

describe('개선점 4: copy_paragraph 직후 교체하면 원본까지 바뀐다', () => {
  it('중간 저장 없이 복제→교체해도 원본 문단은 그대로다', async () => {
    const base = HwpxDocument.createNew('b', 'base');
    base.insertParagraph(0, -1, '원본 문단입니다');
    const basePath = path.join(workDir, 'base.hwpx');
    await saveTo(base, basePath);

    const doc = await reopen('d5', basePath);
    const before = doc.getParagraphs(0);
    const srcIndex = before.findIndex(p => p.text === '원본 문단입니다');
    expect(srcIndex).toBeGreaterThanOrEqual(0);

    // 복제 → 중간 저장 없이 즉시 교체
    expect(doc.copyParagraph(0, srcIndex, 0, srcIndex)).toBe(true);
    doc.updateParagraphText(0, srcIndex + 1, 0, '새 문구로 교체');

    const out = path.join(workDir, 'copy.hwpx');
    await saveTo(doc, out);

    const reopened = await reopen('d6', out);
    const texts = reopened.getParagraphs(0).map(p => p.text);

    // 리뷰 증상: 새 문구가 2곳, 원본 소실
    expect(texts.filter(t => t === '새 문구로 교체')).toHaveLength(1);
    expect(texts).toContain('원본 문단입니다');
  });
});

describe('개선점 2: 저장 경로가 응답에 절대경로로 돌아온다', () => {
  it('createNew 문서의 path는 저장 전까지 신뢰할 수 없음을 표시한다', () => {
    const doc = HwpxDocument.createNew('d7', 'path');
    // 새 문서는 아직 디스크 위치가 없다. 상대 파일명을 흘리면 안 된다.
    expect(doc.path).toBe('');
  });
});

describe('개선점 3: 병합으로 덮인 셀에 쓰면 성공이라 답하고 저장 시 사라진다', () => {
  it('덮인 셀 쓰기는 성공 대신 마스터 셀을 알려주며 거부한다', () => {
    const doc = HwpxDocument.createNew('d8', 'merged');
    doc.insertParagraph(0, -1, '표');
    doc.insertTable(0, 0, 3, 3);
    expect(doc.mergeCells(0, 0, 0, 0, 0, 1)).toBe(true);

    // 리뷰 증상: {"message":"Cell updated"} 뒤 저장본엔 글자 없음
    expect(() => doc.updateTableCell(0, 0, 0, 1, '덮인셀')).toThrow(/covered by the merged cell at \(0, 0\)/);
  });

  it('마스터 셀 쓰기는 정상이고 저장본에 남는다', async () => {
    const doc = HwpxDocument.createNew('d9', 'merged-master');
    doc.insertParagraph(0, -1, '표');
    doc.insertTable(0, 0, 3, 3);
    doc.mergeCells(0, 0, 0, 0, 0, 1);
    expect(doc.updateTableCell(0, 0, 0, 0, '마스터')).toBe(true);

    const out = path.join(workDir, 'merged.hwpx');
    await saveTo(doc, out);
    expect(await sectionXml(out)).toContain('>마스터<');
  });

  it('병합이 없으면 어떤 셀이든 쓸 수 있다', () => {
    const doc = HwpxDocument.createNew('d10', 'plain');
    doc.insertParagraph(0, -1, '표');
    doc.insertTable(0, 0, 3, 3);
    for (let r = 0; r < 3; r++) {
      for (let c = 0; c < 3; c++) {
        expect(doc.updateTableCell(0, 0, r, c, `R${r}C${c}`)).toBe(true);
      }
    }
  });

  it('세로 병합에 덮인 셀도 거부한다', () => {
    const doc = HwpxDocument.createNew('d11', 'vmerge');
    doc.insertParagraph(0, -1, '표');
    doc.insertTable(0, 0, 3, 3);
    doc.mergeCells(0, 0, 0, 0, 1, 0); // (0,0)~(1,0) 세로 병합

    expect(() => doc.updateTableCell(0, 0, 1, 0, '덮인셀')).toThrow(/covered by the merged cell at \(0, 0\)/);
    // 병합 범위 밖은 영향 없다
    expect(doc.updateTableCell(0, 0, 2, 0, '정상')).toBe(true);
    expect(doc.updateTableCell(0, 0, 1, 1, '정상')).toBe(true);
  });
});
