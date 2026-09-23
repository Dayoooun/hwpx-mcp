/**
 * 한/글 실측 렌더 기준 회귀 테스트.
 *
 * 기댓값은 추정이 아니라 Windows 한/글이 그린 PDF 에서 잰 좌표다.
 * 측정 방법과 원본은 src/fixtures/hancom-marker-widths.json 에 기록했다.
 */
import { describe, it, expect, beforeEach, afterEach } from 'vitest';
import * as fs from 'fs';
import * as os from 'os';
import * as path from 'path';
import JSZip from 'jszip';
import { HwpxDocument } from './HwpxDocument';
import { HangingIndentCalculator } from './HangingIndentCalculator';

const measured: { widths_pt: Record<string, number> } = JSON.parse(
  fs.readFileSync(path.join(__dirname, 'fixtures', 'hancom-marker-widths.json'), 'utf8')
);

let workDir: string;
beforeEach(() => { workDir = fs.mkdtempSync(path.join(os.tmpdir(), 'hwpx-metrics-')); });
afterEach(() => { fs.rmSync(workDir, { recursive: true, force: true }); });

async function sectionXml(buf: Buffer): Promise<string> {
  const file = (await JSZip.loadAsync(buf)).file('Contents/section0.xml');
  if (!file) throw new Error('section0.xml missing');
  return file.async('string');
}

describe('내어쓰기 폭이 한/글 실측과 일치한다', () => {
  const calc = new HangingIndentCalculator();

  // 허용 오차 0.5pt ≈ 0.18mm. 인쇄물에서 눈으로 구분되지 않는 수준이다.
  // 수정 전 계산기는 평균 +4.14pt, 최대 +6.30pt 어긋났다.
  for (const [marker, widthPt] of Object.entries(measured.widths_pt)) {
    it(`${JSON.stringify(marker)} → ${widthPt.toFixed(2)}pt (함초롬바탕 10pt)`, () => {
      const got = calc.calculateHangingIndent(`${marker}본문`, 10, '함초롬바탕');
      expect(Math.abs(got - widthPt)).toBeLessThan(0.5);
    });
  }

  it('폰트 크기에 비례한다', () => {
    const at10 = calc.calculateHangingIndent('○ 본문', 10, '함초롬바탕');
    const at12 = calc.calculateHangingIndent('○ 본문', 12, '함초롬바탕');
    expect(at12 / at10).toBeCloseTo(1.2, 3);
  });
});

describe('표 안 표가 부모 셀 안에 들어간다', () => {
  /** 부모 셀 폭(hwpunit)과 그 안 중첩 표의 전체 폭·셀 폭 합을 꺼낸다. */
  function nestedGeometry(xml: string) {
    const outerStart = xml.indexOf('<hp:tbl');
    const innerStart = xml.indexOf('<hp:tbl', outerStart + 1);
    const innerEnd = xml.indexOf('</hp:tbl>', innerStart);
    const inner = xml.slice(innerStart, innerEnd);

    // 중첩 표를 품은 부모 셀: 중첩 표 뒤에 오는 첫 부모 cellSz/cellMargin
    const afterInner = xml.slice(innerEnd);
    const parentSz = afterInner.match(/<hp:cellSz width="(\d+)"/);
    const parentTcOpen = xml.lastIndexOf('<hp:tc ', innerStart);
    const parentUsesOwnMargin = /\bhasMargin="1"/.test(xml.slice(parentTcOpen, xml.indexOf('>', parentTcOpen)));
    // 한/글은 hasMargin="0" 인 셀의 cellMargin 을 무시하고 표의 inMargin 을 쓴다.
    // 실측: A 표(hasMargin=0, inMargin=0)는 cellMargin 141 이 있어도 글자가 셀 선에서 0pt.
    const outerTable = xml.slice(outerStart, innerStart);
    const parentMargin = parentUsesOwnMargin
      ? afterInner.match(/<hp:cellMargin left="(\d+)" right="(\d+)"/)
      : outerTable.match(/<hp:inMargin left="(\d+)" right="(\d+)"/);
    const tableSz = inner.match(/<hp:sz width="(\d+)"/);
    const outMargin = inner.match(/<hp:outMargin left="(\d+)" right="(\d+)"/);
    const firstRow = inner.slice(inner.indexOf('<hp:tr>'), inner.indexOf('</hp:tr>'));
    const cellWidths = [...firstRow.matchAll(/<hp:cellSz width="(\d+)"/g)].map(m => Number(m[1]));

    return {
      parentWidth: Number(parentSz![1]),
      parentPadding: parentMargin ? Number(parentMargin[1]) + Number(parentMargin[2]) : 0,
      tableWidth: Number(tableSz![1]),
      outMargin: Number(outMargin![1]) + Number(outMargin![2]),
      cellWidthSum: cellWidths.reduce((a, b) => a + b, 0),
      cellCount: cellWidths.length,
    };
  }

  it('3열 중첩 표 폭이 부모 셀 안쪽 폭을 넘지 않는다', async () => {
    const doc = HwpxDocument.createNew('n1', 'nested');
    doc.insertTable(0, 0, 2, 2);
    const result = doc.insertNestedTable(0, 0, 1, 1, 3, 3, [
      ['품목', '수량', '금액'],
      ['간판 교체', '1식', '1,200,000원'],
    ]);
    expect(result.success).toBe(true);

    const g = nestedGeometry(await sectionXml(await doc.save()));
    // 실측: 수정 전 중첩 표 24000 > 부모 셀 21260. 금액 열이 부모 테두리 밖으로 나갔다.
    expect(g.tableWidth + g.outMargin).toBeLessThanOrEqual(g.parentWidth - g.parentPadding);
  });

  it('중첩 표 셀 폭의 합이 표 폭과 같다', async () => {
    const doc = HwpxDocument.createNew('n2', 'nested');
    doc.insertTable(0, 0, 2, 2);
    doc.insertNestedTable(0, 0, 1, 1, 2, 4);

    const g = nestedGeometry(await sectionXml(await doc.save()));
    expect(g.cellCount).toBe(4);
    // 한 hwpunit 의 반올림 차이까지만 허용한다.
    expect(Math.abs(g.cellWidthSum - g.tableWidth)).toBeLessThanOrEqual(g.cellCount);
  });

  it('부모 셀이 좁아도 중첩 표가 따라 줄어든다', async () => {
    const doc = HwpxDocument.createNew('n3', 'nested');
    doc.insertTable(0, 0, 2, 4); // 부모 셀 폭 ≈ 42520 / 4
    doc.insertNestedTable(0, 0, 1, 1, 2, 3);

    const g = nestedGeometry(await sectionXml(await doc.save()));
    expect(g.tableWidth + g.outMargin).toBeLessThanOrEqual(g.parentWidth - g.parentPadding);
  });

  it('중첩 표 앞에 빈 줄을 만들지 않는다', async () => {
    const doc = HwpxDocument.createNew('n4', 'nested');
    doc.insertTable(0, 0, 2, 2);
    doc.insertNestedTable(0, 0, 1, 1, 2, 2);

    const xml = await sectionXml(await doc.save());
    const inner = xml.indexOf('<hp:tbl', xml.indexOf('<hp:tbl') + 1);
    const runStart = xml.lastIndexOf('<hp:run', inner);
    // 실측: 수정 전 `<hp:t> </hp:t>` 공백 한 칸이 표 앞에 들어가 한 줄이 비었다.
    expect(xml.slice(runStart, inner)).not.toMatch(/<hp:t>\s+<\/hp:t>/);
  });
});
