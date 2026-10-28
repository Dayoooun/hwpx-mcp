/**
 * 내어쓰기를 걸 때 새로 만드는 문단 모양이 양식의 줄 간격을 지키는지.
 * 실측 결함: 예비창업패키지 공식 양식(줄 간격 130%) 칸에 내어쓰기를 걸자 160%로 바뀌어
 * 표가 다음 쪽으로 밀렸다.
 */
import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { HwpxDocument } from './HwpxDocument';
import { HangingIndentCalculator } from './HangingIndentCalculator';

function paraPr(id: number, spacing: number, hancomSwitch: boolean): string {
  const margin = `<hh:margin><hc:intent value="0" unit="HWPUNIT"/><hc:left value="0" unit="HWPUNIT"/><hc:right value="0" unit="HWPUNIT"/><hc:prev value="0" unit="HWPUNIT"/><hc:next value="0" unit="HWPUNIT"/></hh:margin>`;
  const body = hancomSwitch
    // 한/글이 저장한 파일 모양: hp:switch 안에 case(HwpUnitChar)와 default 두 벌
    ? `<hp:switch><hp:case hp:required-namespace="http://www.hancom.co.kr/hwpml/2016/HwpUnitChar">${margin}<hh:lineSpacing type="PERCENT" value="${spacing}" unit="HWPUNIT"/></hp:case><hp:default>${margin}<hh:lineSpacing type="PERCENT" value="${spacing}" unit="HWPUNIT"/></hp:default></hp:switch>`
    : `${margin}<hh:lineSpacing type="PERCENT" value="${spacing}" unit="HWPUNIT"/>`;
  return `<hh:paraPr id="${id}" tabPrIDRef="0"><hh:align horizontal="JUSTIFY" vertical="BASELINE"/>${body}</hh:paraPr>`;
}

async function makeDoc(): Promise<HwpxDocument> {
  const zip = new JSZip();
  zip.file('mimetype', 'application/hwp+zip');
  zip.file('version.xml', '<?xml version="1.0" encoding="UTF-8"?><hh:version xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" major="1" minor="0"/>');
  zip.file('Contents/content.hpf', `<?xml version="1.0" encoding="UTF-8"?><hpf:package xmlns:hpf="urn:hancom:hwp:file"><hpf:manifest><hpf:item href="Contents/section0.xml" type="application/xml"/><hpf:item href="Contents/header.xml" type="application/xml"/></hpf:manifest></hpf:package>`);
  zip.file('Contents/header.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" xmlns:hc="http://www.hancom.co.kr/hwpml/2011/core" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hh:refList>
    <hh:charProperties itemCnt="1"><hh:charPr id="0" height="1000" textColor="#000000"/></hh:charProperties>
    <hh:paraProperties itemCnt="3">${paraPr(0, 160, false)}${paraPr(1, 130, true)}${paraPr(2, 180, false)}</hh:paraProperties>
  </hh:refList>
</hh:head>`);
  zip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="b1" paraPrIDRef="2" styleIDRef="0"><hp:run charPrIDRef="0"><hp:t>본문</hp:t></hp:run></hp:p>
  <hp:p id="b2" paraPrIDRef="0" styleIDRef="0"><hp:run charPrIDRef="0"><hp:tbl id="100" rowCnt="1" colCnt="2"><hp:tr>
    <hp:tc><hp:subList><hp:p id="c1" paraPrIDRef="1" styleIDRef="0"><hp:run charPrIDRef="0"><hp:t>양식 칸(130%)</hp:t></hp:run></hp:p></hp:subList><hp:cellAddr colAddr="0" rowAddr="0"/><hp:cellSpan colSpan="1" rowSpan="1"/></hp:tc>
    <hp:tc><hp:subList><hp:p id="c2" paraPrIDRef="0" styleIDRef="0"><hp:run charPrIDRef="0"><hp:t>기본 칸(160%)</hp:t></hp:run></hp:p></hp:subList><hp:cellAddr colAddr="1" rowAddr="0"/><hp:cellSpan colSpan="1" rowSpan="1"/></hp:tc>
  </hp:tr></hp:tbl></hp:run></hp:p>
</hs:sec>`);
  const buf = await zip.generateAsync({ type: 'nodebuffer' });
  return HwpxDocument.createFromBuffer('t', 't.hwpx', buf);
}

async function saved(doc: HwpxDocument) {
  const z = await JSZip.loadAsync(await doc.save());
  const header = (await z.file('Contents/header.xml')!.async('string'));
  const section = (await z.file('Contents/section0.xml')!.async('string'));
  const spacingOf = (id: string) =>
    header.match(new RegExp(`<hh:paraPr\\b[^>]*\\sid="${id}"[\\s\\S]*?<hh:lineSpacing\\b[^>]*value="(\\d+)"`))?.[1];
  const intentOf = (id: string) =>
    header.match(new RegExp(`<hh:paraPr\\b[^>]*\\sid="${id}"[\\s\\S]*?<hc:intent value="(-?\\d+)"`))?.[1];
  const refsInCell = (col: number) => {
    const cell = section.split('<hp:tc>')[col + 1].split('</hp:tc>')[0];
    return [...cell.matchAll(/<hp:p\b[^>]*paraPrIDRef="(\d+)"/g)].map(m => m[1]);
  };
  return { header, section, spacingOf, intentOf, refsInCell };
}

describe('내어쓰기 문단 모양은 원래 줄 간격을 지킨다', () => {
  it('한/글 저장 양식 칸(130%)에 □ ○ 여러 줄을 넣어도 130%가 유지된다', async () => {
    const doc = await makeDoc();
    // update_table_cell(auto_hanging_indent) 이 하는 일 그대로: 줄마다 계산기 → setTableCellHangingIndent
    const text = '□ 결론 한 줄\n  ○ 근거 첫째\n  ○ 근거 둘째';
    doc.updateTableCell(0, 0, 0, 0, text);
    const calc = new HangingIndentCalculator();
    text.split('\n').forEach((line, i) => doc.setTableCellHangingIndent(0, 0, 0, 0, i, calc.calculateHangingIndent(line, 10)));
    const s = await saved(doc);
    const refs = s.refsInCell(0);
    expect(refs).toHaveLength(3);
    for (const id of refs) {
      expect(Number(s.intentOf(id))).toBeLessThan(0);   // 내어쓰기는 걸렸고
      expect(s.spacingOf(id)).toBe('130');               // 줄 간격은 양식 그대로
    }
    // 한/글 내어쓰기 모양: 첫 줄은 그대로(left=0), 둘째 줄부터 들여쓰기만큼 안쪽(intent<0).
    // left 를 들여쓰기만큼 주면 □ 줄까지 통째로 밀려 위계가 무너진다(실측: □ 155pt → 170pt).
    const leftOf = (id: string) =>
      s.header.match(new RegExp(`<hh:paraPr\\b[^>]*\\sid="${id}"[\\s\\S]*?<hc:left value="(-?\\d+)"`))?.[1];
    for (const id of refs) expect(leftOf(id)).toBe('0');
  });

  it('같은 들여쓰기라도 줄 간격이 다른 칸은 서로 다른 문단 모양을 받는다', async () => {
    const doc = await makeDoc();
    doc.setTableCellHangingIndent(0, 0, 0, 0, 0, 12);
    doc.setTableCellHangingIndent(0, 0, 0, 1, 0, 12);
    const s = await saved(doc);
    const [a] = s.refsInCell(0);
    const [b] = s.refsInCell(1);
    expect(a).not.toBe(b);
    expect(s.spacingOf(a)).toBe('130');
    expect(s.spacingOf(b)).toBe('160');
  });

  it('본문 문단 내어쓰기도 원래 줄 간격(180%)을 지킨다', async () => {
    const doc = await makeDoc();
    doc.setHangingIndent(0, 0, 12);
    const s = await saved(doc);
    const ref = s.section.match(/<hp:p id="b1" paraPrIDRef="(\d+)"/)![1];
    expect(ref).not.toBe('2');
    expect(Number(s.intentOf(ref))).toBe(-1200);
    expect(s.spacingOf(ref)).toBe('180');
  });
});
