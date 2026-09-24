/**
 * 회귀: 구역 추가·삭제가 저장본에 반영되지 않던 문제 (2026-09-24 발견)
 *
 * 박사님 회신 ③ 을 "insert_section + insert_table 로 2구역 문서를 만들어도 똑같이
 * 재현"된다는 문장대로 옮기다가 찾았다. 0.3.3 에서 insert_section 은 메모리만
 * 바꾸고 section1.xml 을 쓰지 않았다 — 저장하고 다시 열면 구역이 하나로 돌아가고
 * 새 구역에 넣은 내용이 모두 사라졌다. delete_section 은 반대로 파일을 남겼다.
 */
import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { HwpxDocument } from '../../src/HwpxDocument';
import { roundTrip } from '../helpers/hwpx';

async function sectionFiles(buf: Buffer): Promise<string[]> {
  const zip = await JSZip.loadAsync(buf);
  return Object.keys(zip.files).filter(n => /^Contents\/section\d+\.xml$/.test(n)).sort();
}

describe('insert_section', () => {
  it('저장하면 section1.xml 이 생기고 다시 열어도 구역이 둘이다', async () => {
    const doc = HwpxDocument.createNew('s1', 'sections');
    doc.insertParagraph(0, -1, '표지');
    doc.insertSection(0);
    doc.insertParagraph(1, 0, '본문');

    const { buf, doc: back } = await roundTrip(doc);
    expect(await sectionFiles(buf)).toEqual(['Contents/section0.xml', 'Contents/section1.xml']);
    expect(back.content.sections).toHaveLength(2);
    expect(back.getParagraphs(1).map(p => p.text)).toContain('본문');
  });

  it('content.hpf 와 header.xml secCnt 가 구역 수를 따라간다', async () => {
    const doc = HwpxDocument.createNew('s2', 'sections');
    doc.insertSection(0);
    doc.insertSection(1);
    const zip = await JSZip.loadAsync(await doc.save());
    const hpf = await zip.file('Contents/content.hpf')!.async('string');
    const header = await zip.file('Contents/header.xml')!.async('string');
    expect((hpf.match(/<opf:itemref\b[^>]*idref="section\d+"/g) || [])).toHaveLength(3);
    expect(header).toMatch(/secCnt="3"/);
  });

  it('새 구역은 원래 구역의 쪽 설정(secPr)을 물려받는다', async () => {
    const doc = HwpxDocument.createNew('s3', 'sections');
    doc.insertSection(0);
    const zip = await JSZip.loadAsync(await doc.save());
    const pagePr = (x: string) => x.match(/<hp:pagePr\b[^>]*>/)?.[0];
    const s0 = await zip.file('Contents/section0.xml')!.async('string');
    const s1 = await zip.file('Contents/section1.xml')!.async('string');
    expect(pagePr(s1)).toBeDefined();
    expect(pagePr(s1)).toBe(pagePr(s0));
  });

  it('앞 구역 편집을 기록한 뒤 맨 앞에 구역을 넣어도 그 편집은 원래 구역에 들어간다', async () => {
    const doc = HwpxDocument.createNew('s4', 'sections');
    doc.insertParagraph(0, -1, '원래 첫 구역');
    doc.insertSection(-1);
    doc.insertParagraph(0, 0, '새 첫 구역');

    const { doc: back } = await roundTrip(doc);
    expect(back.getParagraphs(0).map(p => p.text)).toContain('새 첫 구역');
    expect(back.getParagraphs(1).map(p => p.text)).toContain('원래 첫 구역');
  });
});

describe('delete_section', () => {
  it('저장하면 그 구역 파일이 없어지고 다시 열어도 돌아오지 않는다', async () => {
    const seed = HwpxDocument.createNew('d1', 'sections');
    seed.insertParagraph(0, -1, '구역0');
    seed.insertSection(0);
    seed.insertParagraph(1, 0, '구역1');
    const { doc } = await roundTrip(seed);

    expect(doc.deleteSection(1)).toBe(true);
    const { buf, doc: back } = await roundTrip(doc);
    expect(await sectionFiles(buf)).toEqual(['Contents/section0.xml']);
    expect(back.content.sections).toHaveLength(1);
    expect(back.getParagraphs(0).map(p => p.text)).toContain('구역0');
  });

  it('가운데 구역을 지우면 뒤 구역이 한 칸 앞당겨진다', async () => {
    const seed = HwpxDocument.createNew('d2', 'sections');
    seed.insertParagraph(0, -1, 'A');
    seed.insertSection(0);
    seed.insertParagraph(1, 0, 'B');
    seed.insertSection(1);
    seed.insertParagraph(2, 0, 'C');
    const { doc } = await roundTrip(seed);

    doc.deleteSection(1);
    const { doc: back } = await roundTrip(doc);
    expect(back.content.sections).toHaveLength(2);
    expect(back.getParagraphs(1).map(p => p.text)).toContain('C');
  });
});
