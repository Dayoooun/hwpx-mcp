import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { HwpxDocument } from './HwpxDocument';

describe('Paragraph index persistence with shapes and tables', () => {
  it('updates the selected paragraph without changing matching text in other elements', async () => {
    const zip = new JSZip();
    zip.file('Contents/header.xml', '<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head"/>');
    zip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:p id="before"><hp:run><hp:t>앞 문단 보존</hp:t></hp:run></hp:p>
  <hp:rect id="shape-before"/>
  <hp:tbl id="table-before" rowCnt="1" colCnt="1"><hp:tr><hp:tc colAddr="0" rowAddr="0"><hp:subList>
    <hp:p id="cell"><hp:run><hp:t> ◦ 동일한 항목</hp:t></hp:run></hp:p>
  </hp:subList></hp:tc></hp:tr></hp:tbl>
  <hp:p id="duplicate"><hp:run><hp:t> ◦ 동일한 항목</hp:t></hp:run></hp:p>
  <hp:line id="shape-between"/>
  <hp:p id="duplicate"><hp:run><hp:t> ◦ 동일한 항목</hp:t></hp:run></hp:p>
  <hp:p id="after"><hp:run><hp:t>뒤 문단 보존</hp:t></hp:run></hp:p>
</hs:sec>`);
    const buffer = await zip.generateAsync({ type: 'nodebuffer' });
    const doc = await HwpxDocument.createFromBuffer('fixture', 'synthetic.hwpx', buffer);
    const paragraphs = doc.getParagraphs(0);
    const candidates = paragraphs.filter(paragraph => paragraph.text === ' ◦ 동일한 항목');
    expect(candidates).toHaveLength(2);
    const target = candidates[1];
    const replacement = '선택한 두 번째 항목만 수정';
    doc.updateParagraphText(0, target.index, 0, replacement);
    expect(doc.getParagraph(0, target.index)?.text).toBe(replacement);

    const saved = await doc.save();
    const reopened = await HwpxDocument.createFromBuffer('reloaded', 'synthetic.hwpx', saved);
    expect(reopened.getParagraphs(0).map(paragraph => paragraph.text)).toEqual(
      paragraphs.map(paragraph => paragraph.index === target.index ? replacement : paragraph.text)
    );
    const cell = reopened.findTable(0, 0)?.rows[0].cells[0];
    expect(cell?.paragraphs.flatMap(paragraph => paragraph.runs).map(run => run.text).join(''))
      .toBe(' ◦ 동일한 항목');
    const savedZip = await JSZip.loadAsync(saved);
    const xml = await savedZip.file('Contents/section0.xml')!.async('string');
    expect(xml).toContain('<hp:rect id="shape-before"/>');
    expect(xml).toContain('<hp:line id="shape-between"/>');
    expect(xml.split(replacement)).toHaveLength(2);
  });
});
