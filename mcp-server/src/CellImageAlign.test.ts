/** 표 셀 그림을 가운데 둘 수 있는지 (insert_image_in_cell align). */
import { describe, it, expect } from 'vitest';
import JSZip from 'jszip';
import { HwpxDocument } from './HwpxDocument';

// 1x1 PNG
const PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==';

async function picPos(align?: 'LEFT' | 'CENTER' | 'RIGHT') {
  const doc = HwpxDocument.createNew('img', 'img');
  doc.insertTable(0, 0, 1, 1);
  expect(doc.insertImageInCell(0, 0, 0, 0, { data: PNG, mimeType: 'image/png', width: 100, height: 75, align })).toBeTruthy();
  const z = await JSZip.loadAsync(await doc.save());
  const xml = await z.file('Contents/section0.xml')!.async('string');
  return xml.match(/<hp:pic\b[\s\S]*?<hp:pos\b[^>]*horzAlign="([A-Z]+)"/)?.[1];
}

describe('insert_image_in_cell align', () => {
  it('CENTER 를 주면 그림이 셀 가운데에 놓인다', async () => {
    expect(await picPos('CENTER')).toBe('CENTER');
  });
  it('주지 않으면 지금처럼 왼쪽', async () => {
    expect(await picPos()).toBe('LEFT');
  });
});
