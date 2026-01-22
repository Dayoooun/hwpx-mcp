import { describe, it, expect, beforeEach } from 'vitest';
import * as fs from 'fs';
import * as path from 'path';
import JSZip from 'jszip';
import { HwpxDocument } from './HwpxDocument';

/**
 * E2E test to reproduce bug: update_paragraph_text changes don't persist after save
 * Especially at higher paragraph indices (10+)
 */

async function createTestHwpxWithManyParagraphs(): Promise<Buffer> {
  const zip = new JSZip();

  const headerXml = `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:docInfo>
    <hh:title>Test Document</hh:title>
  </hh:docInfo>
</hh:head>`;

  // Create section with 15 paragraphs to test high indices
  const paragraphs = Array.from({ length: 15 }, (_, i) =>
    `  <hp:p id="${i + 1}">
    <hp:run><hp:t>Paragraph ${i + 1} - Original Text</hp:t></hp:run>
  </hp:p>`
  ).join('\n');

  const sectionXml = `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
${paragraphs}
</hs:sec>`;

  const contentTypesXml = `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
</Types>`;

  zip.file('Contents/header.xml', headerXml);
  zip.file('Contents/section0.xml', sectionXml);
  zip.file('[Content_Types].xml', contentTypesXml);

  return await zip.generateAsync({ type: 'nodebuffer' });
}

describe('HwpxDocument E2E - update_paragraph_text persistence bug', () => {
  let testFilePath: string;

  beforeEach(async () => {
    testFilePath = path.join(__dirname, '..', 'test-output', 'e2e-paragraph-update.hwpx');
    const buffer = await createTestHwpxWithManyParagraphs();
    fs.writeFileSync(testFilePath, buffer);
  });

  it('should persist paragraph 0 update after save/reload', async () => {
    // 1. Open document
    const originalBuffer = fs.readFileSync(testFilePath);
    const doc = await HwpxDocument.createFromBuffer('test-id-1', testFilePath, originalBuffer);

    // 2. Verify original text
    const paraBefore = doc.getParagraph(0, 0);
    console.log('Paragraph 0 before update:', paraBefore?.text);
    expect(paraBefore?.text).toBe('Paragraph 1 - Original Text');

    // 3. Update paragraph 0
    doc.updateParagraphText(0, 0, 0, 'UPDATED PARAGRAPH 1');

    // 4. Verify memory update
    const paraAfterMemory = doc.getParagraph(0, 0);
    console.log('Paragraph 0 after memory update:', paraAfterMemory?.text);
    expect(paraAfterMemory?.text).toBe('UPDATED PARAGRAPH 1');

    // 5. Save
    const savedBuffer = await doc.save();
    fs.writeFileSync(testFilePath, savedBuffer);

    // 6. Debug: Check saved XML
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');
    console.log('Saved XML (paragraph 0):', savedXml?.substring(0, 500));

    // 7. Close and reload
    const doc2 = await HwpxDocument.createFromBuffer('test-id-2', testFilePath, savedBuffer);

    // 8. Verify persisted value
    const paraAfterReload = doc2.getParagraph(0, 0);
    console.log('Paragraph 0 after reload:', paraAfterReload?.text);

    // THIS SHOULD PASS but verifies the fix works
    expect(paraAfterReload?.text).toBe('UPDATED PARAGRAPH 1');
  });

  it('should persist paragraph 10 update after save/reload (BUG REPRODUCTION)', async () => {
    // This test reproduces the bug at higher indices

    // 1. Open document
    const originalBuffer = fs.readFileSync(testFilePath);
    const doc = await HwpxDocument.createFromBuffer('test-id-1', testFilePath, originalBuffer);

    // 2. Verify original text at index 10
    const paraBefore = doc.getParagraph(0, 10);
    console.log('Paragraph 10 before update:', paraBefore?.text);
    expect(paraBefore?.text).toBe('Paragraph 11 - Original Text');

    // 3. Update paragraph 10
    doc.updateParagraphText(0, 10, 0, 'UPDATED PARAGRAPH 11 - NEW TEXT');

    // 4. Verify memory update
    const paraAfterMemory = doc.getParagraph(0, 10);
    console.log('Paragraph 10 after memory update:', paraAfterMemory?.text);
    expect(paraAfterMemory?.text).toBe('UPDATED PARAGRAPH 11 - NEW TEXT');

    // 5. Save
    const savedBuffer = await doc.save();
    fs.writeFileSync(testFilePath, savedBuffer);

    // 6. Debug: Check saved XML
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');
    console.log('Saved XML snippet (paragraph 10):');

    // Find the paragraph 11 section in XML (id="11")
    const paraMatch = savedXml?.match(/<hp:p id="11"[^>]*>[\s\S]*?<\/hp:p>/);
    console.log('Paragraph 11 XML:', paraMatch?.[0]);

    // 7. Close and reload
    const doc2 = await HwpxDocument.createFromBuffer('test-id-2', testFilePath, savedBuffer);

    // 8. Verify persisted value
    const paraAfterReload = doc2.getParagraph(0, 10);
    console.log('Paragraph 10 after reload:', paraAfterReload?.text);

    // THIS IS THE BUG: Text reverts to original after reload
    expect(paraAfterReload?.text).toBe('UPDATED PARAGRAPH 11 - NEW TEXT');
  });

  it('should persist paragraph 14 update after save/reload (highest index)', async () => {
    // Test at the highest available index

    // 1. Open document
    const originalBuffer = fs.readFileSync(testFilePath);
    const doc = await HwpxDocument.createFromBuffer('test-id-1', testFilePath, originalBuffer);

    // 2. Verify original text at index 14
    const paraBefore = doc.getParagraph(0, 14);
    console.log('Paragraph 14 before update:', paraBefore?.text);
    expect(paraBefore?.text).toBe('Paragraph 15 - Original Text');

    // 3. Update paragraph 14
    doc.updateParagraphText(0, 14, 0, 'FINAL PARAGRAPH UPDATED');

    // 4. Verify memory update
    const paraAfterMemory = doc.getParagraph(0, 14);
    expect(paraAfterMemory?.text).toBe('FINAL PARAGRAPH UPDATED');

    // 5. Save
    const savedBuffer = await doc.save();
    fs.writeFileSync(testFilePath, savedBuffer);

    // 6. Debug: Check saved XML
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');
    const paraMatch = savedXml?.match(/<hp:p id="15"[^>]*>[\s\S]*?<\/hp:p>/);
    console.log('Paragraph 15 XML:', paraMatch?.[0]);

    // 7. Reload
    const doc2 = await HwpxDocument.createFromBuffer('test-id-2', testFilePath, savedBuffer);

    // 8. Verify persisted value
    const paraAfterReload = doc2.getParagraph(0, 14);
    console.log('Paragraph 14 after reload:', paraAfterReload?.text);

    expect(paraAfterReload?.text).toBe('FINAL PARAGRAPH UPDATED');
  });

  it('should persist multiple paragraph updates at various indices', async () => {
    // Test updating multiple paragraphs at different indices

    // 1. Open document
    const originalBuffer = fs.readFileSync(testFilePath);
    const doc = await HwpxDocument.createFromBuffer('test-id-1', testFilePath, originalBuffer);

    // 2. Update paragraphs at indices 0, 5, 10, 14
    const updates = [
      { index: 0, text: 'FIRST UPDATED' },
      { index: 5, text: 'MIDDLE UPDATED' },
      { index: 10, text: 'HIGH INDEX UPDATED' },
      { index: 14, text: 'LAST UPDATED' }
    ];

    for (const { index, text } of updates) {
      doc.updateParagraphText(0, index, 0, text);

      // Verify in memory
      const para = doc.getParagraph(0, index);
      expect(para?.text).toBe(text);
    }

    // 3. Save
    const savedBuffer = await doc.save();

    // 4. Reload
    const doc2 = await HwpxDocument.createFromBuffer('test-id-2', testFilePath, savedBuffer);

    // 5. Verify all updates persisted
    for (const { index, text } of updates) {
      const para = doc2.getParagraph(0, index);
      console.log(`Paragraph ${index} after reload:`, para?.text);
      expect(para?.text).toBe(text);
    }
  });

  it('should use existing real test file if available', async () => {
    // Try to use an existing test file with actual content
    const realTestFile = path.join(__dirname, '..', 'test-output', 'saveload-text-bug.hwpx');

    if (!fs.existsSync(realTestFile)) {
      console.log('Real test file not found, skipping this test');
      return;
    }

    // 1. Open real document
    const originalBuffer = fs.readFileSync(realTestFile);
    const doc = await HwpxDocument.createFromBuffer('real-test', realTestFile, originalBuffer);

    // 2. Get paragraph count
    const paragraphs = doc.getParagraphs(0);
    console.log('Real file paragraph count:', paragraphs?.length);

    if (!paragraphs || paragraphs.length < 11) {
      console.log('Not enough paragraphs in real file');
      return;
    }

    // 3. Get original text at index 10
    const paraBefore = doc.getParagraph(0, 10);
    console.log('Real file - Paragraph 10 before:', paraBefore?.text);
    const originalText = paraBefore?.text || '';

    // 4. Update paragraph 10
    doc.updateParagraphText(0, 10, 0, 'REAL FILE TEST - UPDATED');

    // 5. Save
    const savedBuffer = await doc.save();

    // 6. Reload
    const doc2 = await HwpxDocument.createFromBuffer('real-test-2', realTestFile, savedBuffer);

    // 7. Verify update persisted
    const paraAfterReload = doc2.getParagraph(0, 10);
    console.log('Real file - Paragraph 10 after reload:', paraAfterReload?.text);

    expect(paraAfterReload?.text).toBe('REAL FILE TEST - UPDATED');
    expect(paraAfterReload?.text).not.toBe(originalText);
  });

  it('should handle Korean text update and persist correctly', async () => {
    // Test with Korean characters at high index

    // 1. Open document
    const originalBuffer = fs.readFileSync(testFilePath);
    const doc = await HwpxDocument.createFromBuffer('test-id-1', testFilePath, originalBuffer);

    // 2. Update with Korean text
    const koreanText = '한글 테스트 문단 - 저장 후 유지 검증';
    doc.updateParagraphText(0, 10, 0, koreanText);

    // 3. Save
    const savedBuffer = await doc.save();

    // 4. Reload
    const doc2 = await HwpxDocument.createFromBuffer('test-id-2', testFilePath, savedBuffer);

    // 5. Verify Korean text persisted
    const paraAfterReload = doc2.getParagraph(0, 10);
    console.log('Korean text after reload:', paraAfterReload?.text);
    expect(paraAfterReload?.text).toBe(koreanText);
  });
});

describe('실제 복잡한 문서 테스트 (Real Complex Document)', () => {
  let realDocPath: string;
  let testDocPath: string;

  beforeEach(() => {
    // Path to the real didimddol document
    realDocPath = path.join(__dirname, '..', 'test-output', 'real-didimddol.hwpx');
    // Work on a copy to avoid modifying the original
    testDocPath = path.join(__dirname, '..', 'test-output', 'real-didimddol-test-copy.hwpx');
  });

  it('디딤돌 문서 - paragraph 10 업데이트 후 저장 → 유지되어야 함', async () => {
    // Check if the real document exists
    if (!fs.existsSync(realDocPath)) {
      console.warn('Real didimddol document not found at:', realDocPath);
      console.warn('Skipping this test. Please ensure the file is copied to test-output/');
      return;
    }

    // 1. Copy the real document to work on
    fs.copyFileSync(realDocPath, testDocPath);

    // 2. Open the document
    const originalBuffer = fs.readFileSync(testDocPath);
    const doc = await HwpxDocument.createFromBuffer('didimddol-test-1', testDocPath, originalBuffer);

    // 3. Get paragraphs and verify we have enough
    const paragraphs = doc.getParagraphs(0);
    console.log('디딤돌 문서 - Total paragraphs:', paragraphs?.length);

    if (!paragraphs || paragraphs.length <= 10) {
      console.warn('Not enough paragraphs in the document');
      return;
    }

    // 4. Get paragraph 10's original text
    const paraBefore = doc.getParagraph(0, 10);
    console.log('Paragraph 10 before update:', paraBefore?.text?.substring(0, 100) + '...');
    const originalText = paraBefore?.text || '';

    // Should start with "※ 상기 확인사항의..." based on the task description
    console.log('Original text starts with "※ 상기 확인사항의...":', originalText.includes('※ 상기 확인사항의'));

    // 5. Update paragraph 10
    const updatedText = '테스트 텍스트 - 디딤돌 문서 E2E 검증용';
    doc.updateParagraphText(0, 10, 0, updatedText);

    // 6. Verify in memory
    const paraAfterMemory = doc.getParagraph(0, 10);
    console.log('Paragraph 10 after memory update:', paraAfterMemory?.text);
    expect(paraAfterMemory?.text).toBe(updatedText);

    // 7. Save
    const savedBuffer = await doc.save();
    fs.writeFileSync(testDocPath, savedBuffer);

    // 8. Debug: Check saved XML
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const savedXml = await savedZip.file('Contents/section0.xml')?.async('string');

    // Find paragraph at index 10 (might have varying id)
    const allParas = savedXml?.match(/<hp:p[^>]*>[\s\S]*?<\/hp:p>/g);
    if (allParas && allParas.length > 10) {
      console.log('Saved XML - Paragraph 10:', allParas[10].substring(0, 200) + '...');
    }

    // 9. Reload with a new instance
    const doc2 = await HwpxDocument.createFromBuffer('didimddol-test-2', testDocPath, savedBuffer);

    // 10. Verify paragraph 10 persisted
    const paraAfterReload = doc2.getParagraph(0, 10);
    console.log('Paragraph 10 after reload:', paraAfterReload?.text);

    // THIS IS THE CRITICAL CHECK - the bug causes this to fail
    expect(paraAfterReload?.text).toBe(updatedText);
    expect(paraAfterReload?.text).not.toBe(originalText);

    // 11. Restore original text for cleanup
    doc2.updateParagraphText(0, 10, 0, originalText);
    const restoredBuffer = await doc2.save();
    fs.writeFileSync(testDocPath, restoredBuffer);

    console.log('✓ Test completed - original text restored');
  });

  it('디딤돌 문서 - 여러 문단 동시 업데이트 후 저장 → 모두 유지되어야 함', async () => {
    if (!fs.existsSync(realDocPath)) {
      console.warn('Real didimddol document not found, skipping');
      return;
    }

    // 1. Copy and open
    fs.copyFileSync(realDocPath, testDocPath);
    const originalBuffer = fs.readFileSync(testDocPath);
    const doc = await HwpxDocument.createFromBuffer('didimddol-multi-1', testDocPath, originalBuffer);

    // 2. Get paragraph count
    const paragraphs = doc.getParagraphs(0);
    if (!paragraphs || paragraphs.length <= 15) {
      console.warn('Not enough paragraphs for multi-update test');
      return;
    }

    // 3. Store original texts and update multiple paragraphs
    const updates = [
      { index: 5, text: '업데이트 5번 문단' },
      { index: 10, text: '업데이트 10번 문단' },
      { index: 15, text: '업데이트 15번 문단' }
    ];

    const originals: { index: number; text: string }[] = [];
    const validUpdates: typeof updates = [];

    for (const { index, text } of updates) {
      const para = doc.getParagraph(0, index);
      if (para && para.text) { // Only update if paragraph has text content
        originals.push({ index, text: para.text });
        validUpdates.push({ index, text });
        doc.updateParagraphText(0, index, 0, text);
      } else {
        console.log(`Skipping paragraph ${index} - no text content`);
      }
    }

    // 4. Save
    const savedBuffer = await doc.save();

    // 5. Reload
    const doc2 = await HwpxDocument.createFromBuffer('didimddol-multi-2', testDocPath, savedBuffer);

    // 6. Verify all updates persisted
    for (const { index, text } of validUpdates) {
      const para = doc2.getParagraph(0, index);
      console.log(`Paragraph ${index} after reload:`, para?.text);
      expect(para?.text).toBe(text);
    }

    // 7. Restore originals
    for (const { index, text } of originals) {
      doc2.updateParagraphText(0, index, 0, text);
    }
    const restoredBuffer = await doc2.save();
    fs.writeFileSync(testDocPath, restoredBuffer);

    console.log('✓ Multi-update test completed - originals restored');
  });
});
