import { describe, it, expect, afterAll } from 'vitest';
import { HwpxDocument } from './HwpxDocument';
import * as fs from 'fs';
import * as path from 'path';
import * as os from 'os';

describe('Update paragraph text persistence', () => {
  const testFile = 'D:/rlaek/doc-cursor(26new)/26년-지원사업/초기창업패키지-딥테크특화형/별첨/(별첨1) 2026년도 초기창업패키지(딥테크 특화형) 사업계획서 양식.hwpx';
  let tempFile: string;

  afterAll(() => {
    // Clean up temp file
    if (tempFile && fs.existsSync(tempFile)) {
      fs.unlinkSync(tempFile);
    }
  });

  it('should persist paragraph text update after save and reload', async () => {
    // 1. Open document using createFromBuffer
    const buffer = fs.readFileSync(testFile);
    const doc = await HwpxDocument.createFromBuffer('test-id', testFile, buffer);

    // 2. Find paragraph with " ◦ "
    const paragraphs = doc.getParagraphs(0);
    const targetPara = paragraphs.find(p => p.text.includes(' ◦ '));

    expect(targetPara).toBeDefined();
    console.log(`Found paragraph at index ${targetPara!.index} with text: "${targetPara!.text}"`);
    const originalText = targetPara!.text;

    // 3. Update paragraph text
    const newText = '테스트 업데이트 ' + Date.now();
    doc.updateParagraphText(0, targetPara!.index, 0, newText);

    // Verify in-memory update
    const updatedParas = doc.getParagraphs(0);
    const updatedPara = updatedParas.find(p => p.index === targetPara!.index);
    expect(updatedPara?.text).toBe(newText);
    console.log(`In-memory update verified: "${updatedPara?.text}"`);

    // 4. Save to buffer then write to temp file
    const savedBuffer = await doc.save();
    tempFile = path.join(os.tmpdir(), `test-update-${Date.now()}.hwpx`);
    fs.writeFileSync(tempFile, savedBuffer);
    console.log(`Saved to: ${tempFile}`);

    // 5. Reopen and verify
    const reopenBuffer = fs.readFileSync(tempFile);
    const doc2 = await HwpxDocument.createFromBuffer('test-id-2', tempFile, reopenBuffer);

    const reloadedParas = doc2.getParagraphs(0);
    const reloadedPara = reloadedParas.find(p => p.index === targetPara!.index);

    console.log(`After reload, paragraph at index ${targetPara!.index} has text: "${reloadedPara?.text}"`);
    console.log(`Original was: "${originalText}"`);
    console.log(`Expected: "${newText}"`);

    // 6. Assert the text was persisted
    expect(reloadedPara?.text).toBe(newText);
  });
});
