/**
 * Regression for issue #22: HH style id 0 must not become langId 1042.
 * This test uses a synthetic HWPX archive; no customer documents or Hancom installation.
 */
import { describe, expect, it } from 'vitest';
import JSZip from 'jszip';
import { HwpxDocument } from '../../src/HwpxDocument';

describe('issue #22: exact hh:style id attribute', () => {
  it('reads real id attributes rather than the later langId property', async () => {
    const original = HwpxDocument.createNew('style-id', 'synthetic-style-test');
    const zip = await JSZip.loadAsync(await original.save());
    const file = zip.file('Contents/header.xml');
    expect(file).not.toBeNull();
    const header = await file!.async('string');
    const fakeStyles = [
      '<hh:style id="0" type="PARA" name="Body" paraPrIDRef="0" charPrIDRef="0" nextStyleIDRef="0" langId="1042"/>',
      '<hh:style langId="1033" name="Subtitle" id="7" type="PARA" paraPrIDRef="0" charPrIDRef="0"/>',
      '<hh:style id="9" type="CHAR" name="Emphasis" langId="1042"></hh:style>',
    ].join('');
    const changed = header.replace(/<hh:style\s+[^>]*\/>/, fakeStyles);
    expect(changed).not.toBe(header); // fixture must inject all 3 synthetic styles
    zip.file('Contents/header.xml', changed);
    const reopened = await HwpxDocument.createFromBuffer(
      'style-id-readback', 'style-id.hwpx', await zip.generateAsync({ type: 'nodebuffer' })
    );
    const styles = reopened.getStyles();
    expect(styles.map(s => s.id).sort((a, b) => a - b)).toEqual([0, 7, 9]);
    expect(styles.find(s => s.id === 0)?.name).toBe('Body');
    expect(styles.find(s => s.id === 7)?.name).toBe('Subtitle');
    expect(styles.some(s => s.id === 1042 || s.id === 1033)).toBe(false);
  });
});
