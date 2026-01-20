/**
 * Tests for image insertion in table cells
 */
import { describe, it, expect, beforeEach } from 'vitest';
import { HwpxDocument } from './HwpxDocument';
import JSZip from 'jszip';

describe('HwpxDocument - Cell Image Insertion', () => {
  let doc: HwpxDocument;
  let testBuffer: Buffer;

  // Simple 1x1 red PNG (base64)
  const testImageBase64 = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFBQIAX8jx0gAAAABJRU5ErkJggg==';

  beforeEach(async () => {
    const zip = new JSZip();

    zip.file('version.xml', '<?xml version="1.0" encoding="UTF-8"?><hh:version xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head" major="1" minor="0"/>');

    zip.file('Contents/header.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hh:head xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">
  <hh:refList>
    <hh:charShapeList itemCnt="1">
      <hh:charShape id="0" height="1000" baseSize="1000"/>
    </hh:charShapeList>
    <hh:paraShapeList itemCnt="1">
      <hh:paraShape id="0"/>
    </hh:paraShapeList>
  </hh:refList>
</hh:head>`);

    zip.file('Contents/section0.xml', `<?xml version="1.0" encoding="UTF-8"?>
<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section"
        xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">
  <hp:tbl id="100" rowCnt="2" colCnt="2">
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="10" paraPrIDRef="0" styleIDRef="0">
            <hp:run charPrIDRef="0"><hp:t>Cell 0,0</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="0" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="11" paraPrIDRef="0" styleIDRef="0">
            <hp:run charPrIDRef="0"><hp:t>Cell 0,1</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
    </hp:tr>
    <hp:tr>
      <hp:tc colAddr="0" rowAddr="1" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="20" paraPrIDRef="0" styleIDRef="0">
            <hp:run charPrIDRef="0"><hp:t>Cell 1,0</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
      <hp:tc colAddr="1" rowAddr="1" colSpan="1" rowSpan="1">
        <hp:subList>
          <hp:p id="21" paraPrIDRef="0" styleIDRef="0">
            <hp:run charPrIDRef="0"><hp:t>Cell 1,1</hp:t></hp:run>
          </hp:p>
        </hp:subList>
      </hp:tc>
    </hp:tr>
  </hp:tbl>
</hs:sec>`);

    zip.file('Contents/content.hpf', `<?xml version="1.0" encoding="UTF-8"?>
<hpf:package xmlns:hpf="urn:hancom:hwp:file">
  <hpf:manifest>
    <hpf:item href="Contents/section0.xml" type="application/xml"/>
    <hpf:item href="Contents/header.xml" type="application/xml"/>
  </hpf:manifest>
</hpf:package>`);

    zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Default Extension="png" ContentType="image/png"/>
</Types>`);

    testBuffer = await zip.generateAsync({ type: 'nodebuffer' });
    doc = await HwpxDocument.createFromBuffer('test', 'test.hwpx', testBuffer);
  });

  it('should insert image into table cell', async () => {
    const result = doc.insertImageInCell(0, 0, 0, 0, {
      data: testImageBase64,
      mimeType: 'image/png',
      width: 100,
      height: 100,
    });

    expect(result).not.toBeNull();
    expect(result?.id).toBe('image1');
    expect(result?.actualWidth).toBe(100);
    expect(result?.actualHeight).toBe(100);

    // Save and verify XML
    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

    expect(sectionXml).toBeDefined();
    // Check image is wrapped in hp:run
    expect(sectionXml).toContain('<hp:run charPrIDRef="0"><hp:pic');
    // Check curSz is 0,0 (required for 한글)
    expect(sectionXml).toContain('<hp:curSz width="0" height="0"/>');
    // Check hp:t/ after hp:pic
    expect(sectionXml).toMatch(/<\/hp:pic><hp:t\/><\/hp:run>/);
    // Check image is in BinData folder
    const binDataFile = savedZip.file('BinData/image1.png');
    expect(binDataFile).not.toBeNull();
  });

  it('should insert image into specific cell (row 1, col 1)', async () => {
    const result = doc.insertImageInCell(0, 0, 1, 1, {
      data: testImageBase64,
      mimeType: 'image/png',
      width: 150,
      height: 100,
    });

    expect(result).not.toBeNull();

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

    // Verify image is in the correct cell (colAddr="1" rowAddr="1")
    const cellMatch = sectionXml?.match(/<hp:tc[^>]*colAddr="1"[^>]*rowAddr="1"[^>]*>([\s\S]*?)<\/hp:tc>/);
    expect(cellMatch).toBeTruthy();
    expect(cellMatch?.[1]).toContain('<hp:pic');
  });

  it('should insert multiple images into different cells', async () => {
    const result1 = doc.insertImageInCell(0, 0, 0, 0, {
      data: testImageBase64,
      mimeType: 'image/png',
      width: 100,
      height: 100,
    });

    const result2 = doc.insertImageInCell(0, 0, 1, 1, {
      data: testImageBase64,
      mimeType: 'image/png',
      width: 100,
      height: 100,
    });

    expect(result1?.id).toBe('image1');
    expect(result2?.id).toBe('image2');

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

    // Count hp:pic elements
    const picCount = (sectionXml?.match(/<hp:pic/g) || []).length;
    expect(picCount).toBe(2);

    // Verify both images in BinData
    expect(savedZip.file('BinData/image1.png')).not.toBeNull();
    expect(savedZip.file('BinData/image2.png')).not.toBeNull();
  });

  it('should have correct position attributes for cell image', async () => {
    doc.insertImageInCell(0, 0, 0, 0, {
      data: testImageBase64,
      mimeType: 'image/png',
      width: 100,
      height: 100,
    });

    const savedBuffer = await doc.save();
    const savedZip = await JSZip.loadAsync(savedBuffer);
    const sectionXml = await savedZip.file('Contents/section0.xml')?.async('string');

    // Check position attributes (required for proper display in 한글)
    expect(sectionXml).toContain('treatAsChar="0"');
    expect(sectionXml).toContain('flowWithText="1"');
    expect(sectionXml).toContain('vertRelTo="PARA"');
    expect(sectionXml).toContain('horzRelTo="COLUMN"');
  });

  it('should reject invalid cell coordinates', () => {
    const result = doc.insertImageInCell(0, 0, 99, 99, {
      data: testImageBase64,
      mimeType: 'image/png',
    });

    expect(result).toBeNull();
  });

  it('should reject invalid table index', () => {
    const result = doc.insertImageInCell(0, 99, 0, 0, {
      data: testImageBase64,
      mimeType: 'image/png',
    });

    expect(result).toBeNull();
  });
});
