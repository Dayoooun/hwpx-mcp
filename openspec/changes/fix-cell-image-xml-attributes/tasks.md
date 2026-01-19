# Tasks: Fix cell image XML attributes

## 1. Implementation

- [ ] 1.1 Update `textWrap` attribute from `"SQUARE"` to `"TOP_AND_BOTTOM"` in `applyCellImageInsertsToXml()` (HwpxDocument.ts:3499)
- [ ] 1.2 Update `affectLSpacing` attribute from `"1"` to `"0"` in `hp:pos` element (HwpxDocument.ts:3522)
- [ ] 1.3 Update `holdAnchorAndSO` attribute from `"1"` to `"0"` in `hp:pos` element (HwpxDocument.ts:3522)
- [ ] 1.4 Update `horzRelTo` attribute from `"PARA"` to `"COLUMN"` in `hp:pos` element (HwpxDocument.ts:3522)

## 2. Testing

- [ ] 2.1 Run MCP tool `render_mermaid_in_cell` on test HWPX file
- [ ] 2.2 Open generated HWPX file in 한글 application
- [ ] 2.3 Verify image displays inside table cell correctly
- [ ] 2.4 Run `insert_image_in_cell` tool and verify same behavior

## 3. Validation

- [ ] 3.1 Compare generated XML structure with reference file
- [ ] 3.2 Ensure no regression in VSCode preview
