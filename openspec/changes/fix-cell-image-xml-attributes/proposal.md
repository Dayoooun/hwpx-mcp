# Change: Fix cell image XML attributes for proper display in 한글

## Why

Cell images inserted via `insert_image_in_cell` and `render_mermaid_in_cell` appear correctly in VSCode preview but fail to display in the native 한글 (HWP) application. Analysis of reference files created by 한글 revealed incorrect XML attribute values in our implementation.

## What Changes

- **BREAKING**: Change `textWrap` attribute from `"SQUARE"` to `"TOP_AND_BOTTOM"`
- Change `affectLSpacing` attribute from `"1"` to `"0"` (per HWPML spec default)
- Change `holdAnchorAndSO` attribute from `"1"` to `"0"` (per HWPML spec default)
- Change `horzRelTo` attribute from `"PARA"` to `"COLUMN"` (match 한글 behavior)

## Impact

- Affected specs: `cell-image-insertion` (new capability)
- Affected code: `mcp-server/src/HwpxDocument.ts:3498-3525` (`applyCellImageInsertsToXml` function)

## Analysis Reference

### 한글에서 생성한 셀 내부 이미지 (정상 작동):
```xml
<hp:pic textWrap="TOP_AND_BOTTOM" ...>
  <hp:pos treatAsChar="1" affectLSpacing="0" holdAnchorAndSO="0"
          vertRelTo="PARA" horzRelTo="COLUMN" .../>
</hp:pic>
```

### 현재 MCP 생성 (문제):
```xml
<hp:pic textWrap="SQUARE" ...>
  <hp:pos treatAsChar="1" affectLSpacing="1" holdAnchorAndSO="1"
          vertRelTo="PARA" horzRelTo="PARA" .../>
</hp:pic>
```

### 공식 HWPML 문서 참조:
- `TextWrapType.TopAndBottom`: "좌/우에는 텍스트를 배치하지 않음"
- `affectLSpacing`: 기본값 `false` (TreatAsChar="true"일 때만 사용)
- `holdAnchorAndSO`: 기본값 `false`
- `horzRelTo`: Column = "단" 기준
