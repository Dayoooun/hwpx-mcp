# Cell Image Insertion Capability

## ADDED Requirements

### Requirement: Cell Image XML Structure

The system SHALL generate cell images with XML attributes compatible with the native 한글 (HWP) application.

Cell images MUST use the following attribute values:
- `textWrap="TOP_AND_BOTTOM"` (not SQUARE) for proper text flow
- `affectLSpacing="0"` (HWPML spec default)
- `holdAnchorAndSO="0"` (HWPML spec default)
- `horzRelTo="COLUMN"` (column-relative positioning for cell context)

#### Scenario: Image displays correctly in 한글 application

- **GIVEN** a user inserts an image into a table cell using `insert_image_in_cell` or `render_mermaid_in_cell`
- **WHEN** the generated HWPX file is opened in 한글 application
- **THEN** the image SHALL be visible inside the table cell
- **AND** the image SHALL be positioned relative to the cell column

#### Scenario: Image preserves VSCode preview compatibility

- **GIVEN** an image inserted into a table cell
- **WHEN** the HWPX file is previewed in VSCode
- **THEN** the image SHALL display correctly within the cell

### Requirement: Cell Image Position Attributes

The system SHALL set `hp:pos` element attributes according to HWPML specification defaults when `treatAsChar="1"`.

#### Scenario: Position attributes match HWPML spec

- **GIVEN** a cell image with `treatAsChar="1"`
- **WHEN** the XML is generated
- **THEN** `affectLSpacing` SHALL be `"0"` (default false)
- **AND** `holdAnchorAndSO` SHALL be `"0"` (default false)
- **AND** `horzRelTo` SHALL be `"COLUMN"` for cell context
- **AND** `vertRelTo` SHALL be `"PARA"` for paragraph-relative vertical positioning
