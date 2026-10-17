# PDF Export Capability

## ADDED Requirements

### Requirement: Engine selection by platform

The system SHALL convert HWP/HWPX documents to PDF with Hancom's own layout engine, choosing it from the
platform the server runs on.

#### Scenario: Automatic choice on macOS and Windows

- **GIVEN** `export_to_pdf` is called without `engine` or with `engine: "auto"`
- **WHEN** the server runs on macOS
- **THEN** the 한/글 for Mac engine SHALL be used
- **AND** on Windows the 한/글 COM engine SHALL be used

#### Scenario: Unsupported platform

- **GIVEN** the server runs on a platform other than macOS or Windows
- **WHEN** `export_to_pdf` is called
- **THEN** the call SHALL return an error naming the supported platforms
- **AND** no output file SHALL be created

### Requirement: Safe output

The system SHALL never leave a partial PDF and SHALL never overwrite the source document.

#### Scenario: Conversion fails

- **GIVEN** an output path where a file already exists
- **WHEN** the conversion fails, crashes or times out
- **THEN** the call SHALL return an error describing the cause
- **AND** the existing file SHALL be unchanged

#### Scenario: Output path is the source

- **GIVEN** `output_path` refers to the input file or to the open document's file
- **WHEN** `export_to_pdf` is called
- **THEN** the call SHALL be rejected before converting
- **AND** the source file SHALL be unchanged

### Requirement: Open documents

The system SHALL render an open document as currently edited without saving it.

#### Scenario: Unsaved edits

- **GIVEN** an open HWPX document with unsaved edits
- **WHEN** `export_to_pdf` is called with its `doc_id`
- **THEN** the PDF SHALL contain the unsaved edits
- **AND** the document SHALL still be reported as unsaved
