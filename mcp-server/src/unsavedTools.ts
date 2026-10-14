/**
 * Tools whose edit changes the in-memory model only: save_document has no
 * step that writes it to the HWPX file, so the edit is gone after save.
 *
 * Measured on 0.3.5 (2026-10-14): each tool answered success, then the saved
 * file was searched for the inserted text/shape — absent, both in a document
 * from create_document and in one saved by Hancom Office. Answering success
 * there is the failure this repository keeps hitting ("성공이라 했는데 저장본은
 * 틀림"), so these tools refuse instead until a save step exists for them.
 *
 * tests/module/unsaved-tools.test.ts fails as soon as one of these edits does
 * reach the file — remove the tool from this list then.
 */
export const UNSAVED_TOOLS: ReadonlyMap<string, string> = new Map([
  ['insert_footnote', 'footnotes'],
  ['insert_endnote', 'endnotes'],
  ['insert_memo', 'memos'],
  ['insert_bookmark', 'bookmarks'],
  ['insert_hyperlink', 'hyperlinks'],
  ['insert_equation', 'equations'],
  ['insert_line', 'line shapes'],
  ['insert_rect', 'rectangle shapes'],
  ['insert_ellipse', 'ellipse shapes'],
  ['set_column_def', 'column definitions'],
  ['set_document_metadata', 'document metadata (written only where header.xml already holds the field; Hancom files keep it in content.hpf)'],
]);

export function unsavedToolError(tool: string): string | null {
  const what = UNSAVED_TOOLS.get(tool);
  if (!what) return null;
  return `${tool} is not available: save_document does not write ${what} to the HWPX file yet, ` +
    'so the change would be lost on save. Nothing was changed.';
}

/** Prefix for the tool description, so clients see the limit before calling. */
export const UNSAVED_TOOL_NOTE = '⛔ NOT AVAILABLE — save_document cannot write this change to the file yet; calling it returns an error. ';
