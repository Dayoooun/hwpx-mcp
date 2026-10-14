/**
 * Tools whose edit changes the in-memory model only: save_document has no
 * step that writes it to the HWPX file, so the edit is gone after save.
 *
 * Measured on 0.3.5 (2026-10-14): each tool answered success, the in-memory
 * value changed, and the saved XML was identical to saving the same file
 * without the edit. Answering success there is the failure this repository
 * keeps hitting ("성공이라 했는데 저장본은 틀림"), so these tools refuse
 * instead until a save step exists for them.
 *
 * set_document_metadata is checked per document instead (index.ts): it is
 * saved only for fields that header.xml already holds.
 *
 * tests/module/unsaved-tools.test.ts fails as soon as one of these edits does
 * reach the file — remove the tool from this list then.
 */
export const UNSAVED_TOOLS: ReadonlyMap<string, string> = new Map([
  ['insert_footnote', 'footnotes'],
  ['insert_endnote', 'endnotes'],
  ['insert_memo', 'new memos'],
  ['delete_memo', 'memo deletion'],
  ['insert_bookmark', 'bookmarks'],
  ['insert_hyperlink', 'hyperlinks'],
  ['insert_equation', 'equations'],
  ['insert_line', 'line shapes'],
  ['insert_rect', 'rectangle shapes'],
  ['insert_ellipse', 'ellipse shapes'],
  ['update_image_size', 'image size changes'],
  ['set_column_def', 'column definitions'],
  ['set_page_settings', 'page size and margins'],
  ['apply_style', 'paragraph style changes'],
]);

export function unsavedToolError(tool: string): string | null {
  const what = UNSAVED_TOOLS.get(tool);
  if (!what) return null;
  return `${tool} is not available: save_document does not write ${what} to the HWPX file yet, ` +
    'so the change would be lost on save. Nothing was changed.';
}

/** Prefix for the tool description, so clients see the limit before calling. */
export const UNSAVED_TOOL_NOTE = '⛔ NOT AVAILABLE — save_document cannot write this change to the file yet; calling it returns an error. ';
