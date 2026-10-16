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
  ['set_cell_properties', 'cell size, background or alignment changes'],
  ['set_column_def', 'column definitions'],
  ['set_page_settings', 'page size and margins'],
  ['apply_style', 'paragraph style changes'],
]);

/**
 * Arguments that a partly saved tool accepts but save_document cannot write
 * reliably: the new charPr is built with textColor="#000000" shadeColor="none"
 * and the underline/strikeout of the document's default charPr (id 0), the
 * new paraPr with every margin and the first-line indent at 0. A request can
 * still land by coincidence (black text, margin 0, underline=false when the
 * default has none) but not in general, so turning a setting off is refused
 * as well. The rest of the tool is saved, so only a request that uses one of
 * these is refused — whole, before anything changes.
 */
export const UNSAVED_ARGS: ReadonlyMap<string, readonly string[]> = new Map([
  ['set_text_style', ['font_color', 'background_color', 'underline', 'strikethrough']],
  ['set_paragraph_style', ['margin_left', 'margin_right', 'margin_top', 'margin_bottom', 'first_line_indent']],
]);

export function unsavedToolError(tool: string, args?: Record<string, unknown>): string | null {
  const what = UNSAVED_TOOLS.get(tool);
  if (what) {
    return `${tool} is not available: save_document does not write ${what} to the HWPX file yet, ` +
      'so the change would be lost on save. Nothing was changed.';
  }
  const used = (UNSAVED_ARGS.get(tool) ?? []).filter(name => args?.[name] !== undefined);
  if (used.length > 0) {
    return `${tool} cannot save ${used.join(', ')}: save_document cannot write ` +
      `${used.length > 1 ? 'these settings' : 'this setting'} to the HWPX file reliably yet (turning it off included). ` +
      'Nothing was changed; call again without them.';
  }
  return null;
}

/** Prefix for the tool description, so clients see the limit before calling. */
export const UNSAVED_ARG_NOTE = '⛔ Not saved yet — a call that uses this (true or false) returns an error. ';

export const UNSAVED_TOOL_NOTE = '⛔ NOT AVAILABLE — save_document cannot write this change to the file yet; calling it returns an error. ';
