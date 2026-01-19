#!/usr/bin/env node
import { Server } from '@modelcontextprotocol/sdk/server/index.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import {
  CallToolRequestSchema,
  ListToolsRequestSchema,
} from '@modelcontextprotocol/sdk/types.js';
import * as fs from 'fs';
import * as path from 'path';
import { HwpxDocument, ImagePositionOptions } from './HwpxDocument';

// Version marker for debugging
const MCP_VERSION = 'v2-fixed-xml-replacement';
console.error(`[HWPX MCP] Server starting - ${MCP_VERSION} - ${new Date().toISOString()}`);

// Document storage
const openDocuments = new Map<string, HwpxDocument>();

function generateId(): string {
  return Math.random().toString(36).substring(2, 11);
}

// ============================================================
// Tool Definitions
// ============================================================

const tools = [
  // === Document Management ===
  {
    name: 'open_document',
    description: 'Open an HWPX or HWP document for reading and editing',
    inputSchema: {
      type: 'object',
      properties: {
        file_path: { type: 'string', description: 'Path to the HWPX or HWP file' },
      },
      required: ['file_path'],
    },
  },
  {
    name: 'close_document',
    description: 'Close an open document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID from open_document' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'save_document',
    description: 'Save the document (HWPX only). Supports backup creation and integrity verification.',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        output_path: { type: 'string', description: 'Output path (optional, saves to original if omitted)' },
        create_backup: { type: 'boolean', description: 'Create .bak backup before saving (default: true)' },
        verify_integrity: { type: 'boolean', description: 'Verify saved file integrity (default: true)' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'list_open_documents',
    description: 'List all currently open documents',
    inputSchema: { type: 'object', properties: {} },
  },

  // === Document Info ===
  {
    name: 'get_document_text',
    description: 'Get all text content from the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'get_document_structure',
    description: 'Get document structure (sections, paragraphs, tables, images count)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'get_document_metadata',
    description: 'Get document metadata (title, author, dates, etc.)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'set_document_metadata',
    description: 'Set document metadata (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        title: { type: 'string', description: 'Document title' },
        creator: { type: 'string', description: 'Author name' },
        subject: { type: 'string', description: 'Subject' },
        description: { type: 'string', description: 'Description' },
      },
      required: ['doc_id'],
    },
  },

  // === Paragraph Operations ===
  {
    name: 'get_paragraphs',
    description: 'Get paragraphs from the document with their text and styles',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (optional, all if omitted)' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'get_paragraph',
    description: 'Get a specific paragraph with full details',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index'],
    },
  },
  {
    name: 'insert_paragraph',
    description: 'Insert a new paragraph (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        after_index: { type: 'number', description: 'Insert after this paragraph index (-1 for beginning)' },
        text: { type: 'string', description: 'Paragraph text' },
      },
      required: ['doc_id', 'section_index', 'after_index', 'text'],
    },
  },
  {
    name: 'delete_paragraph',
    description: 'Delete a paragraph (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index to delete' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index'],
    },
  },
  {
    name: 'update_paragraph_text',
    description: 'Update paragraph text content (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
        run_index: { type: 'number', description: 'Run index (default 0)' },
        text: { type: 'string', description: 'New text content' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index', 'text'],
    },
  },
  {
    name: 'append_text_to_paragraph',
    description: 'Append text to an existing paragraph (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
        text: { type: 'string', description: 'Text to append' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index', 'text'],
    },
  },

  // === Character Styling ===
  {
    name: 'set_text_style',
    description: 'Apply character formatting to a paragraph run (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
        run_index: { type: 'number', description: 'Run index (default 0)' },
        bold: { type: 'boolean', description: 'Bold' },
        italic: { type: 'boolean', description: 'Italic' },
        underline: { type: 'boolean', description: 'Underline' },
        strikethrough: { type: 'boolean', description: 'Strikethrough' },
        font_name: { type: 'string', description: 'Font name' },
        font_size: { type: 'number', description: 'Font size in pt' },
        font_color: { type: 'string', description: 'Text color (hex)' },
        background_color: { type: 'string', description: 'Background color (hex)' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index'],
    },
  },
  {
    name: 'get_text_style',
    description: 'Get character formatting of a paragraph',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
        run_index: { type: 'number', description: 'Run index (optional)' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index'],
    },
  },

  // === Paragraph Styling ===
  {
    name: 'set_paragraph_style',
    description: 'Apply paragraph formatting (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
        align: { type: 'string', enum: ['left', 'center', 'right', 'justify', 'distribute'], description: 'Text alignment' },
        line_spacing: { type: 'number', description: 'Line spacing in %' },
        margin_left: { type: 'number', description: 'Left margin in pt' },
        margin_right: { type: 'number', description: 'Right margin in pt' },
        margin_top: { type: 'number', description: 'Top margin in pt' },
        margin_bottom: { type: 'number', description: 'Bottom margin in pt' },
        first_line_indent: { type: 'number', description: 'First line indent in pt' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index'],
    },
  },
  {
    name: 'get_paragraph_style',
    description: 'Get paragraph formatting',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index'],
    },
  },

  // === Hanging Indent (내어쓰기) ===
  {
    name: 'set_hanging_indent',
    description: 'Set hanging indent on a paragraph (HWPX only). Hanging indent pulls the first line left while indenting the rest of the lines.',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
        indent_pt: { type: 'number', description: 'Indent amount in points (positive value)' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index', 'indent_pt'],
    },
  },
  {
    name: 'get_hanging_indent',
    description: 'Get hanging indent value for a paragraph',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index'],
    },
  },
  {
    name: 'remove_hanging_indent',
    description: 'Remove hanging indent from a paragraph (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index'],
    },
  },
  {
    name: 'set_table_cell_hanging_indent',
    description: 'Set hanging indent on a paragraph inside a table cell (HWPX only). Hanging indent pulls the first line left while indenting the rest of the lines.',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index within section' },
        row: { type: 'number', description: 'Row index (0-based)' },
        col: { type: 'number', description: 'Column index (0-based)' },
        paragraph_index: { type: 'number', description: 'Paragraph index within cell (0-based)' },
        indent_pt: { type: 'number', description: 'Indent amount in points (positive value)' },
      },
      required: ['doc_id', 'section_index', 'table_index', 'row', 'col', 'paragraph_index', 'indent_pt'],
    },
  },
  {
    name: 'get_table_cell_hanging_indent',
    description: 'Get hanging indent value for a paragraph inside a table cell',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index within section' },
        row: { type: 'number', description: 'Row index (0-based)' },
        col: { type: 'number', description: 'Column index (0-based)' },
        paragraph_index: { type: 'number', description: 'Paragraph index within cell (0-based)' },
      },
      required: ['doc_id', 'section_index', 'table_index', 'row', 'col', 'paragraph_index'],
    },
  },
  {
    name: 'remove_table_cell_hanging_indent',
    description: 'Remove hanging indent from a paragraph inside a table cell (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index within section' },
        row: { type: 'number', description: 'Row index (0-based)' },
        col: { type: 'number', description: 'Column index (0-based)' },
        paragraph_index: { type: 'number', description: 'Paragraph index within cell (0-based)' },
      },
      required: ['doc_id', 'section_index', 'table_index', 'row', 'col', 'paragraph_index'],
    },
  },

  // === Search & Replace ===
  {
    name: 'search_text',
    description: 'Search for text in the document (includes table cells by default)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        query: { type: 'string', description: 'Text to search for' },
        case_sensitive: { type: 'boolean', description: 'Case sensitive search (default: false)' },
        regex: { type: 'boolean', description: 'Use regular expression (default: false)' },
        include_tables: { type: 'boolean', description: 'Include table cell text in search (default: true)' },
      },
      required: ['doc_id', 'query'],
    },
  },
  {
    name: 'replace_text',
    description: 'Replace text in the document (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        old_text: { type: 'string', description: 'Text to find' },
        new_text: { type: 'string', description: 'Replacement text' },
        case_sensitive: { type: 'boolean', description: 'Case sensitive (default: false)' },
        regex: { type: 'boolean', description: 'Use regular expression (default: false)' },
        replace_all: { type: 'boolean', description: 'Replace all occurrences (default: true)' },
      },
      required: ['doc_id', 'old_text', 'new_text'],
    },
  },
  {
    name: 'batch_replace',
    description: 'Perform multiple text replacements at once (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        replacements: {
          type: 'array',
          items: {
            type: 'object',
            properties: {
              old_text: { type: 'string' },
              new_text: { type: 'string' },
            },
          },
          description: 'Array of {old_text, new_text} pairs',
        },
      },
      required: ['doc_id', 'replacements'],
    },
  },
  {
    name: 'replace_text_in_cell',
    description: 'Replace text within a specific table cell (HWPX only). More targeted than replace_text.',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index within section' },
        row: { type: 'number', description: 'Row index (0-based)' },
        col: { type: 'number', description: 'Column index (0-based)' },
        old_text: { type: 'string', description: 'Text to find' },
        new_text: { type: 'string', description: 'Replacement text' },
        case_sensitive: { type: 'boolean', description: 'Case sensitive (default: false)' },
        regex: { type: 'boolean', description: 'Use regular expression (default: false)' },
        replace_all: { type: 'boolean', description: 'Replace all occurrences (default: true)' },
      },
      required: ['doc_id', 'section_index', 'table_index', 'row', 'col', 'old_text', 'new_text'],
    },
  },

  // === Table Operations ===
  {
    name: 'get_tables',
    description: 'Get all tables from the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'get_table_map',
    description: 'Get table map with headers - maps table indices to their header paragraphs. Returns table info including header text from preceding paragraph, size, empty status, and first row preview.',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'find_empty_tables',
    description: 'Find tables that are empty or contain only placeholder text (dashes, bullets, numbers only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'get_tables_by_section',
    description: 'Get all tables within a specific section',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (0-based)' },
      },
      required: ['doc_id', 'section_index'],
    },
  },
  {
    name: 'find_table_by_header',
    description: 'Find tables by their header text (partial match, case-insensitive)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        search_text: { type: 'string', description: 'Text to search for in table headers' },
      },
      required: ['doc_id', 'search_text'],
    },
  },
  {
    name: 'get_tables_summary',
    description: 'Get summary of multiple tables by index range. Returns compact info: header, size, empty status, and content preview.',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        start_index: { type: 'number', description: 'Start table index (0-based, default: 0)' },
        end_index: { type: 'number', description: 'End table index (inclusive, default: last table)' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'get_document_outline',
    description: 'Get document outline - hierarchical structure showing sections, headings, and tables with their positions',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },

  // === Position/Index Helper Tools ===
  {
    name: 'get_element_index_for_table',
    description: 'Convert a global table index to element index in its section. Use this to find the right position for inserting content near a table.',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        table_index: { type: 'number', description: 'Global table index (0-based, from get_tables or get_table_map)' },
      },
      required: ['doc_id', 'table_index'],
    },
  },
  {
    name: 'find_paragraph_by_text',
    description: 'Find paragraphs containing specific text. Returns element indices with surrounding context.',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        search_text: { type: 'string', description: 'Text to search for (partial match, case-insensitive)' },
        section_index: { type: 'number', description: 'Optional: limit search to specific section' },
      },
      required: ['doc_id', 'search_text'],
    },
  },
  {
    name: 'get_insert_context',
    description: 'Get context around an element index to verify insertion point. Shows elements before/after.',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        element_index: { type: 'number', description: 'Element index to inspect' },
        context_range: { type: 'number', description: 'Number of elements before/after to show (default: 2)' },
      },
      required: ['doc_id', 'section_index', 'element_index'],
    },
  },
  {
    name: 'find_insert_position_after_header',
    description: `Find the right insertion position after text. Searches both independent paragraphs AND table cell contents by default.

IMPORTANT - Check 'found_in' in the result:
- If found_in='paragraph': Use insert_image with section_index and insert_after to insert AFTER the paragraph
- If found_in='table_cell': The text is INSIDE a table cell. Use insert_image_in_cell with table_info (table_index, row, col) to insert the image INSIDE that cell. Do NOT use insert_image as it will place the image OUTSIDE the table.`,
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        header_text: { type: 'string', description: 'Header/title text to search for' },
        search_in: {
          type: 'string',
          enum: ['paragraphs', 'table_cells', 'all'],
          description: 'Where to search: "paragraphs" (independent paragraphs only), "table_cells" (table cell contents only), "all" (both, default). Many Korean documents have content inside table cells, so "all" is recommended.'
        },
      },
      required: ['doc_id', 'header_text'],
    },
  },
  {
    name: 'find_insert_position_after_table',
    description: `Find the right insertion position AFTER a specific table (OUTSIDE the table).

Returns section_index and insert_after value for use with insert_image/render_mermaid.
NOTE: This inserts AFTER the table, not inside it. To insert an image INSIDE a table cell, use insert_image_in_cell directly.`,
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        table_index: { type: 'number', description: 'Global table index (0-based)' },
      },
      required: ['doc_id', 'table_index'],
    },
  },
  {
    name: 'get_table',
    description: 'Get a specific table with full data',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index within section' },
      },
      required: ['doc_id', 'section_index', 'table_index'],
    },
  },
  {
    name: 'get_table_cell',
    description: 'Get content of a specific table cell',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index' },
        row: { type: 'number', description: 'Row index (0-based)' },
        col: { type: 'number', description: 'Column index (0-based)' },
      },
      required: ['doc_id', 'section_index', 'table_index', 'row', 'col'],
    },
  },
  {
    name: 'update_table_cell',
    description: 'Update content of a table cell (HWPX only). Preserves existing charPrIDRef by default.',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index' },
        row: { type: 'number', description: 'Row index' },
        col: { type: 'number', description: 'Column index' },
        text: { type: 'string', description: 'New cell content' },
        char_shape_id: { type: 'number', description: 'Character shape ID to apply (optional, uses existing style if omitted)' },
      },
      required: ['doc_id', 'section_index', 'table_index', 'row', 'col', 'text'],
    },
  },
  {
    name: 'set_cell_properties',
    description: 'Set table cell properties (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index' },
        row: { type: 'number', description: 'Row index' },
        col: { type: 'number', description: 'Column index' },
        width: { type: 'number', description: 'Cell width' },
        height: { type: 'number', description: 'Cell height' },
        background_color: { type: 'string', description: 'Background color (hex)' },
        vertical_align: { type: 'string', enum: ['top', 'middle', 'bottom'], description: 'Vertical alignment' },
      },
      required: ['doc_id', 'section_index', 'table_index', 'row', 'col'],
    },
  },
  {
    name: 'merge_cells',
    description: 'Merge multiple table cells into a single cell (HWPX only). The top-left cell becomes the master cell with increased colSpan/rowSpan.',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index' },
        start_row: { type: 'number', description: 'Starting row index (0-based)' },
        start_col: { type: 'number', description: 'Starting column index (0-based)' },
        end_row: { type: 'number', description: 'Ending row index (0-based, inclusive)' },
        end_col: { type: 'number', description: 'Ending column index (0-based, inclusive)' },
      },
      required: ['doc_id', 'section_index', 'table_index', 'start_row', 'start_col', 'end_row', 'end_col'],
    },
  },
  {
    name: 'split_cell',
    description: 'Split a merged table cell back into individual cells (HWPX only). Only works on cells with colSpan > 1 or rowSpan > 1.',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index' },
        row: { type: 'number', description: 'Row index of the merged cell (0-based)' },
        col: { type: 'number', description: 'Column index of the merged cell (0-based)' },
      },
      required: ['doc_id', 'section_index', 'table_index', 'row', 'col'],
    },
  },
  {
    name: 'insert_table_row',
    description: 'Insert a new row in a table (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index' },
        after_row: { type: 'number', description: 'Insert after this row index (-1 for beginning)' },
        cell_texts: { type: 'array', items: { type: 'string' }, description: 'Text for each cell (optional)' },
      },
      required: ['doc_id', 'section_index', 'table_index', 'after_row'],
    },
  },
  {
    name: 'delete_table_row',
    description: 'Delete a row from a table (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index' },
        row_index: { type: 'number', description: 'Row index to delete' },
      },
      required: ['doc_id', 'section_index', 'table_index', 'row_index'],
    },
  },
  {
    name: 'insert_table_column',
    description: 'Insert a new column in a table (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index' },
        after_col: { type: 'number', description: 'Insert after this column (-1 for beginning)' },
      },
      required: ['doc_id', 'section_index', 'table_index', 'after_col'],
    },
  },
  {
    name: 'delete_table_column',
    description: 'Delete a column from a table (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index' },
        col_index: { type: 'number', description: 'Column index to delete' },
      },
      required: ['doc_id', 'section_index', 'table_index', 'col_index'],
    },
  },
  {
    name: 'get_table_as_csv',
    description: 'Export table content as CSV format',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        table_index: { type: 'number', description: 'Table index' },
        delimiter: { type: 'string', description: 'Delimiter character (default: comma)' },
      },
      required: ['doc_id', 'section_index', 'table_index'],
    },
  },

  // === Page Settings ===
  {
    name: 'get_page_settings',
    description: 'Get page settings (paper size, margins)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'set_page_settings',
    description: 'Set page settings (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        width: { type: 'number', description: 'Page width in pt' },
        height: { type: 'number', description: 'Page height in pt' },
        margin_top: { type: 'number', description: 'Top margin in pt' },
        margin_bottom: { type: 'number', description: 'Bottom margin in pt' },
        margin_left: { type: 'number', description: 'Left margin in pt' },
        margin_right: { type: 'number', description: 'Right margin in pt' },
        orientation: { type: 'string', enum: ['portrait', 'landscape'], description: 'Page orientation' },
      },
      required: ['doc_id'],
    },
  },

  // === Copy/Move ===
  {
    name: 'copy_paragraph',
    description: 'Copy a paragraph to another location (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        source_section: { type: 'number', description: 'Source section index' },
        source_paragraph: { type: 'number', description: 'Source paragraph index' },
        target_section: { type: 'number', description: 'Target section index' },
        target_after: { type: 'number', description: 'Insert after this paragraph in target' },
      },
      required: ['doc_id', 'source_section', 'source_paragraph', 'target_section', 'target_after'],
    },
  },
  {
    name: 'move_paragraph',
    description: 'Move a paragraph to another location (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        source_section: { type: 'number', description: 'Source section index' },
        source_paragraph: { type: 'number', description: 'Source paragraph index' },
        target_section: { type: 'number', description: 'Target section index' },
        target_after: { type: 'number', description: 'Insert after this paragraph in target' },
      },
      required: ['doc_id', 'source_section', 'source_paragraph', 'target_section', 'target_after'],
    },
  },

  // === Statistics ===
  {
    name: 'get_word_count',
    description: 'Get word and character count statistics',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },

  // === Images ===
  {
    name: 'get_images',
    description: 'Get all images in the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },

  // === Export ===
  {
    name: 'export_to_text',
    description: 'Export document to plain text file',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        output_path: { type: 'string', description: 'Output file path' },
      },
      required: ['doc_id', 'output_path'],
    },
  },
  {
    name: 'export_to_html',
    description: 'Export document to HTML file',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        output_path: { type: 'string', description: 'Output file path' },
      },
      required: ['doc_id', 'output_path'],
    },
  },

  // === Undo/Redo ===
  {
    name: 'undo',
    description: 'Undo the last change',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'redo',
    description: 'Redo the last undone change',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },

  // === Table Creation ===
  {
    name: 'insert_table',
    description: 'Insert a new table (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        after_index: { type: 'number', description: 'Insert after this element index (-1 for beginning)' },
        rows: { type: 'number', description: 'Number of rows' },
        cols: { type: 'number', description: 'Number of columns' },
        width: { type: 'number', description: 'Table width (optional)' },
      },
      required: ['doc_id', 'section_index', 'after_index', 'rows', 'cols'],
    },
  },
  {
    name: 'insert_nested_table',
    description: 'Insert a table inside a table cell (nested table, HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        parent_table_index: { type: 'number', description: 'Parent table index' },
        row: { type: 'number', description: 'Row index in parent table (0-based)' },
        col: { type: 'number', description: 'Column index in parent table (0-based)' },
        nested_rows: { type: 'number', description: 'Number of rows in nested table' },
        nested_cols: { type: 'number', description: 'Number of columns in nested table' },
        data: {
          type: 'array',
          description: 'Optional 2D array of cell data for nested table',
          items: {
            type: 'array',
            items: { type: 'string' }
          }
        },
      },
      required: ['doc_id', 'section_index', 'parent_table_index', 'row', 'col', 'nested_rows', 'nested_cols'],
    },
  },

  // === Header/Footer ===
  {
    name: 'get_header',
    description: 'Get header content for a section',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'set_header',
    description: 'Set header content for a section (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
        text: { type: 'string', description: 'Header text content' },
        apply_page_type: { type: 'string', enum: ['both', 'even', 'odd'], description: 'Apply to page type (default: both)' },
      },
      required: ['doc_id', 'text'],
    },
  },
  {
    name: 'get_footer',
    description: 'Get footer content for a section',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'set_footer',
    description: 'Set footer content for a section (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
        text: { type: 'string', description: 'Footer text content' },
        apply_page_type: { type: 'string', enum: ['both', 'even', 'odd'], description: 'Apply to page type (default: both)' },
      },
      required: ['doc_id', 'text'],
    },
  },

  // === Footnotes/Endnotes ===
  {
    name: 'get_footnotes',
    description: 'Get all footnotes in the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'insert_footnote',
    description: 'Insert a footnote at a specific location (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
        text: { type: 'string', description: 'Footnote text content' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index', 'text'],
    },
  },
  {
    name: 'get_endnotes',
    description: 'Get all endnotes in the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'insert_endnote',
    description: 'Insert an endnote at a specific location (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
        text: { type: 'string', description: 'Endnote text content' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index', 'text'],
    },
  },

  // === Bookmarks/Hyperlinks ===
  {
    name: 'get_bookmarks',
    description: 'Get all bookmarks in the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'insert_bookmark',
    description: 'Insert a bookmark at a specific location (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
        name: { type: 'string', description: 'Bookmark name' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index', 'name'],
    },
  },
  {
    name: 'get_hyperlinks',
    description: 'Get all hyperlinks in the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'insert_hyperlink',
    description: 'Insert a hyperlink in a paragraph (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
        url: { type: 'string', description: 'URL for the hyperlink' },
        text: { type: 'string', description: 'Display text for the hyperlink' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index', 'url', 'text'],
    },
  },

  // === Images ===
  {
    name: 'insert_image',
    description: `Insert an image as an independent element in the document (HWPX only). The image is placed OUTSIDE of tables, between paragraphs or after tables.

Use after_table or after_header for easier positioning.

⚠️ WARNING: This tool ALWAYS inserts OUTSIDE tables. Even if after_header finds text inside a table cell, the image will be placed AFTER the table, not inside it.

👉 To insert an image INSIDE a table cell:
1. First use find_insert_position_after_header to check found_in
2. If found_in='table_cell', use insert_image_in_cell with the returned table_info (table_index, row, col)
3. If found_in='paragraph', use this insert_image tool`,
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (auto-detected if using after_table or after_header)' },
        after_index: { type: 'number', description: 'Insert after this element index. Use after_table or after_header instead for easier positioning.' },
        after_table: { type: 'number', description: 'RECOMMENDED: Insert after this table index (0-based global index from get_table_map). Automatically sets section_index and after_index.' },
        after_header: { type: 'string', description: 'RECOMMENDED: Insert after paragraph containing this text. Automatically sets section_index and after_index.' },
        image_path: { type: 'string', description: 'Path to the image file' },
        width: { type: 'number', description: 'Image width in points (optional). If only width is specified with preserve_aspect_ratio=true, height is auto-calculated.' },
        height: { type: 'number', description: 'Image height in points (optional). If only height is specified with preserve_aspect_ratio=true, width is auto-calculated.' },
        preserve_aspect_ratio: { type: 'boolean', description: 'If true, maintains original image aspect ratio. Default: false.' },
        position_type: { type: 'string', enum: ['inline', 'floating'], description: 'Position type: "inline" (flows with text like a character) or "floating" (positioned relative to anchor). Default: floating.' },
        vert_rel_to: { type: 'string', enum: ['para', 'paper'], description: 'Vertical reference point: "para" (paragraph) or "paper" (page). Default: para.' },
        horz_rel_to: { type: 'string', enum: ['column', 'para', 'paper'], description: 'Horizontal reference point: "column", "para" (paragraph), or "paper" (page). Default: column.' },
        vert_align: { type: 'string', enum: ['top', 'center', 'bottom'], description: 'Vertical alignment. Default: top.' },
        horz_align: { type: 'string', enum: ['left', 'center', 'right'], description: 'Horizontal alignment. Default: left.' },
        vert_offset: { type: 'number', description: 'Vertical offset from anchor in points. Default: 0.' },
        horz_offset: { type: 'number', description: 'Horizontal offset from anchor in points. Default: 0.' },
        text_wrap: { type: 'string', enum: ['top_and_bottom', 'square', 'tight', 'behind_text', 'in_front_of_text', 'none'], description: 'Text wrap mode. Default: top_and_bottom.' },
      },
      required: ['doc_id', 'image_path'],
    },
  },
  {
    name: 'update_image_size',
    description: 'Update the size of an existing image (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        image_index: { type: 'number', description: 'Image index within section' },
        width: { type: 'number', description: 'New width' },
        height: { type: 'number', description: 'New height' },
      },
      required: ['doc_id', 'section_index', 'image_index', 'width', 'height'],
    },
  },
  {
    name: 'delete_image',
    description: 'Delete an image from the document (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        image_index: { type: 'number', description: 'Image index within section' },
      },
      required: ['doc_id', 'section_index', 'image_index'],
    },
  },
  {
    name: 'render_mermaid',
    description: `Render a Mermaid diagram and insert it as an independent element OUTSIDE tables (HWPX only). Uses mermaid.ink API.

Use after_table or after_header for easier positioning.

⚠️ WARNING: This tool ALWAYS inserts OUTSIDE tables. Even if after_header finds text inside a table cell, the diagram will be placed AFTER the table, not inside it.

👉 To insert a Mermaid diagram INSIDE a table cell:
1. First use find_insert_position_after_header to check found_in
2. If found_in='table_cell', use render_mermaid_in_cell with the returned table_info (table_index, row, col)
3. If found_in='paragraph', use this render_mermaid tool`,
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        mermaid_code: { type: 'string', description: 'Mermaid diagram code (e.g., "graph TD; A-->B;")' },
        section_index: { type: 'number', description: 'Section index (auto-detected if using after_table or after_header)' },
        after_index: { type: 'number', description: 'Insert after this element index. Use after_table or after_header instead for easier positioning.' },
        after_table: { type: 'number', description: 'RECOMMENDED: Insert after this table index (0-based global index from get_table_map). Automatically sets section_index and after_index.' },
        after_header: { type: 'string', description: 'RECOMMENDED: Insert after paragraph containing this text. Automatically sets section_index and after_index.' },
        width: { type: 'number', description: 'Image width in points (optional). If specified with preserve_aspect_ratio=true, height is auto-calculated.' },
        height: { type: 'number', description: 'Image height in points (optional). If specified with preserve_aspect_ratio=true, width is auto-calculated.' },
        theme: { type: 'string', enum: ['default', 'dark', 'forest', 'neutral'], description: 'Diagram theme (default: default)' },
        background_color: { type: 'string', description: 'Background color (e.g., "#ffffff" or "transparent")' },
        preserve_aspect_ratio: { type: 'boolean', description: 'If true, maintains original image aspect ratio. Default: true for Mermaid diagrams.' },
        position_type: { type: 'string', enum: ['inline', 'floating'], description: 'Position type: "inline" (flows with text) or "floating" (positioned relative to anchor). Default: floating.' },
        vert_rel_to: { type: 'string', enum: ['para', 'paper'], description: 'Vertical reference point: "para" (paragraph) or "paper" (page). Default: para.' },
        horz_rel_to: { type: 'string', enum: ['column', 'para', 'paper'], description: 'Horizontal reference point: "column", "para" (paragraph), or "paper" (page). Default: column.' },
        vert_align: { type: 'string', enum: ['top', 'center', 'bottom'], description: 'Vertical alignment. Default: top.' },
        horz_align: { type: 'string', enum: ['left', 'center', 'right'], description: 'Horizontal alignment. Default: left.' },
        vert_offset: { type: 'number', description: 'Vertical offset from anchor in points. Default: 0.' },
        horz_offset: { type: 'number', description: 'Horizontal offset from anchor in points. Default: 0.' },
        text_wrap: { type: 'string', enum: ['top_and_bottom', 'square', 'tight', 'behind_text', 'in_front_of_text', 'none'], description: 'Text wrap mode. Default: top_and_bottom.' },
      },
      required: ['doc_id', 'mermaid_code'],
    },
  },
  {
    name: 'insert_image_in_cell',
    description: `📍 Insert an image INSIDE a specific table cell (HWPX only). The image appears inline within the cell content.

⚠️ IMPORTANT: Use this tool (NOT insert_image) when inserting images into table cells!

When to use:
1. find_insert_position_after_header returned found_in='table_cell' → use table_info (table_index, row, col)
2. You want to add an image to a specific cell you already know

How to get table_index:
- From find_insert_position_after_header result: table_info.table_index
- Or use get_table_map to list all tables and find the index

Positioning within cell:
- By default, image is inserted at the beginning of the cell
- Use after_text to insert the image after a specific paragraph containing that text`,
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        table_index: { type: 'number', description: 'Global table index (0-based). Get from get_table_map.' },
        row: { type: 'number', description: 'Row index (0-based)' },
        col: { type: 'number', description: 'Column index (0-based)' },
        image_path: { type: 'string', description: 'Path to the image file' },
        width: { type: 'number', description: 'Image width in points (optional, default: 200)' },
        height: { type: 'number', description: 'Image height in points (optional, default: 150)' },
        preserve_aspect_ratio: { type: 'boolean', description: 'If true, maintains original image aspect ratio. Default: false.' },
        after_text: { type: 'string', description: 'Insert the image after the paragraph containing this text. If not found, falls back to beginning of cell.' },
      },
      required: ['doc_id', 'table_index', 'row', 'col', 'image_path'],
    },
  },
  {
    name: 'render_mermaid_in_cell',
    description: `📍 Render a Mermaid diagram and insert it INSIDE a specific table cell (HWPX only). Uses mermaid.ink API.

⚠️ IMPORTANT: Use this tool (NOT render_mermaid) when inserting diagrams into table cells!

When to use:
1. find_insert_position_after_header returned found_in='table_cell' → use table_info (table_index, row, col)
2. You want to add a diagram to a specific cell you already know

How to get table_index:
- From find_insert_position_after_header result: table_info.table_index
- Or use get_table_map to list all tables and find the index

Positioning within cell:
- By default, diagram is inserted at the beginning of the cell
- Use after_text to insert the diagram after a specific paragraph containing that text`,
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        mermaid_code: { type: 'string', description: 'Mermaid diagram code (e.g., "graph TD; A-->B;")' },
        table_index: { type: 'number', description: 'Global table index (0-based). Get from get_table_map.' },
        row: { type: 'number', description: 'Row index (0-based)' },
        col: { type: 'number', description: 'Column index (0-based)' },
        width: { type: 'number', description: 'Image width in points (optional)' },
        height: { type: 'number', description: 'Image height in points (optional)' },
        theme: { type: 'string', enum: ['default', 'dark', 'forest', 'neutral'], description: 'Diagram theme (default: default)' },
        background_color: { type: 'string', description: 'Background color (e.g., "#ffffff" or "transparent")' },
        preserve_aspect_ratio: { type: 'boolean', description: 'If true, maintains original image aspect ratio. Default: true.' },
        after_text: { type: 'string', description: 'Insert the diagram after the paragraph containing this text. If not found, falls back to beginning of cell.' },
      },
      required: ['doc_id', 'mermaid_code', 'table_index', 'row', 'col'],
    },
  },

  // === Drawing Objects ===
  {
    name: 'insert_line',
    description: 'Insert a line drawing object (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        after_index: { type: 'number', description: 'Insert after this element index (-1 for beginning)' },
        x1: { type: 'number', description: 'Start X coordinate' },
        y1: { type: 'number', description: 'Start Y coordinate' },
        x2: { type: 'number', description: 'End X coordinate' },
        y2: { type: 'number', description: 'End Y coordinate' },
        stroke_color: { type: 'string', description: 'Stroke color (hex)' },
        stroke_width: { type: 'number', description: 'Stroke width' },
      },
      required: ['doc_id', 'section_index', 'after_index', 'x1', 'y1', 'x2', 'y2'],
    },
  },
  {
    name: 'insert_rect',
    description: 'Insert a rectangle drawing object (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        after_index: { type: 'number', description: 'Insert after this element index (-1 for beginning)' },
        x: { type: 'number', description: 'X coordinate' },
        y: { type: 'number', description: 'Y coordinate' },
        width: { type: 'number', description: 'Width' },
        height: { type: 'number', description: 'Height' },
        fill_color: { type: 'string', description: 'Fill color (hex)' },
        stroke_color: { type: 'string', description: 'Stroke color (hex)' },
        stroke_width: { type: 'number', description: 'Stroke width' },
      },
      required: ['doc_id', 'section_index', 'after_index', 'x', 'y', 'width', 'height'],
    },
  },
  {
    name: 'insert_ellipse',
    description: 'Insert an ellipse drawing object (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        after_index: { type: 'number', description: 'Insert after this element index (-1 for beginning)' },
        cx: { type: 'number', description: 'Center X coordinate' },
        cy: { type: 'number', description: 'Center Y coordinate' },
        rx: { type: 'number', description: 'Radius X' },
        ry: { type: 'number', description: 'Radius Y' },
        fill_color: { type: 'string', description: 'Fill color (hex)' },
        stroke_color: { type: 'string', description: 'Stroke color (hex)' },
        stroke_width: { type: 'number', description: 'Stroke width' },
      },
      required: ['doc_id', 'section_index', 'after_index', 'cx', 'cy', 'rx', 'ry'],
    },
  },

  // === Equations ===
  {
    name: 'get_equations',
    description: 'Get all equations in the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'insert_equation',
    description: 'Insert an equation (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        after_index: { type: 'number', description: 'Insert after this element index (-1 for beginning)' },
        script: { type: 'string', description: 'Equation script (HWP equation format)' },
      },
      required: ['doc_id', 'section_index', 'after_index', 'script'],
    },
  },

  // === Memos ===
  {
    name: 'get_memos',
    description: 'Get all memos/comments in the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'insert_memo',
    description: 'Insert a memo/comment (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
        author: { type: 'string', description: 'Memo author' },
        content: { type: 'string', description: 'Memo content' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index', 'content'],
    },
  },
  {
    name: 'delete_memo',
    description: 'Delete a memo/comment (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        memo_id: { type: 'string', description: 'Memo ID to delete' },
      },
      required: ['doc_id', 'memo_id'],
    },
  },

  // === Sections ===
  {
    name: 'get_sections',
    description: 'Get all sections in the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'insert_section',
    description: 'Insert a new section (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        after_index: { type: 'number', description: 'Insert after this section index (-1 for beginning)' },
      },
      required: ['doc_id', 'after_index'],
    },
  },
  {
    name: 'delete_section',
    description: 'Delete a section (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index to delete' },
      },
      required: ['doc_id', 'section_index'],
    },
  },
  {
    name: 'get_section_xml',
    description: 'Get raw XML content of a section. Useful for AI-based document manipulation. Returns the complete section XML that can be modified and set back using set_section_xml.',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'set_section_xml',
    description: 'Set (replace) raw XML content of a section (HWPX only). WARNING: This completely replaces the section XML. The XML must be valid HWPML format. Use get_section_xml first to get the current structure, modify it, then set it back.',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
        xml: { type: 'string', description: 'New XML content (must be valid HWPML section XML)' },
        validate: { type: 'boolean', description: 'Validate XML structure before replacing (default: true)' },
      },
      required: ['doc_id', 'xml'],
    },
  },

  // === Styles ===
  {
    name: 'get_styles',
    description: 'Get all defined styles in the document',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'get_char_shapes',
    description: 'Get all character shape definitions',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'get_para_shapes',
    description: 'Get all paragraph shape definitions',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'apply_style',
    description: 'Apply a named style to a paragraph (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        paragraph_index: { type: 'number', description: 'Paragraph index' },
        style_id: { type: 'number', description: 'Style ID to apply' },
      },
      required: ['doc_id', 'section_index', 'paragraph_index', 'style_id'],
    },
  },

  // === Column Definition ===
  {
    name: 'get_column_def',
    description: 'Get column definition for a section',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'set_column_def',
    description: 'Set column definition for a section (HWPX only)',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index (default 0)' },
        count: { type: 'number', description: 'Number of columns' },
        type: { type: 'string', enum: ['newspaper', 'balanced', 'parallel'], description: 'Column type' },
        same_size: { type: 'boolean', description: 'Whether all columns have same width' },
        gap: { type: 'number', description: 'Gap between columns' },
      },
      required: ['doc_id', 'count'],
    },
  },

  // === New Document Creation ===
  {
    name: 'create_document',
    description: 'Create a new empty HWPX document',
    inputSchema: {
      type: 'object',
      properties: {
        title: { type: 'string', description: 'Document title (optional)' },
        creator: { type: 'string', description: 'Document author (optional)' },
      },
    },
  },

  // === XML Analysis and Repair ===
  {
    name: 'analyze_xml',
    description: 'Analyze document XML for issues like tag imbalance, malformed elements, etc. Useful for diagnosing save failures.',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index to analyze (optional, analyzes all sections if not specified)' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'repair_xml',
    description: 'Attempt to repair XML issues in a section. Removes orphan closing tags and fixes table structure.',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index to repair' },
        remove_orphan_close_tags: { type: 'boolean', description: 'Remove orphan closing tags (default: true)' },
        fix_table_structure: { type: 'boolean', description: 'Fix table structure issues (default: true)' },
        backup: { type: 'boolean', description: 'Keep backup of original XML (default: true)' },
      },
      required: ['doc_id', 'section_index'],
    },
  },
  {
    name: 'get_raw_section_xml',
    description: 'Get the raw XML content of a section for manual inspection or editing.',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
      },
      required: ['doc_id', 'section_index'],
    },
  },
  {
    name: 'set_raw_section_xml',
    description: 'Set the raw XML content of a section. Use with caution - validates XML structure by default.',
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        section_index: { type: 'number', description: 'Section index' },
        xml: { type: 'string', description: 'New XML content (must be valid HWPML section XML)' },
        validate: { type: 'boolean', description: 'Validate XML structure before replacing (default: true)' },
      },
      required: ['doc_id', 'section_index', 'xml'],
    },
  },

  // ===== Agentic Document Reading Tools =====
  {
    name: 'chunk_document',
    description: `📖 Split document into overlapping chunks for agentic reading.

Use this for:
- Large document analysis where full text would exceed context limits
- Semantic search across document sections
- Progressive document exploration

Returns array of chunks with:
- Unique chunk ID for reference
- Text content
- Position offsets (global character positions)
- Element type (paragraph/table/mixed)
- Metadata (char count, word count, heading level)

Chunks are cached for performance. Call invalidate_reading_cache after document modifications.`,
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        chunk_size: { type: 'number', description: 'Target chunk size in characters (default: 500)' },
        overlap: { type: 'number', description: 'Overlap between chunks in characters (default: 100)' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'search_chunks',
    description: `🔍 Search document chunks using BM25-based relevance scoring.

Returns chunks ranked by similarity to query with:
- Relevance score (higher = more relevant)
- Matched search terms
- Text snippet around first match
- Full chunk data with position info

Use for finding relevant sections in large documents without reading the entire content.`,
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        query: { type: 'string', description: 'Search query (keywords or phrase)' },
        top_k: { type: 'number', description: 'Number of top results to return (default: 5)' },
        min_score: { type: 'number', description: 'Minimum relevance score threshold (default: 0.1)' },
      },
      required: ['doc_id', 'query'],
    },
  },
  {
    name: 'get_chunk_context',
    description: `📄 Get surrounding chunks for expanded context around a specific chunk.

After finding a relevant chunk with search_chunks, use this to get additional context
by retrieving chunks before and after the target chunk.`,
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        chunk_id: { type: 'string', description: 'ID of the center chunk (from search_chunks or chunk_document)' },
        before: { type: 'number', description: 'Number of chunks before to include (default: 1)' },
        after: { type: 'number', description: 'Number of chunks after to include (default: 1)' },
      },
      required: ['doc_id', 'chunk_id'],
    },
  },
  {
    name: 'extract_toc',
    description: `📋 Extract table of contents based on Korean document formatting conventions.

Detects headings by:
- Roman numerals (I. II. III.)
- Arabic numerals (1. 2. 3.)
- Korean characters (가. 나. 다.)
- Circled numbers (① ② ③)
- Parenthesized numbers ((1) (2) (3))
- Korean consonants (ㄱ. ㄴ. ㄷ.)
- Bullet points (- • ◦)

Returns hierarchical TOC with level, title, section/element indices, and character offsets.`,
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'build_position_index',
    description: `🗂️ Build position index for document elements (headings, paragraphs, tables).

Creates a searchable index of all document elements with:
- Unique ID
- Element type (heading/paragraph/table/image)
- Text preview (first 200 chars)
- Section and element indices
- Character offset
- Heading level (if applicable)
- Table info (rows, cols) for tables

Use get_position_index to retrieve cached index, or call this to rebuild after modifications.`,
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'get_position_index',
    description: `📍 Get cached position index (builds if not available).

Returns all indexed elements. Use search_position_index for filtered queries.`,
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
  {
    name: 'search_position_index',
    description: `🔎 Search position index by text and/or element type.

Filter the position index to find specific headings, paragraphs, or tables by their text content.`,
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        query: { type: 'string', description: 'Text to search for in element content' },
        type: { type: 'string', enum: ['heading', 'paragraph', 'table'], description: 'Filter by element type (optional)' },
      },
      required: ['doc_id', 'query'],
    },
  },
  {
    name: 'get_chunk_at_offset',
    description: `📌 Get the chunk containing a specific character offset.

Use after finding a position in the index to get the full chunk context.`,
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
        offset: { type: 'number', description: 'Character offset in the document' },
      },
      required: ['doc_id', 'offset'],
    },
  },
  {
    name: 'invalidate_reading_cache',
    description: `🔄 Clear cached chunks and position index.

Call this after modifying the document to ensure fresh data on next read operation.`,
    inputSchema: {
      type: 'object',
      properties: {
        doc_id: { type: 'string', description: 'Document ID' },
      },
      required: ['doc_id'],
    },
  },
];

// ============================================================
// Server Setup
// ============================================================

const server = new Server(
  {
    name: 'hwpx-mcp-server',
    version: '0.3.0',
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

server.setRequestHandler(ListToolsRequestSchema, async () => ({ tools }));

// ============================================================
// Tool Handlers
// ============================================================

server.setRequestHandler(CallToolRequestSchema, async (request) => {
  const { name, arguments: args } = request.params;

  try {
    switch (name) {
      // === Document Management ===
      case 'open_document': {
        const filePath = args?.file_path as string;
        if (!filePath) return error('file_path is required');

        const absolutePath = path.resolve(filePath);
        const data = fs.readFileSync(absolutePath);
        const docId = generateId();

        const doc = await HwpxDocument.createFromBuffer(docId, absolutePath, data);
        openDocuments.set(docId, doc);

        return success({
          doc_id: docId,
          format: doc.format,
          path: absolutePath,
          structure: doc.getStructure(),
          metadata: doc.getMetadata(),
        });
      }

      case 'close_document': {
        const docId = args?.doc_id as string;
        if (openDocuments.delete(docId)) {
          return success({ message: 'Document closed' });
        }
        return error('Document not found');
      }

      case 'save_document': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const savePath = (args?.output_path as string) || doc.path;
        const createBackup = args?.create_backup !== false; // default: true
        const verifyIntegrity = args?.verify_integrity !== false; // default: true
        let backupPath: string | null = null;
        const tempPath = savePath + '.tmp';

        // Create backup if file exists and backup is enabled
        if (createBackup && fs.existsSync(savePath)) {
          backupPath = savePath + '.bak';
          try {
            fs.copyFileSync(savePath, backupPath);
          } catch (backupErr) {
            return error(`Failed to create backup: ${backupErr}`);
          }
        }

        try {
          const data = await doc.save();

          // Phase 1: Write to temp file first (atomic write pattern)
          fs.writeFileSync(tempPath, data);

          // Verify integrity on temp file before moving
          if (verifyIntegrity) {
            try {
              const JSZip = require('jszip');
              const savedData = fs.readFileSync(tempPath);
              const zip = await JSZip.loadAsync(savedData);

              // Check essential HWPX structure files
              const requiredFiles = [
                'mimetype',
                'Contents/content.hpf',
                'Contents/header.xml',
                'Contents/section0.xml'
              ];

              const missingFiles: string[] = [];
              for (const requiredFile of requiredFiles) {
                if (!zip.file(requiredFile)) {
                  missingFiles.push(requiredFile);
                }
              }

              if (missingFiles.length > 0) {
                throw new Error(`Missing required files: ${missingFiles.join(', ')}`);
              }

              // Verify all section XML files are valid
              const sectionFiles = Object.keys(zip.files).filter(f => f.match(/^Contents\/section\d+\.xml$/));
              for (const sectionFile of sectionFiles) {
                const file = zip.file(sectionFile);
                if (file) {
                  const xmlContent = await file.async('string');
                  if (!xmlContent || !xmlContent.includes('<?xml')) {
                    throw new Error(`Invalid XML in ${sectionFile}`);
                  }
                  // Check for truncated XML (incomplete tag at end)
                  if (xmlContent.match(/<[^>]*$/)) {
                    throw new Error(`Truncated XML in ${sectionFile}`);
                  }
                  // Check for broken opening tags (< followed by < without >)
                  if (xmlContent.match(/<[^>]*</)) {
                    throw new Error(`Broken tag structure in ${sectionFile}`);
                  }
                }
              }
            } catch (verifyErr) {
              // Clean up temp file
              if (fs.existsSync(tempPath)) {
                fs.unlinkSync(tempPath);
              }
              // Restore from backup if exists
              if (backupPath && fs.existsSync(backupPath)) {
                return error(`Save verification failed, backup preserved: ${verifyErr}`);
              }
              return error(`Save verification failed: ${verifyErr}`);
            }
          }

          // Phase 2: Atomic move - rename temp to final (atomic on same filesystem)
          if (fs.existsSync(savePath)) {
            fs.unlinkSync(savePath);
          }
          fs.renameSync(tempPath, savePath);

          return success({
            message: `Saved to ${savePath}`,
            backup_created: backupPath ? true : false,
            integrity_verified: verifyIntegrity
          });
        } catch (saveErr) {
          // Clean up temp file if exists
          if (fs.existsSync(tempPath)) {
            try { fs.unlinkSync(tempPath); } catch {}
          }
          // Restore from backup if save fails
          if (backupPath && fs.existsSync(backupPath)) {
            fs.copyFileSync(backupPath, savePath);
            return error(`Save failed, restored from backup: ${saveErr}`);
          }
          return error(`Save failed: ${saveErr}`);
        }
      }

      case 'list_open_documents': {
        const docs = Array.from(openDocuments.values()).map(d => ({
          id: d.id,
          path: d.path,
          format: d.format,
          isDirty: d.isDirty,
        }));
        return success({ documents: docs });
      }

      // === Document Info ===
      case 'get_document_text': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ text: doc.getAllText() });
      }

      case 'get_document_structure': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success(doc.getStructure());
      }

      case 'get_document_metadata': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ metadata: doc.getMetadata() });
      }

      case 'set_document_metadata': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const metadata: any = {};
        if (args?.title) metadata.title = args.title;
        if (args?.creator) metadata.creator = args.creator;
        if (args?.subject) metadata.subject = args.subject;
        if (args?.description) metadata.description = args.description;

        doc.setMetadata(metadata);
        return success({ metadata: doc.getMetadata() });
      }

      // === Paragraph Operations ===
      case 'get_paragraphs': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const sectionIndex = args?.section_index as number | undefined;
        const paragraphs = doc.getParagraphs(sectionIndex);
        return success({ paragraphs });
      }

      case 'get_paragraph': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const result = doc.getParagraph(args?.section_index as number, args?.paragraph_index as number);
        if (!result) return error('Paragraph not found');
        return success(result);
      }

      case 'insert_paragraph': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const index = doc.insertParagraph(
          args?.section_index as number,
          args?.after_index as number,
          args?.text as string
        );

        if (index === -1) return error('Failed to insert paragraph');
        return success({ message: 'Paragraph inserted', index });
      }

      case 'delete_paragraph': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.deleteParagraph(args?.section_index as number, args?.paragraph_index as number)) {
          return success({ message: 'Paragraph deleted' });
        }
        return error('Failed to delete paragraph');
      }

      case 'update_paragraph_text': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        doc.updateParagraphText(
          args?.section_index as number,
          args?.paragraph_index as number,
          args?.run_index as number ?? 0,
          args?.text as string
        );
        return success({ message: 'Paragraph updated' });
      }

      case 'append_text_to_paragraph': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        doc.appendTextToParagraph(
          args?.section_index as number,
          args?.paragraph_index as number,
          args?.text as string
        );
        return success({ message: 'Text appended' });
      }

      // === Character Styling ===
      case 'set_text_style': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const style: any = {};
        if (args?.bold !== undefined) style.bold = args.bold;
        if (args?.italic !== undefined) style.italic = args.italic;
        if (args?.underline !== undefined) style.underline = args.underline;
        if (args?.strikethrough !== undefined) style.strikethrough = args.strikethrough;
        if (args?.font_name) style.fontName = args.font_name;
        if (args?.font_size) style.fontSize = args.font_size;
        if (args?.font_color) style.fontColor = args.font_color;
        if (args?.background_color) style.backgroundColor = args.background_color;

        doc.applyCharacterStyle(
          args?.section_index as number,
          args?.paragraph_index as number,
          args?.run_index as number ?? 0,
          style
        );
        return success({ message: 'Text style applied' });
      }

      case 'get_text_style': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const style = doc.getCharacterStyle(
          args?.section_index as number,
          args?.paragraph_index as number,
          args?.run_index as number | undefined
        );
        return success({ style });
      }

      // === Paragraph Styling ===
      case 'set_paragraph_style': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const style: any = {};
        if (args?.align) style.align = args.align;
        if (args?.line_spacing) style.lineSpacing = args.line_spacing;
        if (args?.margin_left) style.marginLeft = args.margin_left;
        if (args?.margin_right) style.marginRight = args.margin_right;
        if (args?.margin_top) style.marginTop = args.margin_top;
        if (args?.margin_bottom) style.marginBottom = args.margin_bottom;
        if (args?.first_line_indent) style.firstLineIndent = args.first_line_indent;

        doc.applyParagraphStyle(
          args?.section_index as number,
          args?.paragraph_index as number,
          style
        );
        return success({ message: 'Paragraph style applied' });
      }

      case 'get_paragraph_style': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const style = doc.getParagraphStyle(
          args?.section_index as number,
          args?.paragraph_index as number
        );
        return success({ style });
      }

      // === Hanging Indent (내어쓰기) ===
      case 'set_hanging_indent': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const result = doc.setHangingIndent(
          args?.section_index as number,
          args?.paragraph_index as number,
          args?.indent_pt as number
        );
        if (!result) return error('Failed to set hanging indent. Check section/paragraph indices and indent value (must be positive).');
        return success({ message: `Hanging indent set to ${args?.indent_pt}pt` });
      }

      case 'get_hanging_indent': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const indent = doc.getHangingIndent(
          args?.section_index as number,
          args?.paragraph_index as number
        );
        if (indent === null) return error('Invalid section or paragraph index');
        return success({ hanging_indent_pt: indent });
      }

      case 'remove_hanging_indent': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const result = doc.removeHangingIndent(
          args?.section_index as number,
          args?.paragraph_index as number
        );
        if (!result) return error('Failed to remove hanging indent. Check section/paragraph indices.');
        return success({ message: 'Hanging indent removed' });
      }

      // === Table Cell Hanging Indent (테이블 셀 내어쓰기) ===
      case 'set_table_cell_hanging_indent': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const result = doc.setTableCellHangingIndent(
          args?.section_index as number,
          args?.table_index as number,
          args?.row as number,
          args?.col as number,
          args?.paragraph_index as number,
          args?.indent_pt as number
        );
        if (!result) return error('Failed to set hanging indent. Check indices and indent value (must be positive).');
        return success({ message: `Table cell hanging indent set to ${args?.indent_pt}pt` });
      }

      case 'get_table_cell_hanging_indent': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const indent = doc.getTableCellHangingIndent(
          args?.section_index as number,
          args?.table_index as number,
          args?.row as number,
          args?.col as number,
          args?.paragraph_index as number
        );
        if (indent === null) return error('Invalid indices (section, table, row, col, or paragraph)');
        return success({ hanging_indent_pt: indent });
      }

      case 'remove_table_cell_hanging_indent': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const result = doc.removeTableCellHangingIndent(
          args?.section_index as number,
          args?.table_index as number,
          args?.row as number,
          args?.col as number,
          args?.paragraph_index as number
        );
        if (!result) return error('Failed to remove hanging indent. Check indices.');
        return success({ message: 'Table cell hanging indent removed' });
      }

      // === Search & Replace ===
      case 'search_text': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const results = doc.searchText(args?.query as string, {
          caseSensitive: args?.case_sensitive as boolean,
          regex: args?.regex as boolean,
          includeTables: args?.include_tables !== false, // default true
        });

        return success({
          query: args?.query,
          total_matches: results.reduce((sum, r) => sum + r.count, 0),
          locations: results,
        });
      }

      case 'replace_text': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const count = doc.replaceText(args?.old_text as string, args?.new_text as string, {
          caseSensitive: args?.case_sensitive as boolean,
          regex: args?.regex as boolean,
          replaceAll: args?.replace_all as boolean ?? true,
        });

        return success({ message: `Replaced ${count} occurrence(s)`, count });
      }

      case 'batch_replace': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const replacements = args?.replacements as Array<{ old_text: string; new_text: string }>;
        if (!replacements) return error('replacements array is required');

        const results: any[] = [];
        for (const { old_text, new_text } of replacements) {
          const count = doc.replaceText(old_text, new_text);
          results.push({ old_text, new_text, count });
        }

        return success({ results });
      }

      case 'replace_text_in_cell': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const result = doc.replaceTextInCell(
          args?.section_index as number,
          args?.table_index as number,
          args?.row as number,
          args?.col as number,
          args?.old_text as string,
          args?.new_text as string,
          {
            caseSensitive: args?.case_sensitive as boolean,
            regex: args?.regex as boolean,
            replaceAll: args?.replace_all as boolean ?? true,
          }
        );

        if (!result.success) {
          return error(result.error || 'Replace failed');
        }

        return success({
          message: `Replaced ${result.count} occurrence(s) in cell [${args?.row}, ${args?.col}]`,
          count: result.count,
        });
      }

      // === Table Operations ===
      case 'get_tables': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ tables: doc.getTables() });
      }

      case 'get_table_map': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ table_map: doc.getTableMap() });
      }

      case 'find_empty_tables': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ empty_tables: doc.findEmptyTables() });
      }

      case 'get_tables_by_section': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        const sectionIndex = args?.section_index as number;
        if (typeof sectionIndex !== 'number') return error('section_index is required');
        return success({ tables: doc.getTablesBySection(sectionIndex) });
      }

      case 'find_table_by_header': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        const searchText = args?.search_text as string;
        if (!searchText) return error('search_text is required');
        return success({ tables: doc.findTableByHeader(searchText) });
      }

      case 'get_tables_summary': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        const startIndex = args?.start_index as number | undefined;
        const endIndex = args?.end_index as number | undefined;
        return success({ tables: doc.getTablesSummary(startIndex, endIndex) });
      }

      case 'get_document_outline': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ outline: doc.getDocumentOutline() });
      }

      // === Position/Index Helper Handlers ===
      case 'get_element_index_for_table': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        const tableIndex = args?.table_index as number;
        if (typeof tableIndex !== 'number') return error('table_index is required');
        const result = doc.getElementIndexForTable(tableIndex);
        if (!result) return error(`Table ${tableIndex} not found`);
        return success(result);
      }

      case 'find_paragraph_by_text': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        const searchText = args?.search_text as string;
        if (!searchText) return error('search_text is required');
        const sectionIndex = args?.section_index as number | undefined;
        const results = doc.findParagraphByText(searchText, sectionIndex);
        return success({ matches: results, count: results.length });
      }

      case 'get_insert_context': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        const sectionIdx = args?.section_index as number;
        const elementIdx = args?.element_index as number;
        if (typeof sectionIdx !== 'number') return error('section_index is required');
        if (typeof elementIdx !== 'number') return error('element_index is required');
        const contextRange = args?.context_range as number | undefined;
        const result = doc.getInsertContext(sectionIdx, elementIdx, contextRange);
        if (!result) return error('Invalid section or element index');
        return success(result);
      }

      case 'find_insert_position_after_header': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        const headerText = args?.header_text as string;
        if (!headerText) return error('header_text is required');
        const searchIn = (args?.search_in as 'paragraphs' | 'table_cells' | 'all') || 'all';
        const result = doc.findInsertPositionAfterHeader(headerText, searchIn);
        if (!result) return error(`Text "${headerText}" not found in ${searchIn === 'all' ? 'paragraphs or table cells' : searchIn}`);
        return success({
          ...result,
          usage_hint: result.found_in === 'table_cell'
            ? `Found in table cell. Use section_index=${result.section_index} and after_index=${result.insert_after} to insert AFTER this table, or use insert_image_in_cell with table_index=${result.table_info?.table_index}, row=${result.table_info?.row}, col=${result.table_info?.col} to insert INSIDE this cell.`
            : `Use section_index=${result.section_index} and after_index=${result.insert_after} in insert_image/render_mermaid`,
        });
      }

      case 'find_insert_position_after_table': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        const tableIndex = args?.table_index as number;
        if (typeof tableIndex !== 'number') return error('table_index is required');
        const result = doc.findInsertPositionAfterTable(tableIndex);
        if (!result) return error(`Table ${tableIndex} not found`);
        return success({
          ...result,
          usage_hint: `Use section_index=${result.section_index} and after_index=${result.insert_after} in insert_image/render_mermaid`,
        });
      }

      case 'get_table': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const table = doc.getTable(args?.section_index as number, args?.table_index as number);
        if (!table) return error('Table not found');
        return success(table);
      }

      case 'get_table_cell': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const cell = doc.getTableCell(
          args?.section_index as number,
          args?.table_index as number,
          args?.row as number,
          args?.col as number
        );
        if (!cell) return error('Cell not found');
        return success(cell);
      }

      case 'update_table_cell': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const charShapeId = args?.char_shape_id as number | undefined;
        if (doc.updateTableCell(
          args?.section_index as number,
          args?.table_index as number,
          args?.row as number,
          args?.col as number,
          args?.text as string,
          charShapeId
        )) {
          return success({ message: 'Cell updated' });
        }
        return error('Failed to update cell');
      }

      case 'set_cell_properties': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const props: any = {};
        if (args?.width) props.width = args.width;
        if (args?.height) props.height = args.height;
        if (args?.background_color) props.backgroundColor = args.background_color;
        if (args?.vertical_align) props.verticalAlign = args.vertical_align;

        if (doc.setCellProperties(
          args?.section_index as number,
          args?.table_index as number,
          args?.row as number,
          args?.col as number,
          props
        )) {
          return success({ message: 'Cell properties updated' });
        }
        return error('Failed to update cell properties');
      }

      case 'merge_cells': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.mergeCells(
          args?.section_index as number,
          args?.table_index as number,
          args?.start_row as number,
          args?.start_col as number,
          args?.end_row as number,
          args?.end_col as number
        )) {
          const colSpan = (args?.end_col as number) - (args?.start_col as number) + 1;
          const rowSpan = (args?.end_row as number) - (args?.start_row as number) + 1;
          return success({
            message: `Cells merged successfully`,
            colSpan,
            rowSpan,
            masterCell: { row: args?.start_row, col: args?.start_col }
          });
        }
        return error('Failed to merge cells. Check that the range is valid and cells are not already merged.');
      }

      case 'split_cell': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.splitCell(
          args?.section_index as number,
          args?.table_index as number,
          args?.row as number,
          args?.col as number
        )) {
          return success({
            message: `Cell split successfully`,
            cell: { row: args?.row, col: args?.col }
          });
        }
        return error('Failed to split cell. Check that the cell is actually merged (colSpan > 1 or rowSpan > 1).');
      }

      case 'insert_table_row': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.insertTableRow(
          args?.section_index as number,
          args?.table_index as number,
          args?.after_row as number,
          args?.cell_texts as string[]
        )) {
          return success({ message: 'Row inserted' });
        }
        return error('Failed to insert row');
      }

      case 'delete_table_row': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.deleteTableRow(
          args?.section_index as number,
          args?.table_index as number,
          args?.row_index as number
        )) {
          return success({ message: 'Row deleted' });
        }
        return error('Failed to delete row');
      }

      case 'insert_table_column': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.insertTableColumn(
          args?.section_index as number,
          args?.table_index as number,
          args?.after_col as number
        )) {
          return success({ message: 'Column inserted' });
        }
        return error('Failed to insert column');
      }

      case 'delete_table_column': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.deleteTableColumn(
          args?.section_index as number,
          args?.table_index as number,
          args?.col_index as number
        )) {
          return success({ message: 'Column deleted' });
        }
        return error('Failed to delete column');
      }

      case 'get_table_as_csv': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const csv = doc.getTableAsCsv(
          args?.section_index as number,
          args?.table_index as number,
          args?.delimiter as string || ','
        );
        if (!csv) return error('Table not found');
        return success({ csv });
      }

      // === Page Settings ===
      case 'get_page_settings': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const settings = doc.getPageSettings(args?.section_index as number || 0);
        return success({ settings });
      }

      case 'set_page_settings': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const settings: any = {};
        if (args?.width) settings.width = args.width;
        if (args?.height) settings.height = args.height;
        if (args?.margin_top) settings.marginTop = args.margin_top;
        if (args?.margin_bottom) settings.marginBottom = args.margin_bottom;
        if (args?.margin_left) settings.marginLeft = args.margin_left;
        if (args?.margin_right) settings.marginRight = args.margin_right;
        if (args?.orientation) settings.orientation = args.orientation;

        if (doc.setPageSettings(args?.section_index as number || 0, settings)) {
          return success({ message: 'Page settings updated' });
        }
        return error('Failed to update page settings');
      }

      // === Copy/Move ===
      case 'copy_paragraph': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.copyParagraph(
          args?.source_section as number,
          args?.source_paragraph as number,
          args?.target_section as number,
          args?.target_after as number
        )) {
          return success({ message: 'Paragraph copied' });
        }
        return error('Failed to copy paragraph');
      }

      case 'move_paragraph': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.moveParagraph(
          args?.source_section as number,
          args?.source_paragraph as number,
          args?.target_section as number,
          args?.target_after as number
        )) {
          return success({ message: 'Paragraph moved' });
        }
        return error('Failed to move paragraph');
      }

      // === Statistics ===
      case 'get_word_count': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success(doc.getWordCount());
      }

      // === Images ===
      case 'get_images': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ images: doc.getImages() });
      }

      // === Export ===
      case 'export_to_text': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const text = doc.getAllText();
        const outputPath = args?.output_path as string;
        fs.writeFileSync(outputPath, text, 'utf-8');
        return success({ message: `Exported to ${outputPath}`, characters: text.length });
      }

      case 'export_to_html': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        let html = '<!DOCTYPE html><html><head><meta charset="UTF-8">';
        html += '<style>body{font-family:sans-serif;max-width:800px;margin:0 auto;padding:20px;}table{border-collapse:collapse;width:100%;}td,th{border:1px solid #ccc;padding:8px;}</style>';
        html += '</head><body>';

        const content = doc.content;
        for (const section of content.sections) {
          for (const element of section.elements) {
            if (element.type === 'paragraph') {
              const text = element.data.runs.map(r => escapeHtml(r.text)).join('');
              html += `<p>${text}</p>`;
            } else if (element.type === 'table') {
              const table = element.data;
              html += '<table>';
              for (const row of table.rows) {
                html += '<tr>';
                for (const cell of row.cells) {
                  const text = cell.paragraphs.map(p => p.runs.map(r => escapeHtml(r.text)).join('')).join('<br>');
                  html += `<td>${text}</td>`;
                }
                html += '</tr>';
              }
              html += '</table>';
            }
          }
        }

        html += '</body></html>';
        const outputPath = args?.output_path as string;
        fs.writeFileSync(outputPath, html, 'utf-8');
        return success({ message: `Exported to ${outputPath}` });
      }

      // === Undo/Redo ===
      case 'undo': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        if (doc.undo()) {
          return success({ message: 'Undo successful', canUndo: doc.canUndo(), canRedo: doc.canRedo() });
        }
        return error('Nothing to undo');
      }

      case 'redo': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        if (doc.redo()) {
          return success({ message: 'Redo successful', canUndo: doc.canUndo(), canRedo: doc.canRedo() });
        }
        return error('Nothing to redo');
      }

      // === Table Creation ===
      case 'insert_table': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const result = doc.insertTable(
          args?.section_index as number,
          args?.after_index as number,
          args?.rows as number,
          args?.cols as number,
          { width: args?.width as number | undefined }
        );
        if (!result) return error('Failed to insert table');
        return success({ message: 'Table inserted', tableIndex: result.tableIndex });
      }

      case 'insert_nested_table': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const success_result = doc.insertNestedTable(
          args?.section_index as number,
          args?.parent_table_index as number,
          args?.row as number,
          args?.col as number,
          args?.nested_rows as number,
          args?.nested_cols as number,
          { data: args?.data as string[][] | undefined }
        );
        if (!success_result) return error('Failed to insert nested table');
        return success({ message: 'Nested table inserted successfully' });
      }

      // === Header/Footer ===
      case 'get_header': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const result = doc.getHeader(args?.section_index as number || 0);
        return success({ header: result });
      }

      case 'set_header': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.setHeader(args?.section_index as number || 0, args?.text as string)) {
          return success({ message: 'Header set successfully' });
        }
        return error('Failed to set header');
      }

      case 'get_footer': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const result = doc.getFooter(args?.section_index as number || 0);
        return success({ footer: result });
      }

      case 'set_footer': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.setFooter(args?.section_index as number || 0, args?.text as string)) {
          return success({ message: 'Footer set successfully' });
        }
        return error('Failed to set footer');
      }

      // === Footnotes/Endnotes ===
      case 'get_footnotes': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ footnotes: doc.getFootnotes() });
      }

      case 'insert_footnote': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const result = doc.insertFootnote(
          args?.section_index as number,
          args?.paragraph_index as number,
          args?.text as string
        );
        if (!result) return error('Failed to insert footnote');
        return success({ message: 'Footnote inserted', id: result.id });
      }

      case 'get_endnotes': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ endnotes: doc.getEndnotes() });
      }

      case 'insert_endnote': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const result = doc.insertEndnote(
          args?.section_index as number,
          args?.paragraph_index as number,
          args?.text as string
        );
        if (!result) return error('Failed to insert endnote');
        return success({ message: 'Endnote inserted', id: result.id });
      }

      // === Bookmarks/Hyperlinks ===
      case 'get_bookmarks': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ bookmarks: doc.getBookmarks() });
      }

      case 'insert_bookmark': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.insertBookmark(
          args?.section_index as number,
          args?.paragraph_index as number,
          args?.name as string
        )) {
          return success({ message: 'Bookmark inserted' });
        }
        return error('Failed to insert bookmark');
      }

      case 'get_hyperlinks': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ hyperlinks: doc.getHyperlinks() });
      }

      case 'insert_hyperlink': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.insertHyperlink(
          args?.section_index as number,
          args?.paragraph_index as number,
          args?.url as string,
          args?.text as string
        )) {
          return success({ message: 'Hyperlink inserted' });
        }
        return error('Failed to insert hyperlink');
      }

      // === Image Operations ===
      case 'insert_image': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const imagePath = args?.image_path as string;
        if (!fs.existsSync(imagePath)) return error('Image file not found');

        // Resolve position using after_table, after_header, or direct indices
        let sectionIndex = args?.section_index as number | undefined;
        let afterIndex = args?.after_index as number | undefined;
        let insertedAfter = '';

        const afterTable = args?.after_table as number | undefined;
        const afterHeader = args?.after_header as string | undefined;

        if (afterTable !== undefined) {
          // Insert after a specific table
          const pos = doc.findInsertPositionAfterTable(afterTable);
          if (!pos) return error(`Table ${afterTable} not found`);
          sectionIndex = pos.section_index;
          afterIndex = pos.insert_after;
          insertedAfter = `table ${afterTable} ("${pos.table_info.header.substring(0, 50)}")`;
        } else if (afterHeader) {
          // Insert after a header paragraph
          const pos = doc.findInsertPositionAfterHeader(afterHeader);
          if (!pos) return error(`Header "${afterHeader}" not found`);
          sectionIndex = pos.section_index;
          afterIndex = pos.insert_after;
          insertedAfter = `header "${pos.header_found.substring(0, 50)}"`;
        } else {
          // Use direct indices
          if (sectionIndex === undefined) return error('section_index is required when not using after_table or after_header');
          if (afterIndex === undefined) return error('after_index is required when not using after_table or after_header');
          insertedAfter = `element ${afterIndex}`;
        }

        const imageData = fs.readFileSync(imagePath);
        const ext = path.extname(imagePath).toLowerCase();
        const mimeTypes: Record<string, string> = {
          '.png': 'image/png',
          '.jpg': 'image/jpeg',
          '.jpeg': 'image/jpeg',
          '.gif': 'image/gif',
          '.bmp': 'image/bmp',
        };

        const preserveAspectRatio = args?.preserve_aspect_ratio as boolean | undefined;

        // Build position options from args
        const position: ImagePositionOptions | undefined = (
          args?.position_type || args?.vert_rel_to || args?.horz_rel_to ||
          args?.vert_align || args?.horz_align || args?.vert_offset !== undefined ||
          args?.horz_offset !== undefined || args?.text_wrap
        ) ? {
          positionType: args?.position_type as 'inline' | 'floating' | undefined,
          vertRelTo: args?.vert_rel_to as 'para' | 'paper' | undefined,
          horzRelTo: args?.horz_rel_to as 'column' | 'para' | 'paper' | undefined,
          vertAlign: args?.vert_align as 'top' | 'center' | 'bottom' | undefined,
          horzAlign: args?.horz_align as 'left' | 'center' | 'right' | undefined,
          vertOffset: args?.vert_offset as number | undefined,
          horzOffset: args?.horz_offset as number | undefined,
          textWrap: args?.text_wrap as 'top_and_bottom' | 'square' | 'tight' | 'behind_text' | 'in_front_of_text' | 'none' | undefined,
        } : undefined;

        const result = doc.insertImage(
          sectionIndex,
          afterIndex,
          {
            data: imageData.toString('base64'),
            mimeType: mimeTypes[ext] || 'image/png',
            width: args?.width as number | undefined,
            height: args?.height as number | undefined,
            preserveAspectRatio,
            position,
            headerText: afterHeader, // Pass header text for precise XML positioning
          }
        );
        if (!result) return error('Failed to insert image');

        // Get context around insertion point for verification
        const context = doc.getInsertContext(sectionIndex, afterIndex + 1, 1);

        return success({
          message: `Image inserted after ${insertedAfter}`,
          id: result.id,
          actualWidth: result.actualWidth,
          actualHeight: result.actualHeight,
          section_index: sectionIndex,
          element_index: afterIndex + 1,
          context: context ? {
            before: context.elements_before.map(e => e.text).join(' → '),
            after: context.elements_after.map(e => e.text).join(' → '),
          } : undefined,
        });
      }

      case 'update_image_size': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        // Find image ID from section and index
        const images = doc.getImages();
        const imageIndex = args?.image_index as number;
        if (imageIndex < 0 || imageIndex >= images.length) return error('Image not found');

        if (doc.updateImageSize(
          images[imageIndex].id,
          args?.width as number,
          args?.height as number
        )) {
          return success({ message: 'Image size updated' });
        }
        return error('Failed to update image size');
      }

      case 'delete_image': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const images = doc.getImages();
        const imageIndex = args?.image_index as number;
        if (imageIndex < 0 || imageIndex >= images.length) return error('Image not found');

        if (doc.deleteImage(images[imageIndex].id)) {
          return success({ message: 'Image deleted' });
        }
        return error('Failed to delete image');
      }

      case 'render_mermaid': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const mermaidCode = args?.mermaid_code as string;
        if (!mermaidCode) return error('Mermaid code is required');

        // Resolve position using after_table, after_header, or direct indices
        let sectionIndex = args?.section_index as number | undefined;
        let afterIndex = args?.after_index as number | undefined;
        let insertedAfter = '';

        const afterTable = args?.after_table as number | undefined;
        const afterHeader = args?.after_header as string | undefined;

        if (afterTable !== undefined) {
          // Insert after a specific table
          const pos = doc.findInsertPositionAfterTable(afterTable);
          if (!pos) return error(`Table ${afterTable} not found`);
          sectionIndex = pos.section_index;
          afterIndex = pos.insert_after;
          insertedAfter = `table ${afterTable} ("${pos.table_info.header.substring(0, 50)}")`;
        } else if (afterHeader) {
          // Insert after a header paragraph
          const pos = doc.findInsertPositionAfterHeader(afterHeader);
          if (!pos) return error(`Header "${afterHeader}" not found`);
          sectionIndex = pos.section_index;
          afterIndex = pos.insert_after;
          insertedAfter = `header "${pos.header_found.substring(0, 50)}"`;
        } else {
          // Use direct indices (default section to 0)
          sectionIndex = sectionIndex ?? 0;
          if (afterIndex === undefined) return error('after_index is required when not using after_table or after_header');
          insertedAfter = `element ${afterIndex}`;
        }

        // Build position options from args
        const positionOptions: ImagePositionOptions | undefined = (
          args?.position_type || args?.vert_rel_to || args?.horz_rel_to ||
          args?.vert_align || args?.horz_align || args?.vert_offset !== undefined ||
          args?.horz_offset !== undefined || args?.text_wrap
        ) ? {
          positionType: args?.position_type as 'inline' | 'floating' | undefined,
          vertRelTo: args?.vert_rel_to as 'para' | 'paper' | undefined,
          horzRelTo: args?.horz_rel_to as 'column' | 'para' | 'paper' | undefined,
          vertAlign: args?.vert_align as 'top' | 'center' | 'bottom' | undefined,
          horzAlign: args?.horz_align as 'left' | 'center' | 'right' | undefined,
          vertOffset: args?.vert_offset as number | undefined,
          horzOffset: args?.horz_offset as number | undefined,
          textWrap: args?.text_wrap as 'top_and_bottom' | 'square' | 'tight' | 'behind_text' | 'in_front_of_text' | 'none' | undefined,
        } : undefined;

        const result = await doc.renderMermaidToImage(mermaidCode, sectionIndex, afterIndex, {
          width: args?.width as number | undefined,
          height: args?.height as number | undefined,
          theme: args?.theme as 'default' | 'dark' | 'forest' | 'neutral' | undefined,
          backgroundColor: args?.background_color as string | undefined,
          preserveAspectRatio: args?.preserve_aspect_ratio as boolean | undefined,
          position: positionOptions,
          headerText: afterHeader, // Pass header text for precise XML positioning
        });

        if (result.success) {
          // Get context around insertion point for verification
          const context = doc.getInsertContext(sectionIndex, afterIndex + 1, 1);

          return success({
            message: `Mermaid diagram inserted after ${insertedAfter}`,
            image_id: result.imageId,
            actualWidth: result.actualWidth,
            actualHeight: result.actualHeight,
            section_index: sectionIndex,
            element_index: afterIndex + 1,
            context: context ? {
              before: context.elements_before.map(e => e.text).join(' → '),
              after: context.elements_after.map(e => e.text).join(' → '),
            } : undefined,
          });
        }
        return error(result.error || 'Failed to render Mermaid diagram');
      }

      case 'insert_image_in_cell': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const imagePath = args?.image_path as string;
        if (!fs.existsSync(imagePath)) return error('Image file not found');

        const globalTblIdx = args?.table_index as number;
        const rowIdx = args?.row as number;
        const colIdx = args?.col as number;

        // Convert global table index to section and local index
        const tableLocation = doc.convertGlobalToLocalTableIndex(globalTblIdx);
        if (!tableLocation) {
          return error(`Table with global index ${globalTblIdx} not found. Use get_table_map to find valid table indices.`);
        }

        const { section_index: secIdx, local_index: localTblIdx } = tableLocation;

        const imageData = fs.readFileSync(imagePath);
        const ext = path.extname(imagePath).toLowerCase();
        const mimeTypes: Record<string, string> = {
          '.png': 'image/png',
          '.jpg': 'image/jpeg',
          '.jpeg': 'image/jpeg',
          '.gif': 'image/gif',
          '.bmp': 'image/bmp',
        };

        const afterText = args?.after_text as string | undefined;

        const result = doc.insertImageInCell(
          secIdx,
          localTblIdx,
          rowIdx,
          colIdx,
          {
            data: imageData.toString('base64'),
            mimeType: mimeTypes[ext] || 'image/png',
            width: args?.width as number | undefined,
            height: args?.height as number | undefined,
            preserveAspectRatio: args?.preserve_aspect_ratio as boolean | undefined,
            afterText,
          }
        );

        if (!result) return error('Failed to insert image in cell. Check row/col indices.');

        // Get cell content for context using getTableCell
        const cellInfo = doc.getTableCell(secIdx, localTblIdx, rowIdx, colIdx);
        const cellText = cellInfo?.text?.substring(0, 30) || '';

        const positionInfo = afterText
          ? `after paragraph containing "${afterText}"`
          : 'at the beginning';

        return success({
          message: `Image inserted in cell [${rowIdx}, ${colIdx}] of table ${globalTblIdx} ${positionInfo}`,
          id: result.id,
          actualWidth: result.actualWidth,
          actualHeight: result.actualHeight,
          cell_content: cellText || '(empty cell)',
        });
      }

      case 'render_mermaid_in_cell': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const mermaidCode = args?.mermaid_code as string;
        if (!mermaidCode) return error('Mermaid code is required');

        const globalTblIdx = args?.table_index as number;
        const rowIdx = args?.row as number;
        const colIdx = args?.col as number;

        // Convert global table index to section and local index
        const tableLocation = doc.convertGlobalToLocalTableIndex(globalTblIdx);
        if (!tableLocation) {
          return error(`Table with global index ${globalTblIdx} not found. Use get_table_map to find valid table indices.`);
        }

        const { section_index: secIdx, local_index: localTblIdx } = tableLocation;

        // Fetch Mermaid diagram from mermaid.ink API using pako compression (same as renderMermaidToImage)
        const pako = await import('pako');
        const theme = args?.theme as string || 'default';
        const bgColor = args?.background_color as string;

        const stateObject = {
          code: mermaidCode,
          mermaid: { theme: theme },
          autoSync: true,
          updateDiagram: true
        };

        const jsonString = JSON.stringify(stateObject);
        const compressed = pako.deflate(jsonString, { level: 9 });
        const base64Code = Buffer.from(compressed)
          .toString('base64')
          .replace(/\+/g, '-')
          .replace(/\//g, '_');

        let url = `https://mermaid.ink/img/pako:${base64Code}?type=png`;
        if (bgColor) {
          const bgColorClean = bgColor.replace(/^#/, '');
          url += `&bgColor=${bgColorClean}`;
        }

        try {
          const response = await fetch(url);

          if (!response.ok) {
            return error(`Failed to render Mermaid diagram: ${response.statusText}`);
          }

          const imageBuffer = Buffer.from(await response.arrayBuffer());

          const afterText = args?.after_text as string | undefined;

          const result = doc.insertImageInCell(
            secIdx,
            localTblIdx,
            rowIdx,
            colIdx,
            {
              data: imageBuffer.toString('base64'),
              mimeType: 'image/png',
              width: args?.width as number | undefined,
              height: args?.height as number | undefined,
              preserveAspectRatio: args?.preserve_aspect_ratio !== false, // default true for Mermaid
              afterText,
            }
          );

          if (!result) return error('Failed to insert Mermaid diagram in cell. Check row/col indices.');

          // Get cell content for context using getTableCell
          const cellInfo = doc.getTableCell(secIdx, localTblIdx, rowIdx, colIdx);
          const cellText = cellInfo?.text?.substring(0, 30) || '';

          const positionInfo = afterText
            ? `after paragraph containing "${afterText}"`
            : 'at the beginning';

          return success({
            message: `Mermaid diagram inserted in cell [${rowIdx}, ${colIdx}] of table ${globalTblIdx} ${positionInfo}`,
            image_id: result.id,
            actualWidth: result.actualWidth,
            actualHeight: result.actualHeight,
            cell_content: cellText || '(empty cell)',
          });
        } catch (err: unknown) {
          const errorMessage = err instanceof Error ? err.message : String(err);
          return error(`Failed to fetch Mermaid diagram: ${errorMessage}`);
        }
      }

      // === Drawing Objects ===
      case 'insert_line': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const result = doc.insertLine(
          args?.section_index as number,
          args?.x1 as number,
          args?.y1 as number,
          args?.x2 as number,
          args?.y2 as number,
          {
            color: args?.stroke_color as string,
            width: args?.stroke_width as number,
          }
        );
        if (!result) return error('Failed to insert line');
        return success({ message: 'Line inserted', id: result.id });
      }

      case 'insert_rect': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const result = doc.insertRect(
          args?.section_index as number,
          args?.x as number,
          args?.y as number,
          args?.width as number,
          args?.height as number,
          {
            fillColor: args?.fill_color as string,
            strokeColor: args?.stroke_color as string,
          }
        );
        if (!result) return error('Failed to insert rectangle');
        return success({ message: 'Rectangle inserted', id: result.id });
      }

      case 'insert_ellipse': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const result = doc.insertEllipse(
          args?.section_index as number,
          args?.cx as number,
          args?.cy as number,
          args?.rx as number,
          args?.ry as number,
          {
            fillColor: args?.fill_color as string,
            strokeColor: args?.stroke_color as string,
          }
        );
        if (!result) return error('Failed to insert ellipse');
        return success({ message: 'Ellipse inserted', id: result.id });
      }

      // === Equations ===
      case 'get_equations': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ equations: doc.getEquations() });
      }

      case 'insert_equation': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const result = doc.insertEquation(
          args?.section_index as number,
          args?.after_index as number,
          args?.script as string
        );
        if (!result) return error('Failed to insert equation');
        return success({ message: 'Equation inserted', id: result.id });
      }

      // === Memos ===
      case 'get_memos': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ memos: doc.getMemos() });
      }

      case 'insert_memo': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const result = doc.insertMemo(
          args?.section_index as number,
          args?.paragraph_index as number,
          args?.content as string,
          args?.author as string
        );
        if (!result) return error('Failed to insert memo');
        return success({ message: 'Memo inserted', id: result.id });
      }

      case 'delete_memo': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.deleteMemo(args?.memo_id as string)) {
          return success({ message: 'Memo deleted' });
        }
        return error('Failed to delete memo');
      }

      // === Sections ===
      case 'get_sections': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ sections: doc.getSections() });
      }

      case 'insert_section': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const newIndex = doc.insertSection(args?.after_index as number);
        return success({ message: 'Section inserted', index: newIndex });
      }

      case 'delete_section': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.deleteSection(args?.section_index as number)) {
          return success({ message: 'Section deleted' });
        }
        return error('Failed to delete section');
      }

      case 'get_section_xml': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const sectionIndex = (args?.section_index as number) ?? 0;
        const xml = await doc.getSectionXml(sectionIndex);
        if (xml === null) {
          return error(`Section ${sectionIndex} not found or document is HWP format`);
        }
        return success({ section_index: sectionIndex, xml });
      }

      case 'set_section_xml': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const sectionIndex = (args?.section_index as number) ?? 0;
        const xml = args?.xml as string;
        const validate = (args?.validate as boolean) ?? true;

        if (!xml) {
          return error('XML content is required');
        }

        const result = await doc.setSectionXml(sectionIndex, xml, validate);
        if (result.success) {
          return success({ message: `Section ${sectionIndex} XML replaced successfully` });
        }
        return error(result.error || 'Failed to set section XML');
      }

      // === Styles ===
      case 'get_styles': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ styles: doc.getStyles() });
      }

      case 'get_char_shapes': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ charShapes: doc.getCharShapes() });
      }

      case 'get_para_shapes': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ paraShapes: doc.getParaShapes() });
      }

      case 'apply_style': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.applyStyle(
          args?.section_index as number,
          args?.paragraph_index as number,
          args?.style_id as number
        )) {
          return success({ message: 'Style applied' });
        }
        return error('Failed to apply style');
      }

      // === Column Definition ===
      case 'get_column_def': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        return success({ columnDef: doc.getColumnDef(args?.section_index as number || 0) });
      }

      case 'set_column_def': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        if (doc.setColumnDef(
          args?.section_index as number || 0,
          args?.count as number,
          args?.gap as number
        )) {
          return success({ message: 'Column definition set' });
        }
        return error('Failed to set column definition');
      }

      // === Create New Document ===
      case 'create_document': {
        const docId = generateId();
        const doc = HwpxDocument.createNew(docId, args?.title as string, args?.creator as string);
        openDocuments.set(docId, doc);
        return success({
          doc_id: docId,
          format: 'hwpx',
          message: 'New document created',
        });
      }

      // === XML Analysis and Repair ===
      case 'analyze_xml': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const sectionIndex = args?.section_index as number | undefined;
        const result = await doc.analyzeXml(sectionIndex);

        return success({
          has_issues: result.hasIssues,
          summary: result.summary,
          sections: result.sections.map(s => ({
            section_index: s.sectionIndex,
            issues: s.issues,
            tag_counts: s.tagCounts,
          })),
        });
      }

      case 'repair_xml': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const sectionIndex = args?.section_index as number;
        if (sectionIndex === undefined) return error('section_index is required');

        const result = await doc.repairXml(sectionIndex, {
          removeOrphanCloseTags: args?.remove_orphan_close_tags as boolean | undefined,
          fixTableStructure: args?.fix_table_structure as boolean | undefined,
          backup: args?.backup as boolean | undefined,
        });

        return success({
          success: result.success,
          message: result.message,
          repairs_applied: result.repairsApplied,
          has_original_backup: !!result.originalXml,
        });
      }

      case 'get_raw_section_xml': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const sectionIndex = args?.section_index as number;
        if (sectionIndex === undefined) return error('section_index is required');

        const xml = await doc.getRawSectionXml(sectionIndex);
        if (xml === null) return error(`Section ${sectionIndex} not found`);

        return success({
          section_index: sectionIndex,
          xml_length: xml.length,
          xml: xml,
        });
      }

      case 'set_raw_section_xml': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');
        if (doc.format === 'hwp') return error('HWP files are read-only');

        const sectionIndex = args?.section_index as number;
        const xml = args?.xml as string;
        const validate = args?.validate !== false; // default: true

        if (sectionIndex === undefined) return error('section_index is required');
        if (!xml) return error('xml is required');

        const result = await doc.setRawSectionXml(sectionIndex, xml, validate);

        if (result.success) {
          return success({
            success: true,
            message: result.message,
          });
        } else {
          return success({
            success: false,
            message: result.message,
            issues: result.issues,
          });
        }
      }

      // ===== Agentic Document Reading Handlers =====

      case 'chunk_document': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const chunkSize = (args?.chunk_size as number) || 500;
        const overlap = (args?.overlap as number) || 100;

        const chunks = doc.chunkDocument(chunkSize, overlap);
        return success({
          total_chunks: chunks.length,
          chunk_size: chunkSize,
          overlap: overlap,
          chunks: chunks.map(c => ({
            id: c.id,
            text: c.text,
            start_offset: c.startOffset,
            end_offset: c.endOffset,
            section_index: c.sectionIndex,
            element_type: c.elementType,
            element_index: c.elementIndex,
            metadata: c.metadata,
          })),
        });
      }

      case 'search_chunks': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const query = args?.query as string;
        if (!query) return error('query is required');

        const topK = (args?.top_k as number) || 5;
        const minScore = (args?.min_score as number) || 0.1;

        const results = doc.searchChunks(query, topK, minScore);
        return success({
          query,
          total_results: results.length,
          results: results.map(r => ({
            chunk_id: r.chunk.id,
            score: r.score,
            matched_terms: r.matchedTerms,
            snippet: r.snippet,
            chunk: {
              text: r.chunk.text,
              start_offset: r.chunk.startOffset,
              end_offset: r.chunk.endOffset,
              section_index: r.chunk.sectionIndex,
              element_type: r.chunk.elementType,
              metadata: r.chunk.metadata,
            },
          })),
        });
      }

      case 'get_chunk_context': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const chunkId = args?.chunk_id as string;
        if (!chunkId) return error('chunk_id is required');

        const before = (args?.before as number) || 1;
        const after = (args?.after as number) || 1;

        const context = doc.getChunkContext(chunkId, before, after);
        return success({
          center_index: context.centerIndex,
          total_chunks: context.chunks.length,
          chunks: context.chunks.map(c => ({
            id: c.id,
            text: c.text,
            start_offset: c.startOffset,
            end_offset: c.endOffset,
            section_index: c.sectionIndex,
            element_type: c.elementType,
            metadata: c.metadata,
          })),
        });
      }

      case 'extract_toc': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const toc = doc.extractToc();
        return success({
          total_entries: toc.length,
          toc: toc.map(t => ({
            level: t.level,
            title: t.title,
            section_index: t.sectionIndex,
            element_index: t.elementIndex,
            offset: t.offset,
          })),
        });
      }

      case 'build_position_index': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const index = doc.buildPositionIndex();
        return success({
          total_entries: index.length,
          index: index.map(e => ({
            id: e.id,
            type: e.type,
            text: e.text,
            section_index: e.sectionIndex,
            element_index: e.elementIndex,
            offset: e.offset,
            level: e.level,
            table_info: e.tableInfo,
          })),
        });
      }

      case 'get_position_index': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const index = doc.getPositionIndex();
        return success({
          total_entries: index.length,
          index: index.map(e => ({
            id: e.id,
            type: e.type,
            text: e.text,
            section_index: e.sectionIndex,
            element_index: e.elementIndex,
            offset: e.offset,
            level: e.level,
            table_info: e.tableInfo,
          })),
        });
      }

      case 'search_position_index': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const query = args?.query as string;
        if (!query) return error('query is required');

        const type = args?.type as 'heading' | 'paragraph' | 'table' | undefined;

        const results = doc.searchPositionIndex(query, type);
        return success({
          query,
          type_filter: type || 'all',
          total_results: results.length,
          results: results.map(e => ({
            id: e.id,
            type: e.type,
            text: e.text,
            section_index: e.sectionIndex,
            element_index: e.elementIndex,
            offset: e.offset,
            level: e.level,
            table_info: e.tableInfo,
          })),
        });
      }

      case 'get_chunk_at_offset': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        const offset = args?.offset as number;
        if (offset === undefined) return error('offset is required');

        const chunk = doc.getChunkAtOffset(offset);
        if (!chunk) {
          return success({ found: false, message: 'No chunk found at this offset' });
        }
        return success({
          found: true,
          chunk: {
            id: chunk.id,
            text: chunk.text,
            start_offset: chunk.startOffset,
            end_offset: chunk.endOffset,
            section_index: chunk.sectionIndex,
            element_type: chunk.elementType,
            metadata: chunk.metadata,
          },
        });
      }

      case 'invalidate_reading_cache': {
        const doc = getDoc(args?.doc_id as string);
        if (!doc) return error('Document not found');

        doc.invalidateReadingCache();
        return success({ success: true, message: 'Reading cache invalidated' });
      }

      default:
        return error(`Unknown tool: ${name}`);
    }
  } catch (err: any) {
    return error(err.message);
  }
});

// ============================================================
// Helper Functions
// ============================================================

function getDoc(docId: string): HwpxDocument | undefined {
  return openDocuments.get(docId);
}

function success(data: any) {
  return { content: [{ type: 'text', text: JSON.stringify(data, null, 2) }] };
}

function error(message: string) {
  return { content: [{ type: 'text', text: JSON.stringify({ error: message }) }] };
}

function escapeHtml(text: string): string {
  return text
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// ============================================================
// Main
// ============================================================

async function main() {
  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch(console.error);
