/**
 * Represents a paragraph in the document read-model.
 * Paragraphs are the primary addressable unit for edits.
 *
 * @example
 * ```typescript
 * const doc = await loadTrackable(buffer);
 * for (const para of doc.paragraphs) {
 *   console.log(`${para.id}: ${para.text}`);
 * }
 * ```
 */
export interface Paragraph {
	/**
	 * Unique identifier for this paragraph.
	 * Format: `{part}:{hexId}` where part is 'body', 'header1', 'footer1', 'footnotes', or 'endnotes'.
	 * @example 'body:4F2A8B91'
	 */
	id: string;

	/** Plain text content of the paragraph (concatenated from all runs) */
	text: string;

	/** Formatted segments within the paragraph */
	runs: Run[];

	/** Paragraph style name (e.g., 'Heading1', 'Normal'), if any */
	style?: string;

	/** Numbering info if this paragraph is part of a numbered or bulleted list */
	numbering?: {
		/** Indentation level (0 = top level) */
		level: number;
		/** Reference to the numbering definition */
		numId: string;
	};

	/** Table location if this paragraph is inside a table cell */
	table?: {
		/** Which table in the document (0-indexed) */
		tableIndex: number;
		/** Row index within the table (0-indexed) */
		row: number;
		/** Cell index within the row (0-indexed) */
		cell: number;
	};
}

/**
 * A formatted text segment within a paragraph.
 * A paragraph consists of one or more runs, each with its own formatting.
 */
export interface Run {
	/** The text content of this run */
	text: string;
	/** Whether the text is bold */
	bold?: boolean;
	/** Whether the text is italic */
	italic?: boolean;
	/** Whether the text is underlined */
	underline?: boolean;
	/** Whether the text has strikethrough */
	strike?: boolean;
	/** Character style name, if any */
	style?: string;
}

/**
 * Content specification for insertions and replacements.
 * Can be a plain string or an object with text and formatting options.
 *
 * @example
 * ```typescript
 * // Simple string
 * const simple: ContentRun = 'Hello, World!';
 *
 * // With formatting
 * const formatted: ContentRun = { text: 'Important', bold: true };
 * ```
 */
export type ContentRun =
	| string
	| {
			text: string;
			bold?: boolean;
			italic?: boolean;
			underline?: boolean;
	  };

/**
 * Replace the entire content of a paragraph.
 * The old content is wrapped in `<w:del>` and new content in `<w:ins>`.
 *
 * @example
 * ```typescript
 * const edit: ReplaceParagraphEdit = {
 *   type: 'replaceParagraph',
 *   paraId: 'body:4F2A8B91',
 *   content: [{ text: 'New content', bold: true }]
 * };
 * ```
 */
export interface ReplaceParagraphEdit {
	type: 'replaceParagraph';
	/** ID of the paragraph to replace */
	paraId: string;
	/** New content for the paragraph */
	content: ContentRun[];
	/** Keep paragraph properties (alignment, spacing, indentation). Default: true */
	preserveStyle?: boolean;
}

/**
 * Insert a new paragraph after the specified paragraph.
 * The new paragraph is wrapped in `<w:ins>`.
 *
 * @example
 * ```typescript
 * const edit: InsertAfterEdit = {
 *   type: 'insertAfter',
 *   paraId: 'body:4F2A8B91',
 *   content: ['New paragraph text']
 * };
 * ```
 */
export interface InsertAfterEdit {
	type: 'insertAfter';
	/** ID of the paragraph to insert after */
	paraId: string;
	/** Content for the new paragraph */
	content: ContentRun[];
	/** Style to apply to the new paragraph. Inherits from reference paragraph if omitted. */
	style?: string;
}

/**
 * Insert a new paragraph before the specified paragraph.
 * The new paragraph is wrapped in `<w:ins>`.
 */
export interface InsertBeforeEdit {
	type: 'insertBefore';
	/** ID of the paragraph to insert before */
	paraId: string;
	/** Content for the new paragraph */
	content: ContentRun[];
	/** Style to apply to the new paragraph. Inherits from reference paragraph if omitted. */
	style?: string;
}

/**
 * Delete a single paragraph.
 * The paragraph content is wrapped in `<w:del>`.
 */
export interface DeleteParagraphEdit {
	type: 'deleteParagraph';
	/** ID of the paragraph to delete */
	paraId: string;
}

/**
 * Delete a range of consecutive paragraphs (inclusive).
 * Each paragraph in the range has its content wrapped in `<w:del>`.
 */
export interface DeleteRangeEdit {
	type: 'deleteRange';
	/** ID of the first paragraph to delete */
	from: string;
	/** ID of the last paragraph to delete (inclusive) */
	to: string;
}

/**
 * Replace text within a paragraph.
 * The found text is wrapped in `<w:del>` and replacement in `<w:ins>`.
 *
 * @example
 * ```typescript
 * const edit: ReplaceTextEdit = {
 *   type: 'replaceText',
 *   paraId: 'body:4F2A8B91',
 *   find: 'indemnify',
 *   replace: 'hold harmless'
 * };
 * ```
 */
export interface ReplaceTextEdit {
	type: 'replaceText';
	/** ID of the paragraph containing the text */
	paraId: string;
	/** Text to find */
	find: string;
	/** Text to replace with (empty string for deletion) */
	replace: string;
	/** Which occurrence to replace (default: 1 = first). Ignored if `all` is true. */
	occurrence?: number;
	/** Replace all occurrences. Default: false */
	all?: boolean;
}

/**
 * Union type for all edit operations.
 * Use this when building arrays of edits to apply.
 */
export type Edit =
	| ReplaceParagraphEdit
	| InsertAfterEdit
	| InsertBeforeEdit
	| DeleteParagraphEdit
	| DeleteRangeEdit
	| ReplaceTextEdit;

/**
 * Options for applying edits with tracked changes.
 */
export interface EditOptions {
	/** Author name to record in tracked changes */
	author: string;
	/** Date to record in tracked changes. Defaults to now. */
	date?: Date;
	/** Continue applying edits if one fails. Default: true */
	continueOnError?: boolean;
}

/**
 * Result of applying edits to a document.
 */
export interface EditResult {
	/** The modified DOCX file as a Buffer */
	buffer: Buffer;
	/** Edits that were successfully applied */
	applied: AppliedEdit[];
	/** Edits that failed to apply */
	failed: FailedEdit[];
}

/**
 * Information about a successfully applied edit.
 */
export interface AppliedEdit {
	/** The edit that was applied */
	edit: Edit;
	/** The w:id values assigned to the tracked changes (for accept/reject) */
	changeIds: string[];
}

/**
 * Information about a failed edit.
 */
export interface FailedEdit {
	/** The edit that failed */
	edit: Edit;
	/** Human-readable reason for the failure */
	reason: string;
}

/**
 * Represents an existing tracked change in the document.
 */
export interface TrackedChange {
	/** Unique ID of the tracked change (from w:id attribute) */
	id: string;
	/** Type of change: 'insertion' or 'deletion' */
	type: 'insertion' | 'deletion';
	/** Author who made the change */
	author: string;
	/** When the change was made */
	date: Date;
	/** The text content of the change */
	text: string;
}

/**
 * Specifies how to resolve (accept or reject) a tracked change.
 *
 * @example
 * ```typescript
 * await doc.resolveChanges([
 *   { changeId: '1', action: 'accept' },
 *   { changeId: '2', action: 'reject' }
 * ]);
 * ```
 */
export interface ChangeResolution {
	/** ID of the tracked change to resolve */
	changeId: string;
	/** Whether to accept or reject the change */
	action: 'accept' | 'reject';
}
