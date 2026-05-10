/**
 * Error codes for docx-track-changes errors.
 *
 * - `INVALID_ARCHIVE` - The file is not a valid ZIP archive
 * - `MISSING_PART` - Required document parts (like word/document.xml) are missing
 * - `INVALID_XML` - XML parsing failed
 * - `INVALID_EDIT` - An edit operation was invalid (e.g., referencing non-existent paragraph)
 */
export type DocxErrorCode =
	| 'INVALID_ARCHIVE'
	| 'MISSING_PART'
	| 'INVALID_XML'
	| 'INVALID_EDIT';

/**
 * Custom error class for docx-track-changes errors.
 *
 * All errors thrown by this library are instances of DocxError,
 * allowing you to distinguish library errors from other exceptions.
 *
 * @example
 * ```typescript
 * try {
 *   const doc = await loadTrackable(buffer);
 * } catch (error) {
 *   if (error instanceof DocxError) {
 *     console.error(`DOCX Error [${error.code}]: ${error.message}`);
 *   }
 * }
 * ```
 */
export class DocxError extends Error {
	/** The error code indicating the type of error */
	readonly code: DocxErrorCode;

	constructor(code: DocxErrorCode, message: string) {
		super(message);
		this.name = 'DocxError';
		this.code = code;
	}
}
