import { readFile } from 'node:fs/promises';
import { TrackableDocument } from './TrackableDocument.js';
import { loadZip } from './utils/zip.js';

/**
 * Load a DOCX file from a buffer and return a TrackableDocument.
 *
 * This is the primary entry point for working with DOCX files.
 * The returned document provides a read model with stable paragraph IDs
 * and methods to apply tracked changes.
 *
 * @param buffer - The DOCX file contents as a Buffer or ArrayBuffer
 * @returns A TrackableDocument ready for reading and editing
 * @throws {DocxError} If the file is not a valid DOCX archive
 *
 * @example
 * ```typescript
 * import { loadTrackable } from 'docx-track-changes';
 * import { readFile } from 'fs/promises';
 *
 * const buffer = await readFile('document.docx');
 * const doc = await loadTrackable(buffer);
 * console.log(doc.body.map(p => p.text));
 * ```
 */
export async function loadTrackable(
	buffer: Buffer | ArrayBuffer
): Promise<TrackableDocument> {
	const zip = await loadZip(buffer);
	return TrackableDocument.fromZip(zip);
}

/**
 * Load a DOCX file from a file path and return a TrackableDocument.
 *
 * Convenience function that reads the file from disk before loading.
 *
 * @param path - Path to the DOCX file
 * @returns A TrackableDocument ready for reading and editing
 * @throws {DocxError} If the file is not a valid DOCX archive
 *
 * @example
 * ```typescript
 * const doc = await loadTrackableFile('contract.docx');
 * ```
 */
export async function loadTrackableFile(path: string): Promise<TrackableDocument> {
	const buffer = await readFile(path);
	return loadTrackable(buffer);
}
