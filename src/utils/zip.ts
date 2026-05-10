import JSZip from 'jszip';
import { DocxError } from '../errors.js';

export type DocxZip = JSZip;

/**
 * Load a DOCX file as a ZIP archive
 */
export async function loadZip(buffer: Buffer | ArrayBuffer): Promise<DocxZip> {
	try {
		const zip = await JSZip.loadAsync(buffer);

		// Validate it's a DOCX by checking for required parts
		if (!zip.file('word/document.xml')) {
			throw new DocxError('MISSING_PART', 'Document is missing word/document.xml');
		}
		if (!zip.file('[Content_Types].xml')) {
			throw new DocxError('MISSING_PART', 'Document is missing [Content_Types].xml');
		}

		return zip;
	} catch (error) {
		if (error instanceof DocxError) {
			throw error;
		}
		throw new DocxError(
			'INVALID_ARCHIVE',
			`File is not a valid ZIP archive: ${error instanceof Error ? error.message : String(error)}`
		);
	}
}

/**
 * Save a ZIP archive back to a buffer
 */
export async function saveZip(zip: DocxZip): Promise<Buffer> {
	return zip.generateAsync({
		type: 'nodebuffer',
		compression: 'DEFLATE',
		compressionOptions: { level: 6 },
	});
}

/**
 * Get a document part (XML file) from the ZIP
 */
export async function getDocumentPart(zip: DocxZip, path: string): Promise<string | null> {
	const file = zip.file(path);
	if (!file) {
		return null;
	}
	return file.async('string');
}

/**
 * Set a document part (XML file) in the ZIP
 */
export function setDocumentPart(zip: DocxZip, path: string, content: string): void {
	zip.file(path, content);
}

/**
 * Check if a part exists in the ZIP
 */
export function hasPart(zip: DocxZip, path: string): boolean {
	return zip.file(path) !== null;
}

/**
 * List all files in the ZIP
 */
export function listParts(zip: DocxZip): string[] {
	const paths: string[] = [];
	zip.forEach((relativePath) => {
		paths.push(relativePath);
	});
	return paths;
}
