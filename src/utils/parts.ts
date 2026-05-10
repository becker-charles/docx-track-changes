import type { DocxZip } from './zip.js';
import { getDocumentPart } from './zip.js';
import { XMLParser } from 'fast-xml-parser';

/**
 * Path to the main document part
 */
export const MAIN_DOCUMENT_PATH = 'word/document.xml';

/**
 * Document parts info
 */
export interface DocumentParts {
	main: string;
	headers: string[];
	footers: string[];
	footnotes: string | null;
	endnotes: string | null;
}

// Relationship types
const REL_TYPES = {
	header: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header',
	footer: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer',
	footnotes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes',
	endnotes: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes',
} as const;

/**
 * Get all document part paths from a DOCX file
 */
export async function getPartPaths(zip: DocxZip): Promise<DocumentParts> {
	const result: DocumentParts = {
		main: MAIN_DOCUMENT_PATH,
		headers: [],
		footers: [],
		footnotes: null,
		endnotes: null,
	};

	// Read the document relationships file
	const relsPath = 'word/_rels/document.xml.rels';
	const relsXml = await getDocumentPart(zip, relsPath);

	if (!relsXml) {
		return result;
	}

	const parser = new XMLParser({
		ignoreAttributes: false,
		attributeNamePrefix: '@_',
	});

	const parsed = parser.parse(relsXml);
	const relationships = parsed?.Relationships?.Relationship;

	if (!relationships) {
		return result;
	}

	// Normalize to array
	const rels = Array.isArray(relationships) ? relationships : [relationships];

	for (const rel of rels) {
		const type = rel['@_Type'];
		const target = rel['@_Target'];

		if (!type || !target) continue;

		// Resolve relative path
		const fullPath = target.startsWith('/') ? target.slice(1) : `word/${target}`;

		if (type === REL_TYPES.header) {
			result.headers.push(fullPath);
		} else if (type === REL_TYPES.footer) {
			result.footers.push(fullPath);
		} else if (type === REL_TYPES.footnotes) {
			result.footnotes = fullPath;
		} else if (type === REL_TYPES.endnotes) {
			result.endnotes = fullPath;
		}
	}

	// Sort headers and footers for consistent ordering
	result.headers.sort();
	result.footers.sort();

	return result;
}
