import type { Paragraph } from '../types.js';
import { findElements, getChildren, getTagName, type ParsedXml } from './parser.js';
import { extractRuns } from './runs.js';
import { ensureParaIds } from './ids.js';

/**
 * Context for tracking table position
 */
interface TableContext {
	tableIndex: number;
	row: number;
	cell: number;
}

/**
 * Extract all paragraphs from a parsed document part
 */
export function extractParagraphs(
	parsed: ParsedXml,
	partPrefix: string
): Paragraph[] {
	const paragraphs: Paragraph[] = [];

	// First, ensure all paragraphs have IDs
	const idMap = ensureParaIds(parsed, partPrefix);

	// Find all paragraph elements
	const pElements = findElements(parsed, 'w:p');

	for (const { element } of pElements) {
		const id = idMap.get(element);
		if (!id) continue; // Should never happen

		// Get table context if applicable
		const tableContext = getTableContext(element, parsed);

		// Extract runs and build paragraph
		const runs = extractRuns(element);
		const text = runs.map(r => r.text).join('');

		// Get paragraph properties
		const { style, numbering } = extractParagraphProperties(element);

		const para: Paragraph = {
			id,
			text,
			runs,
		};

		if (style) para.style = style;
		if (numbering) para.numbering = numbering;
		if (tableContext) para.table = tableContext;

		paragraphs.push(para);
	}

	return paragraphs;
}

/**
 * Extract paragraph properties (style, numbering)
 */
function extractParagraphProperties(pElement: Record<string, unknown>): {
	style?: string;
	numbering?: { level: number; numId: string };
} {
	const result: {
		style?: string;
		numbering?: { level: number; numId: string };
	} = {};

	const children = getChildren(pElement);

	for (const child of children) {
		if (typeof child !== 'object' || child === null) continue;

		const childObj = child as Record<string, unknown>;
		const tagName = getTagName(childObj);

		if (tagName === 'w:pPr') {
			const pPrChildren = getChildren(childObj);

			for (const pPrChild of pPrChildren) {
				if (typeof pPrChild !== 'object' || pPrChild === null) continue;

				const pPrChildObj = pPrChild as Record<string, unknown>;
				const pPrTagName = getTagName(pPrChildObj);

				if (pPrTagName === 'w:pStyle') {
					const attrs = pPrChildObj[':@'] as Record<string, string> | undefined;
					if (attrs?.['@_w:val']) {
						result.style = attrs['@_w:val'];
					}
				} else if (pPrTagName === 'w:numPr') {
					const numPr = extractNumberingProperties(pPrChildObj);
					if (numPr) {
						result.numbering = numPr;
					}
				}
			}
			break; // Only one pPr element
		}
	}

	return result;
}

/**
 * Extract numbering properties from w:numPr
 */
function extractNumberingProperties(
	numPrElement: Record<string, unknown>
): { level: number; numId: string } | null {
	let level: number | null = null;
	let numId: string | null = null;

	const children = getChildren(numPrElement);

	for (const child of children) {
		if (typeof child !== 'object' || child === null) continue;

		const childObj = child as Record<string, unknown>;
		const tagName = getTagName(childObj);
		const attrs = childObj[':@'] as Record<string, string> | undefined;

		if (tagName === 'w:ilvl' && attrs?.['@_w:val']) {
			level = parseInt(attrs['@_w:val'], 10);
		} else if (tagName === 'w:numId' && attrs?.['@_w:val']) {
			numId = attrs['@_w:val'];
		}
	}

	if (level !== null && numId !== null) {
		return { level, numId };
	}
	return null;
}

/**
 * Get table context for a paragraph (if it's inside a table)
 */
function getTableContext(
	pElement: Record<string, unknown>,
	parsed: ParsedXml
): TableContext | undefined {
	// Find all tables and their cells
	const tables = findElements(parsed, 'w:tbl');

	for (let tableIndex = 0; tableIndex < tables.length; tableIndex++) {
		const table = tables[tableIndex];
		if (!table) continue;

		const rows = findElements([table.element], 'w:tr');

		for (let rowIdx = 0; rowIdx < rows.length; rowIdx++) {
			const row = rows[rowIdx];
			if (!row) continue;

			const cells = findElements([row.element], 'w:tc');

			for (let cellIdx = 0; cellIdx < cells.length; cellIdx++) {
				const cell = cells[cellIdx];
				if (!cell) continue;

				// Find paragraphs in this cell
				const cellParas = findElements([cell.element], 'w:p');

				for (const cellPara of cellParas) {
					if (cellPara.element === pElement) {
						return {
							tableIndex,
							row: rowIdx,
							cell: cellIdx,
						};
					}
				}
			}
		}
	}

	return undefined;
}
