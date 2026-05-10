import type { InsertAfterEdit, InsertBeforeEdit, AppliedEdit } from '../types.js';
import type { ParsedXml } from '../readModel/parser.js';
import { getChildren } from '../readModel/parser.js';
import { findParagraphById } from './deleteParagraph.js';
import { wrapParagraphInsert } from '../tracking/wrapInsert.js';
import { generateChangeId } from '../tracking/changeIds.js';
import { mintParaId } from '../readModel/ids.js';
import { cloneElement } from '../tracking/wrapDelete.js';

/**
 * Extract paragraph properties (w:pPr) from a paragraph element.
 * Returns a cloned copy of the pPr children, or undefined if none.
 */
function extractParagraphProperties(paragraph: Record<string, unknown>): unknown[] | undefined {
	const children = getChildren(paragraph);

	for (const child of children) {
		if (typeof child !== 'object' || child === null) continue;

		const childObj = child as Record<string, unknown>;
		if ('w:pPr' in childObj) {
			// Clone the pPr children
			const pPrChildren = childObj['w:pPr'] as unknown[];
			return pPrChildren.map(c => {
				if (typeof c === 'object' && c !== null) {
					return cloneElement(c as Record<string, unknown>);
				}
				return c;
			});
		}
	}

	return undefined;
}

/**
 * Insert a new paragraph after the specified paragraph.
 * The new paragraph is wrapped in <w:ins>.
 *
 * @param parsed - The parsed XML document
 * @param edit - The insert after edit
 * @param author - The author name for the tracked change
 * @param date - The date of the change
 * @returns The applied edit result with change IDs
 */
export function insertAfter(
	parsed: ParsedXml,
	edit: InsertAfterEdit,
	author: string,
	date: Date
): AppliedEdit {
	const found = findParagraphById(parsed, edit.paraId);

	if (!found) {
		throw new Error(`Paragraph not found: ${edit.paraId}`);
	}

	const { element, parent, index } = found;
	const changeId = generateChangeId();
	const newParaId = mintParaId();
	const newTextId = mintParaId();

	// Inherit paragraph properties from reference if no style specified
	let pPr: unknown[] | undefined;
	if (!edit.style) {
		pPr = extractParagraphProperties(element);
	} else {
		// Create a pPr with the specified style
		pPr = [{ 'w:pStyle': [], ':@': { '@_w:val': edit.style } }];
	}

	// Create the wrapped paragraph
	const wrappedParagraph = wrapParagraphInsert(
		edit.content,
		changeId,
		author,
		date,
		newParaId,
		newTextId,
		pPr
	);

	// Insert after the reference paragraph
	parent.splice(index + 1, 0, wrappedParagraph);

	return {
		edit,
		changeIds: [changeId],
	};
}

/**
 * Insert a new paragraph before the specified paragraph.
 * The new paragraph is wrapped in <w:ins>.
 *
 * @param parsed - The parsed XML document
 * @param edit - The insert before edit
 * @param author - The author name for the tracked change
 * @param date - The date of the change
 * @returns The applied edit result with change IDs
 */
export function insertBefore(
	parsed: ParsedXml,
	edit: InsertBeforeEdit,
	author: string,
	date: Date
): AppliedEdit {
	const found = findParagraphById(parsed, edit.paraId);

	if (!found) {
		throw new Error(`Paragraph not found: ${edit.paraId}`);
	}

	const { element, parent, index } = found;
	const changeId = generateChangeId();
	const newParaId = mintParaId();
	const newTextId = mintParaId();

	// Inherit paragraph properties from reference if no style specified
	let pPr: unknown[] | undefined;
	if (!edit.style) {
		pPr = extractParagraphProperties(element);
	} else {
		// Create a pPr with the specified style
		pPr = [{ 'w:pStyle': [], ':@': { '@_w:val': edit.style } }];
	}

	// Create the wrapped paragraph
	const wrappedParagraph = wrapParagraphInsert(
		edit.content,
		changeId,
		author,
		date,
		newParaId,
		newTextId,
		pPr
	);

	// Insert before the reference paragraph
	parent.splice(index, 0, wrappedParagraph);

	return {
		edit,
		changeIds: [changeId],
	};
}

/**
 * Generic insert function that dispatches to insertAfter or insertBefore
 */
export function insertParagraph(
	parsed: ParsedXml,
	edit: InsertAfterEdit | InsertBeforeEdit,
	author: string,
	date: Date
): AppliedEdit {
	if (edit.type === 'insertAfter') {
		return insertAfter(parsed, edit, author, date);
	}
	return insertBefore(parsed, edit, author, date);
}
