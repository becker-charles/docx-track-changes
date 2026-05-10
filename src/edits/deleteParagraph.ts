import type { DeleteParagraphEdit, DeleteRangeEdit, AppliedEdit } from '../types.js';
import type { ParsedXml } from '../readModel/parser.js';
import { findElements, getAttributes, getChildren } from '../readModel/parser.js';
import { wrapParagraphContentForDelete } from '../tracking/wrapDelete.js';
import { generateChangeId } from '../tracking/changeIds.js';

/**
 * Find a paragraph element by its full ID (e.g., 'body:ABC123')
 * Returns the element, its parent array, and index in that array
 */
export function findParagraphById(
	parsed: ParsedXml,
	fullId: string
): { element: Record<string, unknown>; parent: unknown[]; index: number } | null {
	// Extract the raw paraId from the full ID (e.g., 'body:ABC123' -> 'ABC123')
	const colonIndex = fullId.indexOf(':');
	const rawId = colonIndex >= 0 ? fullId.slice(colonIndex + 1) : fullId;

	const paragraphs = findElements(parsed, 'w:p');

	for (const { element, parent, index } of paragraphs) {
		const attrs = getAttributes(element);
		if (attrs['w14:paraId'] === rawId) {
			return { element, parent, index };
		}
	}

	return null;
}

/**
 * Delete a single paragraph with tracked changes.
 * Wraps the paragraph content in <w:del> elements.
 *
 * @param parsed - The parsed XML document
 * @param edit - The delete paragraph edit
 * @param author - The author name for the tracked change
 * @param date - The date of the change
 * @returns The applied edit result with change IDs
 */
export function deleteParagraph(
	parsed: ParsedXml,
	edit: DeleteParagraphEdit,
	author: string,
	date: Date
): AppliedEdit {
	const found = findParagraphById(parsed, edit.paraId);

	if (!found) {
		throw new Error(`Paragraph not found: ${edit.paraId}`);
	}

	const { element } = found;
	const changeId = generateChangeId();

	// Get the paragraph's children
	const children = getChildren(element);

	// Wrap the content in w:del (keeping pPr outside)
	const newChildren = wrapParagraphContentForDelete(children, changeId, author, date);

	// Replace the paragraph's children
	element['w:p'] = newChildren;

	return {
		edit,
		changeIds: [changeId],
	};
}

/**
 * Delete a range of consecutive paragraphs with tracked changes.
 * All paragraphs from 'from' to 'to' (inclusive) are marked as deleted.
 *
 * @param parsed - The parsed XML document
 * @param edit - The delete range edit
 * @param author - The author name for the tracked change
 * @param date - The date of the change
 * @returns The applied edit result with change IDs
 */
export function deleteRange(
	parsed: ParsedXml,
	edit: DeleteRangeEdit,
	author: string,
	date: Date
): AppliedEdit {
	// Find both endpoints
	const fromResult = findParagraphById(parsed, edit.from);
	const toResult = findParagraphById(parsed, edit.to);

	if (!fromResult) {
		throw new Error(`Range start paragraph not found: ${edit.from}`);
	}
	if (!toResult) {
		throw new Error(`Range end paragraph not found: ${edit.to}`);
	}

	// Verify they're in the same parent
	if (fromResult.parent !== toResult.parent) {
		throw new Error('Range start and end must be in the same document section');
	}

	// Verify the range is valid (from comes before to)
	if (fromResult.index > toResult.index) {
		throw new Error('Range start must come before range end');
	}

	const changeIds: string[] = [];

	// Delete each paragraph in the range
	for (let i = fromResult.index; i <= toResult.index; i++) {
		const paragraph = fromResult.parent[i] as Record<string, unknown>;

		// Skip if not a paragraph
		if (!paragraph || !paragraph['w:p']) continue;

		const changeId = generateChangeId();
		const children = getChildren(paragraph);

		// Wrap the content in w:del
		const newChildren = wrapParagraphContentForDelete(children, changeId, author, date);
		paragraph['w:p'] = newChildren;

		changeIds.push(changeId);
	}

	return {
		edit,
		changeIds,
	};
}
