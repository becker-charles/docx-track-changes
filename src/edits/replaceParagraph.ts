import type { ReplaceParagraphEdit, AppliedEdit } from '../types.js';
import type { ParsedXml } from '../readModel/parser.js';
import { getChildren } from '../readModel/parser.js';
import { findParagraphById } from './deleteParagraph.js';
import { wrapParagraphContentForDelete, cloneElement } from '../tracking/wrapDelete.js';
import { wrapInsert } from '../tracking/wrapInsert.js';
import { generateChangeId } from '../tracking/changeIds.js';

/**
 * Replace the content of a paragraph with tracked changes.
 * The old content is wrapped in <w:del>, new content in <w:ins>.
 *
 * @param parsed - The parsed XML document
 * @param edit - The replace paragraph edit
 * @param author - The author name for the tracked change
 * @param date - The date of the change
 * @returns The applied edit result with change IDs
 */
export function replaceParagraph(
	parsed: ParsedXml,
	edit: ReplaceParagraphEdit,
	author: string,
	date: Date
): AppliedEdit {
	const found = findParagraphById(parsed, edit.paraId);

	if (!found) {
		throw new Error(`Paragraph not found: ${edit.paraId}`);
	}

	const { element } = found;
	const preserveStyle = edit.preserveStyle !== false; // Default true

	const changeIds: string[] = [];
	const newChildren: unknown[] = [];

	// Get the paragraph's children
	const children = getChildren(element);

	// If preserving style, keep the w:pPr
	if (preserveStyle) {
		for (const child of children) {
			if (typeof child !== 'object' || child === null) continue;
			const childObj = child as Record<string, unknown>;
			if ('w:pPr' in childObj) {
				newChildren.push(cloneElement(childObj));
				break;
			}
		}
	}

	// Collect runs to delete (everything except pPr)
	const runsToDelete: unknown[] = [];
	for (const child of children) {
		if (typeof child !== 'object' || child === null) continue;
		const childObj = child as Record<string, unknown>;
		if (!('w:pPr' in childObj)) {
			runsToDelete.push(childObj);
		}
	}

	// Wrap old content in w:del (if there's content to delete)
	if (runsToDelete.length > 0) {
		const delChangeId = generateChangeId();
		const delChildren = wrapParagraphContentForDelete(runsToDelete, delChangeId, author, date);
		// The wrapParagraphContentForDelete returns [w:del element], extract it
		newChildren.push(...delChildren);
		changeIds.push(delChangeId);
	}

	// Wrap new content in w:ins
	const insChangeId = generateChangeId();
	const insElement = wrapInsert(edit.content, insChangeId, author, date);
	newChildren.push(insElement);
	changeIds.push(insChangeId);

	// Replace the paragraph's children
	element['w:p'] = newChildren;

	return {
		edit,
		changeIds,
	};
}
