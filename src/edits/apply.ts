import type { Edit, EditOptions, AppliedEdit, FailedEdit } from '../types.js';
import type { ParsedXml } from '../readModel/parser.js';
import { deleteParagraph, deleteRange } from './deleteParagraph.js';
import { insertAfter, insertBefore } from './insertParagraph.js';
import { replaceParagraph } from './replaceParagraph.js';
import { replaceText } from './replaceText.js';

/**
 * Apply a single edit to the document.
 *
 * @param parsed - The parsed XML document
 * @param edit - The edit to apply
 * @param author - The author name for the tracked change
 * @param date - The date of the change
 * @returns The applied edit result
 * @throws Error if the edit cannot be applied
 */
export function applySingleEdit(
	parsed: ParsedXml,
	edit: Edit,
	author: string,
	date: Date
): AppliedEdit {
	switch (edit.type) {
		case 'deleteParagraph':
			return deleteParagraph(parsed, edit, author, date);

		case 'deleteRange':
			return deleteRange(parsed, edit, author, date);

		case 'insertAfter':
			return insertAfter(parsed, edit, author, date);

		case 'insertBefore':
			return insertBefore(parsed, edit, author, date);

		case 'replaceParagraph':
			return replaceParagraph(parsed, edit, author, date);

		case 'replaceText':
			return replaceText(parsed, edit, author, date);

		default:
			throw new Error(`Unknown edit type: ${(edit as Edit).type}`);
	}
}

/**
 * Apply a list of edits to the document.
 *
 * @param parsed - The parsed XML document
 * @param edits - The edits to apply
 * @param options - Edit options (author, date, continueOnError)
 * @returns Object containing arrays of applied and failed edits
 */
export function applyEdits(
	parsed: ParsedXml,
	edits: Edit[],
	options: EditOptions
): { applied: AppliedEdit[]; failed: FailedEdit[] } {
	const applied: AppliedEdit[] = [];
	const failed: FailedEdit[] = [];

	const author = options.author;
	const date = options.date ?? new Date();
	const continueOnError = options.continueOnError !== false; // Default true

	for (const edit of edits) {
		try {
			const result = applySingleEdit(parsed, edit, author, date);
			applied.push(result);
		} catch (error) {
			const reason = error instanceof Error ? error.message : String(error);
			failed.push({ edit, reason });

			if (!continueOnError) {
				// Stop processing further edits
				break;
			}
		}
	}

	return { applied, failed };
}
