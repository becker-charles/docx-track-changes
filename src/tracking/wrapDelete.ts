import { formatDateForWord } from './wrapInsert.js';
import { getTagName, getChildren } from '../readModel/parser.js';

/**
 * Create a w:del element wrapper in preserveOrder format
 */
export function createDelElement(
	changeId: string,
	author: string,
	date: Date
): Record<string, unknown> {
	return {
		'w:del': [],
		':@': {
			'@_w:id': changeId,
			'@_w:author': author,
			'@_w:date': formatDateForWord(date),
		},
	};
}

/**
 * Convert w:t elements to w:delText within a run element (deep clone)
 * This is required for tracked deletions - Word uses delText instead of t
 */
export function convertToDelText(element: Record<string, unknown>): Record<string, unknown> {
	const tagName = getTagName(element);
	if (!tagName) return { ...element };

	const children = getChildren(element);
	const convertedChildren: unknown[] = [];

	for (const child of children) {
		if (typeof child !== 'object' || child === null) {
			convertedChildren.push(child);
			continue;
		}

		const childObj = child as Record<string, unknown>;
		const childTagName = getTagName(childObj);

		if (childTagName === 'w:t') {
			// Convert w:t to w:delText
			const newChild: Record<string, unknown> = {
				'w:delText': childObj['w:t'],
			};
			// Preserve attributes (like xml:space)
			if (childObj[':@']) {
				newChild[':@'] = childObj[':@'];
			}
			convertedChildren.push(newChild);
		} else if (childTagName === 'w:r') {
			// Recursively convert runs
			convertedChildren.push(convertToDelText(childObj));
		} else {
			// Keep other elements as-is (clone to avoid mutation)
			convertedChildren.push(cloneElement(childObj));
		}
	}

	const result: Record<string, unknown> = {
		[tagName]: convertedChildren,
	};

	// Preserve attributes
	if (element[':@']) {
		result[':@'] = element[':@'];
	}

	return result;
}

/**
 * Deep clone an element (for preserveOrder format)
 */
export function cloneElement(element: Record<string, unknown>): Record<string, unknown> {
	const tagName = getTagName(element);
	if (!tagName) {
		// Handle text nodes and other special cases
		const result: Record<string, unknown> = {};
		for (const key of Object.keys(element)) {
			result[key] = element[key];
		}
		return result;
	}

	const children = getChildren(element);
	const clonedChildren: unknown[] = [];

	for (const child of children) {
		if (typeof child !== 'object' || child === null) {
			clonedChildren.push(child);
		} else if ('#text' in (child as Record<string, unknown>)) {
			clonedChildren.push({ '#text': (child as Record<string, unknown>)['#text'] });
		} else {
			clonedChildren.push(cloneElement(child as Record<string, unknown>));
		}
	}

	const result: Record<string, unknown> = {
		[tagName]: clonedChildren,
	};

	if (element[':@']) {
		result[':@'] = { ...(element[':@'] as Record<string, unknown>) };
	}

	return result;
}

/**
 * Wrap runs in a <w:del> element for tracked deletions.
 * Converts w:t to w:delText as required by Word.
 *
 * @param runs - Array of run elements to wrap (from getChildren of a paragraph)
 * @param changeId - The unique change ID for this deletion
 * @param author - The author name for the tracked change
 * @param date - The date of the change
 * @returns A w:del element in preserveOrder format
 */
export function wrapDelete(
	runs: unknown[],
	changeId: string,
	author: string,
	date: Date
): Record<string, unknown> {
	const delElement = createDelElement(changeId, author, date);

	for (const run of runs) {
		if (typeof run !== 'object' || run === null) continue;

		const runObj = run as Record<string, unknown>;
		const tagName = getTagName(runObj);

		if (tagName === 'w:r') {
			// Convert w:t to w:delText within the run
			const convertedRun = convertToDelText(runObj);
			(delElement['w:del'] as unknown[]).push(convertedRun);
		} else if (tagName === 'w:pPr') {
			// Skip paragraph properties - they stay outside the del
			continue;
		} else {
			// For other elements (bookmarks, etc.), clone them
			(delElement['w:del'] as unknown[]).push(cloneElement(runObj));
		}
	}

	return delElement;
}

/**
 * Wrap paragraph content for deletion (keeps pPr outside, wraps runs in del)
 * Returns the new children array for the paragraph
 */
export function wrapParagraphContentForDelete(
	paragraphChildren: unknown[],
	changeId: string,
	author: string,
	date: Date
): unknown[] {
	const newChildren: unknown[] = [];
	const runsToDelete: unknown[] = [];

	for (const child of paragraphChildren) {
		if (typeof child !== 'object' || child === null) continue;

		const childObj = child as Record<string, unknown>;
		const tagName = getTagName(childObj);

		if (tagName === 'w:pPr') {
			// Keep paragraph properties at the start
			newChildren.push(cloneElement(childObj));
		} else {
			// Collect runs and other content for deletion
			runsToDelete.push(childObj);
		}
	}

	// Wrap all non-pPr content in a single del element
	if (runsToDelete.length > 0) {
		const delElement = wrapDelete(runsToDelete, changeId, author, date);
		newChildren.push(delElement);
	}

	return newChildren;
}
