import type { ContentRun } from '../types.js';

/**
 * Format a date as ISO 8601 for Word XML (e.g., "2026-05-10T12:00:00Z")
 */
export function formatDateForWord(date: Date): string {
	return date.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

/**
 * Create a w:ins element wrapper in preserveOrder format
 */
export function createInsElement(
	changeId: string,
	author: string,
	date: Date
): Record<string, unknown> {
	return {
		'w:ins': [],
		':@': {
			'@_w:id': changeId,
			'@_w:author': author,
			'@_w:date': formatDateForWord(date),
		},
	};
}

/**
 * Create a w:t element with text content
 */
function createTextElement(text: string): Record<string, unknown> {
	// If text has leading/trailing whitespace, need xml:space="preserve"
	const needsPreserve = text !== text.trim();
	const element: Record<string, unknown> = {
		'w:t': [{ '#text': text }],
	};
	if (needsPreserve) {
		element[':@'] = { '@_xml:space': 'preserve' };
	}
	return element;
}

/**
 * Create a w:r (run) element with optional formatting
 */
export function createRunElement(run: ContentRun): Record<string, unknown> {
	const children: unknown[] = [];

	// Add run properties if there's formatting
	if (typeof run !== 'string') {
		const rPrChildren: unknown[] = [];

		if (run.bold) {
			rPrChildren.push({ 'w:b': [] });
		}
		if (run.italic) {
			rPrChildren.push({ 'w:i': [] });
		}
		if (run.underline) {
			rPrChildren.push({ 'w:u': [], ':@': { '@_w:val': 'single' } });
		}

		if (rPrChildren.length > 0) {
			children.push({ 'w:rPr': rPrChildren });
		}
	}

	// Add text element
	const text = typeof run === 'string' ? run : run.text;
	children.push(createTextElement(text));

	return { 'w:r': children };
}

/**
 * Create a new paragraph element with content
 */
export function createParagraphElement(
	content: ContentRun[],
	paraId: string,
	textId: string,
	pPr?: unknown[]
): Record<string, unknown> {
	const children: unknown[] = [];

	// Add paragraph properties if provided
	if (pPr && pPr.length > 0) {
		children.push({ 'w:pPr': pPr });
	}

	// Add runs
	for (const run of content) {
		children.push(createRunElement(run));
	}

	return {
		'w:p': children,
		':@': {
			'@_w14:paraId': paraId,
			'@_w14:textId': textId,
		},
	};
}

/**
 * Wrap content in a <w:ins> element for tracked insertions
 *
 * @param content - Array of content runs to wrap, or an already-constructed element
 * @param changeId - The unique change ID for this insertion
 * @param author - The author name for the tracked change
 * @param date - The date of the change
 * @returns A w:ins element in preserveOrder format
 */
export function wrapInsert(
	content: ContentRun[] | Record<string, unknown>,
	changeId: string,
	author: string,
	date: Date
): Record<string, unknown> {
	const insElement = createInsElement(changeId, author, date);

	if (Array.isArray(content)) {
		// Content is an array of ContentRuns - convert to run elements
		const runs = content.map(run => createRunElement(run));
		(insElement['w:ins'] as unknown[]).push(...runs);
	} else {
		// Content is already a constructed element (e.g., a paragraph)
		(insElement['w:ins'] as unknown[]).push(content);
	}

	return insElement;
}

/**
 * Wrap a paragraph element in a <w:ins> for paragraph insertion tracking
 */
export function wrapParagraphInsert(
	content: ContentRun[],
	changeId: string,
	author: string,
	date: Date,
	paraId: string,
	textId: string,
	pPr?: unknown[]
): Record<string, unknown> {
	const paragraph = createParagraphElement(content, paraId, textId, pPr);
	return wrapInsert(paragraph, changeId, author, date);
}
