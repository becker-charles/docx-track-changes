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
 *
 * @param run - The content run (string or object with formatting)
 * @param inheritedRPr - Optional inherited run properties (w:rPr children) to apply to plain strings
 */
export function createRunElement(
	run: ContentRun,
	inheritedRPr?: unknown[]
): Record<string, unknown> {
	const children: unknown[] = [];

	// Add run properties based on whether this is a plain string or formatted object
	if (typeof run === 'string') {
		// Plain string: use inherited run properties if available
		if (inheritedRPr && inheritedRPr.length > 0) {
			children.push({ 'w:rPr': cloneRPrChildren(inheritedRPr) });
		}
	} else {
		// Formatted object: explicit formatting overrides inherited
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
 * Deep clone an array of run property children
 */
function cloneRPrChildren(rPrChildren: unknown[]): unknown[] {
	return JSON.parse(JSON.stringify(rPrChildren));
}

/**
 * Create a new paragraph element with content
 *
 * @param content - Array of content runs
 * @param paraId - Paragraph ID
 * @param textId - Text ID
 * @param pPr - Optional paragraph properties
 * @param inheritedRPr - Optional inherited run properties for plain strings
 */
export function createParagraphElement(
	content: ContentRun[],
	paraId: string,
	textId: string,
	pPr?: unknown[],
	inheritedRPr?: unknown[]
): Record<string, unknown> {
	const children: unknown[] = [];

	// Add paragraph properties if provided
	if (pPr && pPr.length > 0) {
		children.push({ 'w:pPr': pPr });
	}

	// Add runs
	for (const run of content) {
		children.push(createRunElement(run, inheritedRPr));
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
 * @param inheritedRPr - Optional inherited run properties for plain strings
 * @returns A w:ins element in preserveOrder format
 */
export function wrapInsert(
	content: ContentRun[] | Record<string, unknown>,
	changeId: string,
	author: string,
	date: Date,
	inheritedRPr?: unknown[]
): Record<string, unknown> {
	const insElement = createInsElement(changeId, author, date);

	if (Array.isArray(content)) {
		// Content is an array of ContentRuns - convert to run elements
		const runs = content.map(run => createRunElement(run, inheritedRPr));
		(insElement['w:ins'] as unknown[]).push(...runs);
	} else {
		// Content is already a constructed element (e.g., a paragraph)
		(insElement['w:ins'] as unknown[]).push(content);
	}

	return insElement;
}

/**
 * Wrap a paragraph element in a <w:ins> for paragraph insertion tracking
 *
 * @param content - Array of content runs
 * @param changeId - The unique change ID for this insertion
 * @param author - The author name for the tracked change
 * @param date - The date of the change
 * @param paraId - Paragraph ID
 * @param textId - Text ID
 * @param pPr - Optional paragraph properties
 * @param inheritedRPr - Optional inherited run properties for plain strings
 */
export function wrapParagraphInsert(
	content: ContentRun[],
	changeId: string,
	author: string,
	date: Date,
	paraId: string,
	textId: string,
	pPr?: unknown[],
	inheritedRPr?: unknown[]
): Record<string, unknown> {
	const paragraph = createParagraphElement(content, paraId, textId, pPr, inheritedRPr);
	return wrapInsert(paragraph, changeId, author, date);
}

/**
 * Extract run properties (w:rPr children) from the first run in a paragraph.
 * Used to inherit formatting when replacing/inserting with plain strings.
 *
 * @param paragraphChildren - The children array from a w:p element
 * @returns The w:rPr children array from the first run, or undefined if none
 */
export function extractFirstRunProperties(paragraphChildren: unknown[]): unknown[] | undefined {
	for (const child of paragraphChildren) {
		if (typeof child !== 'object' || child === null) continue;

		const childObj = child as Record<string, unknown>;

		// Check for w:r (direct run)
		if ('w:r' in childObj) {
			return extractRPrFromRun(childObj);
		}

		// Check for w:ins (tracked insertion containing runs)
		if ('w:ins' in childObj) {
			const insChildren = childObj['w:ins'] as unknown[];
			for (const insChild of insChildren) {
				if (typeof insChild !== 'object' || insChild === null) continue;
				const insChildObj = insChild as Record<string, unknown>;
				if ('w:r' in insChildObj) {
					return extractRPrFromRun(insChildObj);
				}
			}
		}

		// Check for w:hyperlink (containing runs)
		if ('w:hyperlink' in childObj) {
			const hlChildren = childObj['w:hyperlink'] as unknown[];
			for (const hlChild of hlChildren) {
				if (typeof hlChild !== 'object' || hlChild === null) continue;
				const hlChildObj = hlChild as Record<string, unknown>;
				if ('w:r' in hlChildObj) {
					return extractRPrFromRun(hlChildObj);
				}
			}
		}
	}

	return undefined;
}

/**
 * Extract w:rPr children from a run element
 */
function extractRPrFromRun(runElement: Record<string, unknown>): unknown[] | undefined {
	const runChildren = runElement['w:r'] as unknown[];
	if (!Array.isArray(runChildren)) return undefined;

	for (const child of runChildren) {
		if (typeof child !== 'object' || child === null) continue;
		const childObj = child as Record<string, unknown>;
		if ('w:rPr' in childObj) {
			const rPrChildren = childObj['w:rPr'] as unknown[];
			if (Array.isArray(rPrChildren) && rPrChildren.length > 0) {
				return cloneRPrChildren(rPrChildren);
			}
		}
	}

	return undefined;
}
