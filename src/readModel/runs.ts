import type { Run } from '../types.js';
import { getChildren, getTagName } from './parser.js';

/**
 * Extract runs (formatted text segments) from a paragraph element
 */
export function extractRuns(paragraphElement: Record<string, unknown>): Run[] {
	const runs: Run[] = [];
	const children = getChildren(paragraphElement);

	for (const child of children) {
		if (typeof child !== 'object' || child === null) continue;

		const childObj = child as Record<string, unknown>;
		const tagName = getTagName(childObj);

		if (tagName === 'w:r') {
			const run = extractSingleRun(childObj);
			if (run) {
				runs.push(run);
			}
		} else if (tagName === 'w:hyperlink') {
			// Extract runs from inside hyperlinks
			const hyperlinkRuns = extractRunsFromContainer(childObj);
			runs.push(...hyperlinkRuns);
		} else if (tagName === 'w:ins') {
			// Extract runs from tracked insertions (treat as current content)
			const insRuns = extractRunsFromContainer(childObj);
			runs.push(...insRuns);
		}
		// Note: w:del content is NOT included (it's deleted text)
	}

	return runs;
}

/**
 * Extract runs from a container element (hyperlink, ins, etc.)
 */
function extractRunsFromContainer(container: Record<string, unknown>): Run[] {
	const runs: Run[] = [];
	const children = getChildren(container);

	for (const child of children) {
		if (typeof child !== 'object' || child === null) continue;

		const childObj = child as Record<string, unknown>;
		const tagName = getTagName(childObj);

		if (tagName === 'w:r') {
			const run = extractSingleRun(childObj);
			if (run) {
				runs.push(run);
			}
		}
	}

	return runs;
}

/**
 * Extract a single run's text and formatting
 */
function extractSingleRun(runElement: Record<string, unknown>): Run | null {
	const children = getChildren(runElement);
	let text = '';
	let bold = false;
	let italic = false;
	let underline = false;
	let strike = false;
	let style: string | undefined;

	for (const child of children) {
		if (typeof child !== 'object' || child === null) continue;

		const childObj = child as Record<string, unknown>;
		const tagName = getTagName(childObj);

		if (tagName === 'w:t') {
			// Text content
			text += extractTextFromT(childObj);
		} else if (tagName === 'w:rPr') {
			// Run properties (formatting)
			const props = extractRunProperties(childObj);
			bold = props.bold;
			italic = props.italic;
			underline = props.underline;
			strike = props.strike;
			style = props.style;
		} else if (tagName === 'w:tab') {
			text += '\t';
		} else if (tagName === 'w:br') {
			text += '\n';
		}
	}

	// Skip empty runs
	if (!text) {
		return null;
	}

	const run: Run = { text };
	if (bold) run.bold = true;
	if (italic) run.italic = true;
	if (underline) run.underline = true;
	if (strike) run.strike = true;
	if (style) run.style = style;

	return run;
}

/**
 * Extract text from a w:t element
 */
function extractTextFromT(tElement: Record<string, unknown>): string {
	const children = getChildren(tElement);
	let text = '';

	for (const child of children) {
		if (typeof child === 'object' && child !== null && '#text' in child) {
			text += String((child as Record<string, unknown>)['#text']);
		}
	}

	return text;
}

/**
 * Extract formatting properties from w:rPr element
 */
function extractRunProperties(rPrElement: Record<string, unknown>): {
	bold: boolean;
	italic: boolean;
	underline: boolean;
	strike: boolean;
	style?: string;
} {
	const result = {
		bold: false,
		italic: false,
		underline: false,
		strike: false,
		style: undefined as string | undefined,
	};

	const children = getChildren(rPrElement);

	for (const child of children) {
		if (typeof child !== 'object' || child === null) continue;

		const childObj = child as Record<string, unknown>;
		const tagName = getTagName(childObj);

		switch (tagName) {
			case 'w:b':
				result.bold = !hasValFalse(childObj);
				break;
			case 'w:i':
				result.italic = !hasValFalse(childObj);
				break;
			case 'w:u':
				result.underline = !hasValNone(childObj);
				break;
			case 'w:strike':
				result.strike = !hasValFalse(childObj);
				break;
			case 'w:rStyle': {
				const attrs = childObj[':@'] as Record<string, string> | undefined;
				if (attrs?.['@_w:val']) {
					result.style = attrs['@_w:val'];
				}
				break;
			}
		}
	}

	return result;
}

/**
 * Check if an element has w:val="false" or w:val="0"
 */
function hasValFalse(element: Record<string, unknown>): boolean {
	const attrs = element[':@'] as Record<string, string> | undefined;
	if (!attrs) return false;
	const val = attrs['@_w:val'];
	return val === 'false' || val === '0';
}

/**
 * Check if an element has w:val="none"
 */
function hasValNone(element: Record<string, unknown>): boolean {
	const attrs = element[':@'] as Record<string, string> | undefined;
	if (!attrs) return false;
	return attrs['@_w:val'] === 'none';
}
