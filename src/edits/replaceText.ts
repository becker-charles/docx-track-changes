import type { ReplaceTextEdit, AppliedEdit } from '../types.js';
import type { ParsedXml } from '../readModel/parser.js';
import { getChildren, getTagName } from '../readModel/parser.js';
import { findParagraphById } from './deleteParagraph.js';
import { createInsElement } from '../tracking/wrapInsert.js';
import { createDelElement, cloneElement } from '../tracking/wrapDelete.js';
import { generateChangeId } from '../tracking/changeIds.js';

/**
 * Represents a text segment with its run context for processing
 */
interface TextSegment {
	text: string;
	runIndex: number;
	runElement: Record<string, unknown>;
	rPr: unknown[] | null; // run properties (formatting)
}

/**
 * Extract the rPr (run properties) from a run element
 */
function extractRPr(runElement: Record<string, unknown>): unknown[] | null {
	const children = getChildren(runElement);
	for (const child of children) {
		if (typeof child !== 'object' || child === null) continue;
		const childObj = child as Record<string, unknown>;
		if (getTagName(childObj) === 'w:rPr') {
			return getChildren(childObj);
		}
	}
	return null;
}

/**
 * Clone rPr children for use in a new run
 */
function cloneRPr(rPr: unknown[] | null): unknown[] {
	if (!rPr || rPr.length === 0) return [];
	return rPr.map(child => {
		if (typeof child !== 'object' || child === null) return child;
		return cloneElement(child as Record<string, unknown>);
	});
}

/**
 * Create a w:t element with text content
 */
function createTextElement(text: string): Record<string, unknown> {
	const needsPreserve = text !== text.trim() || text.includes(' ');
	const element: Record<string, unknown> = {
		'w:t': [{ '#text': text }],
	};
	if (needsPreserve) {
		element[':@'] = { '@_xml:space': 'preserve' };
	}
	return element;
}

/**
 * Create a w:delText element with text content
 */
function createDelTextElement(text: string): Record<string, unknown> {
	const needsPreserve = text !== text.trim() || text.includes(' ');
	const element: Record<string, unknown> = {
		'w:delText': [{ '#text': text }],
	};
	if (needsPreserve) {
		element[':@'] = { '@_xml:space': 'preserve' };
	}
	return element;
}

/**
 * Create a run element with optional formatting
 */
function createRun(text: string, rPr: unknown[] | null): Record<string, unknown> {
	const children: unknown[] = [];

	if (rPr && rPr.length > 0) {
		children.push({ 'w:rPr': cloneRPr(rPr) });
	}

	children.push(createTextElement(text));

	return { 'w:r': children };
}

/**
 * Create a run element for deletion (with w:delText instead of w:t)
 */
function createDelRun(text: string, rPr: unknown[] | null): Record<string, unknown> {
	const children: unknown[] = [];

	if (rPr && rPr.length > 0) {
		children.push({ 'w:rPr': cloneRPr(rPr) });
	}

	children.push(createDelTextElement(text));

	return { 'w:r': children };
}

/**
 * Get all text content from a run element
 */
function getRunText(runElement: Record<string, unknown>): string {
	const children = getChildren(runElement);
	let text = '';

	for (const child of children) {
		if (typeof child !== 'object' || child === null) continue;
		const childObj = child as Record<string, unknown>;
		const tagName = getTagName(childObj);

		if (tagName === 'w:t') {
			const tChildren = getChildren(childObj);
			for (const tChild of tChildren) {
				if (
					typeof tChild === 'object' &&
					tChild !== null &&
					'#text' in (tChild as Record<string, unknown>)
				) {
					text += String((tChild as Record<string, unknown>)['#text']);
				}
			}
		} else if (tagName === 'w:tab') {
			text += '\t';
		} else if (tagName === 'w:br') {
			text += '\n';
		}
	}

	return text;
}

/**
 * Collect all text segments from paragraph runs for searching
 */
function collectTextSegments(paragraphChildren: unknown[]): {
	segments: TextSegment[];
	fullText: string;
} {
	const segments: TextSegment[] = [];
	let fullText = '';

	for (let i = 0; i < paragraphChildren.length; i++) {
		const child = paragraphChildren[i];
		if (typeof child !== 'object' || child === null) continue;

		const childObj = child as Record<string, unknown>;
		const tagName = getTagName(childObj);

		if (tagName === 'w:r') {
			const runText = getRunText(childObj);
			if (runText) {
				const rPr = extractRPr(childObj);
				segments.push({
					text: runText,
					runIndex: i,
					runElement: childObj,
					rPr,
				});
				fullText += runText;
			}
		} else if (tagName === 'w:hyperlink' || tagName === 'w:ins') {
			// Handle runs inside hyperlinks and existing insertions
			const containerChildren = getChildren(childObj);
			for (const containerChild of containerChildren) {
				if (typeof containerChild !== 'object' || containerChild === null) continue;
				const containerChildObj = containerChild as Record<string, unknown>;
				if (getTagName(containerChildObj) === 'w:r') {
					const runText = getRunText(containerChildObj);
					if (runText) {
						const rPr = extractRPr(containerChildObj);
						segments.push({
							text: runText,
							runIndex: i,
							runElement: containerChildObj,
							rPr,
						});
						fullText += runText;
					}
				}
			}
		}
		// Skip w:del content (already deleted text)
	}

	return { segments, fullText };
}

/**
 * Find match positions in the concatenated text
 */
function findMatches(
	fullText: string,
	searchText: string,
	occurrence: number,
	all: boolean
): Array<{ start: number; end: number }> {
	const matches: Array<{ start: number; end: number }> = [];
	let pos = 0;

	while (pos <= fullText.length - searchText.length) {
		const idx = fullText.indexOf(searchText, pos);
		if (idx === -1) break;

		matches.push({ start: idx, end: idx + searchText.length });
		pos = idx + 1;
	}

	if (matches.length === 0) {
		return [];
	}

	if (all) {
		return matches;
	}

	// Return specific occurrence (1-indexed)
	const occIdx = occurrence - 1;
	if (occIdx < 0 || occIdx >= matches.length) {
		return [];
	}

	const match = matches[occIdx];
	return match ? [match] : [];
}

/**
 * Map a text position to its run and offset within that run
 */
function mapPositionToRun(
	segments: TextSegment[],
	position: number
): { segmentIndex: number; offset: number } | null {
	let currentPos = 0;

	for (let i = 0; i < segments.length; i++) {
		const seg = segments[i];
		if (!seg) continue;
		const segEnd = currentPos + seg.text.length;
		if (position < segEnd) {
			return { segmentIndex: i, offset: position - currentPos };
		}
		currentPos = segEnd;
	}

	return null;
}

/**
 * Replace text within a paragraph with tracked changes.
 *
 * This handles the complex case of text that may span multiple runs,
 * splitting runs at match boundaries and preserving formatting.
 */
export function replaceText(
	parsed: ParsedXml,
	edit: ReplaceTextEdit,
	author: string,
	date: Date
): AppliedEdit {
	const found = findParagraphById(parsed, edit.paraId);

	if (!found) {
		throw new Error(`Paragraph not found: ${edit.paraId}`);
	}

	const { element } = found;
	const children = getChildren(element);

	// Collect text segments from runs
	const { segments, fullText } = collectTextSegments(children);

	if (segments.length === 0) {
		throw new Error(`Paragraph has no text content: ${edit.paraId}`);
	}

	// Find matches
	const occurrence = edit.occurrence ?? 1;
	const all = edit.all ?? false;
	const matches = findMatches(fullText, edit.find, occurrence, all);

	if (matches.length === 0) {
		throw new Error(
			`Text "${edit.find}" not found in paragraph ${edit.paraId}` +
				(occurrence > 1 ? ` (occurrence ${occurrence})` : '')
		);
	}

	const changeIds: string[] = [];

	// Process matches in reverse order so indices remain valid
	for (let matchIdx = matches.length - 1; matchIdx >= 0; matchIdx--) {
		const match = matches[matchIdx];
		if (!match) continue;

		const delChangeId = generateChangeId();
		const insChangeId = generateChangeId();
		changeIds.unshift(delChangeId, insChangeId);

		// Map match boundaries to runs
		const startMapping = mapPositionToRun(segments, match.start);
		const endMapping = mapPositionToRun(segments, match.end - 1); // -1 because end is exclusive

		if (!startMapping || !endMapping) {
			throw new Error('Failed to map text position to run');
		}

		// Build the replacement content
		processMatch(
			children,
			segments,
			startMapping,
			endMapping,
			edit.replace,
			delChangeId,
			insChangeId,
			author,
			date
		);
	}

	return {
		edit,
		changeIds,
	};
}

/**
 * Process a single match and modify the paragraph children
 */
function processMatch(
	children: unknown[],
	segments: TextSegment[],
	startMapping: { segmentIndex: number; offset: number },
	endMapping: { segmentIndex: number; offset: number },
	replacementText: string,
	delChangeId: string,
	insChangeId: string,
	author: string,
	date: Date
): void {
	const startSeg = segments[startMapping.segmentIndex];

	if (!startSeg) {
		throw new Error('Start segment not found');
	}

	// Simple case: match is within a single run
	if (startMapping.segmentIndex === endMapping.segmentIndex) {
		replaceInSingleRun(
			children,
			segments,
			startSeg,
			startMapping.offset,
			endMapping.offset + 1, // +1 because endMapping.offset is inclusive
			replacementText,
			delChangeId,
			insChangeId,
			author,
			date
		);
	} else {
		// Complex case: match spans multiple runs
		replaceAcrossRuns(
			children,
			segments,
			startMapping,
			endMapping,
			replacementText,
			delChangeId,
			insChangeId,
			author,
			date
		);
	}
}

/**
 * Replace text within a single run
 */
function replaceInSingleRun(
	children: unknown[],
	segments: TextSegment[],
	segment: TextSegment,
	startOffset: number,
	endOffset: number,
	newText: string,
	delChangeId: string,
	insChangeId: string,
	author: string,
	date: Date
): void {
	const runText = segment.text;
	const rPr = segment.rPr;

	const beforeText = runText.slice(0, startOffset);
	const matchedText = runText.slice(startOffset, endOffset);
	const afterText = runText.slice(endOffset);

	// Build new content to replace the run
	const newContent: unknown[] = [];

	// Text before the match (if any)
	if (beforeText) {
		newContent.push(createRun(beforeText, rPr));
	}

	// Deletion wrapper
	const delElement = createDelElement(delChangeId, author, date);
	(delElement['w:del'] as unknown[]).push(createDelRun(matchedText, rPr));
	newContent.push(delElement);

	// Insertion wrapper (if replacement text is not empty)
	if (newText) {
		const insElement = createInsElement(insChangeId, author, date);
		(insElement['w:ins'] as unknown[]).push(createRun(newText, rPr));
		newContent.push(insElement);
	}

	// Text after the match (if any)
	if (afterText) {
		newContent.push(createRun(afterText, rPr));
	}

	// Replace the run in children
	const runIndex = segment.runIndex;
	children.splice(runIndex, 1, ...newContent);

	// Update segment indices for subsequent operations
	// (segments after this one need their runIndex adjusted)
	const indexDelta = newContent.length - 1;
	for (const seg of segments) {
		if (seg.runIndex > runIndex) {
			seg.runIndex += indexDelta;
		}
	}
}

/**
 * Replace text that spans multiple runs
 */
function replaceAcrossRuns(
	children: unknown[],
	segments: TextSegment[],
	startMapping: { segmentIndex: number; offset: number },
	endMapping: { segmentIndex: number; offset: number },
	newText: string,
	delChangeId: string,
	insChangeId: string,
	author: string,
	date: Date
): void {
	// Collect all segments involved in this match
	const involvedSegments = segments.slice(
		startMapping.segmentIndex,
		endMapping.segmentIndex + 1
	);

	if (involvedSegments.length === 0) {
		throw new Error('No segments involved in match');
	}

	const firstSeg = involvedSegments[0];
	const lastSeg = involvedSegments[involvedSegments.length - 1];

	if (!firstSeg || !lastSeg) {
		throw new Error('Could not find first or last segment');
	}

	// Use the formatting from the first segment for the replacement
	const rPr = firstSeg.rPr;

	// Build the deletion content from all involved segments
	const delElement = createDelElement(delChangeId, author, date);

	for (let i = 0; i < involvedSegments.length; i++) {
		const seg = involvedSegments[i];
		if (!seg) continue;

		let textToDelete: string;

		if (i === 0) {
			// First segment: from startOffset to end
			textToDelete = seg.text.slice(startMapping.offset);
		} else if (i === involvedSegments.length - 1) {
			// Last segment: from start to endOffset (inclusive)
			textToDelete = seg.text.slice(0, endMapping.offset + 1);
		} else {
			// Middle segments: entire text
			textToDelete = seg.text;
		}

		if (textToDelete) {
			(delElement['w:del'] as unknown[]).push(createDelRun(textToDelete, seg.rPr));
		}
	}

	// Build the new content
	const newContent: unknown[] = [];

	// Text before the match in the first run (if any)
	const beforeText = firstSeg.text.slice(0, startMapping.offset);
	if (beforeText) {
		newContent.push(createRun(beforeText, firstSeg.rPr));
	}

	// The deletion
	newContent.push(delElement);

	// The insertion (if replacement text is not empty)
	if (newText) {
		const insElement = createInsElement(insChangeId, author, date);
		(insElement['w:ins'] as unknown[]).push(createRun(newText, rPr));
		newContent.push(insElement);
	}

	// Text after the match in the last run (if any)
	const afterText = lastSeg.text.slice(endMapping.offset + 1);
	if (afterText) {
		newContent.push(createRun(afterText, lastSeg.rPr));
	}

	// Find the range of indices to replace in children
	const startRunIndex = firstSeg.runIndex;
	const endRunIndex = lastSeg.runIndex;
	const removeCount = endRunIndex - startRunIndex + 1;

	// Replace the runs in children
	children.splice(startRunIndex, removeCount, ...newContent);

	// Update segment indices for subsequent operations
	const indexDelta = newContent.length - removeCount;
	for (const seg of segments) {
		if (seg.runIndex > endRunIndex) {
			seg.runIndex += indexDelta;
		}
	}
}
