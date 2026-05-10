import { findElements, getAttributes } from '../readModel/parser.js';
import type { ParsedXml } from '../readModel/parser.js';

/**
 * Track the next available change ID for a document
 */
let nextId = 1;

/**
 * Generate a unique change ID for w:ins/w:del elements
 */
export function generateChangeId(): string {
	return String(nextId++);
}

/**
 * Get the next change ID without incrementing
 */
export function getNextChangeId(): string {
	return String(nextId);
}

/**
 * Reset the change ID counter (useful for testing)
 */
export function resetChangeIds(startFrom = 1): void {
	nextId = startFrom;
}

/**
 * Find the highest existing change ID in a document and set the counter accordingly.
 * Call this when loading a document that may already have tracked changes.
 */
export function initializeChangeIds(existingIds: string[]): void {
	let maxId = 0;

	for (const id of existingIds) {
		const numId = parseInt(id, 10);
		if (!isNaN(numId) && numId > maxId) {
			maxId = numId;
		}
	}

	nextId = maxId + 1;
}

/**
 * Find all existing change IDs (w:id attributes on w:ins and w:del elements)
 * in a parsed XML document
 */
export function findExistingChangeIds(parsed: ParsedXml): string[] {
	const ids: string[] = [];

	// Find all w:ins elements
	const insElements = findElements(parsed, 'w:ins');
	for (const { element } of insElements) {
		const attrs = getAttributes(element);
		if (attrs['w:id']) {
			ids.push(attrs['w:id']);
		}
	}

	// Find all w:del elements
	const delElements = findElements(parsed, 'w:del');
	for (const { element } of delElements) {
		const attrs = getAttributes(element);
		if (attrs['w:id']) {
			ids.push(attrs['w:id']);
		}
	}

	return ids;
}

/**
 * Initialize change IDs from a parsed document.
 * Finds the highest existing ID and sets the counter to start after it.
 */
export function initializeChangeIdsFromDocument(parsed: ParsedXml): void {
	const existingIds = findExistingChangeIds(parsed);
	initializeChangeIds(existingIds);
}
