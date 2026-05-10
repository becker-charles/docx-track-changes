import type { TrackedChange, ChangeResolution } from '../types.js';
import {
	findElements,
	getAttributes,
	getChildren,
	getTagName,
	type ParsedXml,
} from '../readModel/parser.js';
import { cloneElement } from './wrapDelete.js';

/**
 * Extract text content from an element, handling both w:t and w:delText
 */
function extractTextFromElement(element: Record<string, unknown>): string {
	let text = '';

	function traverse(nodes: unknown[]): void {
		for (const node of nodes) {
			if (typeof node !== 'object' || node === null) continue;

			const nodeObj = node as Record<string, unknown>;
			const tagName = getTagName(nodeObj);

			// Check for text elements
			if (tagName === 'w:t' || tagName === 'w:delText') {
				const children = getChildren(nodeObj);
				for (const child of children) {
					if (
						typeof child === 'object' &&
						child !== null &&
						'#text' in (child as Record<string, unknown>)
					) {
						text += String((child as Record<string, unknown>)['#text']);
					}
				}
			}

			// Recurse into children
			const children = getChildren(nodeObj);
			if (children.length > 0) {
				traverse(children);
			}
		}
	}

	const children = getChildren(element);
	traverse(children);
	return text;
}

/**
 * Get all tracked changes from a parsed XML document part
 */
export function getTrackedChangesFromPart(parsed: ParsedXml): TrackedChange[] {
	const changes: TrackedChange[] = [];

	// Find all w:ins elements
	const insertions = findElements(parsed, 'w:ins');
	for (const { element } of insertions) {
		const attrs = getAttributes(element);
		const id = attrs['w:id'];
		const author = attrs['w:author'] || 'Unknown';
		const dateStr = attrs['w:date'];

		if (!id) continue;

		const text = extractTextFromElement(element);
		changes.push({
			id,
			type: 'insertion',
			author,
			date: dateStr ? new Date(dateStr) : new Date(),
			text,
		});
	}

	// Find all w:del elements
	const deletions = findElements(parsed, 'w:del');
	for (const { element } of deletions) {
		const attrs = getAttributes(element);
		const id = attrs['w:id'];
		const author = attrs['w:author'] || 'Unknown';
		const dateStr = attrs['w:date'];

		if (!id) continue;

		const text = extractTextFromElement(element);
		changes.push({
			id,
			type: 'deletion',
			author,
			date: dateStr ? new Date(dateStr) : new Date(),
			text,
		});
	}

	return changes;
}

/**
 * Convert w:delText elements back to w:t (for rejecting deletions)
 */
function convertDelTextToText(element: Record<string, unknown>): Record<string, unknown> {
	const tagName = getTagName(element);
	if (!tagName) {
		// Handle text nodes and other special cases
		const result: Record<string, unknown> = {};
		for (const key of Object.keys(element)) {
			result[key] = element[key];
		}
		return result;
	}

	// If this is a w:delText, convert it to w:t
	if (tagName === 'w:delText') {
		const result: Record<string, unknown> = {
			'w:t': element['w:delText'],
		};
		if (element[':@']) {
			result[':@'] = { ...(element[':@'] as Record<string, unknown>) };
		}
		return result;
	}

	// Recursively process children
	const children = getChildren(element);
	const convertedChildren: unknown[] = [];

	for (const child of children) {
		if (typeof child !== 'object' || child === null) {
			convertedChildren.push(child);
		} else if ('#text' in (child as Record<string, unknown>)) {
			convertedChildren.push({ '#text': (child as Record<string, unknown>)['#text'] });
		} else {
			convertedChildren.push(convertDelTextToText(child as Record<string, unknown>));
		}
	}

	const result: Record<string, unknown> = {
		[tagName]: convertedChildren,
	};

	if (element[':@']) {
		result[':@'] = { ...(element[':@'] as Record<string, unknown>) };
	}

	return result;
}

/**
 * Accept a tracked change by ID.
 * - For insertions: unwrap content (remove w:ins wrapper, keep children)
 * - For deletions: remove the entire w:del element
 *
 * @returns true if the change was found and processed
 */
export function acceptChange(parsed: ParsedXml, changeId: string): boolean {
	// Try to find and accept an insertion
	const insertions = findElements(parsed, 'w:ins');
	for (const { element, parent, index } of insertions) {
		const attrs = getAttributes(element);
		if (attrs['w:id'] === changeId) {
			// Unwrap: replace the w:ins element with its children
			const children = getChildren(element);
			const clonedChildren = children.map((child) => {
				if (typeof child === 'object' && child !== null) {
					if ('#text' in (child as Record<string, unknown>)) {
						return { '#text': (child as Record<string, unknown>)['#text'] };
					}
					return cloneElement(child as Record<string, unknown>);
				}
				return child;
			});

			// Replace the w:ins element with its children
			parent.splice(index, 1, ...clonedChildren);
			return true;
		}
	}

	// Try to find and accept a deletion
	const deletions = findElements(parsed, 'w:del');
	for (const { element, parent, index } of deletions) {
		const attrs = getAttributes(element);
		if (attrs['w:id'] === changeId) {
			// Remove the entire w:del element (the deletion is accepted, text is gone)
			parent.splice(index, 1);
			return true;
		}
	}

	return false;
}

/**
 * Reject a tracked change by ID.
 * - For insertions: remove the entire w:ins element
 * - For deletions: unwrap content and convert w:delText back to w:t
 *
 * @returns true if the change was found and processed
 */
export function rejectChange(parsed: ParsedXml, changeId: string): boolean {
	// Try to find and reject an insertion
	const insertions = findElements(parsed, 'w:ins');
	for (const { element, parent, index } of insertions) {
		const attrs = getAttributes(element);
		if (attrs['w:id'] === changeId) {
			// Remove the entire w:ins element (reject the insertion)
			parent.splice(index, 1);
			return true;
		}
	}

	// Try to find and reject a deletion
	const deletions = findElements(parsed, 'w:del');
	for (const { element, parent, index } of deletions) {
		const attrs = getAttributes(element);
		if (attrs['w:id'] === changeId) {
			// Unwrap and convert w:delText back to w:t
			const children = getChildren(element);
			const convertedChildren = children.map((child) => {
				if (typeof child === 'object' && child !== null) {
					if ('#text' in (child as Record<string, unknown>)) {
						return { '#text': (child as Record<string, unknown>)['#text'] };
					}
					return convertDelTextToText(child as Record<string, unknown>);
				}
				return child;
			});

			// Replace the w:del element with its converted children
			parent.splice(index, 1, ...convertedChildren);
			return true;
		}
	}

	return false;
}

/**
 * Resolve multiple tracked changes
 */
export function resolveChanges(
	parsed: ParsedXml,
	resolutions: ChangeResolution[]
): { accepted: string[]; rejected: string[]; notFound: string[] } {
	const accepted: string[] = [];
	const rejected: string[] = [];
	const notFound: string[] = [];

	// Process resolutions - need to handle nested changes by processing innermost first
	// Sort by processing deletions before insertions when accepting,
	// and insertions before deletions when rejecting
	// For simplicity, we'll process in order but re-find elements each time

	for (const resolution of resolutions) {
		const { changeId, action } = resolution;

		if (action === 'accept') {
			if (acceptChange(parsed, changeId)) {
				accepted.push(changeId);
			} else {
				notFound.push(changeId);
			}
		} else {
			if (rejectChange(parsed, changeId)) {
				rejected.push(changeId);
			} else {
				notFound.push(changeId);
			}
		}
	}

	return { accepted, rejected, notFound };
}

/**
 * Accept all tracked changes in the document
 */
export function acceptAllChangesInPart(parsed: ParsedXml): { acceptedCount: number } {
	let acceptedCount = 0;

	// Keep processing until no more changes are found
	// This handles nested changes
	let hasChanges = true;
	while (hasChanges) {
		hasChanges = false;

		// Accept all insertions (unwrap them)
		let insertions = findElements(parsed, 'w:ins');
		while (insertions.length > 0) {
			const first = insertions[0]!;
			const { element, parent, index } = first;
			const children = getChildren(element);
			const clonedChildren = children.map((child) => {
				if (typeof child === 'object' && child !== null) {
					if ('#text' in (child as Record<string, unknown>)) {
						return { '#text': (child as Record<string, unknown>)['#text'] };
					}
					return cloneElement(child as Record<string, unknown>);
				}
				return child;
			});
			parent.splice(index, 1, ...clonedChildren);
			acceptedCount++;
			hasChanges = true;

			// Re-find after modification
			insertions = findElements(parsed, 'w:ins');
		}

		// Remove all deletions
		let deletions = findElements(parsed, 'w:del');
		while (deletions.length > 0) {
			const first = deletions[0]!;
			const { parent, index } = first;
			parent.splice(index, 1);
			acceptedCount++;
			hasChanges = true;

			// Re-find after modification
			deletions = findElements(parsed, 'w:del');
		}
	}

	return { acceptedCount };
}

/**
 * Reject all tracked changes in the document
 */
export function rejectAllChangesInPart(parsed: ParsedXml): { rejectedCount: number } {
	let rejectedCount = 0;

	// Keep processing until no more changes are found
	let hasChanges = true;
	while (hasChanges) {
		hasChanges = false;

		// Remove all insertions
		let insertions = findElements(parsed, 'w:ins');
		while (insertions.length > 0) {
			const first = insertions[0]!;
			const { parent, index } = first;
			parent.splice(index, 1);
			rejectedCount++;
			hasChanges = true;

			// Re-find after modification
			insertions = findElements(parsed, 'w:ins');
		}

		// Unwrap all deletions (convert delText back to text)
		let deletions = findElements(parsed, 'w:del');
		while (deletions.length > 0) {
			const first = deletions[0]!;
			const { element, parent, index } = first;
			const children = getChildren(element);
			const convertedChildren = children.map((child) => {
				if (typeof child === 'object' && child !== null) {
					if ('#text' in (child as Record<string, unknown>)) {
						return { '#text': (child as Record<string, unknown>)['#text'] };
					}
					return convertDelTextToText(child as Record<string, unknown>);
				}
				return child;
			});
			parent.splice(index, 1, ...convertedChildren);
			rejectedCount++;
			hasChanges = true;

			// Re-find after modification
			deletions = findElements(parsed, 'w:del');
		}
	}

	return { rejectedCount };
}
