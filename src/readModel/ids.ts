import { getAttributes, setAttribute, findElements, type ParsedXml } from './parser.js';
import { W14_NS } from '../xml/namespaces.js';

/**
 * The w14:paraId attribute name as it appears in parsed XML
 */
const PARA_ID_ATTR = 'w14:paraId';
const TEXT_ID_ATTR = 'w14:textId';

/**
 * Track used IDs to ensure uniqueness
 */
const usedIds = new Set<string>();

/**
 * Read existing w14:paraId attributes from all paragraph elements
 * Returns a Map from element to its paraId
 */
export function readParaIds(
	parsed: ParsedXml
): Map<Record<string, unknown>, string> {
	const result = new Map<Record<string, unknown>, string>();
	const paragraphs = findElements(parsed, 'w:p');

	for (const { element } of paragraphs) {
		const attrs = getAttributes(element);
		const paraId = attrs[PARA_ID_ATTR];
		if (paraId) {
			result.set(element, paraId);
			usedIds.add(paraId);
		}
	}

	return result;
}

/**
 * Mint a new unique paragraph ID (8-character hex)
 */
export function mintParaId(): string {
	let id: string;
	do {
		// Generate 8-character uppercase hex ID
		id = Array.from({ length: 8 }, () =>
			Math.floor(Math.random() * 16).toString(16).toUpperCase()
		).join('');
	} while (usedIds.has(id));

	usedIds.add(id);
	return id;
}

/**
 * Reset the used IDs tracking (useful for testing)
 */
export function resetUsedIds(): void {
	usedIds.clear();
}

/**
 * Detect and re-mint duplicate IDs
 * Returns a Map of old ID -> new ID for any that were changed
 */
export function deduplicateIds(
	parsed: ParsedXml
): Map<string, string> {
	const remapped = new Map<string, string>();
	const seenIds = new Set<string>();
	const paragraphs = findElements(parsed, 'w:p');

	for (const { element } of paragraphs) {
		const attrs = getAttributes(element);
		const paraId = attrs[PARA_ID_ATTR];

		if (paraId) {
			if (seenIds.has(paraId)) {
				// Duplicate found - mint new ID
				const newId = mintParaId();
				setAttribute(element, PARA_ID_ATTR, newId);
				remapped.set(paraId, newId);
			} else {
				seenIds.add(paraId);
			}
		}
	}

	return remapped;
}

/**
 * Ensure all paragraphs have IDs (read existing, mint missing, deduplicate)
 * Modifies the parsed XML in place and returns the ID map
 */
export function ensureParaIds(
	parsed: ParsedXml,
	partPrefix: string
): Map<Record<string, unknown>, string> {
	const result = new Map<Record<string, unknown>, string>();
	const seenIds = new Set<string>();
	const paragraphs = findElements(parsed, 'w:p');

	for (const { element } of paragraphs) {
		const attrs = getAttributes(element);
		let paraId = attrs[PARA_ID_ATTR];

		// Mint if missing
		if (!paraId) {
			paraId = mintParaId();
			setAttribute(element, PARA_ID_ATTR, paraId);
			// Also add a textId if missing
			if (!attrs[TEXT_ID_ATTR]) {
				setAttribute(element, TEXT_ID_ATTR, mintParaId());
			}
		}

		// Deduplicate
		if (seenIds.has(paraId)) {
			paraId = mintParaId();
			setAttribute(element, PARA_ID_ATTR, paraId);
		}

		seenIds.add(paraId);

		// Store with part prefix
		const fullId = `${partPrefix}:${paraId}`;
		result.set(element, fullId);
	}

	return result;
}

/**
 * Ensure the w14 namespace is declared on the document root
 */
export function ensureW14Namespace(parsed: ParsedXml): void {
	// Find the root element (usually w:document for main doc)
	for (const node of parsed) {
		if (typeof node !== 'object' || node === null) continue;

		const nodeObj = node as Record<string, unknown>;
		for (const key of Object.keys(nodeObj)) {
			if (key.startsWith('w:') || key === 'w:document') {
				const attrs = getAttributes(nodeObj);
				if (!attrs['xmlns:w14']) {
					setAttribute(nodeObj, 'xmlns:w14', W14_NS);
				}
				return;
			}
		}
	}
}
