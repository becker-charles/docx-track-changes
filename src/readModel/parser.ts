import { XMLParser, XMLBuilder } from 'fast-xml-parser';
import { DocxError } from '../errors.js';

/**
 * Parser options for reading Word XML with full fidelity
 */
const PARSER_OPTIONS = {
	ignoreAttributes: false,
	attributeNamePrefix: '@_',
	preserveOrder: true,
	commentPropName: '#comment',
	cdataPropName: '#cdata',
	textNodeName: '#text',
	trimValues: false,
	parseTagValue: false,
	parseAttributeValue: false,
};

/**
 * Builder options for writing XML back with preserved structure
 */
const BUILDER_OPTIONS = {
	ignoreAttributes: false,
	attributeNamePrefix: '@_',
	preserveOrder: true,
	commentPropName: '#comment',
	cdataPropName: '#cdata',
	textNodeName: '#text',
	format: false,
	suppressEmptyNode: false,
	suppressBooleanAttributes: false,
};

export type ParsedXml = unknown[];

/**
 * Parse Word XML into a structure we can manipulate
 */
export function parseXml(xml: string): ParsedXml {
	try {
		const parser = new XMLParser(PARSER_OPTIONS);
		const result = parser.parse(xml);
		return result;
	} catch (error) {
		throw new DocxError(
			'INVALID_XML',
			`Failed to parse XML: ${error instanceof Error ? error.message : String(error)}`
		);
	}
}

/**
 * Build XML string from parsed structure
 */
export function buildXml(parsed: ParsedXml): string {
	const builder = new XMLBuilder(BUILDER_OPTIONS);
	return builder.build(parsed);
}

/**
 * Find all elements matching a tag name in the parsed XML
 */
export function findElements(
	parsed: ParsedXml,
	tagName: string
): Array<{ element: Record<string, unknown>; parent: unknown[]; index: number }> {
	const results: Array<{ element: Record<string, unknown>; parent: unknown[]; index: number }> = [];

	function traverse(nodes: unknown[], _parent: unknown[] | null = null): void {
		if (!Array.isArray(nodes)) return;

		for (let i = 0; i < nodes.length; i++) {
			const node = nodes[i];
			if (typeof node !== 'object' || node === null) continue;

			const nodeObj = node as Record<string, unknown>;
			for (const key of Object.keys(nodeObj)) {
				if (key.startsWith('@_') || key.startsWith('#') || key === ':@') continue;

				if (key === tagName) {
					results.push({ element: nodeObj, parent: nodes, index: i });
				}

				// Recurse into children
				const children = nodeObj[key];
				if (Array.isArray(children)) {
					traverse(children, nodes);
				}
			}
		}
	}

	traverse(parsed);
	return results;
}

/**
 * Get attributes from an element in preserveOrder mode
 */
export function getAttributes(element: Record<string, unknown>): Record<string, string> {
	const attrs = element[':@'] as Record<string, string> | undefined;
	if (!attrs) return {};

	const result: Record<string, string> = {};
	for (const [key, value] of Object.entries(attrs)) {
		if (key.startsWith('@_')) {
			result[key.slice(2)] = value;
		}
	}
	return result;
}

/**
 * Set an attribute on an element in preserveOrder mode
 */
export function setAttribute(
	element: Record<string, unknown>,
	name: string,
	value: string
): void {
	if (!element[':@']) {
		element[':@'] = {};
	}
	(element[':@'] as Record<string, string>)[`@_${name}`] = value;
}

/**
 * Get the tag name of an element in preserveOrder mode
 */
export function getTagName(element: Record<string, unknown>): string | null {
	for (const key of Object.keys(element)) {
		if (!key.startsWith('@_') && !key.startsWith('#') && key !== ':@') {
			return key;
		}
	}
	return null;
}

/**
 * Get the children of an element in preserveOrder mode
 */
export function getChildren(element: Record<string, unknown>): unknown[] {
	const tagName = getTagName(element);
	if (!tagName) return [];
	const children = element[tagName];
	return Array.isArray(children) ? children : [];
}

/**
 * Get text content from an element (concatenated from all text nodes)
 */
export function getTextContent(element: Record<string, unknown>): string {
	let text = '';

	function traverse(nodes: unknown[]): void {
		for (const node of nodes) {
			if (typeof node !== 'object' || node === null) continue;

			const nodeObj = node as Record<string, unknown>;

			// Check for text node
			if ('#text' in nodeObj) {
				text += String(nodeObj['#text']);
			}

			// Recurse into children
			for (const key of Object.keys(nodeObj)) {
				if (key.startsWith('@_') || key.startsWith('#') || key === ':@') continue;
				const children = nodeObj[key];
				if (Array.isArray(children)) {
					traverse(children);
				}
			}
		}
	}

	const children = getChildren(element);
	traverse(children);

	return text;
}
