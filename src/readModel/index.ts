// Read-model related exports
export {
	parseXml,
	buildXml,
	findElements,
	getAttributes,
	setAttribute,
	getTagName,
	getChildren,
	getTextContent,
	type ParsedXml,
} from './parser.js';
export { extractParagraphs } from './paragraphs.js';
export { extractRuns } from './runs.js';
export {
	readParaIds,
	mintParaId,
	deduplicateIds,
	ensureParaIds,
	ensureW14Namespace,
	resetUsedIds,
} from './ids.js';
