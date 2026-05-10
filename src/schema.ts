/**
 * JSON Schema for the Edit type, useful for LLM structured output validation
 */
export const editSchema = {
	$schema: 'https://json-schema.org/draft/2020-12/schema',
	title: 'Edit',
	description: 'An edit operation for a DOCX document',
	oneOf: [
		{
			type: 'object',
			title: 'ReplaceParagraphEdit',
			properties: {
				type: { const: 'replaceParagraph' },
				paraId: { type: 'string', description: 'Paragraph ID to replace' },
				content: { $ref: '#/$defs/contentRunArray' },
				preserveStyle: { type: 'boolean', default: true },
			},
			required: ['type', 'paraId', 'content'],
			additionalProperties: false,
		},
		{
			type: 'object',
			title: 'InsertAfterEdit',
			properties: {
				type: { const: 'insertAfter' },
				paraId: { type: 'string', description: 'Paragraph ID to insert after' },
				content: { $ref: '#/$defs/contentRunArray' },
				style: { type: 'string' },
			},
			required: ['type', 'paraId', 'content'],
			additionalProperties: false,
		},
		{
			type: 'object',
			title: 'InsertBeforeEdit',
			properties: {
				type: { const: 'insertBefore' },
				paraId: { type: 'string', description: 'Paragraph ID to insert before' },
				content: { $ref: '#/$defs/contentRunArray' },
				style: { type: 'string' },
			},
			required: ['type', 'paraId', 'content'],
			additionalProperties: false,
		},
		{
			type: 'object',
			title: 'DeleteParagraphEdit',
			properties: {
				type: { const: 'deleteParagraph' },
				paraId: { type: 'string', description: 'Paragraph ID to delete' },
			},
			required: ['type', 'paraId'],
			additionalProperties: false,
		},
		{
			type: 'object',
			title: 'DeleteRangeEdit',
			properties: {
				type: { const: 'deleteRange' },
				from: { type: 'string', description: 'Starting paragraph ID (inclusive)' },
				to: { type: 'string', description: 'Ending paragraph ID (inclusive)' },
			},
			required: ['type', 'from', 'to'],
			additionalProperties: false,
		},
		{
			type: 'object',
			title: 'ReplaceTextEdit',
			properties: {
				type: { const: 'replaceText' },
				paraId: { type: 'string', description: 'Paragraph ID containing the text' },
				find: { type: 'string', description: 'Text to find' },
				replace: { type: 'string', description: 'Text to replace with' },
				occurrence: { type: 'integer', minimum: 1, default: 1 },
				all: { type: 'boolean', default: false },
			},
			required: ['type', 'paraId', 'find', 'replace'],
			additionalProperties: false,
		},
	],
	$defs: {
		contentRun: {
			oneOf: [
				{ type: 'string' },
				{
					type: 'object',
					properties: {
						text: { type: 'string' },
						bold: { type: 'boolean' },
						italic: { type: 'boolean' },
						underline: { type: 'boolean' },
					},
					required: ['text'],
					additionalProperties: false,
				},
			],
		},
		contentRunArray: {
			type: 'array',
			items: { $ref: '#/$defs/contentRun' },
			minItems: 1,
		},
	},
} as const;

/**
 * JSON Schema for an array of Edit operations
 */
export const editsArraySchema = {
	$schema: 'https://json-schema.org/draft/2020-12/schema',
	title: 'Edits',
	description: 'An array of edit operations for a DOCX document',
	type: 'array',
	items: editSchema,
} as const;
