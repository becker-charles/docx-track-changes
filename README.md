# docx-track-changes

TypeScript library for applying tracked changes to existing Word documents.

## Features

- **Read model**: Load a DOCX and get an addressable structure where every paragraph has a stable ID
- **ID-based edits**: Replace, insert, or delete content by referencing paragraph IDs
- **Word-native output**: Changes are written as `<w:ins>` and `<w:del>` XML elements

## Installation

```bash
npm install docx-track-changes
```

## Quick Start

### Load and inspect a document

```typescript
import { loadTrackable } from 'docx-track-changes';
import { readFile } from 'fs/promises';

const buffer = await readFile('contract.docx');
const doc = await loadTrackable(buffer);

// Inspect paragraphs
for (const para of doc.body) {
	console.log(`${para.id}: ${para.text}`);
}
```

### Apply edits with tracked changes

```typescript
const para = doc.body.find((p) => p.text.includes('payment terms'));

const result = await doc.applyTrackedEdits(
	[
		{
			type: 'replaceParagraph',
			paraId: para.id,
			content: ['Updated payment terms: Net 30 days.']
		}
	],
	{ author: 'Legal Review' }
);

await writeFile('contract-edited.docx', result.buffer);
```

### Available edit operations

```typescript
// Replace paragraph content
{ type: 'replaceParagraph', paraId: 'body:ABC123', content: ['New text'] }

// Insert after a paragraph
{ type: 'insertAfter', paraId: 'body:ABC123', content: ['New paragraph'] }

// Insert before a paragraph
{ type: 'insertBefore', paraId: 'body:ABC123', content: ['New paragraph'] }

// Delete a paragraph
{ type: 'deleteParagraph', paraId: 'body:ABC123' }

// Delete a range of paragraphs
{ type: 'deleteRange', from: 'body:ABC123', to: 'body:DEF456' }

// Replace text within a paragraph
{ type: 'replaceText', paraId: 'body:ABC123', find: 'old text', replace: 'new text' }
```

### Accept or reject tracked changes

```typescript
// Get existing tracked changes
const changes = doc.getTrackedChanges();

// Accept or reject specific changes
await doc.resolveChanges([
	{ changeId: changes[0].id, action: 'accept' },
	{ changeId: changes[1].id, action: 'reject' }
]);

// Or accept/reject all
await doc.acceptAllChanges();
await doc.rejectAllChanges();
```

### Formatted content

```typescript
await doc.applyTrackedEdits(
	[
		{
			type: 'insertAfter',
			paraId: para.id,
			content: ['Normal text ', { text: 'bold text', bold: true }, ' and ', { text: 'italic', italic: true }]
		}
	],
	{ author: 'Editor' }
);
```

### Run formatting inheritance

Plain strings in content arrays automatically inherit run-level formatting (font family, font size, color, bold, italic, etc.) from the first run of the original/reference paragraph. This preserves document styling consistency.

```typescript
// Original paragraph has Arial 14pt red bold text
// Plain strings inherit that formatting:
await doc.applyTrackedEdits(
	[{
		type: 'replaceParagraph',
		paraId: para.id,
		content: ['This text inherits Arial 14pt red bold']
	}],
	{ author: 'Editor' }
);

// Explicit ContentRun formatting overrides inherited:
await doc.applyTrackedEdits(
	[{
		type: 'replaceParagraph',
		paraId: para.id,
		content: [
			'Inherits original formatting ',
			{ text: 'explicit italic only', italic: true }  // overrides, not bold
		]
	}],
	{ author: 'Editor' }
);
```

This behavior applies to `replaceParagraph`, `insertAfter`, and `insertBefore` operations.

## API

### `loadTrackable(buffer: Buffer | ArrayBuffer)`

Load a DOCX file and return a `TrackableDocument`.

### `loadTrackableFile(path: string)`

Load a DOCX file from disk.

### `TrackableDocument`

| Property     | Description                                |
| ------------ | ------------------------------------------ |
| `paragraphs` | All paragraphs in the document             |
| `body`       | Paragraphs in the main document body       |
| `headers`    | Paragraphs in headers (keyed by header ID) |
| `footers`    | Paragraphs in footers (keyed by footer ID) |
| `footnotes`  | Paragraphs in footnotes                    |
| `endnotes`   | Paragraphs in endnotes                     |

| Method                              | Description                       |
| ----------------------------------- | --------------------------------- |
| `getText()`                         | Get the full document text        |
| `getParagraph(id)`                  | Get a paragraph by ID             |
| `applyTrackedEdits(edits, options)` | Apply edits with tracked changes  |
| `getBuffer()`                       | Get the document as a Buffer      |
| `getTrackedChanges()`               | List existing tracked changes     |
| `resolveChanges(resolutions)`       | Accept or reject specific changes |
| `acceptAllChanges()`                | Accept all tracked changes        |
| `rejectAllChanges()`                | Reject all tracked changes        |

### Paragraph

```typescript
interface Paragraph {
	id: string; // e.g., 'body:4F2A8B91'
	text: string; // Plain text content
	runs: Run[]; // Formatted segments
	style?: string; // Paragraph style name
	numbering?: { level: number; numId: string };
	table?: { tableIndex: number; row: number; cell: number };
}
```

## LLM Integration

The library exports a JSON schema for edit operations:

```typescript
import { editSchema } from 'docx-track-changes';

// Use with OpenAI, Anthropic, etc. for structured output
const response = await openai.chat.completions.create({
  model: 'gpt-4',
  messages: [...],
  response_format: { type: 'json_schema', json_schema: editSchema }
});
```

## License

MIT
