import type {
	Paragraph,
	Edit,
	EditOptions,
	EditResult,
	TrackedChange,
	ChangeResolution,
} from './types.js';
import type { DocxZip } from './utils/zip.js';
import { getDocumentPart, setDocumentPart, saveZip } from './utils/zip.js';
import { getPartPaths, type DocumentParts } from './utils/parts.js';
import {
	parseXml,
	buildXml,
	extractParagraphs,
	ensureW14Namespace,
	resetUsedIds,
	type ParsedXml,
} from './readModel/index.js';
import { applyEdits } from './edits/apply.js';
import {
	resetChangeIds,
	initializeChangeIdsFromDocument,
} from './tracking/changeIds.js';
import {
	getTrackedChangesFromPart,
	resolveChanges as resolveChangesInPart,
	acceptAllChangesInPart,
	rejectAllChangesInPart,
} from './tracking/resolve.js';

/**
 * Internal state for a document part
 */
interface PartState {
	path: string;
	parsed: ParsedXml;
	modified: boolean;
}

/**
 * Main class representing a loaded DOCX document that can be edited with tracked changes.
 *
 * TrackableDocument provides:
 * - A read model with stable paragraph IDs for addressing content
 * - Methods to apply edits that produce Word-native tracked changes (`<w:ins>`, `<w:del>`)
 * - Methods to accept or reject existing tracked changes
 *
 * @example
 * ```typescript
 * import { loadTrackable } from 'docx-track-changes';
 *
 * // Load a document
 * const doc = await loadTrackable(buffer);
 *
 * // Find a paragraph and replace it
 * const para = doc.body.find(p => p.text.includes('old text'));
 * const result = await doc.applyTrackedEdits(
 *   [{ type: 'replaceParagraph', paraId: para.id, content: ['new text'] }],
 *   { author: 'Editor' }
 * );
 *
 * // Save the modified document
 * await writeFile('output.docx', result.buffer);
 * ```
 */
export class TrackableDocument {
	private zip: DocxZip;
	private parts: DocumentParts;
	private partStates: Map<string, PartState> = new Map();
	private paragraphsById: Map<string, Paragraph> = new Map();

	/**
	 * All paragraphs in the document (body, headers, footers, footnotes, endnotes).
	 * Use this for searching across the entire document.
	 */
	readonly paragraphs: Paragraph[] = [];

	/**
	 * Paragraphs in the main document body.
	 * This is the primary content area of the document.
	 */
	readonly body: Paragraph[] = [];

	/**
	 * Paragraphs in headers, keyed by header ID (e.g., 'header1', 'header2').
	 * Word documents can have multiple headers for different sections.
	 */
	readonly headers: Record<string, Paragraph[]> = {};

	/**
	 * Paragraphs in footers, keyed by footer ID (e.g., 'footer1', 'footer2').
	 */
	readonly footers: Record<string, Paragraph[]> = {};

	/** Paragraphs in footnotes */
	readonly footnotes: Paragraph[] = [];

	/** Paragraphs in endnotes */
	readonly endnotes: Paragraph[] = [];

	private constructor(zip: DocxZip, parts: DocumentParts) {
		this.zip = zip;
		this.parts = parts;
	}

	/**
	 * Create a TrackableDocument from a loaded ZIP
	 */
	static async fromZip(zip: DocxZip): Promise<TrackableDocument> {
		// Reset ID tracking for fresh load
		resetUsedIds();
		resetChangeIds();

		const parts = await getPartPaths(zip);
		const doc = new TrackableDocument(zip, parts);
		await doc.loadAllParts();
		return doc;
	}

	/**
	 * Load and parse all document parts
	 */
	private async loadAllParts(): Promise<void> {
		// Load main document
		await this.loadPart(this.parts.main, 'body');

		// Load headers
		for (const headerPath of this.parts.headers) {
			const headerId = this.extractPartId(headerPath, 'header');
			await this.loadPart(headerPath, headerId);
		}

		// Load footers
		for (const footerPath of this.parts.footers) {
			const footerId = this.extractPartId(footerPath, 'footer');
			await this.loadPart(footerPath, footerId);
		}

		// Load footnotes
		if (this.parts.footnotes) {
			await this.loadPart(this.parts.footnotes, 'footnotes');
		}

		// Load endnotes
		if (this.parts.endnotes) {
			await this.loadPart(this.parts.endnotes, 'endnotes');
		}

		// Build the combined paragraphs list
		this.buildParagraphIndex();

		// Initialize change ID counter from existing tracked changes
		this.initializeChangeIdsFromAllParts();
	}

	/**
	 * Initialize change ID counter from all document parts
	 */
	private initializeChangeIdsFromAllParts(): void {
		// Start with the main document
		const mainState = this.partStates.get(this.parts.main);
		if (mainState) {
			initializeChangeIdsFromDocument(mainState.parsed);
		}
	}

	/**
	 * Load and parse a single document part
	 */
	private async loadPart(path: string, partPrefix: string): Promise<void> {
		const xml = await getDocumentPart(this.zip, path);
		if (!xml) return;

		const parsed = parseXml(xml);

		// Ensure w14 namespace is declared
		ensureW14Namespace(parsed);

		// Extract paragraphs
		const paragraphs = extractParagraphs(parsed, partPrefix);

		// Store part state
		this.partStates.set(path, {
			path,
			parsed,
			modified: false,
		});

		// Store paragraphs in appropriate collection
		if (partPrefix === 'body') {
			this.body.push(...paragraphs);
		} else if (partPrefix.startsWith('header')) {
			this.headers[partPrefix] = paragraphs;
		} else if (partPrefix.startsWith('footer')) {
			this.footers[partPrefix] = paragraphs;
		} else if (partPrefix === 'footnotes') {
			this.footnotes.push(...paragraphs);
		} else if (partPrefix === 'endnotes') {
			this.endnotes.push(...paragraphs);
		}
	}

	/**
	 * Extract part ID from path (e.g., 'word/header1.xml' -> 'header1')
	 */
	private extractPartId(path: string, prefix: string): string {
		const match = path.match(new RegExp(`${prefix}(\\d+)\\.xml$`));
		if (match) {
			return `${prefix}${match[1]}`;
		}
		return prefix;
	}

	/**
	 * Build the combined paragraph list and index
	 */
	private buildParagraphIndex(): void {
		// Clear existing
		this.paragraphs.length = 0;
		this.paragraphsById.clear();

		// Add body paragraphs
		for (const p of this.body) {
			this.paragraphs.push(p);
			this.paragraphsById.set(p.id, p);
		}

		// Add header paragraphs
		for (const headerParas of Object.values(this.headers)) {
			for (const p of headerParas) {
				this.paragraphs.push(p);
				this.paragraphsById.set(p.id, p);
			}
		}

		// Add footer paragraphs
		for (const footerParas of Object.values(this.footers)) {
			for (const p of footerParas) {
				this.paragraphs.push(p);
				this.paragraphsById.set(p.id, p);
			}
		}

		// Add footnotes
		for (const p of this.footnotes) {
			this.paragraphs.push(p);
			this.paragraphsById.set(p.id, p);
		}

		// Add endnotes
		for (const p of this.endnotes) {
			this.paragraphs.push(p);
			this.paragraphsById.set(p.id, p);
		}
	}

	/**
	 * Get the full text of the document body as a single string.
	 * Paragraphs are joined with newline characters.
	 *
	 * @param options.view - 'current' (default) shows text as it would appear with all changes accepted.
	 *                       'original' is not yet implemented.
	 * @returns The document text with paragraphs separated by newlines
	 */
	getText(options?: { view?: 'current' | 'original' }): string {
		const view = options?.view ?? 'current';

		// For now, only support current view
		// TODO: Implement 'original' view that excludes tracked insertions and includes deletions
		if (view === 'original') {
			throw new Error('Original view not yet implemented');
		}

		return this.body.map(p => p.text).join('\n');
	}

	/**
	 * Get a paragraph by its ID.
	 *
	 * @param id - The paragraph ID (e.g., 'body:4F2A8B91')
	 * @returns The paragraph if found, undefined otherwise
	 */
	getParagraph(id: string): Paragraph | undefined {
		return this.paragraphsById.get(id);
	}

	/**
	 * Determine which part a paragraph belongs to based on its ID prefix
	 */
	private getPartPathForParagraph(paraId: string): string {
		const prefix = paraId.split(':')[0];

		if (prefix === 'body') {
			return this.parts.main;
		} else if (prefix?.startsWith('header')) {
			const headerPath = this.parts.headers.find(p => p.includes(`${prefix}.xml`));
			return headerPath ?? this.parts.main;
		} else if (prefix?.startsWith('footer')) {
			const footerPath = this.parts.footers.find(p => p.includes(`${prefix}.xml`));
			return footerPath ?? this.parts.main;
		} else if (prefix === 'footnotes' && this.parts.footnotes) {
			return this.parts.footnotes;
		} else if (prefix === 'endnotes' && this.parts.endnotes) {
			return this.parts.endnotes;
		}

		// Default to main document
		return this.parts.main;
	}

	/**
	 * Apply edits to the document with tracked changes.
	 *
	 * Changes are recorded as Word-native tracked changes (`<w:ins>` and `<w:del>`)
	 * that can be accepted or rejected in Microsoft Word or any compatible editor.
	 *
	 * @param edits - Array of edit operations to apply
	 * @param options - Author name, date, and error handling options
	 * @returns The modified document buffer and details of applied/failed edits
	 *
	 * @example
	 * ```typescript
	 * const result = await doc.applyTrackedEdits([
	 *   { type: 'deleteParagraph', paraId: 'body:12345678' },
	 *   { type: 'insertAfter', paraId: 'body:ABCD1234', content: ['New paragraph'] },
	 *   { type: 'replaceText', paraId: 'body:DEADBEEF', find: 'old', replace: 'new' }
	 * ], { author: 'John Doe', date: new Date() });
	 * ```
	 */
	async applyTrackedEdits(
		edits: Edit[],
		options: EditOptions
	): Promise<EditResult> {
		// Group edits by the part they target
		const editsByPart = new Map<string, Edit[]>();

		for (const edit of edits) {
			// Get the paragraph ID from the edit
			let paraId: string;
			if ('paraId' in edit) {
				paraId = edit.paraId;
			} else if ('from' in edit) {
				// For deleteRange, use the 'from' paragraph
				paraId = edit.from;
			} else {
				// Should not happen with current edit types
				continue;
			}

			const partPath = this.getPartPathForParagraph(paraId);
			if (!editsByPart.has(partPath)) {
				editsByPart.set(partPath, []);
			}
			editsByPart.get(partPath)!.push(edit);
		}

		// Apply edits to each part
		const allApplied: EditResult['applied'] = [];
		const allFailed: EditResult['failed'] = [];

		for (const [partPath, partEdits] of editsByPart) {
			const partState = this.partStates.get(partPath);
			if (!partState) {
				// Part not found - mark all edits as failed
				for (const edit of partEdits) {
					allFailed.push({ edit, reason: `Document part not found: ${partPath}` });
				}
				continue;
			}

			const { applied, failed } = applyEdits(partState.parsed, partEdits, options);
			allApplied.push(...applied);
			allFailed.push(...failed);

			// Mark part as modified if any edits were applied
			if (applied.length > 0) {
				partState.modified = true;
			}
		}

		// Generate the output buffer
		const buffer = await this.getBuffer();

		return {
			buffer,
			applied: allApplied,
			failed: allFailed,
		};
	}

	/**
	 * Get the document as a Buffer.
	 *
	 * This includes any minted paragraph IDs and applied edits.
	 * Use this when you need the raw DOCX bytes to save to disk or send over network.
	 *
	 * @returns The DOCX file as a Buffer
	 */
	async getBuffer(): Promise<Buffer> {
		// Write modified parts back to ZIP
		for (const [path, state] of this.partStates) {
			// Always write back (IDs may have been minted)
			const xml = buildXml(state.parsed);
			setDocumentPart(this.zip, path, xml);
		}

		return saveZip(this.zip);
	}

	/**
	 * Get all existing tracked changes in the document.
	 *
	 * This includes both insertions (`<w:ins>`) and deletions (`<w:del>`)
	 * from all document parts (body, headers, footers, footnotes, endnotes).
	 *
	 * @returns Array of tracked changes with their IDs, types, authors, and content
	 */
	getTrackedChanges(): TrackedChange[] {
		const allChanges: TrackedChange[] = [];

		// Collect changes from all parts
		for (const [_path, state] of this.partStates) {
			const changes = getTrackedChangesFromPart(state.parsed);
			allChanges.push(...changes);
		}

		return allChanges;
	}

	/**
	 * Accept or reject specific tracked changes by their IDs.
	 *
	 * - Accepting an insertion keeps the inserted text and removes the `<w:ins>` wrapper
	 * - Accepting a deletion removes the deleted text entirely
	 * - Rejecting an insertion removes the inserted text
	 * - Rejecting a deletion restores the deleted text as normal content
	 *
	 * @param resolutions - Array of change IDs and their resolution actions
	 * @returns The modified document buffer and details of resolved changes
	 *
	 * @example
	 * ```typescript
	 * const changes = doc.getTrackedChanges();
	 * const result = await doc.resolveChanges([
	 *   { changeId: changes[0].id, action: 'accept' },
	 *   { changeId: changes[1].id, action: 'reject' }
	 * ]);
	 * ```
	 */
	async resolveChanges(
		resolutions: ChangeResolution[]
	): Promise<EditResult> {
		const allApplied: EditResult['applied'] = [];
		const allFailed: EditResult['failed'] = [];

		// Process resolutions in each part
		for (const [_path, state] of this.partStates) {
			const result = resolveChangesInPart(state.parsed, resolutions);

			// Track successes
			for (const changeId of result.accepted) {
				allApplied.push({
					edit: { type: 'deleteParagraph', paraId: `resolved:${changeId}` } as Edit,
					changeIds: [changeId],
				});
				state.modified = true;
			}
			for (const changeId of result.rejected) {
				allApplied.push({
					edit: { type: 'deleteParagraph', paraId: `resolved:${changeId}` } as Edit,
					changeIds: [changeId],
				});
				state.modified = true;
			}
		}

		// Determine which changes weren't found anywhere
		const resolvedIds = new Set(allApplied.flatMap((a) => a.changeIds));
		for (const resolution of resolutions) {
			if (!resolvedIds.has(resolution.changeId)) {
				allFailed.push({
					edit: { type: 'deleteParagraph', paraId: `resolve:${resolution.changeId}` } as Edit,
					reason: `Change not found: ${resolution.changeId}`,
				});
			}
		}

		const buffer = await this.getBuffer();
		return { buffer, applied: allApplied, failed: allFailed };
	}

	/**
	 * Accept all tracked changes in the document.
	 *
	 * This is equivalent to "Accept All Changes" in Microsoft Word.
	 * All insertions are kept and deletions are finalized.
	 *
	 * @returns The modified document buffer
	 */
	async acceptAllChanges(): Promise<EditResult> {
		const allApplied: EditResult['applied'] = [];

		for (const [_path, state] of this.partStates) {
			const result = acceptAllChangesInPart(state.parsed);
			if (result.acceptedCount > 0) {
				state.modified = true;
				// We don't have individual change IDs here, just count
				allApplied.push({
					edit: { type: 'deleteParagraph', paraId: 'acceptAll' } as Edit,
					changeIds: [`accepted:${result.acceptedCount}`],
				});
			}
		}

		const buffer = await this.getBuffer();
		return { buffer, applied: allApplied, failed: [] };
	}

	/**
	 * Reject all tracked changes in the document.
	 *
	 * This is equivalent to "Reject All Changes" in Microsoft Word.
	 * All insertions are removed and deletions are restored.
	 *
	 * @returns The modified document buffer
	 */
	async rejectAllChanges(): Promise<EditResult> {
		const allApplied: EditResult['applied'] = [];

		for (const [_path, state] of this.partStates) {
			const result = rejectAllChangesInPart(state.parsed);
			if (result.rejectedCount > 0) {
				state.modified = true;
				allApplied.push({
					edit: { type: 'deleteParagraph', paraId: 'rejectAll' } as Edit,
					changeIds: [`rejected:${result.rejectedCount}`],
				});
			}
		}

		const buffer = await this.getBuffer();
		return { buffer, applied: allApplied, failed: [] };
	}
}
