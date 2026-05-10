import { describe, it, expect, beforeEach } from 'vitest';
import JSZip from 'jszip';
import { loadTrackable } from '../src/index.js';
import { resetUsedIds } from '../src/readModel/ids.js';
import { resetChangeIds } from '../src/tracking/changeIds.js';
import { createSimpleDoc } from './helpers/createDocx.js';

/**
 * Get the document XML from a DOCX buffer
 */
async function getDocumentXml(buffer: Buffer): Promise<string> {
	const zip = await JSZip.loadAsync(buffer);
	const docXml = await zip.file('word/document.xml')?.async('string');
	return docXml ?? '';
}

async function createMinimalDocx(documentXml: string): Promise<Buffer> {
	const zip = new JSZip();
	zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);
	zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);
	zip.file('word/document.xml', documentXml);
	zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`);
	return zip.generateAsync({ type: 'nodebuffer' });
}

describe('getTrackedChanges', () => {
	beforeEach(() => {
		resetUsedIds();
		resetChangeIds();
	});

	it('should handle changes without w:id attribute', async () => {
		const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:ins w:author="Test"><w:r><w:t>No id</w:t></w:r></w:ins></w:p>
  </w:body>
</w:document>`;
		const buffer = await createMinimalDocx(docXml);
		const doc = await loadTrackable(buffer);
		const changes = doc.getTrackedChanges();
		// Changes without ID should be skipped
		expect(changes.filter(c => c.id)).toHaveLength(0);
	});

	it('should handle changes without date attribute', async () => {
		const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:ins w:id="1" w:author="Test"><w:r><w:t>No date</w:t></w:r></w:ins></w:p>
  </w:body>
</w:document>`;
		const buffer = await createMinimalDocx(docXml);
		const doc = await loadTrackable(buffer);
		const changes = doc.getTrackedChanges();
		expect(changes).toHaveLength(1);
		expect(changes[0]?.date).toBeInstanceOf(Date);
	});

	it('should handle changes without author attribute', async () => {
		const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:del w:id="1"><w:r><w:delText>No author</w:delText></w:r></w:del></w:p>
  </w:body>
</w:document>`;
		const buffer = await createMinimalDocx(docXml);
		const doc = await loadTrackable(buffer);
		const changes = doc.getTrackedChanges();
		expect(changes).toHaveLength(1);
		expect(changes[0]?.author).toBe('Unknown');
	});

	it('should return empty array when no tracked changes exist', async () => {
		const buffer = await createSimpleDoc(['Plain paragraph']);
		const doc = await loadTrackable(buffer);

		const changes = doc.getTrackedChanges();
		expect(changes).toEqual([]);
	});

	it('should list insertions after applying insertAfter edit', async () => {
		const buffer = await createSimpleDoc(['First paragraph']);
		const doc = await loadTrackable(buffer);

		const firstPara = doc.paragraphs[0];
		expect(firstPara).toBeDefined();

		await doc.applyTrackedEdits(
			[{ type: 'insertAfter', paraId: firstPara!.id, content: ['Inserted text'] }],
			{ author: 'Test Author' }
		);

		// Reload to get fresh state
		const result = await doc.getBuffer();
		const doc2 = await loadTrackable(result);

		const changes = doc2.getTrackedChanges();
		expect(changes.length).toBeGreaterThanOrEqual(1);

		const insertion = changes.find(c => c.type === 'insertion');
		expect(insertion).toBeDefined();
		expect(insertion!.author).toBe('Test Author');
		expect(insertion!.text).toContain('Inserted text');
	});

	it('should list deletions after applying deleteParagraph edit', async () => {
		const buffer = await createSimpleDoc(['First paragraph', 'Second to delete']);
		const doc = await loadTrackable(buffer);

		const secondPara = doc.paragraphs.find(p => p.text.includes('Second'));
		expect(secondPara).toBeDefined();

		await doc.applyTrackedEdits(
			[{ type: 'deleteParagraph', paraId: secondPara!.id }],
			{ author: 'Delete Author' }
		);

		// Reload to get fresh state
		const result = await doc.getBuffer();
		const doc2 = await loadTrackable(result);

		const changes = doc2.getTrackedChanges();
		expect(changes.length).toBeGreaterThanOrEqual(1);

		const deletion = changes.find(c => c.type === 'deletion');
		expect(deletion).toBeDefined();
		expect(deletion!.author).toBe('Delete Author');
		expect(deletion!.text).toContain('Second to delete');
	});
});

describe('acceptAllChanges', () => {
	beforeEach(() => {
		resetUsedIds();
		resetChangeIds();
	});

	it('should accept an insertion (keep the inserted text, remove w:ins wrapper)', async () => {
		const buffer = await createSimpleDoc(['First paragraph']);
		const doc = await loadTrackable(buffer);

		const firstPara = doc.paragraphs[0];
		await doc.applyTrackedEdits(
			[{ type: 'insertAfter', paraId: firstPara!.id, content: ['Inserted paragraph'] }],
			{ author: 'Test' }
		);

		// Get buffer with tracked changes
		const withChanges = await doc.getBuffer();

		// Reload and accept all
		const doc2 = await loadTrackable(withChanges);
		expect(doc2.getTrackedChanges().length).toBeGreaterThan(0);

		const result = await doc2.acceptAllChanges();

		// Verify no more tracked changes
		const xml = await getDocumentXml(result.buffer);
		expect(xml).not.toContain('w:ins');
		expect(xml).not.toContain('w:del');

		// Verify the inserted text is still there
		expect(xml).toContain('Inserted paragraph');
	});

	it('should accept a deletion (remove the deleted text entirely)', async () => {
		const buffer = await createSimpleDoc(['Keep this', 'Delete this']);
		const doc = await loadTrackable(buffer);

		const toDelete = doc.paragraphs.find(p => p.text.includes('Delete this'));
		await doc.applyTrackedEdits(
			[{ type: 'deleteParagraph', paraId: toDelete!.id }],
			{ author: 'Test' }
		);

		const withChanges = await doc.getBuffer();

		// Reload and accept all
		const doc2 = await loadTrackable(withChanges);
		const result = await doc2.acceptAllChanges();

		const xml = await getDocumentXml(result.buffer);
		expect(xml).not.toContain('w:ins');
		expect(xml).not.toContain('w:del');
		expect(xml).not.toContain('w:delText');
		expect(xml).toContain('Keep this');
		// The deleted text should be gone
		expect(xml).not.toContain('Delete this');
	});
});

describe('rejectAllChanges', () => {
	beforeEach(() => {
		resetUsedIds();
		resetChangeIds();
	});

	it('should reject an insertion (remove the inserted text)', async () => {
		const buffer = await createSimpleDoc(['Original paragraph']);
		const doc = await loadTrackable(buffer);

		const firstPara = doc.paragraphs[0];
		await doc.applyTrackedEdits(
			[{ type: 'insertAfter', paraId: firstPara!.id, content: ['Should be removed'] }],
			{ author: 'Test' }
		);

		const withChanges = await doc.getBuffer();

		// Reload and reject all
		const doc2 = await loadTrackable(withChanges);
		const result = await doc2.rejectAllChanges();

		const xml = await getDocumentXml(result.buffer);
		expect(xml).not.toContain('w:ins');
		expect(xml).not.toContain('w:del');
		expect(xml).toContain('Original paragraph');
		// The inserted text should be gone
		expect(xml).not.toContain('Should be removed');
	});

	it('should reject a deletion (restore the deleted text)', async () => {
		const buffer = await createSimpleDoc(['Keep this', 'Restore this']);
		const doc = await loadTrackable(buffer);

		const toDelete = doc.paragraphs.find(p => p.text.includes('Restore this'));
		await doc.applyTrackedEdits(
			[{ type: 'deleteParagraph', paraId: toDelete!.id }],
			{ author: 'Test' }
		);

		const withChanges = await doc.getBuffer();

		// Reload and reject all
		const doc2 = await loadTrackable(withChanges);
		const result = await doc2.rejectAllChanges();

		const xml = await getDocumentXml(result.buffer);
		expect(xml).not.toContain('w:ins');
		expect(xml).not.toContain('w:del');
		expect(xml).not.toContain('w:delText');
		expect(xml).toContain('Keep this');
		// The deleted text should be restored as normal w:t
		expect(xml).toContain('Restore this');
		expect(xml).toContain('<w:t');
	});
});

describe('resolveChanges', () => {
	beforeEach(() => {
		resetUsedIds();
		resetChangeIds();
	});

	it('should accept a specific change by ID', async () => {
		const buffer = await createSimpleDoc(['First', 'Second']);
		const doc = await loadTrackable(buffer);

		// Create an insertion
		const firstPara = doc.paragraphs[0];
		await doc.applyTrackedEdits(
			[{ type: 'insertAfter', paraId: firstPara!.id, content: ['Inserted'] }],
			{ author: 'Test' }
		);

		const withChanges = await doc.getBuffer();
		const doc2 = await loadTrackable(withChanges);

		const changes = doc2.getTrackedChanges();
		expect(changes.length).toBeGreaterThan(0);

		const insertionChange = changes.find(c => c.type === 'insertion');
		expect(insertionChange).toBeDefined();

		// Accept just that one change
		const result = await doc2.resolveChanges([
			{ changeId: insertionChange!.id, action: 'accept' }
		]);

		expect(result.applied.length).toBeGreaterThan(0);
		expect(result.failed).toHaveLength(0);

		const xml = await getDocumentXml(result.buffer);
		expect(xml).not.toContain('w:ins');
		expect(xml).toContain('Inserted');
	});

	it('should reject a specific change by ID', async () => {
		const buffer = await createSimpleDoc(['Original', 'To delete']);
		const doc = await loadTrackable(buffer);

		// Create a deletion
		const secondPara = doc.paragraphs.find(p => p.text.includes('To delete'));
		await doc.applyTrackedEdits(
			[{ type: 'deleteParagraph', paraId: secondPara!.id }],
			{ author: 'Test' }
		);

		const withChanges = await doc.getBuffer();
		const doc2 = await loadTrackable(withChanges);

		const changes = doc2.getTrackedChanges();
		const deletionChange = changes.find(c => c.type === 'deletion');
		expect(deletionChange).toBeDefined();

		// Reject the deletion (restore the text)
		const result = await doc2.resolveChanges([
			{ changeId: deletionChange!.id, action: 'reject' }
		]);

		expect(result.applied.length).toBeGreaterThan(0);

		const xml = await getDocumentXml(result.buffer);
		expect(xml).not.toContain('w:del');
		expect(xml).toContain('To delete');
	});

	it('should report not found for invalid change ID', async () => {
		const buffer = await createSimpleDoc(['Paragraph']);
		const doc = await loadTrackable(buffer);

		const result = await doc.resolveChanges([
			{ changeId: 'invalid-id-999', action: 'accept' }
		]);

		expect(result.failed.length).toBeGreaterThan(0);
		expect(result.failed[0]?.reason).toContain('not found');
	});
});

describe('replaceText with accept/reject', () => {
	beforeEach(() => {
		resetUsedIds();
		resetChangeIds();
	});

	it('should accept a text replacement (keep new text)', async () => {
		const buffer = await createSimpleDoc(['The quick brown fox']);
		const doc = await loadTrackable(buffer);

		const para = doc.paragraphs[0];
		await doc.applyTrackedEdits(
			[{
				type: 'replaceText',
				paraId: para!.id,
				find: 'quick',
				replace: 'slow'
			}],
			{ author: 'Editor' }
		);

		const withChanges = await doc.getBuffer();
		const doc2 = await loadTrackable(withChanges);

		// Should have both insertion and deletion
		const changes = doc2.getTrackedChanges();
		expect(changes.length).toBe(2);

		await doc2.acceptAllChanges();
		const result = await doc2.getBuffer();

		const xml = await getDocumentXml(result);
		expect(xml).toContain('slow');
		expect(xml).not.toContain('quick');
		expect(xml).not.toContain('w:ins');
		expect(xml).not.toContain('w:del');
	});

	it('should reject a text replacement (restore original text)', async () => {
		const buffer = await createSimpleDoc(['The quick brown fox']);
		const doc = await loadTrackable(buffer);

		const para = doc.paragraphs[0];
		await doc.applyTrackedEdits(
			[{
				type: 'replaceText',
				paraId: para!.id,
				find: 'quick',
				replace: 'slow'
			}],
			{ author: 'Editor' }
		);

		const withChanges = await doc.getBuffer();
		const doc2 = await loadTrackable(withChanges);

		await doc2.rejectAllChanges();
		const result = await doc2.getBuffer();

		const xml = await getDocumentXml(result);
		expect(xml).toContain('quick');
		expect(xml).not.toContain('slow');
		expect(xml).not.toContain('w:ins');
		expect(xml).not.toContain('w:del');
	});
});

describe('replaceParagraph with accept/reject', () => {
	beforeEach(() => {
		resetUsedIds();
		resetChangeIds();
	});

	it('should accept paragraph replacement', async () => {
		const buffer = await createSimpleDoc(['Old content here']);
		const doc = await loadTrackable(buffer);

		const para = doc.paragraphs[0];
		await doc.applyTrackedEdits(
			[{
				type: 'replaceParagraph',
				paraId: para!.id,
				content: ['New content here']
			}],
			{ author: 'Editor' }
		);

		const withChanges = await doc.getBuffer();
		const doc2 = await loadTrackable(withChanges);

		await doc2.acceptAllChanges();
		const result = await doc2.getBuffer();

		const xml = await getDocumentXml(result);
		expect(xml).toContain('New content here');
		expect(xml).not.toContain('Old content here');
	});

	it('should reject paragraph replacement', async () => {
		const buffer = await createSimpleDoc(['Original content']);
		const doc = await loadTrackable(buffer);

		const para = doc.paragraphs[0];
		await doc.applyTrackedEdits(
			[{
				type: 'replaceParagraph',
				paraId: para!.id,
				content: ['Replacement content']
			}],
			{ author: 'Editor' }
		);

		const withChanges = await doc.getBuffer();
		const doc2 = await loadTrackable(withChanges);

		await doc2.rejectAllChanges();
		const result = await doc2.getBuffer();

		const xml = await getDocumentXml(result);
		expect(xml).toContain('Original content');
		expect(xml).not.toContain('Replacement content');
	});
});
