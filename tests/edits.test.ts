import { describe, it, expect, beforeEach } from 'vitest';
import JSZip from 'jszip';
import { loadTrackable } from '../src/index.js';
import { resetUsedIds } from '../src/readModel/ids.js';
import { resetChangeIds } from '../src/tracking/changeIds.js';
import {
	createSimpleDoc,
	createFormattedDoc,
	createDocWithTable,
	createLegalDoc,
	createMixedFormattingDoc,
} from './helpers/createDocx.js';

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

describe('deleteParagraph', () => {
	beforeEach(() => {
		resetUsedIds();
		resetChangeIds();
	});

	it('should handle paragraphs with bookmarks', async () => {
		const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="BOOKMARK1">
      <w:bookmarkStart w:id="0" w:name="test"/>
      <w:r><w:t>Text with bookmark</w:t></w:r>
      <w:bookmarkEnd w:id="0"/>
    </w:p>
  </w:body>
</w:document>`;
		const buffer = await createMinimalDocx(docXml);
		const doc = await loadTrackable(buffer);
		const para = doc.paragraphs[0];

		const result = await doc.applyTrackedEdits(
			[{ type: 'deleteParagraph', paraId: para!.id }],
			{ author: 'Test' }
		);

		expect(result.applied).toHaveLength(1);
		const xml = await getDocumentXml(result.buffer);
		expect(xml).toContain('w:del');
	});

	it('should handle paragraphs with pPr (paragraph properties)', async () => {
		const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="STYLED01">
      <w:pPr><w:pStyle w:val="Heading1"/></w:pPr>
      <w:r><w:t>Styled heading</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;
		const buffer = await createMinimalDocx(docXml);
		const doc = await loadTrackable(buffer);
		const para = doc.paragraphs[0];

		const result = await doc.applyTrackedEdits(
			[{ type: 'deleteParagraph', paraId: para!.id }],
			{ author: 'Test' }
		);

		expect(result.applied).toHaveLength(1);
		const xml = await getDocumentXml(result.buffer);
		expect(xml).toContain('w:del');
		// pPr should be preserved outside the del
		expect(xml).toContain('w:pStyle');
	});

	it('should delete a paragraph with tracked changes', async () => {
		const buffer = await createSimpleDoc([
			'First paragraph',
			'Second paragraph to delete',
			'Third paragraph',
		]);

		const doc = await loadTrackable(buffer);

		// Find the paragraph to delete
		const paraToDelete = doc.paragraphs.find(p =>
			p.text.includes('Second paragraph to delete')
		);
		expect(paraToDelete).toBeDefined();

		const result = await doc.applyTrackedEdits(
			[{ type: 'deleteParagraph', paraId: paraToDelete!.id }],
			{ author: 'Test Author' }
		);

		expect(result.applied).toHaveLength(1);
		expect(result.failed).toHaveLength(0);
		expect(result.applied[0]?.changeIds).toHaveLength(1);

		// Check the output XML contains w:del
		const outputXml = await getDocumentXml(result.buffer);
		expect(outputXml).toContain('w:del');
		expect(outputXml).toContain('w:delText');
		expect(outputXml).toContain('Test Author');
	});

	it('should fail when paragraph not found', async () => {
		const buffer = await createSimpleDoc(['Only paragraph']);

		const doc = await loadTrackable(buffer);

		const result = await doc.applyTrackedEdits(
			[{ type: 'deleteParagraph', paraId: 'body:NOTFOUND' }],
			{ author: 'Test Author' }
		);

		expect(result.applied).toHaveLength(0);
		expect(result.failed).toHaveLength(1);
		expect(result.failed[0]?.reason).toContain('not found');
	});
});

describe('insertAfter', () => {
	beforeEach(() => {
		resetUsedIds();
		resetChangeIds();
	});

	it('should fail when paragraph not found', async () => {
		const buffer = await createSimpleDoc(['Test']);
		const doc = await loadTrackable(buffer);

		const result = await doc.applyTrackedEdits(
			[{ type: 'insertAfter', paraId: 'body:NOTFOUND', content: ['New'] }],
			{ author: 'Test' }
		);

		expect(result.failed).toHaveLength(1);
		expect(result.failed[0]?.reason).toContain('not found');
	});

	it('should use explicit style when provided', async () => {
		const buffer = await createSimpleDoc(['Reference']);
		const doc = await loadTrackable(buffer);
		const para = doc.body.find(p => p.text === 'Reference');

		const result = await doc.applyTrackedEdits(
			[{ type: 'insertAfter', paraId: para!.id, content: ['Styled'], style: 'Heading1' }],
			{ author: 'Test' }
		);

		expect(result.applied).toHaveLength(1);
		const xml = await getDocumentXml(result.buffer);
		expect(xml).toContain('w:pStyle');
		expect(xml).toContain('Heading1');
	});

	it('should insert a paragraph after the specified paragraph', async () => {
		const buffer = await createSimpleDoc([
			'First paragraph',
			'Second paragraph',
		]);

		const doc = await loadTrackable(buffer);

		const firstPara = doc.paragraphs.find(p => p.text.includes('First paragraph'));
		expect(firstPara).toBeDefined();

		const result = await doc.applyTrackedEdits(
			[{
				type: 'insertAfter',
				paraId: firstPara!.id,
				content: ['Inserted paragraph']
			}],
			{ author: 'Test Author' }
		);

		expect(result.applied).toHaveLength(1);
		expect(result.failed).toHaveLength(0);

		// Check the output XML contains w:ins with the new content
		const outputXml = await getDocumentXml(result.buffer);
		expect(outputXml).toContain('w:ins');
		expect(outputXml).toContain('Test Author');
		expect(outputXml).toContain('Inserted paragraph');
	});

	it('should insert formatted content', async () => {
		const buffer = await createSimpleDoc(['Reference paragraph']);

		const doc = await loadTrackable(buffer);

		const refPara = doc.paragraphs.find(p => p.text.includes('Reference'));
		expect(refPara).toBeDefined();

		const result = await doc.applyTrackedEdits(
			[{
				type: 'insertAfter',
				paraId: refPara!.id,
				content: [
					'Normal text ',
					{ text: 'bold text', bold: true },
					' and ',
					{ text: 'italic', italic: true }
				]
			}],
			{ author: 'Test Author' }
		);

		expect(result.applied).toHaveLength(1);

		const outputXml = await getDocumentXml(result.buffer);
		expect(outputXml).toContain('w:ins');
		expect(outputXml).toContain('Normal text');
		expect(outputXml).toContain('bold text');
		expect(outputXml).toContain('w:b');
		expect(outputXml).toContain('w:i');
	});
});

describe('insertBefore', () => {
	beforeEach(() => {
		resetUsedIds();
		resetChangeIds();
	});

	it('should fail when paragraph not found', async () => {
		const buffer = await createSimpleDoc(['Test']);
		const doc = await loadTrackable(buffer);

		const result = await doc.applyTrackedEdits(
			[{ type: 'insertBefore', paraId: 'body:NOTFOUND', content: ['New'] }],
			{ author: 'Test' }
		);

		expect(result.failed).toHaveLength(1);
		expect(result.failed[0]?.reason).toContain('not found');
	});

	it('should use explicit style when provided', async () => {
		const buffer = await createSimpleDoc(['Reference']);
		const doc = await loadTrackable(buffer);
		const para = doc.body.find(p => p.text === 'Reference');

		const result = await doc.applyTrackedEdits(
			[{ type: 'insertBefore', paraId: para!.id, content: ['Styled'], style: 'Heading2' }],
			{ author: 'Test' }
		);

		expect(result.applied).toHaveLength(1);
		const xml = await getDocumentXml(result.buffer);
		expect(xml).toContain('w:pStyle');
		expect(xml).toContain('Heading2');
	});

	it('should insert a paragraph before the specified paragraph', async () => {
		const buffer = await createSimpleDoc([
			'First paragraph',
			'Second paragraph',
		]);

		const doc = await loadTrackable(buffer);

		const secondPara = doc.paragraphs.find(p => p.text.includes('Second'));
		expect(secondPara).toBeDefined();

		const result = await doc.applyTrackedEdits(
			[{
				type: 'insertBefore',
				paraId: secondPara!.id,
				content: ['Inserted before second']
			}],
			{ author: 'Test Author' }
		);

		expect(result.applied).toHaveLength(1);
		expect(result.failed).toHaveLength(0);

		const outputXml = await getDocumentXml(result.buffer);
		expect(outputXml).toContain('w:ins');
		expect(outputXml).toContain('Inserted before second');
	});
});

describe('deleteRange', () => {
	beforeEach(() => {
		resetUsedIds();
		resetChangeIds();
	});

	it('should fail when from paragraph not found', async () => {
		const buffer = await createSimpleDoc(['Para 1', 'Para 2']);
		const doc = await loadTrackable(buffer);
		const para2 = doc.body.find(p => p.text === 'Para 2');

		const result = await doc.applyTrackedEdits(
			[{ type: 'deleteRange', from: 'body:NOTFOUND', to: para2!.id }],
			{ author: 'Test' }
		);

		expect(result.failed).toHaveLength(1);
	});

	it('should fail when to paragraph not found', async () => {
		const buffer = await createSimpleDoc(['Para 1', 'Para 2']);
		const doc = await loadTrackable(buffer);
		const para1 = doc.body.find(p => p.text === 'Para 1');

		const result = await doc.applyTrackedEdits(
			[{ type: 'deleteRange', from: para1!.id, to: 'body:NOTFOUND' }],
			{ author: 'Test' }
		);

		expect(result.failed).toHaveLength(1);
	});

	it('should fail when range is invalid (from after to)', async () => {
		const buffer = await createSimpleDoc(['Para 1', 'Para 2', 'Para 3']);
		const doc = await loadTrackable(buffer);
		const paras = doc.body.filter(p => p.text.startsWith('Para'));

		const result = await doc.applyTrackedEdits(
			[{ type: 'deleteRange', from: paras[2]!.id, to: paras[0]!.id }],
			{ author: 'Test' }
		);

		expect(result.failed).toHaveLength(1);
		expect(result.failed[0]?.reason).toContain('must come before');
	});

	it('should delete a range of consecutive paragraphs', async () => {
		const buffer = await createSimpleDoc([
			'First - keep',
			'Second - delete',
			'Third - delete',
			'Fourth - delete',
			'Fifth - keep',
		]);

		const doc = await loadTrackable(buffer);

		const secondPara = doc.paragraphs.find(p => p.text.includes('Second - delete'));
		const fourthPara = doc.paragraphs.find(p => p.text.includes('Fourth - delete'));
		expect(secondPara).toBeDefined();
		expect(fourthPara).toBeDefined();

		const result = await doc.applyTrackedEdits(
			[{
				type: 'deleteRange',
				from: secondPara!.id,
				to: fourthPara!.id
			}],
			{ author: 'Test Author' }
		);

		expect(result.applied).toHaveLength(1);
		expect(result.failed).toHaveLength(0);
		// Should have 3 change IDs (one per deleted paragraph)
		expect(result.applied[0]?.changeIds).toHaveLength(3);

		const outputXml = await getDocumentXml(result.buffer);
		// All three deleted paragraphs should have w:del
		expect(outputXml.match(/w:del/g)?.length).toBeGreaterThanOrEqual(3);
	});
});

describe('replaceParagraph', () => {
	beforeEach(() => {
		resetUsedIds();
		resetChangeIds();
	});

	it('should fail when paragraph not found', async () => {
		const buffer = await createSimpleDoc(['Test']);
		const doc = await loadTrackable(buffer);

		const result = await doc.applyTrackedEdits(
			[{ type: 'replaceParagraph', paraId: 'body:NOTFOUND', content: ['New'] }],
			{ author: 'Test' }
		);

		expect(result.failed).toHaveLength(1);
		expect(result.failed[0]?.reason).toContain('not found');
	});

	it('should discard paragraph properties when preserveStyle is false', async () => {
		const buffer = await createSimpleDoc(['Original content']);
		const doc = await loadTrackable(buffer);
		const para = doc.body.find(p => p.text === 'Original content');

		const result = await doc.applyTrackedEdits(
			[{
				type: 'replaceParagraph',
				paraId: para!.id,
				content: ['New content'],
				preserveStyle: false,
			}],
			{ author: 'Test' }
		);

		expect(result.applied).toHaveLength(1);
	});

	it('should replace paragraph content with tracked changes', async () => {
		const buffer = await createSimpleDoc(['Old content to replace']);

		const doc = await loadTrackable(buffer);

		const para = doc.paragraphs.find(p => p.text.includes('Old content'));
		expect(para).toBeDefined();

		const result = await doc.applyTrackedEdits(
			[{
				type: 'replaceParagraph',
				paraId: para!.id,
				content: ['New replacement content']
			}],
			{ author: 'Test Author' }
		);

		expect(result.applied).toHaveLength(1);
		expect(result.failed).toHaveLength(0);
		// Should have 2 change IDs (one for delete, one for insert)
		expect(result.applied[0]?.changeIds).toHaveLength(2);

		const outputXml = await getDocumentXml(result.buffer);
		// Should have both w:del and w:ins
		expect(outputXml).toContain('w:del');
		expect(outputXml).toContain('w:ins');
		expect(outputXml).toContain('New replacement content');
	});

	it('should preserve paragraph properties by default', async () => {
		const buffer = await createSimpleDoc(['Original content']);

		const doc = await loadTrackable(buffer);

		const para = doc.paragraphs.find(p => p.text.includes('Original'));
		expect(para).toBeDefined();

		const result = await doc.applyTrackedEdits(
			[{
				type: 'replaceParagraph',
				paraId: para!.id,
				content: ['New content']
			}],
			{ author: 'Test Author' }
		);

		expect(result.applied).toHaveLength(1);

		const outputXml = await getDocumentXml(result.buffer);
		// Should have both delete and insert
		expect(outputXml).toContain('w:del');
		expect(outputXml).toContain('w:ins');
		expect(outputXml).toContain('New content');
		// The paragraph structure should be maintained
		expect(outputXml).toContain('w:p');
	});
});

describe('multiple edits', () => {
	beforeEach(() => {
		resetUsedIds();
		resetChangeIds();
	});

	it('should apply multiple edits in sequence', async () => {
		const buffer = await createSimpleDoc([
			'First paragraph',
			'Second paragraph',
			'Third paragraph',
		]);

		const doc = await loadTrackable(buffer);

		const firstPara = doc.paragraphs.find(p => p.text.includes('First'));
		const secondPara = doc.paragraphs.find(p => p.text.includes('Second'));
		const thirdPara = doc.paragraphs.find(p => p.text.includes('Third'));

		const result = await doc.applyTrackedEdits(
			[
				{ type: 'deleteParagraph', paraId: secondPara!.id },
				{ type: 'insertAfter', paraId: firstPara!.id, content: ['Inserted'] },
				{ type: 'replaceParagraph', paraId: thirdPara!.id, content: ['Replaced'] }
			],
			{ author: 'Test Author' }
		);

		expect(result.applied).toHaveLength(3);
		expect(result.failed).toHaveLength(0);

		const outputXml = await getDocumentXml(result.buffer);
		expect(outputXml).toContain('w:del');
		expect(outputXml).toContain('w:ins');
		expect(outputXml).toContain('Inserted');
		expect(outputXml).toContain('Replaced');
	});

	it('should continue on error by default', async () => {
		const buffer = await createSimpleDoc([
			'First paragraph',
			'Second paragraph',
		]);

		const doc = await loadTrackable(buffer);

		const firstPara = doc.paragraphs.find(p => p.text.includes('First'));
		const secondPara = doc.paragraphs.find(p => p.text.includes('Second'));

		const result = await doc.applyTrackedEdits(
			[
				{ type: 'deleteParagraph', paraId: firstPara!.id },
				{ type: 'deleteParagraph', paraId: 'body:NOTFOUND' }, // This will fail
				{ type: 'deleteParagraph', paraId: secondPara!.id }
			],
			{ author: 'Test Author' }
		);

		// First and third should succeed, second should fail
		expect(result.applied).toHaveLength(2);
		expect(result.failed).toHaveLength(1);
		expect(result.failed[0]?.reason).toContain('not found');
	});

	it('should stop on error when continueOnError is false', async () => {
		const buffer = await createSimpleDoc([
			'First paragraph',
			'Second paragraph',
		]);

		const doc = await loadTrackable(buffer);

		const firstPara = doc.paragraphs.find(p => p.text.includes('First'));

		const result = await doc.applyTrackedEdits(
			[
				{ type: 'deleteParagraph', paraId: firstPara!.id },
				{ type: 'deleteParagraph', paraId: 'body:NOTFOUND' }, // This will fail
				{ type: 'deleteParagraph', paraId: 'body:ANOTHER' }
			],
			{ author: 'Test Author', continueOnError: false }
		);

		// First should succeed, second fails, third not attempted
		expect(result.applied).toHaveLength(1);
		expect(result.failed).toHaveLength(1);
	});
});

describe('edit with custom date', () => {
	beforeEach(() => {
		resetUsedIds();
		resetChangeIds();
	});

	it('should use provided date in tracked changes', async () => {
		const buffer = await createSimpleDoc(['Test paragraph']);

		const doc = await loadTrackable(buffer);

		const para = doc.paragraphs.find(p => p.text.includes('Test'));
		const customDate = new Date('2025-01-15T10:30:00Z');

		const result = await doc.applyTrackedEdits(
			[{ type: 'deleteParagraph', paraId: para!.id }],
			{ author: 'Test Author', date: customDate }
		);

		expect(result.applied).toHaveLength(1);

		const outputXml = await getDocumentXml(result.buffer);
		expect(outputXml).toContain('2025-01-15T10:30:00Z');
	});
});

describe('table cell editing', () => {
	beforeEach(() => {
		resetUsedIds();
		resetChangeIds();
	});

	it('should edit table cell content', async () => {
		const buffer = await createDocWithTable({
			tableData: [
				['Original cell text', 'Other cell'],
			],
		});

		const doc = await loadTrackable(buffer);

		const cellPara = doc.paragraphs.find(
			p => p.table !== undefined && p.text.includes('Original cell text')
		);
		expect(cellPara).toBeDefined();

		const result = await doc.applyTrackedEdits(
			[{
				type: 'replaceParagraph',
				paraId: cellPara!.id,
				content: ['Updated cell text'],
			}],
			{ author: 'Test Author' }
		);

		expect(result.applied).toHaveLength(1);

		const outputXml = await getDocumentXml(result.buffer);
		expect(outputXml).toContain('Updated cell text');
	});
});

describe('legal document scenario', () => {
	beforeEach(() => {
		resetUsedIds();
		resetChangeIds();
	});

	it('should handle a realistic legal document workflow', async () => {
		const buffer = await createLegalDoc({
			title: 'SERVICE AGREEMENT',
			parties: [
				'This Agreement is entered into by:',
				'Party A: Acme Corporation',
				'Party B: Widget Inc.',
			],
			clauses: [
				{
					heading: '1. Services',
					text: 'Party A agrees to provide consulting services to Party B.',
				},
				{
					heading: '2. Payment',
					text: 'Party B shall pay Party A $10,000 per month.',
				},
				{
					heading: '3. Term',
					text: 'This agreement shall remain in effect for 12 months.',
				},
			],
		});

		const doc = await loadTrackable(buffer);

		// Find the payment clause
		const paymentPara = doc.paragraphs.find(p =>
			p.text.includes('$10,000 per month')
		);
		expect(paymentPara).toBeDefined();

		// Modify the payment amount
		const result = await doc.applyTrackedEdits(
			[{
				type: 'replaceParagraph',
				paraId: paymentPara!.id,
				content: ['Party B shall pay Party A $15,000 per month.'],
			}],
			{ author: 'Legal Review', date: new Date('2025-06-01') }
		);

		expect(result.applied).toHaveLength(1);

		const outputXml = await getDocumentXml(result.buffer);
		expect(outputXml).toContain('w:del');
		expect(outputXml).toContain('w:ins');
		expect(outputXml).toContain('$15,000');
		expect(outputXml).toContain('Legal Review');
	});
});

describe('round-trip', () => {
	beforeEach(() => {
		resetUsedIds();
		resetChangeIds();
	});

	it('should maintain document integrity after load-edit-save cycle', async () => {
		const buffer = await createFormattedDoc([
			{ text: 'Important Heading', heading: 'Heading1' },
			'First body paragraph with some text.',
			{ text: 'Bold paragraph', bold: true },
			'Final paragraph.',
		]);

		// Load and edit
		const doc = await loadTrackable(buffer);
		const originalParaCount = doc.paragraphs.length;

		const firstBodyPara = doc.paragraphs.find(p =>
			p.text.includes('First body paragraph')
		);

		const result = await doc.applyTrackedEdits(
			[{
				type: 'insertAfter',
				paraId: firstBodyPara!.id,
				content: ['Newly inserted paragraph'],
			}],
			{ author: 'Editor' }
		);

		// Reload the edited document
		const reloadedDoc = await loadTrackable(result.buffer);

		// Should have all original paragraphs plus the inserted one
		expect(reloadedDoc.paragraphs.length).toBe(originalParaCount + 1);
		expect(reloadedDoc.getText()).toContain('Newly inserted paragraph');
	});
});

describe('replaceText', () => {
	beforeEach(() => {
		resetUsedIds();
		resetChangeIds();
	});

	it('should fail when paragraph not found', async () => {
		const buffer = await createSimpleDoc(['Test text']);
		const doc = await loadTrackable(buffer);

		const result = await doc.applyTrackedEdits(
			[{ type: 'replaceText', paraId: 'body:NOTFOUND', find: 'Test', replace: 'New' }],
			{ author: 'Test' }
		);

		expect(result.failed).toHaveLength(1);
		expect(result.failed[0]?.reason).toContain('not found');
	});

	it('should fail when occurrence exceeds available matches', async () => {
		const buffer = await createSimpleDoc(['The cat sat.']);
		const doc = await loadTrackable(buffer);
		const para = doc.body.find(p => p.text.includes('cat'));

		const result = await doc.applyTrackedEdits(
			[{ type: 'replaceText', paraId: para!.id, find: 'cat', replace: 'dog', occurrence: 5 }],
			{ author: 'Test' }
		);

		expect(result.failed).toHaveLength(1);
	});

	it('should replace text within a paragraph with tracked changes', async () => {
		const buffer = await createSimpleDoc([
			'The party agrees to indemnify the other party.',
		]);

		const doc = await loadTrackable(buffer);

		const para = doc.paragraphs.find(p => p.text.includes('indemnify'));
		expect(para).toBeDefined();

		const result = await doc.applyTrackedEdits(
			[{
				type: 'replaceText',
				paraId: para!.id,
				find: 'indemnify',
				replace: 'hold harmless',
			}],
			{ author: 'Legal Review' }
		);

		expect(result.applied).toHaveLength(1);
		expect(result.failed).toHaveLength(0);
		// Should have 2 change IDs (one for delete, one for insert)
		expect(result.applied[0]?.changeIds).toHaveLength(2);

		const outputXml = await getDocumentXml(result.buffer);
		expect(outputXml).toContain('w:del');
		expect(outputXml).toContain('w:ins');
		expect(outputXml).toContain('w:delText');
		expect(outputXml).toContain('indemnify');
		expect(outputXml).toContain('hold harmless');
		expect(outputXml).toContain('Legal Review');
	});

	it('should fail when text is not found', async () => {
		const buffer = await createSimpleDoc(['Hello world']);

		const doc = await loadTrackable(buffer);

		const para = doc.paragraphs.find(p => p.text.includes('Hello'));
		expect(para).toBeDefined();

		const result = await doc.applyTrackedEdits(
			[{
				type: 'replaceText',
				paraId: para!.id,
				find: 'goodbye',
				replace: 'farewell',
			}],
			{ author: 'Test' }
		);

		expect(result.applied).toHaveLength(0);
		expect(result.failed).toHaveLength(1);
		expect(result.failed[0]?.reason).toContain('not found');
	});

	it('should replace specific occurrence', async () => {
		const buffer = await createSimpleDoc([
			'The cat sat on the mat. The cat was happy.',
		]);

		const doc = await loadTrackable(buffer);

		const para = doc.paragraphs.find(p => p.text.includes('cat'));
		expect(para).toBeDefined();

		// Replace the second "cat"
		const result = await doc.applyTrackedEdits(
			[{
				type: 'replaceText',
				paraId: para!.id,
				find: 'cat',
				replace: 'dog',
				occurrence: 2,
			}],
			{ author: 'Editor' }
		);

		expect(result.applied).toHaveLength(1);
		expect(result.failed).toHaveLength(0);

		// Reload and check content
		const reloadedDoc = await loadTrackable(result.buffer);
		const text = reloadedDoc.getText();
		// First "cat" should still be there (now with the insertion after it)
		// The structure will show the tracked changes
		expect(text).toContain('cat');
		expect(text).toContain('dog');
	});

	it('should replace all occurrences when all: true', async () => {
		const buffer = await createSimpleDoc([
			'The cat and the cat and the cat.',
		]);

		const doc = await loadTrackable(buffer);

		const para = doc.paragraphs.find(p => p.text.includes('cat'));
		expect(para).toBeDefined();

		const result = await doc.applyTrackedEdits(
			[{
				type: 'replaceText',
				paraId: para!.id,
				find: 'cat',
				replace: 'dog',
				all: true,
			}],
			{ author: 'Editor' }
		);

		expect(result.applied).toHaveLength(1);
		expect(result.failed).toHaveLength(0);
		// Should have 6 change IDs (2 per replacement: 1 delete + 1 insert) × 3 occurrences
		expect(result.applied[0]?.changeIds).toHaveLength(6);

		const outputXml = await getDocumentXml(result.buffer);
		// Should have 3 deletions and 3 insertions (match opening tags only)
		const delMatches = outputXml.match(/<w:del /g);
		const insMatches = outputXml.match(/<w:ins /g);
		expect(delMatches?.length).toBe(3);
		expect(insMatches?.length).toBe(3);
	});

	it('should preserve text before and after the match', async () => {
		const buffer = await createSimpleDoc([
			'Hello wonderful world!',
		]);

		const doc = await loadTrackable(buffer);

		const para = doc.paragraphs.find(p => p.text.includes('wonderful'));
		expect(para).toBeDefined();

		const result = await doc.applyTrackedEdits(
			[{
				type: 'replaceText',
				paraId: para!.id,
				find: 'wonderful',
				replace: 'beautiful',
			}],
			{ author: 'Editor' }
		);

		expect(result.applied).toHaveLength(1);

		// Reload and verify the surrounding text is preserved
		const reloadedDoc = await loadTrackable(result.buffer);
		const text = reloadedDoc.getText();
		expect(text).toContain('Hello');
		expect(text).toContain('world');
		expect(text).toContain('beautiful');
	});

	it('should handle replacement at the start of text', async () => {
		const buffer = await createSimpleDoc([
			'Hello world',
		]);

		const doc = await loadTrackable(buffer);

		const para = doc.paragraphs.find(p => p.text.includes('Hello'));
		expect(para).toBeDefined();

		const result = await doc.applyTrackedEdits(
			[{
				type: 'replaceText',
				paraId: para!.id,
				find: 'Hello',
				replace: 'Goodbye',
			}],
			{ author: 'Editor' }
		);

		expect(result.applied).toHaveLength(1);

		const reloadedDoc = await loadTrackable(result.buffer);
		expect(reloadedDoc.getText()).toContain('Goodbye');
		expect(reloadedDoc.getText()).toContain('world');
	});

	it('should handle replacement at the end of text', async () => {
		const buffer = await createSimpleDoc([
			'Hello world',
		]);

		const doc = await loadTrackable(buffer);

		const para = doc.paragraphs.find(p => p.text.includes('world'));
		expect(para).toBeDefined();

		const result = await doc.applyTrackedEdits(
			[{
				type: 'replaceText',
				paraId: para!.id,
				find: 'world',
				replace: 'universe',
			}],
			{ author: 'Editor' }
		);

		expect(result.applied).toHaveLength(1);

		const reloadedDoc = await loadTrackable(result.buffer);
		expect(reloadedDoc.getText()).toContain('Hello');
		expect(reloadedDoc.getText()).toContain('universe');
	});

	it('should handle deletion (empty replacement)', async () => {
		const buffer = await createSimpleDoc([
			'Remove this word please.',
		]);

		const doc = await loadTrackable(buffer);

		const para = doc.paragraphs.find(p => p.text.includes('this word'));
		expect(para).toBeDefined();

		const result = await doc.applyTrackedEdits(
			[{
				type: 'replaceText',
				paraId: para!.id,
				find: 'this word ',
				replace: '',
			}],
			{ author: 'Editor' }
		);

		expect(result.applied).toHaveLength(1);

		const outputXml = await getDocumentXml(result.buffer);
		expect(outputXml).toContain('w:del');
		expect(outputXml).toContain('this word');
	});

	it('should handle text spanning multiple runs with different formatting', async () => {
		// Create a paragraph with: "Hello " (normal) + "beautiful" (bold) + " world" (normal)
		const buffer = await createMixedFormattingDoc([
			[
				{ text: 'Hello ' },
				{ text: 'beautiful', bold: true },
				{ text: ' world' },
			],
		]);

		const doc = await loadTrackable(buffer);

		const para = doc.paragraphs.find(p => p.text.includes('beautiful'));
		expect(para).toBeDefined();

		// Replace "o beautiful w" which spans 3 runs
		const result = await doc.applyTrackedEdits(
			[{
				type: 'replaceText',
				paraId: para!.id,
				find: 'o beautiful w',
				replace: 'o amazing w',
			}],
			{ author: 'Editor' }
		);

		expect(result.applied).toHaveLength(1);
		expect(result.failed).toHaveLength(0);

		const outputXml = await getDocumentXml(result.buffer);
		expect(outputXml).toContain('w:del');
		expect(outputXml).toContain('w:ins');
		expect(outputXml).toContain('o amazing w');
	});

	it('should preserve formatting when replacing within a formatted run', async () => {
		// Create a paragraph with bold text
		const buffer = await createMixedFormattingDoc([
			[
				{ text: 'This is ', bold: true },
				{ text: 'important', bold: true },
				{ text: ' text.', bold: true },
			],
		]);

		const doc = await loadTrackable(buffer);

		const para = doc.paragraphs.find(p => p.text.includes('important'));
		expect(para).toBeDefined();

		const result = await doc.applyTrackedEdits(
			[{
				type: 'replaceText',
				paraId: para!.id,
				find: 'important',
				replace: 'critical',
			}],
			{ author: 'Editor' }
		);

		expect(result.applied).toHaveLength(1);

		const outputXml = await getDocumentXml(result.buffer);
		// The replacement text should also be bold
		expect(outputXml).toContain('critical');
		expect(outputXml).toContain('w:b');
	});
});
