import { describe, it, expect, beforeEach } from 'vitest';
import JSZip from 'jszip';
import { loadTrackable } from '../src/index.js';
import { resetUsedIds } from '../src/readModel/ids.js';

/**
 * Create a minimal valid DOCX buffer for testing
 */
async function createMinimalDocx(documentXml: string): Promise<Buffer> {
	const zip = new JSZip();

	// [Content_Types].xml
	zip.file(
		'[Content_Types].xml',
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
	);

	// _rels/.rels
	zip.file(
		'_rels/.rels',
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
	);

	// word/document.xml
	zip.file('word/document.xml', documentXml);

	// word/_rels/document.xml.rels (empty but valid)
	zip.file(
		'word/_rels/document.xml.rels',
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`
	);

	return zip.generateAsync({ type: 'nodebuffer' });
}

describe('loadTrackable', () => {
	beforeEach(() => {
		resetUsedIds();
	});

	it('should handle documents with footnotes/endnotes relationships', async () => {
		const zip = new JSZip();
		zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/footnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml"/>
  <Override PartName="/word/endnotes.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml"/>
</Types>`);
		zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);
		zip.file('word/document.xml', `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>Main text</w:t></w:r></w:p></w:body>
</w:document>`);
		zip.file('word/_rels/document.xml.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" Target="footnotes.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" Target="endnotes.xml"/>
</Relationships>`);
		zip.file('word/footnotes.xml', `<?xml version="1.0"?>
<w:footnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:footnote w:id="1"><w:p><w:r><w:t>Footnote text</w:t></w:r></w:p></w:footnote>
</w:footnotes>`);
		zip.file('word/endnotes.xml', `<?xml version="1.0"?>
<w:endnotes xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:endnote w:id="1"><w:p><w:r><w:t>Endnote text</w:t></w:r></w:p></w:endnote>
</w:endnotes>`);

		const buffer = await zip.generateAsync({ type: 'nodebuffer' });
		const doc = await loadTrackable(buffer);

		// Main body should be loaded
		expect(doc.body.length).toBeGreaterThan(0);
		expect(doc.body[0]?.text).toBe('Main text');
	});

	it('should handle documents without relationship file', async () => {
		const zip = new JSZip();
		zip.file('[Content_Types].xml', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`);
		zip.file('_rels/.rels', `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`);
		zip.file('word/document.xml', `<?xml version="1.0"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>Test</w:t></w:r></w:p></w:body>
</w:document>`);
		// Note: no word/_rels/document.xml.rels

		const buffer = await zip.generateAsync({ type: 'nodebuffer' });
		const doc = await loadTrackable(buffer);

		expect(doc.paragraphs.length).toBeGreaterThan(0);
	});

	it('should load a simple DOCX file', async () => {
		const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Hello, World!</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

		const buffer = await createMinimalDocx(docXml);
		const doc = await loadTrackable(buffer);

		expect(doc.paragraphs).toHaveLength(1);
		expect(doc.paragraphs[0]?.text).toBe('Hello, World!');
	});

	it('should extract paragraph IDs from w14:paraId attributes', async () => {
		const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="12345678" w14:textId="AABBCCDD">
      <w:r>
        <w:t>Paragraph with existing ID</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

		const buffer = await createMinimalDocx(docXml);
		const doc = await loadTrackable(buffer);

		expect(doc.paragraphs).toHaveLength(1);
		expect(doc.paragraphs[0]?.id).toBe('body:12345678');
	});

	it('should mint IDs for paragraphs without w14:paraId', async () => {
		const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>No ID here</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

		const buffer = await createMinimalDocx(docXml);
		const doc = await loadTrackable(buffer);

		expect(doc.paragraphs).toHaveLength(1);
		// Should have a minted ID in format body:XXXXXXXX
		expect(doc.paragraphs[0]?.id).toMatch(/^body:[A-F0-9]{8}$/);
	});

	it('should deduplicate IDs', async () => {
		const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="DUPLICATE1" w14:textId="11111111">
      <w:r><w:t>First</w:t></w:r>
    </w:p>
    <w:p w14:paraId="DUPLICATE1" w14:textId="22222222">
      <w:r><w:t>Second with same ID</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

		const buffer = await createMinimalDocx(docXml);
		const doc = await loadTrackable(buffer);

		expect(doc.paragraphs).toHaveLength(2);
		// First keeps original
		expect(doc.paragraphs[0]?.id).toBe('body:DUPLICATE1');
		// Second gets a new minted ID
		expect(doc.paragraphs[1]?.id).not.toBe('body:DUPLICATE1');
		expect(doc.paragraphs[1]?.id).toMatch(/^body:[A-F0-9]{8}$/);
	});

	it('should extract multiple paragraphs', async () => {
		const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>First paragraph</w:t></w:r></w:p>
    <w:p><w:r><w:t>Second paragraph</w:t></w:r></w:p>
    <w:p><w:r><w:t>Third paragraph</w:t></w:r></w:p>
  </w:body>
</w:document>`;

		const buffer = await createMinimalDocx(docXml);
		const doc = await loadTrackable(buffer);

		expect(doc.paragraphs).toHaveLength(3);
		expect(doc.paragraphs[0]?.text).toBe('First paragraph');
		expect(doc.paragraphs[1]?.text).toBe('Second paragraph');
		expect(doc.paragraphs[2]?.text).toBe('Third paragraph');
	});

	it('should extract formatted runs', async () => {
		const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:t>Normal </w:t>
      </w:r>
      <w:r>
        <w:rPr><w:b/></w:rPr>
        <w:t>bold </w:t>
      </w:r>
      <w:r>
        <w:rPr><w:i/></w:rPr>
        <w:t>italic</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

		const buffer = await createMinimalDocx(docXml);
		const doc = await loadTrackable(buffer);

		expect(doc.paragraphs).toHaveLength(1);
		const para = doc.paragraphs[0]!;

		expect(para.runs).toHaveLength(3);
		expect(para.runs[0]).toEqual({ text: 'Normal ' });
		expect(para.runs[1]).toEqual({ text: 'bold ', bold: true });
		expect(para.runs[2]).toEqual({ text: 'italic', italic: true });
	});
});

describe('special paragraph content', () => {
	beforeEach(() => {
		resetUsedIds();
	});

	it('should handle tabs in text', async () => {
		const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Before</w:t><w:tab/><w:t>After</w:t></w:r></w:p>
  </w:body>
</w:document>`;
		const buffer = await createMinimalDocx(docXml);
		const doc = await loadTrackable(buffer);
		expect(doc.paragraphs[0]?.text).toContain('\t');
	});

	it('should handle line breaks in text', async () => {
		const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Line1</w:t><w:br/><w:t>Line2</w:t></w:r></w:p>
  </w:body>
</w:document>`;
		const buffer = await createMinimalDocx(docXml);
		const doc = await loadTrackable(buffer);
		expect(doc.paragraphs[0]?.text).toContain('\n');
	});

	it('should handle explicit false formatting values', async () => {
		const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:rPr><w:b w:val="false"/><w:i w:val="0"/></w:rPr><w:t>Not bold or italic</w:t></w:r></w:p>
  </w:body>
</w:document>`;
		const buffer = await createMinimalDocx(docXml);
		const doc = await loadTrackable(buffer);
		expect(doc.paragraphs[0]?.runs[0]?.bold).toBeFalsy();
		expect(doc.paragraphs[0]?.runs[0]?.italic).toBeFalsy();
	});

	it('should handle underline with val=none', async () => {
		const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:rPr><w:u w:val="none"/></w:rPr><w:t>Not underlined</w:t></w:r></w:p>
  </w:body>
</w:document>`;
		const buffer = await createMinimalDocx(docXml);
		const doc = await loadTrackable(buffer);
		expect(doc.paragraphs[0]?.runs[0]?.underline).toBeFalsy();
	});

	it('should extract character style from runs', async () => {
		const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:rPr><w:rStyle w:val="Strong"/></w:rPr><w:t>Styled text</w:t></w:r></w:p>
  </w:body>
</w:document>`;
		const buffer = await createMinimalDocx(docXml);
		const doc = await loadTrackable(buffer);
		expect(doc.paragraphs[0]?.runs[0]?.style).toBe('Strong');
	});

	it('should handle hyperlinks containing runs', async () => {
		const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:hyperlink><w:r><w:t>Link text</w:t></w:r></w:hyperlink></w:p>
  </w:body>
</w:document>`;
		const buffer = await createMinimalDocx(docXml);
		const doc = await loadTrackable(buffer);
		expect(doc.paragraphs[0]?.text).toBe('Link text');
	});

	it('should handle existing tracked insertions in read model', async () => {
		const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:ins w:id="1" w:author="Test"><w:r><w:t>Inserted</w:t></w:r></w:ins></w:p>
  </w:body>
</w:document>`;
		const buffer = await createMinimalDocx(docXml);
		const doc = await loadTrackable(buffer);
		expect(doc.paragraphs[0]?.text).toBe('Inserted');
	});

	it('should handle paragraphs with numbering', async () => {
		const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>List item</w:t></w:r></w:p>
  </w:body>
</w:document>`;
		const buffer = await createMinimalDocx(docXml);
		const doc = await loadTrackable(buffer);
		expect(doc.paragraphs[0]?.numbering).toEqual({ level: 0, numId: '1' });
	});

	it('should handle paragraphs with styles', async () => {
		const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Heading</w:t></w:r></w:p>
  </w:body>
</w:document>`;
		const buffer = await createMinimalDocx(docXml);
		const doc = await loadTrackable(buffer);
		expect(doc.paragraphs[0]?.style).toBe('Heading1');
	});
});

describe('TrackableDocument', () => {
	beforeEach(() => {
		resetUsedIds();
	});

	describe('getText', () => {
		it('should return full document text', async () => {
			const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Line one</w:t></w:r></w:p>
    <w:p><w:r><w:t>Line two</w:t></w:r></w:p>
  </w:body>
</w:document>`;

			const buffer = await createMinimalDocx(docXml);
			const doc = await loadTrackable(buffer);

			expect(doc.getText()).toBe('Line one\nLine two');
		});
	});

	describe('getParagraph', () => {
		it('should find paragraph by ID', async () => {
			const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
            xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:body>
    <w:p w14:paraId="FINDME00" w14:textId="11111111">
      <w:r><w:t>Find me!</w:t></w:r>
    </w:p>
  </w:body>
</w:document>`;

			const buffer = await createMinimalDocx(docXml);
			const doc = await loadTrackable(buffer);

			const para = doc.getParagraph('body:FINDME00');
			expect(para).toBeDefined();
			expect(para?.text).toBe('Find me!');
		});

		it('should return undefined for unknown ID', async () => {
			const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Exists</w:t></w:r></w:p>
  </w:body>
</w:document>`;

			const buffer = await createMinimalDocx(docXml);
			const doc = await loadTrackable(buffer);

			expect(doc.getParagraph('body:NOTFOUND')).toBeUndefined();
		});
	});

	describe('getBuffer', () => {
		it('should return valid DOCX buffer', async () => {
			const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Test</w:t></w:r></w:p>
  </w:body>
</w:document>`;

			const buffer = await createMinimalDocx(docXml);
			const doc = await loadTrackable(buffer);

			const outputBuffer = await doc.getBuffer();

			// Should be a valid ZIP
			const zip = await JSZip.loadAsync(outputBuffer);
			expect(zip.file('word/document.xml')).toBeTruthy();
		});

		it('should include minted IDs in output', async () => {
			const docXml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>No ID originally</w:t></w:r></w:p>
  </w:body>
</w:document>`;

			const buffer = await createMinimalDocx(docXml);
			const doc = await loadTrackable(buffer);
			const mintedId = doc.paragraphs[0]?.id.split(':')[1]; // Get the hex part

			const outputBuffer = await doc.getBuffer();
			const zip = await JSZip.loadAsync(outputBuffer);
			const outputXml = await zip.file('word/document.xml')?.async('string');

			// The minted ID should appear in the output XML
			expect(outputXml).toContain(mintedId);
		});
	});
});
