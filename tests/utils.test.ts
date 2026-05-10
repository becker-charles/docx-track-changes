/**
 * Tests for utility functions: errors, zip, parser, xml
 */
import { describe, it, expect, beforeEach } from 'vitest';
import JSZip from 'jszip';
import { writeFile, unlink } from 'node:fs/promises';
import { join } from 'node:path';
import { tmpdir } from 'node:os';
import { loadTrackableFile, DocxError } from '../src/index.js';
import { resetUsedIds, readParaIds, deduplicateIds, mintParaId } from '../src/readModel/ids.js';
import { parseXml, buildXml, getTagName, getTextContent, getChildren, getAttributes, setAttribute } from '../src/readModel/parser.js';
import { loadZip, hasPart, listParts } from '../src/utils/zip.js';
import { cloneElement, convertToDelText } from '../src/tracking/wrapDelete.js';
import { createSimpleDoc } from './helpers/createDocx.js';

describe('DocxError', () => {
	it('should create an error with code and message', () => {
		const error = new DocxError('INVALID_ARCHIVE', 'Test error message');
		expect(error).toBeInstanceOf(Error);
		expect(error.code).toBe('INVALID_ARCHIVE');
		expect(error.message).toBe('Test error message');
		expect(error.name).toBe('DocxError');
	});

	it('should support all error codes', () => {
		const codes: Array<'INVALID_ARCHIVE' | 'MISSING_PART' | 'INVALID_XML' | 'INVALID_EDIT'> = [
			'INVALID_ARCHIVE', 'MISSING_PART', 'INVALID_XML', 'INVALID_EDIT',
		];
		for (const code of codes) {
			const error = new DocxError(code, `Error with code ${code}`);
			expect(error.code).toBe(code);
		}
	});
});

describe('loadTrackableFile', () => {
	beforeEach(() => resetUsedIds());

	it('should load a DOCX file from disk', async () => {
		const buffer = await createSimpleDoc(['Test paragraph']);
		const tempPath = join(tmpdir(), `test-${Date.now()}.docx`);
		try {
			await writeFile(tempPath, buffer);
			const doc = await loadTrackableFile(tempPath);
			const testPara = doc.body.find(p => p.text === 'Test paragraph');
			expect(testPara).toBeDefined();
		} finally {
			await unlink(tempPath).catch(() => {});
		}
	});

	it('should throw when file does not exist', async () => {
		await expect(loadTrackableFile('/nonexistent/path.docx')).rejects.toThrow();
	});
});

describe('loadZip error handling', () => {
	it('should throw INVALID_ARCHIVE for non-ZIP data', async () => {
		const invalidData = Buffer.from('This is not a ZIP file');
		await expect(loadZip(invalidData)).rejects.toMatchObject({ code: 'INVALID_ARCHIVE' });
	});

	it('should throw MISSING_PART when document.xml is missing', async () => {
		const zip = new JSZip();
		zip.file('[Content_Types].xml', '<?xml version="1.0"?><Types></Types>');
		const buffer = await zip.generateAsync({ type: 'nodebuffer' });
		await expect(loadZip(buffer)).rejects.toMatchObject({ code: 'MISSING_PART' });
	});

	it('should throw MISSING_PART when [Content_Types].xml is missing', async () => {
		const zip = new JSZip();
		zip.file('word/document.xml', '<?xml version="1.0"?><w:document></w:document>');
		const buffer = await zip.generateAsync({ type: 'nodebuffer' });
		await expect(loadZip(buffer)).rejects.toMatchObject({ code: 'MISSING_PART' });
	});
});

describe('ZIP utility functions', () => {
	it('hasPart should return true/false correctly', async () => {
		const buffer = await createSimpleDoc(['Test']);
		const zip = await JSZip.loadAsync(buffer);
		expect(hasPart(zip, 'word/document.xml')).toBe(true);
		expect(hasPart(zip, 'nonexistent.xml')).toBe(false);
	});

	it('listParts should return all file paths', async () => {
		const buffer = await createSimpleDoc(['Test']);
		const zip = await JSZip.loadAsync(buffer);
		const parts = listParts(zip);
		expect(parts).toContain('word/document.xml');
		expect(parts).toContain('[Content_Types].xml');
	});
});

describe('XML parser utilities', () => {
	it('getTagName should return null for element without tag', () => {
		expect(getTagName({ ':@': { '@_attr': 'value' } })).toBeNull();
	});

	it('getTagName should return the tag name', () => {
		expect(getTagName({ 'w:p': [], ':@': {} })).toBe('w:p');
	});

	it('getChildren should return empty array for element without children', () => {
		expect(getChildren({ ':@': {} })).toEqual([]);
	});

	it('getTextContent should extract text from nested elements', () => {
		const xml = `<?xml version="1.0"?><w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:r><w:t>Hello </w:t></w:r><w:r><w:t>World</w:t></w:r></w:p>`;
		const parsed = parseXml(xml);
		const pElement = parsed.find((node: unknown) =>
			typeof node === 'object' && node !== null && 'w:p' in (node as Record<string, unknown>)
		) as Record<string, unknown>;
		expect(getTextContent(pElement)).toBe('Hello World');
	});

	it('getAttributes should return empty object when no attributes', () => {
		expect(getAttributes({ 'w:p': [] })).toEqual({});
	});

	it('setAttribute should create :@ if not present', () => {
		const element: Record<string, unknown> = { 'w:p': [] };
		setAttribute(element, 'w:val', 'test');
		expect((element[':@'] as Record<string, string>)['@_w:val']).toBe('test');
	});

	it('buildXml should produce valid output', () => {
		const xml = '<?xml version="1.0"?><root><child>text</child></root>';
		const parsed = parseXml(xml);
		const rebuilt = buildXml(parsed);
		const reparsed = parseXml(rebuilt);
		expect(reparsed).toEqual(parsed);
	});
});

describe('ID functions', () => {
	beforeEach(() => resetUsedIds());

	it('mintParaId should generate unique 8-character hex IDs', () => {
		const id1 = mintParaId();
		const id2 = mintParaId();
		expect(id1).toMatch(/^[A-F0-9]{8}$/);
		expect(id2).toMatch(/^[A-F0-9]{8}$/);
		expect(id1).not.toBe(id2);
	});

	it('readParaIds should extract IDs from parsed XML', () => {
		const xml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="ABCD1234"><w:r><w:t>Test</w:t></w:r></w:p></w:body></w:document>`;
		const parsed = parseXml(xml);
		const ids = readParaIds(parsed);
		expect(Array.from(ids.values())).toContain('ABCD1234');
	});

	it('deduplicateIds should handle duplicate IDs', () => {
		const xml = `<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"><w:body><w:p w14:paraId="SAME1234"><w:r><w:t>First</w:t></w:r></w:p><w:p w14:paraId="SAME1234"><w:r><w:t>Second</w:t></w:r></w:p></w:body></w:document>`;
		const parsed = parseXml(xml);
		const remapped = deduplicateIds(parsed);
		expect(remapped.has('SAME1234')).toBe(true);
	});
});

describe('cloneElement', () => {
	it('should deep clone an element with children', () => {
		const original = {
			'w:r': [{ 'w:t': [{ '#text': 'Hello' }] }],
			':@': { '@_w:val': 'test' },
		};
		const cloned = cloneElement(original);
		expect(cloned).not.toBe(original);
		expect(cloned['w:r']).not.toBe(original['w:r']);
		expect(cloned[':@']).toEqual({ '@_w:val': 'test' });
	});

	it('should handle text nodes in children', () => {
		const original = { 'w:t': [{ '#text': 'Content' }] };
		const cloned = cloneElement(original);
		expect((cloned['w:t'] as unknown[])[0]).toEqual({ '#text': 'Content' });
	});
});

describe('convertToDelText', () => {
	it('should convert w:t to w:delText', () => {
		const run = { 'w:r': [{ 'w:t': [{ '#text': 'Hello' }] }] };
		const converted = convertToDelText(run);
		const children = converted['w:r'] as unknown[];
		expect((children[0] as Record<string, unknown>)['w:delText']).toBeDefined();
	});

	it('should preserve attributes when converting', () => {
		const run = { 'w:r': [{ 'w:t': [{ '#text': 'Hello' }], ':@': { '@_xml:space': 'preserve' } }] };
		const converted = convertToDelText(run);
		const textNode = (converted['w:r'] as unknown[])[0] as Record<string, unknown>;
		expect(textNode[':@']).toEqual({ '@_xml:space': 'preserve' });
	});
});
