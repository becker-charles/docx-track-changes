/**
 * Test helpers for creating realistic DOCX documents using the `docx` package.
 *
 * These helpers generate documents that closely match what Microsoft Word produces,
 * making tests more realistic than hand-crafted XML strings.
 */

import {
	Document,
	Packer,
	Paragraph,
	TextRun,
	HeadingLevel,
	Table,
	TableRow,
	TableCell,
	Header,
	Footer,
	PageNumber,
	NumberFormat,
	AlignmentType,
	WidthType,
	BorderStyle,
} from 'docx';

/**
 * Options for creating a simple test document
 */
export interface SimpleDocOptions {
	paragraphs: Array<string | ParagraphOptions>;
}

export interface ParagraphOptions {
	text: string;
	heading?: 'Heading1' | 'Heading2' | 'Heading3';
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
}

/**
 * Create a simple document with plain text paragraphs
 */
export async function createSimpleDoc(texts: string[]): Promise<Buffer> {
	const doc = new Document({
		sections: [
			{
				children: texts.map(
					text =>
						new Paragraph({
							children: [new TextRun(text)],
						})
				),
			},
		],
	});

	return Buffer.from(await Packer.toBuffer(doc));
}

/**
 * Create a document with formatted paragraphs
 */
export async function createFormattedDoc(
	paragraphs: Array<string | ParagraphOptions>
): Promise<Buffer> {
	const doc = new Document({
		sections: [
			{
				children: paragraphs.map(p => {
					if (typeof p === 'string') {
						return new Paragraph({
							children: [new TextRun(p)],
						});
					}

					const headingLevel = p.heading
						? HeadingLevel[p.heading.toUpperCase() as keyof typeof HeadingLevel]
						: undefined;

					return new Paragraph({
						heading: headingLevel,
						children: [
							new TextRun({
								text: p.text,
								bold: p.bold,
								italics: p.italic,
								underline: p.underline ? {} : undefined,
							}),
						],
					});
				}),
			},
		],
	});

	return Buffer.from(await Packer.toBuffer(doc));
}

/**
 * Create a document with mixed formatting within paragraphs
 */
export interface MixedRun {
	text: string;
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
}

export async function createMixedFormattingDoc(
	paragraphs: MixedRun[][]
): Promise<Buffer> {
	const doc = new Document({
		sections: [
			{
				children: paragraphs.map(
					runs =>
						new Paragraph({
							children: runs.map(
								run =>
									new TextRun({
										text: run.text,
										bold: run.bold,
										italics: run.italic,
										underline: run.underline ? {} : undefined,
									})
							),
						})
				),
			},
		],
	});

	return Buffer.from(await Packer.toBuffer(doc));
}

/**
 * Create a document with headers and footers
 */
export async function createDocWithHeaderFooter(options: {
	headerText?: string;
	footerText?: string;
	bodyParagraphs: string[];
}): Promise<Buffer> {
	const doc = new Document({
		sections: [
			{
				headers: options.headerText
					? {
							default: new Header({
								children: [
									new Paragraph({
										children: [new TextRun(options.headerText)],
									}),
								],
							}),
						}
					: undefined,
				footers: options.footerText
					? {
							default: new Footer({
								children: [
									new Paragraph({
										children: [
											new TextRun(options.footerText),
											new TextRun(' - Page '),
											new TextRun({
												children: [PageNumber.CURRENT],
											}),
										],
									}),
								],
							}),
						}
					: undefined,
				children: options.bodyParagraphs.map(
					text =>
						new Paragraph({
							children: [new TextRun(text)],
						})
				),
			},
		],
	});

	return Buffer.from(await Packer.toBuffer(doc));
}

/**
 * Create a document with a simple table
 */
export async function createDocWithTable(options: {
	beforeTable?: string[];
	tableData: string[][];
	afterTable?: string[];
}): Promise<Buffer> {
	const children: (Paragraph | Table)[] = [];

	// Paragraphs before table
	if (options.beforeTable) {
		for (const text of options.beforeTable) {
			children.push(
				new Paragraph({
					children: [new TextRun(text)],
				})
			);
		}
	}

	// Table
	const table = new Table({
		width: {
			size: 100,
			type: WidthType.PERCENTAGE,
		},
		rows: options.tableData.map(
			row =>
				new TableRow({
					children: row.map(
						cellText =>
							new TableCell({
								children: [
									new Paragraph({
										children: [new TextRun(cellText)],
									}),
								],
								borders: {
									top: { style: BorderStyle.SINGLE, size: 1 },
									bottom: { style: BorderStyle.SINGLE, size: 1 },
									left: { style: BorderStyle.SINGLE, size: 1 },
									right: { style: BorderStyle.SINGLE, size: 1 },
								},
							})
					),
				})
		),
	});
	children.push(table);

	// Paragraphs after table
	if (options.afterTable) {
		for (const text of options.afterTable) {
			children.push(
				new Paragraph({
					children: [new TextRun(text)],
				})
			);
		}
	}

	const doc = new Document({
		sections: [{ children }],
	});

	return Buffer.from(await Packer.toBuffer(doc));
}

/**
 * Create a document with numbered/bulleted lists
 */
export async function createDocWithList(options: {
	title?: string;
	items: string[];
	numbered?: boolean;
}): Promise<Buffer> {
	const children: Paragraph[] = [];

	if (options.title) {
		children.push(
			new Paragraph({
				children: [new TextRun(options.title)],
			})
		);
	}

	// Create list paragraphs
	for (let i = 0; i < options.items.length; i++) {
		children.push(
			new Paragraph({
				numbering: {
					reference: options.numbered ? 'numbered-list' : 'bullet-list',
					level: 0,
				},
				children: [new TextRun(options.items[i])],
			})
		);
	}

	const doc = new Document({
		numbering: {
			config: [
				{
					reference: 'numbered-list',
					levels: [
						{
							level: 0,
							format: NumberFormat.DECIMAL,
							text: '%1.',
							alignment: AlignmentType.LEFT,
						},
					],
				},
				{
					reference: 'bullet-list',
					levels: [
						{
							level: 0,
							format: NumberFormat.BULLET,
							text: '\u2022',
							alignment: AlignmentType.LEFT,
						},
					],
				},
			],
		},
		sections: [{ children }],
	});

	return Buffer.from(await Packer.toBuffer(doc));
}

/**
 * Extended run options including font properties
 */
export interface RichRunOptions {
	text: string;
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	font?: string;
	size?: number; // in half-points (e.g., 24 = 12pt)
	color?: string; // hex color without # (e.g., "FF0000")
}

/**
 * Create a document with rich formatting (font, size, color)
 */
export async function createRichFormattingDoc(
	paragraphs: RichRunOptions[][]
): Promise<Buffer> {
	const doc = new Document({
		sections: [
			{
				children: paragraphs.map(
					runs =>
						new Paragraph({
							children: runs.map(
								run =>
									new TextRun({
										text: run.text,
										bold: run.bold,
										italics: run.italic,
										underline: run.underline ? {} : undefined,
										font: run.font,
										size: run.size,
										color: run.color,
									})
							),
						})
				),
			},
		],
	});

	return Buffer.from(await Packer.toBuffer(doc));
}

/**
 * Create a legal-style document (common use case)
 */
export async function createLegalDoc(options: {
	title: string;
	parties: string[];
	clauses: Array<{ heading: string; text: string }>;
}): Promise<Buffer> {
	const children: Paragraph[] = [];

	// Title
	children.push(
		new Paragraph({
			heading: HeadingLevel.TITLE,
			alignment: AlignmentType.CENTER,
			children: [new TextRun({ text: options.title, bold: true })],
		})
	);

	// Parties
	children.push(
		new Paragraph({
			children: [new TextRun('')],
		})
	);
	for (const party of options.parties) {
		children.push(
			new Paragraph({
				children: [new TextRun(party)],
			})
		);
	}

	// Clauses
	for (const clause of options.clauses) {
		children.push(
			new Paragraph({
				children: [new TextRun('')],
			})
		);
		children.push(
			new Paragraph({
				heading: HeadingLevel.HEADING_1,
				children: [new TextRun({ text: clause.heading, bold: true })],
			})
		);
		children.push(
			new Paragraph({
				children: [new TextRun(clause.text)],
			})
		);
	}

	const doc = new Document({
		sections: [{ children }],
	});

	return Buffer.from(await Packer.toBuffer(doc));
}
