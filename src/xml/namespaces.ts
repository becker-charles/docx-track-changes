/**
 * OOXML namespace constants
 */

/** Main WordprocessingML namespace */
export const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main';

/** Word 2010 namespace (for paraId, textId) */
export const W14_NS = 'http://schemas.microsoft.com/office/word/2010/wordml';

/** All OOXML namespaces used in Word documents */
export const NAMESPACES = {
	w: W_NS,
	w14: W14_NS,
	r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
	wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
	a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
	pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
	mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
} as const;
