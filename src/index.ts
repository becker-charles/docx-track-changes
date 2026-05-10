// Public API exports

export { loadTrackable, loadTrackableFile } from './load.js';
export { TrackableDocument } from './TrackableDocument.js';

// Types
export type {
	Paragraph,
	Run,
	Edit,
	ReplaceParagraphEdit,
	InsertAfterEdit,
	InsertBeforeEdit,
	DeleteParagraphEdit,
	DeleteRangeEdit,
	ReplaceTextEdit,
	ContentRun,
	EditOptions,
	EditResult,
	AppliedEdit,
	FailedEdit,
	TrackedChange,
	ChangeResolution,
} from './types.js';

// JSON Schema for LLM integration
export { editSchema } from './schema.js';

// Error types
export { DocxError, DocxErrorCode } from './errors.js';
