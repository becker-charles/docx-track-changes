// Tracked changes utilities
export {
	wrapInsert,
	wrapParagraphInsert,
	createInsElement,
	createRunElement,
	createParagraphElement,
	formatDateForWord,
} from './wrapInsert.js';
export {
	wrapDelete,
	wrapParagraphContentForDelete,
	createDelElement,
	convertToDelText,
	cloneElement,
} from './wrapDelete.js';
export {
	generateChangeId,
	getNextChangeId,
	resetChangeIds,
	initializeChangeIds,
	findExistingChangeIds,
	initializeChangeIdsFromDocument,
} from './changeIds.js';
export {
	getTrackedChangesFromPart,
	acceptChange,
	rejectChange,
	resolveChanges,
	acceptAllChangesInPart,
	rejectAllChangesInPart,
} from './resolve.js';
