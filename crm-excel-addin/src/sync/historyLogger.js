// src/sync/historyLogger.js

import { HISTORY_SHEET_NAME } from "../utils/constants.js";
import { ensureSheetExists } from "../utils/sheetHelper.js";

/**
 * Appends a history row to the FeedbackHistory sheet.
 * Always appends, never overwrites.
 *
 * @param {Excel.RequestContext} context
 * @param {number} dataIndex - 0-based data index from MasterData
 * @param {string} loanId
 * @param {Object} payload
 */
export async function appendHistory(context, dataIndex, loanId, payload) {
  const headers = [
    "Loan ID", "Row Index", "Agent ID", "Timestamp",
    "Availability", "Std Feedback", "Detail",
    "REF1 Avail", "REF1 FB", "REF2 Avail", "REF2 FB", "Type"
  ];

  const sheet = await ensureSheetExists(context, HISTORY_SHEET_NAME, headers);
  const usedRange = sheet.getUsedRange();
  usedRange.load("rowCount");
  await context.sync();

  const nextRow = usedRange.rowCount > 0 ? usedRange.rowCount : 1;
  const targetRange = sheet.getRangeByIndexes(nextRow, 0, 1, 12);

  targetRange.values = [[
    loanId,
    dataIndex,
    payload.agentId,
    payload.timestamp,
    payload.availability,
    payload.standardFeedback,
    payload.detailedFeedback,
    payload.ref1Availability,
    payload.ref1Feedback,
    payload.ref2Availability,
    payload.ref2Feedback,
    payload.type
  ]];

  // No need for context.sync() here as the parent writeBack function does it
}
