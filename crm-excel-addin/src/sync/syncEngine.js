// src/sync/syncEngine.js

import { MASTER_SHEET_NAME, COL_LOAN_ID, COL_AVAILABILITY, COL_STD_FEEDBACK, COL_DETAIL_FEEDBACK, COL_REF1_AVAIL, COL_REF1_FEEDBACK, COL_REF2_AVAIL, COL_REF2_FEEDBACK, COL_TYPE, COL_LAST_UPDATED, COL_UPDATED_BY, COL_HISTORY } from "../utils/constants.js";
import { acquireLock, releaseLock } from "./lockManager.js";
import { appendHistory } from "./historyLogger.js";

/**
 * Performs a two-phase write back to the master sheet and history sheet.
 * @param {number} dataIndex - 0-based data index
 * @param {Object} feedbackPayload
 */
export async function writeBack(dataIndex, feedbackPayload) {
  // Phase A: Lock check. Normally lock is already acquired when form opens,
  // but we double check / re-acquire to be safe before writing.
  await acquireLock(dataIndex, feedbackPayload.agentId);

  try {
    // Phase B: Atomic Write
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(MASTER_SHEET_NAME);
      // dataIndex is 0-based excluding header. So Excel Row = dataIndex + 2
      const rowNum = dataIndex + 2;

      // Load Q column to ensure we get up to index 16 (History string)
      const rowRange = sheet.getRange(`A${rowNum}:Q${rowNum}`);
      rowRange.load("values");
      await context.sync();

      const existing = rowRange.values[0];
      const loanId = existing[COL_LOAN_ID];

      // Mutate existing array
      existing[COL_AVAILABILITY] = feedbackPayload.availability;
      existing[COL_STD_FEEDBACK] = feedbackPayload.standardFeedback;
      existing[COL_DETAIL_FEEDBACK] = feedbackPayload.detailedFeedback;
      existing[COL_REF1_AVAIL] = feedbackPayload.ref1Availability;
      existing[COL_REF1_FEEDBACK] = feedbackPayload.ref1Feedback;
      existing[COL_REF2_AVAIL] = feedbackPayload.ref2Availability;
      existing[COL_REF2_FEEDBACK] = feedbackPayload.ref2Feedback;
      existing[COL_TYPE] = feedbackPayload.type;
      existing[COL_LAST_UPDATED] = feedbackPayload.timestamp;
      existing[COL_UPDATED_BY] = feedbackPayload.agentId;

      // History string mutation
      const newHistoryEntry = `${feedbackPayload.timestamp} (${feedbackPayload.agentId})`;
      const currentHistoryStr = existing[COL_HISTORY] || "";
      existing[COL_HISTORY] = currentHistoryStr ? `${currentHistoryStr}; ${newHistoryEntry}` : newHistoryEntry;

      // Write back to Master
      rowRange.values = [existing];

      // Append to History Sheet
      await appendHistory(context, dataIndex, loanId, feedbackPayload);

      await context.sync();
    });
  } finally {
    // Always release lock after write attempt (success or fail)
    await releaseLock(dataIndex);
  }
}
