// src/sync/lockManager.js

import { LOCK_SHEET_NAME } from "../utils/constants.js";
import { ensureSheetExists } from "../utils/sheetHelper.js";

/**
 * Acquires a soft lock on a row for a specific agent.
 * @param {number} rowIndex
 * @param {string} agentId
 * @returns {Promise<boolean>} true if lock acquired, throws error if locked
 */
export async function acquireLock(rowIndex, agentId) {
  let locked = false;

  await Excel.run(async (context) => {
    const sheet = await ensureSheetExists(context, LOCK_SHEET_NAME, ["Row Index", "Agent ID", "Lock Time"]);
    const usedRange = sheet.getUsedRange();
    usedRange.load("values, rowCount");
    await context.sync();

    const now = new Date().getTime();
    let lockFound = false;

    if (usedRange.rowCount > 1) {
      for (let i = 1; i < usedRange.rowCount; i++) {
        const row = usedRange.values[i];
        if (row[0] === rowIndex) {
          lockFound = true;
          const lockedBy = row[1];
          const lockTime = new Date(row[2]).getTime();

          // Check if lock is stale (older than 5 min)
          if (now - lockTime < 5 * 60 * 1000) {
            if (lockedBy !== agentId) {
              throw new Error("Row is being updated by another agent. Please try again in a moment.");
            }
          } else {
            // Overwrite stale lock
            const cellAgent = sheet.getCell(i, 1);
            const cellTime = sheet.getCell(i, 2);
            cellAgent.values = [[agentId]];
            cellTime.values = [[new Date().toISOString()]];
            locked = true;
          }
          break;
        }
      }
    }

    if (!lockFound && !locked) {
      // Append new lock
      const newRow = [[rowIndex, agentId, new Date().toISOString()]];
      const nextRowIdx = usedRange.rowCount > 0 ? usedRange.rowCount : 1;
      const range = sheet.getRangeByIndexes(nextRowIdx, 0, 1, 3);
      range.values = newRow;
      locked = true;
    }

    await context.sync();
  });

  return locked;
}

/**
 * Releases a lock on a row.
 * @param {number} rowIndex
 */
export async function releaseLock(rowIndex) {
  await Excel.run(async (context) => {
    const sheet = await ensureSheetExists(context, LOCK_SHEET_NAME, ["Row Index", "Agent ID", "Lock Time"]);
    const usedRange = sheet.getUsedRange();
    usedRange.load("values, rowCount");
    await context.sync();

    if (usedRange.rowCount > 1) {
      for (let i = 1; i < usedRange.rowCount; i++) {
        if (usedRange.values[i][0] === rowIndex) {
          // Clear lock by deleting row or nulling it out.
          // For safety of indices, we null out the values
          const range = sheet.getRangeByIndexes(i, 0, 1, 3);
          range.values = [["", "", ""]];
          break;
        }
      }
    }
    await context.sync();
  });
}

/**
 * Auto-release locks older than maxAgeMinutes
 * @param {number} maxAgeMinutes
 */
export async function cleanStaleLocks(maxAgeMinutes = 5) {
  try {
    await Excel.run(async (context) => {
      const sheet = await ensureSheetExists(context, LOCK_SHEET_NAME, ["Row Index", "Agent ID", "Lock Time"]);
      const usedRange = sheet.getUsedRange();
      usedRange.load("values, rowCount");
      await context.sync();

      if (usedRange.rowCount > 1) {
        const now = new Date().getTime();
        for (let i = 1; i < usedRange.rowCount; i++) {
          const lockTimeStr = usedRange.values[i][2];
          if (lockTimeStr) {
            const lockTime = new Date(lockTimeStr).getTime();
            if (now - lockTime > maxAgeMinutes * 60 * 1000) {
              const range = sheet.getRangeByIndexes(i, 0, 1, 3);
              range.values = [["", "", ""]];
            }
          }
        }
      }
      await context.sync();
    });
  } catch (e) {
    console.error("Failed to clean stale locks", e);
  }
}
