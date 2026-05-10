// src/admin/assignmentEngine.js

import { MASTER_SHEET_NAME, COL_AGENT_ID } from "../utils/constants.js";
import { showToast } from "../utils/sheetHelper.js";

/**
 * Assigns a single row to an agent.
 * @param {number} rowIndex - 0-based data row index (excludes header)
 * @param {string} agentId
 */
export async function assignSingle(rowIndex, agentId) {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(MASTER_SHEET_NAME);
      // Data starts at row 2 in Excel (index 1), so rowIndex + 1 is the API row index
      const cell = sheet.getCell(rowIndex + 1, COL_AGENT_ID);
      cell.values = [[agentId]];
      await context.sync();
    });
    showToast(`Row ${rowIndex + 1} assigned to ${agentId}`, "success");
  } catch (error) {
    console.error(error);
    showToast(`Failed to assign row: ${error.message}`, "error");
  }
}

/**
 * Assigns multiple rows to an agent.
 * @param {number[]} rowIndices - Array of 0-based data row indices
 * @param {string} agentId
 */
export async function assignBulk(rowIndices, agentId) {
  if (!rowIndices || rowIndices.length === 0) {
    showToast("No rows selected.", "error");
    return;
  }

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(MASTER_SHEET_NAME);
      for (const rowIndex of rowIndices) {
        const cell = sheet.getCell(rowIndex + 1, COL_AGENT_ID);
        cell.values = [[agentId]];
      }
      await context.sync();
    });
    showToast(`Assigned ${rowIndices.length} rows to ${agentId}`, "success");
  } catch (error) {
    console.error(error);
    showToast(`Failed to bulk assign: ${error.message}`, "error");
  }
}

/**
 * Assigns rows based on a simple column value filter.
 * @param {Object} filterConfig - { colIndex: number, matchValue: string }
 * @param {string} agentId
 */
export async function assignByFilter(filterConfig, agentId) {
  if (!filterConfig || filterConfig.colIndex === undefined || !filterConfig.matchValue) {
    showToast("Invalid filter configuration.", "error");
    return;
  }

  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(MASTER_SHEET_NAME);
      const usedRange = sheet.getUsedRange();
      usedRange.load("values, rowCount");
      await context.sync();

      if (usedRange.rowCount <= 1) {
        throw new Error("No data found.");
      }

      const matchVal = filterConfig.matchValue.toString().toLowerCase().trim();
      let matchCount = 0;

      // Skip header row
      for (let i = 1; i < usedRange.rowCount; i++) {
        const rowValue = usedRange.values[i][filterConfig.colIndex];
        if (rowValue && rowValue.toString().toLowerCase().trim() === matchVal) {
          // Assign agent
          const cell = sheet.getCell(i, COL_AGENT_ID);
          cell.values = [[agentId]];
          matchCount++;
        }
      }

      if (matchCount > 0) {
        await context.sync();
        showToast(`Assigned ${matchCount} matching rows to ${agentId}`, "success");
      } else {
        showToast(`No rows matched the filter.`, "error");
      }
    });
  } catch (error) {
    console.error(error);
    showToast(`Failed to assign by filter: ${error.message}`, "error");
  }
}
