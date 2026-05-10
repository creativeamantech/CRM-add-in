// src/utils/sheetHelper.js

/**
 * Ensures a sheet exists, creates it if it doesn't.
 * @param {Excel.RequestContext} context
 * @param {string} sheetName
 * @param {string[]} headers (optional)
 * @returns {Excel.Worksheet}
 */
export async function ensureSheetExists(context, sheetName, headers = null) {
  const sheets = context.workbook.worksheets;
  sheets.load("items/name");
  await context.sync();

  let sheet = sheets.items.find(s => s.name === sheetName);

  if (!sheet) {
    sheet = sheets.add(sheetName);
    if (headers && headers.length > 0) {
      const headerRange = sheet.getRangeByIndexes(0, 0, 1, headers.length);
      headerRange.values = [headers];
      headerRange.format.font.bold = true;
      headerRange.format.autofitColumns();
    }
    await context.sync();
  }

  return sheet;
}

/**
 * Displays a toast message by adding a div to the #toast-container.
 * @param {string} message
 * @param {string} type - "success" or "error"
 */
export function showToast(message, type = "success") {
  const container = document.getElementById("toast-container");
  if (!container) return;

  const toast = document.createElement("div");
  toast.className = `toast toast-${type}`;
  toast.innerText = message;

  container.appendChild(toast);

  // Trigger animation
  setTimeout(() => {
    toast.classList.add("show");
  }, 10);

  // Auto dismiss after 4 seconds
  setTimeout(() => {
    toast.classList.remove("show");
    setTimeout(() => {
      toast.remove();
    }, 300); // Wait for fade-out transition
  }, 4000);
}
