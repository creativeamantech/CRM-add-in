// src/admin/adminPanel.js

import { MASTER_SHEET_NAME, COL_LOAN_ID, COL_CLIENT_NAME, COL_CITY, COL_AMOUNT, COL_PHONE, COL_AGENT_ID } from "../utils/constants.js";
import { USERS } from "../auth/roleConfig.js";
import { assignSingle, assignBulk, assignByFilter } from "./assignmentEngine.js";
import { showToast } from "../utils/sheetHelper.js";

let allRowsCache = [];
let currentDisplayIndex = 0;
const CHUNK_SIZE = 50;

/**
 * Initializes the Admin Panel UI
 */
export async function initAdminPanel() {
  document.getElementById("admin-panel").hidden = false;
  document.getElementById("agent-view").hidden = true;

  // Setup Tabs
  const tabs = document.querySelectorAll(".admin-tab");
  tabs.forEach(tab => {
    tab.addEventListener("click", (e) => {
      // Deactivate all
      tabs.forEach(t => t.classList.remove("active"));
      document.querySelectorAll(".admin-tab-content").forEach(c => c.hidden = true);

      // Activate clicked
      const targetId = e.target.getAttribute("data-target");
      e.target.classList.add("active");
      document.getElementById(targetId).hidden = false;
    });
  });

  // Populate Dropdowns
  populateAgentDropdowns();

  // Load Data
  await loadDataFromSheet();

  // Bind Bulk Assign Event
  document.getElementById("btn-bulk-assign").addEventListener("click", async () => {
    const checkboxes = document.querySelectorAll(".bulk-row-checkbox:checked");
    const selectedIndices = Array.from(checkboxes).map(cb => parseInt(cb.getAttribute("data-index")));
    const agentId = document.getElementById("bulk-agent-select").value;

    if (!agentId) return showToast("Select an agent first.", "error");

    await assignBulk(selectedIndices, agentId);
    await loadDataFromSheet(); // refresh view
  });

  // Bind Select All Event
  document.getElementById("bulk-select-all").addEventListener("change", (e) => {
    const isChecked = e.target.checked;
    document.querySelectorAll(".bulk-row-checkbox").forEach(cb => cb.checked = isChecked);
  });

  // Bind Filter Preview
  document.getElementById("btn-filter-preview").addEventListener("click", () => {
    const colIndex = parseInt(document.getElementById("filter-col-select").value);
    const matchVal = document.getElementById("filter-val-input").value.toLowerCase().trim();
    if (!matchVal) return showToast("Enter a filter value.", "error");

    const matchCount = allRowsCache.filter(row => {
      const val = row.values[colIndex];
      return val && val.toString().toLowerCase().trim() === matchVal;
    }).length;

    document.getElementById("filter-preview-result").innerText = `Found ${matchCount} matching rows.`;
  });

  // Bind Filter Apply
  document.getElementById("btn-filter-apply").addEventListener("click", async () => {
    const colIndex = parseInt(document.getElementById("filter-col-select").value);
    const matchVal = document.getElementById("filter-val-input").value.trim();
    const agentId = document.getElementById("filter-agent-select").value;

    if (!matchVal) return showToast("Enter a filter value.", "error");
    if (!agentId) return showToast("Select an agent first.", "error");

    await assignByFilter({ colIndex, matchValue: matchVal }, agentId);
    await loadDataFromSheet(); // refresh view
  });

  // Bind Load More button
  document.getElementById("btn-load-more").addEventListener("click", () => renderNextChunk());
}

/**
 * Loads data from Excel and prepares the cache.
 */
async function loadDataFromSheet() {
  try {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(MASTER_SHEET_NAME);
      const usedRange = sheet.getUsedRange();
      usedRange.load("values, rowCount");
      await context.sync();

      allRowsCache = [];
      if (usedRange.rowCount > 1) {
        // Skip header (index 0)
        for (let i = 1; i < usedRange.rowCount; i++) {
          allRowsCache.push({
            dataIndex: i - 1, // 0-based data index
            values: usedRange.values[i]
          });
        }
      }
    });

    currentDisplayIndex = 0;
    document.getElementById("table-full-body").innerHTML = "";
    document.getElementById("table-bulk-body").innerHTML = "";
    document.getElementById("bulk-select-all").checked = false;

    renderNextChunk();
  } catch (error) {
    console.error(error);
    showToast(`Error loading data: ${error.message}`, "error");
  }
}

/**
 * Renders the next 50 rows into the UI tables.
 */
function renderNextChunk() {
  if (currentDisplayIndex >= allRowsCache.length) {
    document.getElementById("btn-load-more").hidden = true;
    return;
  }

  const chunk = allRowsCache.slice(currentDisplayIndex, currentDisplayIndex + CHUNK_SIZE);
  const fullBody = document.getElementById("table-full-body");
  const bulkBody = document.getElementById("table-bulk-body");

  chunk.forEach(row => {
    const trFull = document.createElement("tr");
    trFull.innerHTML = `
      <td>${row.values[COL_LOAN_ID] || ""}</td>
      <td>${row.values[COL_CLIENT_NAME] || ""}</td>
      <td>${row.values[COL_CITY] || ""}</td>
      <td>
        <select class="single-assign-select" data-index="${row.dataIndex}">
          <option value="">Unassigned</option>
          ${getAgentOptionsHTML(row.values[COL_AGENT_ID])}
        </select>
      </td>
    `;
    fullBody.appendChild(trFull);

    const trBulk = document.createElement("tr");
    trBulk.innerHTML = `
      <td><input type="checkbox" class="bulk-row-checkbox" data-index="${row.dataIndex}"></td>
      <td>${row.values[COL_LOAN_ID] || ""}</td>
      <td>${row.values[COL_CLIENT_NAME] || ""}</td>
      <td>${row.values[COL_AGENT_ID] || "Unassigned"}</td>
    `;
    bulkBody.appendChild(trBulk);
  });

  // Attach change event for single assigns
  fullBody.querySelectorAll(".single-assign-select").forEach(select => {
    // Only attach once
    if (select.dataset.bound === "true") return;
    select.dataset.bound = "true";

    select.addEventListener("change", async (e) => {
      const agentId = e.target.value;
      const dataIndex = parseInt(e.target.getAttribute("data-index"));
      if (agentId) {
        await assignSingle(dataIndex, agentId);
      }
    });
  });

  currentDisplayIndex += CHUNK_SIZE;
  document.getElementById("btn-load-more").hidden = currentDisplayIndex >= allRowsCache.length;
}

/**
 * Returns HTML string of <option> tags for agents
 */
function getAgentOptionsHTML(selectedId) {
  let html = "";
  for (const [id, user] of Object.entries(USERS)) {
    if (user.role === "agent") {
      const selected = (selectedId === id) ? "selected" : "";
      html += `<option value="${id}" ${selected}>${id} - ${user.name}</option>`;
    }
  }
  return html;
}

/**
 * Populates global agent dropdowns
 */
function populateAgentDropdowns() {
  const options = `<option value="">Select Agent...</option>` + getAgentOptionsHTML("");
  document.getElementById("bulk-agent-select").innerHTML = options;
  document.getElementById("filter-agent-select").innerHTML = options;
}
