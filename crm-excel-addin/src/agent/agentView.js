// src/agent/agentView.js

import { MASTER_SHEET_NAME, COL_LOAN_ID, COL_CLIENT_NAME, COL_CITY, COL_AMOUNT, COL_PHONE, COL_AGENT_ID, COL_LAST_UPDATED } from "../utils/constants.js";
import { getSession } from "../auth/auth.js";
import { showToast } from "../utils/sheetHelper.js";
import { openFeedbackForm } from "./feedbackForm.js";
import { acquireLock } from "../sync/lockManager.js";

/**
 * Initializes the Agent View UI
 */
export async function initAgentView() {
  document.getElementById("admin-panel").hidden = true;
  document.getElementById("agent-view").hidden = false;
  document.getElementById("case-detail-panel").hidden = true;

  await loadAgentCases();

  // Refresh button
  const refreshBtn = document.getElementById("btn-agent-refresh");
  if (refreshBtn) {
    refreshBtn.addEventListener("click", loadAgentCases);
  }
}

/**
 * Loads cases assigned to the current agent
 */
export async function loadAgentCases() {
  const session = getSession();
  if (!session || session.role !== "agent") return;

  const listContainer = document.getElementById("agent-card-list");
  listContainer.innerHTML = "<p>Loading cases...</p>";

  try {
    let assignedRows = [];

    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getItem(MASTER_SHEET_NAME);
      const usedRange = sheet.getUsedRange();
      usedRange.load("values, rowCount");
      await context.sync();

      if (usedRange.rowCount > 1) {
        for (let i = 1; i < usedRange.rowCount; i++) {
          const rowValues = usedRange.values[i];
          if (rowValues[COL_AGENT_ID] && rowValues[COL_AGENT_ID].toString().toUpperCase() === session.agentId) {
            assignedRows.push({
              dataIndex: i - 1, // 0-based data index (row index minus header)
              values: rowValues
            });
          }
        }
      }
    });

    renderCards(assignedRows);
  } catch (error) {
    console.error(error);
    showToast(`Error loading cases: ${error.message}`, "error");
    listContainer.innerHTML = "<p>Error loading cases.</p>";
  }
}

/**
 * Renders the list of cards
 */
function renderCards(rows) {
  const listContainer = document.getElementById("agent-card-list");
  listContainer.innerHTML = "";

  if (rows.length === 0) {
    listContainer.innerHTML = "<p>No cases assigned to you.</p>";
    return;
  }

  rows.forEach(row => {
    const lastUpdate = row.values[COL_LAST_UPDATED] ? new Date(row.values[COL_LAST_UPDATED]).toLocaleString() : "Never";

    const card = document.createElement("div");
    card.className = "fluent-card";
    card.innerHTML = `
      <div class="card-header">
        <strong>${row.values[COL_LOAN_ID] || "N/A"}</strong>
      </div>
      <div class="card-body">
        <p><strong>Name:</strong> ${row.values[COL_CLIENT_NAME] || "N/A"}</p>
        <p><strong>Info:</strong> ${row.values[COL_CITY] || "-"} | ${row.values[COL_AMOUNT] || "-"} | ${row.values[COL_PHONE] || "-"}</p>
        <p class="text-muted"><small>Last Update: ${lastUpdate}</small></p>
      </div>
      <div class="card-footer">
        <button class="fluent-btn primary btn-view-case" data-index="${row.dataIndex}">View & Update</button>
      </div>
    `;
    listContainer.appendChild(card);
  });

  // Attach events
  document.querySelectorAll(".btn-view-case").forEach(btn => {
    btn.addEventListener("click", async (e) => {
      const btnEl = e.target;
      const dataIndex = parseInt(btnEl.getAttribute("data-index"));
      const session = getSession();

      btnEl.disabled = true;
      btnEl.innerText = "Loading...";

      try {
        // Try to acquire application-level lock
        const locked = await acquireLock(dataIndex, session.agentId);
        if (locked) {
          const rowData = rows.find(r => r.dataIndex === dataIndex);
          openFeedbackForm(rowData);
        }
      } catch (error) {
        showToast(error.message, "error");
      } finally {
        btnEl.disabled = false;
        btnEl.innerText = "View & Update";
      }
    });
  });
}
