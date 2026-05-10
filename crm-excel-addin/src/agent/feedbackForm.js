// src/agent/feedbackForm.js

import { COL_LOAN_ID, COL_CLIENT_NAME, COL_CITY, COL_AMOUNT, COL_PHONE, COL_AVAILABILITY, COL_STD_FEEDBACK, COL_DETAIL_FEEDBACK, COL_REF1_AVAIL, COL_REF1_FEEDBACK, COL_REF2_AVAIL, COL_REF2_FEEDBACK, COL_TYPE } from "../utils/constants.js";
import { getSession } from "../auth/auth.js";
import { writeBack } from "../sync/syncEngine.js";
import { releaseLock } from "../sync/lockManager.js";
import { showToast } from "../utils/sheetHelper.js";

let currentRowData = null;

/**
 * Opens the Case Detail Panel and populates it.
 * @param {Object} rowData { dataIndex, values }
 */
export function openFeedbackForm(rowData) {
  currentRowData = rowData;

  document.getElementById("agent-view").hidden = true;
  document.getElementById("case-detail-panel").hidden = false;

  // Populate Header
  const vals = rowData.values;
  document.getElementById("case-header-loan").innerText = vals[COL_LOAN_ID] || "N/A";
  document.getElementById("case-header-name").innerText = vals[COL_CLIENT_NAME] || "N/A";
  document.getElementById("case-header-info").innerText = `${vals[COL_CITY] || "-"} | ${vals[COL_AMOUNT] || "-"} | ${vals[COL_PHONE] || "-"}`;

  // Populate Form with existing values if any
  document.getElementById("ff-availability").value = vals[COL_AVAILABILITY] || "";
  document.getElementById("ff-std-feedback").value = vals[COL_STD_FEEDBACK] || "";
  document.getElementById("ff-detail-feedback").value = vals[COL_DETAIL_FEEDBACK] || "";
  document.getElementById("ff-ref1-avail").value = vals[COL_REF1_AVAIL] || "";
  document.getElementById("ff-ref1-fb").value = vals[COL_REF1_FEEDBACK] || "";
  document.getElementById("ff-ref2-avail").value = vals[COL_REF2_AVAIL] || "";
  document.getElementById("ff-ref2-fb").value = vals[COL_REF2_FEEDBACK] || "";
  document.getElementById("ff-type").value = vals[COL_TYPE] || "";

  clearValidationErrors();
}

/**
 * Validates the form and returns payload if valid.
 */
function validateForm() {
  clearValidationErrors();
  let isValid = true;

  const availability = document.getElementById("ff-availability").value;
  const stdFeedback = document.getElementById("ff-std-feedback").value;
  const detailFeedback = document.getElementById("ff-detail-feedback").value.trim();
  const ref1Avail = document.getElementById("ff-ref1-avail").value;
  const ref1Fb = document.getElementById("ff-ref1-fb").value.trim();
  const ref2Avail = document.getElementById("ff-ref2-avail").value;
  const ref2Fb = document.getElementById("ff-ref2-fb").value.trim();
  const type = document.getElementById("ff-type").value;

  if (!availability) {
    showError("ff-availability", "Availability is required.");
    isValid = false;
  }
  if (!stdFeedback) {
    showError("ff-std-feedback", "Standard Feedback is required.");
    isValid = false;
  }
  if (!type) {
    showError("ff-type", "Type is required.");
    isValid = false;
  }

  if (!isValid) return null;

  const session = getSession();

  return {
    availability,
    standardFeedback: stdFeedback,
    detailedFeedback: detailFeedback,
    ref1Availability: ref1Avail,
    ref1Feedback: ref1Fb,
    ref2Availability: ref2Avail,
    ref2Feedback: ref2Fb,
    type,
    timestamp: new Date().toISOString(),
    agentId: session.agentId
  };
}

function showError(fieldId, msg) {
  const errEl = document.getElementById(`${fieldId}-error`);
  if (errEl) {
    errEl.innerText = msg;
    errEl.hidden = false;
  }
}

function clearValidationErrors() {
  document.querySelectorAll(".ff-error").forEach(el => {
    el.innerText = "";
    el.hidden = true;
  });
}

/**
 * Event handler for submission
 */
export async function handleFeedbackSubmit() {
  const payload = validateForm();
  if (!payload || !currentRowData) return;

  const btn = document.getElementById("btn-submit-feedback");
  btn.disabled = true;
  btn.innerText = "Submitting...";

  try {
    await writeBack(currentRowData.dataIndex, payload);
    showToast("Feedback submitted successfully!", "success");
    // Do NOT clear form per requirement, keep visible
  } catch (error) {
    showToast(`Error: ${error.message}`, "error");
  } finally {
    btn.disabled = false;
    btn.innerText = "Submit Feedback";
  }
}

/**
 * Event handler for Back to List
 */
export async function closeFeedbackForm() {
  if (currentRowData) {
    await releaseLock(currentRowData.dataIndex);
    currentRowData = null;
  }

  document.getElementById("case-detail-panel").hidden = true;
  document.getElementById("agent-view").hidden = false;

  // Optionally refresh the list to show updated greyed out status
  // import { loadAgentCases } from "./agentView.js"; // circular dep mitigation via global or event
  // window.dispatchEvent(new Event('refreshAgentList'));
}

// Attach events when script loads
document.addEventListener("DOMContentLoaded", () => {
  const submitBtn = document.getElementById("btn-submit-feedback");
  if (submitBtn) submitBtn.addEventListener("click", handleFeedbackSubmit);

  const backBtn = document.getElementById("btn-back-to-list");
  if (backBtn) backBtn.addEventListener("click", async () => {
    await closeFeedbackForm();
    const { loadAgentCases } = await import("./agentView.js");
    await loadAgentCases();
  });
});
