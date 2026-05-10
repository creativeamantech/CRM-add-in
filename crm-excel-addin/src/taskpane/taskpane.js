// src/taskpane/taskpane.js

import "./taskpane.css";
import { getSession, login, logout } from "../auth/auth.js";
import { initAdminPanel } from "../admin/adminPanel.js";
import { initAgentView } from "../agent/agentView.js";

// Ensure Office.js is ready before proceeding
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.addEventListener("DOMContentLoaded", initApp);
  }
});

function initApp() {
  bindAuthEvents();
  checkAuthState();
}

function checkAuthState() {
  const session = getSession();
  const overlay = document.getElementById("login-overlay");
  const shell = document.getElementById("app-shell");

  if (!session) {
    overlay.hidden = false;
    shell.hidden = true;
  } else {
    overlay.hidden = true;
    shell.hidden = false;

    // Update Header
    document.getElementById("role-badge").innerText = session.role;
    document.getElementById("user-name").innerText = session.name;

    // Route based on role
    if (session.role === "admin") {
      initAdminPanel();
    } else if (session.role === "agent") {
      initAgentView();
    }
  }
}

function bindAuthEvents() {
  document.getElementById("btn-login").addEventListener("click", () => {
    const id = document.getElementById("login-id").value;
    const pin = document.getElementById("login-pin").value;
    const errEl = document.getElementById("login-error");

    errEl.hidden = true;

    if (!id || !pin) {
      errEl.innerText = "Please enter both ID and PIN.";
      errEl.hidden = false;
      return;
    }

    const result = login(id, pin);
    if (result.success) {
      document.getElementById("login-id").value = "";
      document.getElementById("login-pin").value = "";
      checkAuthState();
    } else {
      errEl.innerText = result.message;
      errEl.hidden = false;
    }
  });

  document.getElementById("btn-logout").addEventListener("click", () => {
    logout();
    checkAuthState();
  });
}
