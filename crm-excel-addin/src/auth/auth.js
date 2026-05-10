// src/auth/auth.js

import { USERS } from "./roleConfig.js";

/**
 * Checks if there's an active session in localStorage.
 * @returns {Object|null} session object or null
 */
export function getSession() {
  const session = localStorage.getItem("crmSession");
  if (session) {
    try {
      return JSON.parse(session);
    } catch (e) {
      console.error("Invalid session format in localStorage.");
      return null;
    }
  }
  return null;
}

/**
 * Authenticates a user against the hardcoded roleConfig.
 * @param {string} userId - The agent ID or admin ID (e.g., AGT001)
 * @param {string} pin - The PIN for the user
 * @returns {Object} { success: boolean, message?: string, user?: Object }
 */
export function login(userId, pin) {
  const upperId = userId.trim().toUpperCase();
  const user = USERS[upperId];

  if (!user) {
    return { success: false, message: "Invalid User ID." };
  }

  if (user.pin !== pin.trim()) {
    return { success: false, message: "Incorrect PIN." };
  }

  const sessionData = {
    userId: upperId,
    agentId: upperId, // aliased for convenience
    role: user.role,
    name: user.name
  };

  localStorage.setItem("crmSession", JSON.stringify(sessionData));

  return { success: true, user: sessionData };
}

/**
 * Clears the session from localStorage.
 */
export function logout() {
  localStorage.removeItem("crmSession");
}
