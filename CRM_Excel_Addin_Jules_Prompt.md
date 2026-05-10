# Jules Agent Prompt: Enterprise Multi-User CRM Add-in for Microsoft 365 Excel

---

## 🧠 Agent Context & Mission

You are **Jules**, an autonomous AI software engineer. Your mission is to fully design, scaffold, and implement a **production-ready Microsoft 365 Excel Task Pane Add-in** that transforms Excel into a synchronized, multi-user CRM system. This is an enterprise-grade build — not a prototype. Every file you generate must be complete, functional, and immediately deployable without manual edits.

You will work **phase by phase**. After completing each phase, pause and summarize what was built, what decisions were made, and what comes next. Do not skip phases or merge them.

---

## 📁 Required Project Structure

Generate the following directory tree exactly:

```
crm-excel-addin/
├── manifest.xml
├── package.json
├── webpack.config.js
├── .env.example
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html
│   │   ├── taskpane.css
│   │   └── taskpane.js
│   ├── auth/
│   │   ├── auth.js               # Role resolution & session management
│   │   └── roleConfig.js         # Agent ID → Name mapping registry
│   ├── admin/
│   │   ├── adminPanel.js         # Admin UI logic
│   │   └── assignmentEngine.js   # Individual, bulk & filter-based allocation
│   ├── agent/
│   │   ├── agentView.js          # Filtered row rendering per agent
│   │   └── feedbackForm.js       # Feedback form renderer & validator
│   ├── sync/
│   │   ├── syncEngine.js         # Office.js write-back & conflict handler
│   │   ├── historyLogger.js      # Audit trail & append logic
│   │   └── lockManager.js        # Row-level soft lock for concurrency
│   └── utils/
│       ├── sheetHelper.js        # Utility wrappers for Office.js sheet ops
│       └── constants.js          # Column indices, sheet names, enums
├── assets/
│   └── icon-32.png
│       icon-80.png
└── README.md
```

---

## 🔐 Phase 1 — Authentication & Role Resolution

**File:** `src/auth/auth.js`

Implement a role-resolution system using **Office.js `Office.context.mailbox` or a custom login modal** (since Excel does not natively provide SSO identity for all tenants). The system must:

- On add-in load, check `localStorage` for a saved session token (`crmSession`).
- If none found, display a **login modal overlay** (not a page redirect) with:
  - An `Agent ID` input field (e.g., `AGT001` through `AGT010`).
  - A `PIN` or `Password` field.
  - A "Login" CTA button.
- Validate credentials against a **hardcoded config object** in `roleConfig.js` (to be replaceable by an API call later).
- After successful login:
  - Persist `{ agentId, role, name }` to `localStorage` as `crmSession`.
  - Conditionally render either the **Admin Panel** (if `role === 'admin'`) or the **Agent View** (if `role === 'agent'`).
- Provide a visible **Logout** button that clears `crmSession` and resets the UI.

**File:** `src/auth/roleConfig.js`

```js
// Example structure — Jules must populate with 10 agents + 1 admin
export const USERS = {
  ADMIN01: { name: "Admin", role: "admin", pin: "admin@123" },
  AGT001:  { name: "Ravi Kumar",    role: "agent", pin: "1234" },
  AGT002:  { name: "Priya Sharma",  role: "agent", pin: "1234" },
  // ... up to AGT010
};
```

---

## 🛠️ Phase 2 — Master Sheet Schema & Constants

**File:** `src/utils/constants.js`

Define all column indices (0-based) for the master Excel sheet. Jules must assume the following schema and make all column references use these constants (never hardcoded numbers):

| Column | Field Name          | Constant Name        |
|--------|---------------------|----------------------|
| 0      | Loan ID             | `COL_LOAN_ID`        |
| 1      | Client Name         | `COL_CLIENT_NAME`    |
| 2      | City                | `COL_CITY`           |
| 3      | Amount              | `COL_AMOUNT`         |
| 4      | Phone               | `COL_PHONE`          |
| 5      | Assigned Agent ID   | `COL_AGENT_ID`       |
| 6      | Availability        | `COL_AVAILABILITY`   |
| 7      | Standard Feedback   | `COL_STD_FEEDBACK`   |
| 8      | Detailed Feedback   | `COL_DETAIL_FEEDBACK`|
| 9      | REF 1 Availability  | `COL_REF1_AVAIL`     |
| 10     | REF 1 Feedback      | `COL_REF1_FEEDBACK`  |
| 11     | REF 2 Availability  | `COL_REF2_AVAIL`     |
| 12     | REF 2 Feedback      | `COL_REF2_FEEDBACK`  |
| 13     | Type                | `COL_TYPE`           |
| 14     | Last Updated        | `COL_LAST_UPDATED`   |
| 15     | Last Updated By     | `COL_UPDATED_BY`     |
| 16     | Feedback History    | `COL_HISTORY`        |

Also define:
- `MASTER_SHEET_NAME = "MasterData"`
- `HISTORY_SHEET_NAME = "FeedbackHistory"`
- `LOCK_SHEET_NAME = "RowLocks"`

---

## 👑 Phase 3 — Admin Panel

**File:** `src/admin/adminPanel.js`

The Admin Panel must render inside the Task Pane (not a new window). Build a tabbed UI with three tabs:

### Tab 1: Full Data View
- Load the entire `MasterData` sheet using `Excel.run()`.
- Render all rows as a **virtualized scrollable table** (render max 50 rows at a time with a "Load More" button — do not load all rows into the DOM at once).
- Each row shows: Loan ID, Client Name, City, Amount, Phone, and a **dropdown to assign an agent**.
- Changing the dropdown immediately calls `assignmentEngine.assignSingle(rowIndex, agentId)`.

### Tab 2: Bulk Assignment
- Render the same table but with a **checkbox column** on the left.
- A "Select All" master checkbox at the top.
- A fixed-bottom action bar containing:
  - A **"Assign To" agent dropdown**.
  - A **"Assign Selected" button**.
- On click, call `assignmentEngine.assignBulk(selectedRowIndices, agentId)`.

### Tab 3: Filter-Based Bulk Assignment
- A **dynamic filter builder** with:
  - A "Filter By" column selector (e.g., City, Amount Range, etc.)
  - A value input or range input.
  - A "Preview Matches" button — shows a count of matching rows.
  - An "Assign All Matches To" agent dropdown.
  - An "Apply Assignment" CTA.
- Call `assignmentEngine.assignByFilter(filterConfig, agentId)`.

**File:** `src/admin/assignmentEngine.js`

Implement three exported async functions:

```js
export async function assignSingle(rowIndex, agentId) { ... }
export async function assignBulk(rowIndices, agentId) { ... }
export async function assignByFilter(filterConfig, agentId) { ... }
```

All three must use `Excel.run()` with `context.sync()` and handle errors gracefully with user-visible toast notifications.

---

## 👤 Phase 4 — Agent View

**File:** `src/agent/agentView.js`

On login as an agent:

1. Use `Excel.run()` to read the entire `MasterData` sheet.
2. **Filter rows client-side** where `COL_AGENT_ID === currentUser.agentId`.
3. Render only matched rows as a **card list** (not a table). Each card shows:
   - Loan ID (bold, primary)
   - Client Name
   - City | Amount | Phone
   - Last feedback submitted (if any), greyed out
   - A **"View & Update"** button
4. Clicking "View & Update" opens the **Case Info Panel** within the Task Pane:
   - A non-closable header strip showing: Loan ID, Client Name, City, Amount, Phone.
   - Below: the full Feedback Form.
   - A "Back to List" link at the top.

---

## 📝 Phase 5 — Feedback Form (Strict Schema)

**File:** `src/agent/feedbackForm.js`

Render a form with **exactly these 8 fields in this order**. Do not alter field options or input types under any circumstance:

```
1. Availability              → <select>: Yes | No | Third Party
2. Standard Feedback         → <select>: Paid | RNR | PTP | CB | RTP | Non Contactable |
                                          Multiple PTP | Already Settled | Disputed Case |
                                          Loan Closed | Settlement | Agent Issue | Dealer Issue |
                                          Sensitive Case | Switch Off | Busy | Not in Service |
                                          Wrong No. | Not Connected
3. Detailed Feedback         → <textarea rows="4"> (free text, multi-line)
4. REF 1 Availability        → <select>: Yes | No | Third Party
5. REF 1 Feedback            → <input type="text">
6. REF 2 Availability        → <select>: Yes | No | Third Party
7. REF 2 Feedback            → <input type="text">
8. Type                      → <select>: Paid | Hold | Remove
```

**Validation Rules (enforce before submission):**
- Fields 1, 2, and 8 are **required** — block submission if not selected.
- Fields 3–7 are optional but trim whitespace before saving.
- On failed validation, show inline error messages below each offending field in red.

**On Submit:**
1. Collect all 8 field values.
2. Build a `feedbackPayload` object with values + `timestamp` (ISO 8601) + `agentId`.
3. Call `syncEngine.writeBack(rowIndex, feedbackPayload)`.
4. Show a success toast on resolve; show error toast on reject.
5. Do NOT clear the form on submit — keep values visible so agent can verify.

---

## 🔄 Phase 6 — Sync Engine, History Logger & Lock Manager (CRITICAL)

### `src/sync/syncEngine.js`

Implement `writeBack(rowIndex, feedbackPayload)` as a two-phase write:

**Phase A — Soft Lock Acquisition:**
1. Call `lockManager.acquireLock(rowIndex, agentId)`.
2. If lock is held by another agent, reject with: `"Row is being updated by another agent. Please try again in a moment."`.
3. If lock acquired, proceed to Phase B.

**Phase B — Atomic Write:**
```js
await Excel.run(async (context) => {
  const sheet = context.workbook.worksheets.getItem(MASTER_SHEET_NAME);
  const row = sheet.getRange(`A${rowIndex + 2}:Q${rowIndex + 2}`); // +2 for header
  row.load("values");
  await context.sync();

  const existing = row.values[0];

  // Write current feedback fields (cols 6–15)
  existing[COL_AVAILABILITY]    = feedbackPayload.availability;
  existing[COL_STD_FEEDBACK]    = feedbackPayload.standardFeedback;
  // ... all 8 fields
  existing[COL_LAST_UPDATED]    = feedbackPayload.timestamp;
  existing[COL_UPDATED_BY]      = feedbackPayload.agentId;

  row.values = [existing];

  // Append to history column
  await historyLogger.appendHistory(context, rowIndex, feedbackPayload);

  await context.sync();
});
```

After write: call `lockManager.releaseLock(rowIndex)`.

### `src/sync/historyLogger.js`

Maintain a **dedicated `FeedbackHistory` sheet** with columns:

| Loan ID | Row Index | Agent ID | Timestamp | Availability | Std Feedback | Detail | REF1 Avail | REF1 FB | REF2 Avail | REF2 FB | Type |

Implement:
```js
export async function appendHistory(context, rowIndex, payload) {
  // Always appends a new row — never overwrites
  // Finds last used row in FeedbackHistory sheet and writes one row below it
}
```

Also update `COL_HISTORY` in `MasterData` as a **semicolon-delimited** string of timestamps:
```
"2025-06-01T10:22:00Z (AGT001); 2025-06-02T14:05:00Z (AGT003)"
```

### `src/sync/lockManager.js`

Use a dedicated `RowLocks` sheet with columns: `Row Index | Agent ID | Lock Time`.

Implement:
```js
export async function acquireLock(rowIndex, agentId) { ... }
export async function releaseLock(rowIndex) { ... }
export async function cleanStaleLocks(maxAgeMinutes = 5) { ... } // Auto-release locks older than 5 min
```

Call `cleanStaleLocks()` on every add-in load.

---

## 🎨 Phase 7 — UI, HTML & CSS (Fluent Design System)

**File:** `src/taskpane/taskpane.html`

Structure:
```html
<body>
  <div id="login-overlay">...</div>
  <div id="app-shell" hidden>
    <header id="top-bar">
      <!-- Logo | Role Badge | Agent Name | Logout Button -->
    </header>
    <div id="admin-panel" hidden>...</div>
    <div id="agent-view" hidden>...</div>
    <div id="case-detail-panel" hidden>...</div>
  </div>
  <div id="toast-container">...</div>
</body>
```

**File:** `src/taskpane/taskpane.css`

Implement using **Microsoft Fluent Design tokens**:
- Font: `Segoe UI`, 14px base.
- Colors: `#0078d4` (primary blue), `#f3f2f1` (background), `#323130` (text), `#d13438` (error red).
- All interactive elements must have `:focus-visible` outlines for accessibility.
- Cards: `border-radius: 4px`, `box-shadow: 0 1.6px 3.6px rgba(0,0,0,0.13)`.
- Tabs: use a simple underline-style tab bar, no JavaScript frameworks — pure CSS + JS toggle.
- Toast notifications: fixed bottom-right, fade-in/fade-out animation, auto-dismiss after 4 seconds.
- Task pane width constraint: max `350px`, all elements must render correctly at this width.

---

## 📦 Phase 8 — Manifest & Build Config

**File:** `manifest.xml`

Generate a fully valid **Office Add-in Manifest v1.1** for Excel:
- `DisplayName`: "CRM Manager Pro"
- `DefaultLocale`: en-IN
- Permissions: `ReadWriteDocument`
- Task Pane source URL: `https://localhost:3000/taskpane.html`
- Icons: reference `assets/icon-32.png` and `assets/icon-80.png`
- Add a `VersionOverrides` block with a ribbon button to open/close the Task Pane.

**File:** `package.json`

Dependencies must include:
```json
{
  "dependencies": {
    "@microsoft/office-js": "latest"
  },
  "devDependencies": {
    "webpack": "^5",
    "webpack-cli": "^5",
    "webpack-dev-server": "^4",
    "copy-webpack-plugin": "^11",
    "html-webpack-plugin": "^5",
    "css-loader": "^6",
    "style-loader": "^3"
  },
  "scripts": {
    "start": "webpack serve --mode development --https",
    "build": "webpack --mode production"
  }
}
```

**File:** `webpack.config.js`

Configure `webpack-dev-server` with:
- `https: true` (required for Office Add-in sideloading)
- `port: 3000`
- `contentBase: dist/`
- Copy `manifest.xml` and `assets/` to `dist/` via `CopyWebpackPlugin`.

---

## 📋 Phase 9 — README

**File:** `README.md`

Write a complete README with:
1. **Project Overview** — What this add-in does.
2. **Prerequisites** — Node.js v18+, Microsoft 365 account, Office Add-in dev tools.
3. **Setup Instructions** — `npm install` → `npm start` → sideload manifest in Excel.
4. **How to Sideload** — Step-by-step for Windows (Insert → Add-ins → Upload My Add-in → browse `manifest.xml`).
5. **User Guide** — Admin login flow, agent login flow, how to assign data, how to submit feedback.
6. **Sheet Schema** — Table documenting all 17 columns in `MasterData`.
7. **Troubleshooting** — Common errors (HTTPS cert issues, CORS, Office.js not loading).

---

## ⚠️ Jules Execution Rules

1. **Generate every file completely.** No `// TODO`, no `...`, no placeholder stubs. Every function must have a full implementation.
2. **Do not use any external UI framework** (no React, Vue, Bootstrap). Pure HTML, CSS, and vanilla JavaScript with Office.js only.
3. **All Office.js operations must be wrapped in `Excel.run(async context => { ... })`** with proper `await context.sync()` calls and `try/catch` error handling.
4. **All async functions must handle errors** and surface them as user-visible toast messages — never silent failures.
5. **Column indices must only be referenced via constants** from `constants.js`. Never use a raw number like `row.values[0][5]`.
6. **The feedback history must be append-only.** No function may overwrite or delete an existing history entry under any condition.
7. **After completing each phase**, output a brief phase summary:
   - Files generated
   - Key decisions made
   - Any assumptions (and why)
8. **After all phases are complete**, generate a final checklist confirming every deliverable is present and functional.

---

## ✅ Final Deliverable Checklist (Jules Must Verify)

- [ ] `manifest.xml` — Valid, sideloadable, correct permissions
- [ ] `package.json` + `webpack.config.js` — Project builds with `npm start`
- [ ] Login modal with role routing (Admin vs Agent)
- [ ] Admin Tab 1: Full data table with per-row agent assignment
- [ ] Admin Tab 2: Checkbox bulk assignment
- [ ] Admin Tab 3: Filter-based bulk assignment with preview
- [ ] Agent View: Filtered card list + Case Detail panel
- [ ] Feedback Form: All 8 fields, exact options, validation enforced
- [ ] `syncEngine.js`: Two-phase write with lock acquisition
- [ ] `historyLogger.js`: Append-only log in `FeedbackHistory` sheet
- [ ] `lockManager.js`: Acquire, release, stale-lock cleanup
- [ ] Fluent Design UI: Correct colors, fonts, card styles, toast system
- [ ] `README.md`: Full setup, sideload, and usage guide
