# Empty Dashboard — Troubleshooting Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Restore dashboard to show PTA and PSU activity cards

**Architecture:** The dashboard loads data via `getSessionData()` → `getKpiInfoList()` + `getPelanStrategikData()` + `getKpiStatus()` Apps Script functions, fed from Google Sheets. Frontend renders cards client-side. Empty state means either a backend call failed, Sheets data is missing, or a JS error prevents rendering.

**Tech Stack:** Google Apps Script (Code.js) + HTML/JS (index.html) + Google Sheets

---
### Task 1: Identify which deployment URL is broken

**Files:** (observation only, no edits)

- [ ] **Step 1: Ask user which URL shows empty**

Ask: "Are you seeing empty on the Sites page (`https://sites.google.com/unisza.edu.my/pps/tindakan/psu-pta`) or on the direct Apps Script URL (`https://script.google.com/macros/s/AKfycbwg6WXCk7OZFXQztWzeRdLAZIGzI-OQ9oQP_yk6lbOy75Jnov-XEY8sTAOXQ-InDpTJ/exec`) or on the new @69 URL?"

- [ ] **Step 2: Open the broken URL in a browser**

Open the URL the user confirms. Open browser DevTools (F12) → Console tab.

- [ ] **Step 3: Check console for JS errors**

Look for: red errors, network failures (Failed to load resource), CSP violations. Report all messages.

- [ ] **Step 4: Check Network tab for `google.script.run` calls**

Filter to `usercode` or look for POST requests to the script. Verify `getSessionData`, `getKpiInfoList`, `getPelanStrategikData`, `getKpiStatus` are called and return data.

---

### Task 2: Verify Script Properties

**Files:**
- Read: `Code.js:6-7` (FOLDER_SISTEM_KECIL_ID, SPREADSHEET_ID)

- [ ] **Step 1: Check SPREADSHEET_ID in Script Properties**

Run in clasp or Apps Script editor:

```javascript
PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')
```

If empty/null, this is the root cause. The sheet ID must be set via File → Project properties → Script properties in the Apps Script editor.

- [ ] **Step 2: Check ADMIN_EMAILS Script Property**

Run:

```javascript
PropertiesService.getScriptProperties().getProperty('ADMIN_EMAILS')
```

If value is `["SET_ADMIN_EMAILS_IN_SCRIPT_PROPERTIES"]`, the admin list was never set. Set it to a JSON array of real @unisza.edu.my emails.

- [ ] **Step 3: Check FOLDER_SISTEM_KECIL_ID**

Run:

```javascript
PropertiesService.getScriptProperties().getProperty('FOLDER_SISTEM_KECIL_ID')
```

---

### Task 3: Verify Sheet Data

**Files:** (observation only)

- [ ] **Step 1: Open the Google Sheet**

Use the SPREADSHEET_ID to open the sheet. Check these sheets exist:
- `Maklumat_Lonjakan` (PTA data source)
- `Pelan_Strategik_Universiti` (PSU data source)
- `Status_Semasa` (KPI status)
- `Log_Kemaskini` (entry logs)

- [ ] **Step 2: Check Maklumat_Lonjakan has rows**

Minimal content: header row + at least 1 data row with ID in column A. If empty, `getKpiInfoList()` returns `[]` → no PTA cards rendered.

- [ ] **Step 3: Check Pelan_Strategik_Universiti has rows**

If empty, `getPelanStrategikData()` returns `[]` → no PSU cards rendered.

- [ ] **Step 4: Run data fetch directly in Apps Script editor**

Run this in the editor:

```javascript
function debugData() {
  const kpiInfo = getKpiInfoList();
  const psuInfo = getPelanStrategikData();
  const statusData = getKpiStatus();
  console.log('KPI Info count:', kpiInfo.length);
  console.log('PSU Info count:', psuInfo.length);
  console.log('Status keys:', Object.keys(statusData).length);
  return { kpiCount: kpiInfo.length, psuCount: psuInfo.length, statusCount: Object.keys(statusData).length };
}
```

If counts are 0, the issue is in the Sheet or the data functions.

---

### Task 4: Check if Senarai Admin sheet was created by @68

**Files:** (observation only)

- [ ] **Step 1: Check for stray "Senarai Admin" sheet**

Look in the Google Sheet for a `Senarai Admin` tab created by the @68 deployment's `getAdminList()`. If it exists and contains the deploying user's email, it's harmless but confirm.

- [ ] **Step 2: Verify it didn't interfere with sheet indexing**

The `Senarai Admin` sheet being present shouldn't affect anything — `getKpiInfoList()` and `getPelanStrategikData()` use explicit sheet names (`Maklumat_Lonjakan`, `Pelan_Strategik_Universiti`).

---

### Task 5: Fix root cause

**Files:**
- Modify: `Code.js` (if code fix needed)
- Modify: `index.html` (if frontend fix needed)

- [ ] **Step 1: Apply the fix**

Based on findings from Tasks 1-4:
- If SPREADSHEET_ID missing → set it in Script Properties
- If Sheets empty → restore data from backup or run `setupPelanStrategikUniversiti()`
- If JS error → fix the code
- If wrong URL → point user to correct URL

- [ ] **Step 2: Push fix to Apps Script**

```bash
npx @google/clasp push
npx @google/clasp deploy -d "fix: restore dashboard data loading"
```

- [ ] **Step 3: Commit fix**

```bash
git add -A && git commit -m "fix: restore dashboard data loading"
```

---

### Task 6: Verify fix

**Files:** (observation only)

- [ ] **Step 1: Open the deployment URL in browser**

Confirm PTA and PSU cards render, data loads, no console errors.

- [ ] **Step 2: Compare with the working screenshot**

Verify the restored dashboard matches the user's screenshot of the last working state.

- [ ] **Step 3: Confirm user is satisfied**

Ask user to confirm the dashboard is restored.
