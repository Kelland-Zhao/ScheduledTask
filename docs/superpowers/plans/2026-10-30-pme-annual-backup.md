# PME Annual Backup Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Archive the previous year’s complete PME dataset and KPI block as one gated operation, alerting on unresolved records and notifying on successful changes.

**Architecture:** Add one standalone Apps Script module with a single `backupPmeData(e)` entry point. Preflight all PME rows by header names and `YearMonth`; block every destination write and source deletion while any status or YearMonth prevents confirming annual completeness. When clear, idempotently merge PME rows by serial number, snapshot the KPI block as values with formats, verify both destinations, delete source rows, and send/queue email notifications.

**Tech Stack:** Google Apps Script V8, SpreadsheetApp, LockService, PropertiesService, GmailApp.

**Spec:** `docs/superpowers/specs/2026-10-30-pme-annual-backup-design.md`

## Global Constraints

- Archive year is selected using `保养月份/ YearMonth`, never `PM_Date`.
- Only `Production Status=生产` or `未达144` with checked exemption and nonblank reason counts as processed.
- Any unresolved target-year record blocks all PME and KPI destination writes and all source deletions.
- Blank/invalid `YearMonth` anywhere in `Master Data` blocks annual archive completion.
- Only delete PME source records after both destination datasets verify successfully.
- Send reminder emails to `kelland_zhao@colpal.com` for blockers and success emails after verified archive changes.
- Do not alter KPI eligibility or recalculate source KPI in the backup module.
- Keep existing user changes to `35 - PME计算.js` and unrelated untracked files untouched.
- No automated tests are added or run under the current session instruction; use syntax and diff checks only.

## Review Focus

- A pending previous-year row blocks the entire PME and KPI archive and prevents all source deletion.
- Invalid or blank `YearMonth` fails closed without falling back to `PM_Date`.
- Missing or duplicate source/target serial numbers prevent an unverifiable PME archive from deleting source data.
- Partial target writes or verification failures leave all PME source rows intact and remain safe to retry.
- Reminder and success email retries do not create repeated success messages when archived content has not changed.

---

### Task 1: Add PME preflight, annual gate, and blocker reminders

**Files:**
- Create: `36 - PME数据备份.js`

**Interfaces:**
- Consumes: `SpreadsheetApp`, `LockService`, `GmailApp`, `writeLog`, the approved source/destination spreadsheet IDs and headers.
- Produces: `backupPmeData(e)`; private helpers prefixed `_pmeBackup_` for YearMonth parsing, header lookup, annual completeness checks, and blocker email rendering.

- [x] Define source/target IDs, sheet names, required headers, recipient, and lock-protected `backupPmeData(e)` entry point.
- [x] Resolve the previous year using `Asia/Shanghai`; allow scheduled calls only on Mondays while allowing a manual menu call any day.
- [x] Read `Master Data` headers and rows by header name; fail closed on missing/duplicate required headers, blank/invalid `YearMonth`, or any target-year row not resolved as Production or validly exempted `<144`.
- [x] Send one consolidated blocker email per run, containing row identifiers and source link; log mail failures without bypassing the annual gate.
- [x] Run `node --check '36 - PME数据备份.js'` and inspect the entry point for lock release and trigger-window behavior.

### Task 2: Implement idempotent PME/KPI archive and completion notifications

**Files:**
- Modify: `36 - PME数据备份.js`
- Modify: `35 - PME计算.js` only to share the ScriptLock while recalculation writes `Master Data`

**Interfaces:**
- Consumes: Task 1 preflight result and exact `Master Data` header map.
- Produces: private helpers for annual-tab PME merge, header-discovered KPI block extraction, destination verification, source deletion, and email notification state.

- [x] Validate unique nonblank `Serial No.` across candidate source rows and target annual tab; abort PME/KPI writes on identity ambiguity.
- [x] Merge the complete target-year PME rows into the year tab by `Serial No.`, preserve all fields, and set checkbox validation only for Boolean exemption cells.
- [x] Find the target year KPI block by `Process`, month headers, and annual column; write static values plus cell formatting to the existing `KPI` tab, replacing only the same-year block.
- [x] Re-read and verify target headers, PME identities/fields, and KPI block values/formats; only then delete the validated source rows from bottom to top.
- [x] Before each contiguous source deletion, re-read the full row block and compare it against the approved source snapshot.
- [x] Validate that existing PME archive rows have valid `YearMonth` values for the target year.
- [x] Send a success email only after source deletion succeeds; track the last notified archive content version per year so no-op reruns stay quiet and failed notifications retry.
- [x] Run `node --check '36 - PME数据备份.js'`, `git diff --check`, and inspect all failure paths for source-row preservation.

### Task 3: Prepare schedule/menu integration for reviewed deployment

**Files:**
- Modify: service-account config spreadsheet `1VH9I4-vRA3GlKo2dx6LVsr29HKdzCylxwROdsfoLwtM` (`定时设置`, `菜单设置`) after final deployment approval.
- Modify: deployed Apps Script project ID `1It7AdBDwWS62EmLpaV0U0UDTE5P1kuS9oFVxMhemmlAhIPQV-gn2IVNE` only after final deployment approval.

**Interfaces:**
- Consumes: deployed `backupPmeData(e)` and the existing seven-column `定时设置` / three-column `菜单设置` formats.
- Produces: one weekly Monday rule at 09:25 and one enabled manual menu item `PME数据备份`.

- [x] Re-read the live config through the service account and confirm an available time; do not switch accounts if the TLS read fails.
- [x] Compare the remote Apps Script project to the local source in an isolated pull before selecting full-project deployment.
- [x] Expand the config tabs with the service-account append-row operation, then write and verify the weekly Monday rule and `PME数据备份` manual menu entry.
- [x] Push the reviewed source using the work-account `clasp` login; re-pull the remote project and verify all 39 source files match.
- [x] Re-read the written config rows and verify the dynamic menu/config integration.

### Task 4: Record the project change

**Files:**
- Modify: `docs/项目记录.md`

- [x] Add a concise project history entry for PME annual backup behavior and its schedule/menu integration status.
- [x] Run whitespace and diff checks; do not run automated tests.
