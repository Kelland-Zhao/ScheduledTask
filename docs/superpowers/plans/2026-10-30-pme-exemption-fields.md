# PME Exemption Fields Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make 35 号 PME calculations write and retain `Exempt?` and `Exemption Reason` according to the approved field rules.

**Architecture:** Keep the existing `Master Data` layout and resolve the two manual fields by their header names. Carry extended row values by PME record identity through full and incremental writes, then set or remove the checkbox validation according to the updated `Production Status`.

**Tech Stack:** Google Apps Script V8, SpreadsheetApp.

**Spec:** `docs/superpowers/specs/2026-10-30-pme-exemption-fields-design.md`

## Global Constraints

- Preserve the first 15 PME calculation headers and calculations.
- New PME rows receive an unchecked checkbox in `豁免与否/ Exempt?` and a blank `豁免原因/ Exemption Reason`.
- On an existing row update to `生产`, write `NA` to both fields and remove checkbox validation from `Exempt?`.
- On an existing row update to `未达144`, preserve an existing checked exemption only when its reason is non-empty; otherwise set the checkbox to unchecked and clear the reason.
- `未达144` remains excluded from KPI under the existing eligibility logic.

## Review Focus

- New rows, including the first calculation result, must follow the new-row default rule.
- A transition from `未达144` to `生产` must write both `NA` values and remove checkbox validation.
- A transition from `生产` to `未达144` must restore checkbox validation and write unchecked/blank unless a valid checked waiver exists.
- Existing checked waivers with a non-empty reason must survive full and incremental recalculation while the status remains `未达144`.
- Row deletions, reidentification, or normalized Workcenter matching must not move manual values to another PME record.

---

### Task 1: Preserve and write manual archive fields in PME calculations

**Files:**
- Modify: `35 - PME计算.js`

**Interfaces:**
- Consume existing `Master Data` headers `序列号/ Serial No.`, `Production Status`, `豁免与否/ Exempt?`, and `豁免原因/ Exemption Reason`.
- Produce `Master Data` rows whose manual fields follow the matching PME record through full and incremental calculations.

- [x] **Step 1: Resolve the two manual columns by header name and require their presence**

Keep the first 15 `_pme_MASTER_HEADERS` entries unchanged. Add an exact header lookup for `豁免与否/ Exempt?` and `豁免原因/ Exemption Reason`; fail clearly if either header is absent.

- [x] **Step 2: Apply the approved new-row and update rules while merging rows**

For a new record, set `Exempt?` to boolean `false` and the reason to `""`. For an existing record with updated status `生产`, set both values to `"NA"`. For status `未达144`, retain an existing checked value and non-empty reason; otherwise set `false` and `""`. Preserve other manual fields beyond the first 15 columns by matching the existing PME record using serial number, with the existing date-plus-normalized-Workcenter identity as fallback.

- [x] **Step 3: Write full row width in both calculation modes and apply checkbox validation**

Update `_pme_writeMasterDataRows`, `_pme_buildIncrementalMasterDataWritePlan`, and `_pme_writeMasterDataIncremental` to compare and write the merged row width, including the two manual fields. Ensure new and `未达144` rows use checkbox validation in the exemption column; remove that validation for updated `生产` rows before storing `"NA"`. Keep the existing number formatting behavior for calculated columns.

- [x] **Step 4: Run a syntax check and inspect the write paths**

Run `node --check '35 - PME计算.js'`. Then inspect the full and incremental branches to confirm the five Review Focus conditions follow the approved values and that `_pme_calculateKpiEligibility` remains unchanged. No automated tests are added or run.
