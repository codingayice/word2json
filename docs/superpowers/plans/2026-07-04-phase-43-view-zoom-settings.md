# Phase 43 View Zoom Settings Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve document view and zoom settings so DOCX files reopen closer to the original user-facing view state.

**Architecture:** Add optional `DocumentSettings.view` with a view mode plus zoom preset/percent values. The writer emits `w:view` and `w:zoom` into `word/settings.xml`; the reader parses the same supported subset back into JSON.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add `DocumentViewSettings` and reference it from `DocumentSettings.view`.
- `src/docx-writer.ts`: emit `w:view` and `w:zoom` inside `settingsXml`.
- `src/docx-reader.ts`: parse `w:view` and `w:zoom` from `word/settings.xml`.
- `tests/docx-core.test.ts`: add writer and round-trip tests for view settings.

## Task 1: Writer View Settings

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add a test named `writes view settings` with:

```ts
settings: {
  view: {
    mode: "print",
    zoom: { preset: "fullPage", percent: 125 },
  },
}
```

Assert that `word/settings.xml` contains:

```xml
<w:view w:val="print"/>
<w:zoom w:val="fullPage" w:percent="125"/>
```

- [x] **Step 2: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "view settings"`

Expected: FAIL because view settings are not written yet.

- [x] **Step 3: Implement minimal writer support**

Add:

```ts
export type DocumentViewSettings = {
  mode?: "none" | "print" | "outline" | "masterPages" | "normal" | "web";
  zoom?: {
    preset?: "none" | "fullPage" | "bestFit" | "textFit";
    percent?: number;
  };
};
```

and `view?: DocumentViewSettings` to `DocumentSettings`. Emit `w:view` when `mode` exists and `w:zoom` when `zoom.preset` or `zoom.percent` exists.

- [x] **Step 4: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "view settings"`

Expected: writer test PASS.

## Task 2: Reader Round-Trip Support

**Files:**
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing round-trip test**

Add a test named `round-trips view settings` using the same `settings.view` object.

- [x] **Step 2: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "view settings"`

Expected: FAIL because the reader drops view settings.

- [x] **Step 3: Implement minimal reader support**

Parse:

```ts
w:view w:val -> view.mode
w:zoom w:val -> view.zoom.preset
w:zoom w:percent -> view.zoom.percent
```

Only preserve known view modes and zoom presets.

- [x] **Step 4: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "view settings"`

Expected: writer and round-trip tests PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 43 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

Run: `git diff --check`

Expected: no whitespace errors.

- [x] **Step 2: Commit**

```bash
git add docs/superpowers/plans/2026-07-04-phase-43-view-zoom-settings.md src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts
git commit -m "feat: add phase 43 view zoom settings"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-43-view-zoom-settings
```

Expected: branch `phase-43-view-zoom-settings` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers writing, parsing, and round-trip of `w:view` and `w:zoom` settings.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: JSON uses `settings.view`; OOXML uses `w:view` and `w:zoom`.
