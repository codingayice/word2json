# Phase 41 Compatibility Settings Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve Word compatibility settings that influence layout, pagination, and rendering behavior across Word versions.

**Architecture:** Add optional `DocumentSettings.compatibility` containing a `compatMode` shortcut plus ordered `compatSettings`. The writer emits these values under `w:settings/w:compat`; the reader parses them back from `word/settings.xml`.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add `DocumentCompatibilitySettings` and `DocumentCompatSetting`, referenced from `DocumentSettings.compatibility`.
- `src/docx-writer.ts`: emit `w:compat` with `w:compatSetting` entries inside `settingsXml`.
- `src/docx-reader.ts`: parse `w:compat` entries into `settings.compatibility`.
- `tests/docx-core.test.ts`: add writer and round-trip tests for compatibility settings.

## Task 1: Writer Compatibility Settings

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add a test named `writes compatibility settings` with:

```ts
settings: {
  compatibility: {
    compatMode: "15",
    settings: [
      { name: "overrideTableStyleFontSizeAndJustification", uri: "http://schemas.microsoft.com/office/word", value: "1" },
      { name: "useWord2013TrackBottomHyphenation", uri: "http://schemas.microsoft.com/office/word", value: "0" },
    ],
  },
}
```

Assert that `word/settings.xml` contains:

```xml
<w:compat>
<w:compatSetting w:name="compatibilityMode" w:uri="http://schemas.microsoft.com/office/word" w:val="15"/>
<w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:uri="http://schemas.microsoft.com/office/word" w:val="1"/>
<w:compatSetting w:name="useWord2013TrackBottomHyphenation" w:uri="http://schemas.microsoft.com/office/word" w:val="0"/>
</w:compat>
```

- [x] **Step 2: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "compatibility settings"`

Expected: FAIL because `w:compat` is not written yet.

- [x] **Step 3: Implement minimal writer support**

Add:

```ts
export type DocumentCompatibilitySettings = {
  compatMode?: string;
  settings?: DocumentCompatSetting[];
};

export type DocumentCompatSetting = {
  name: string;
  uri: string;
  value: string;
};
```

and `compatibility?: DocumentCompatibilitySettings` to `DocumentSettings`. Emit `w:compat` with a `compatibilityMode` entry when `compatMode` exists, followed by `settings` entries in order.

- [x] **Step 4: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "compatibility settings"`

Expected: writer test PASS.

## Task 2: Reader Round-Trip Support

**Files:**
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing round-trip test**

Add a test named `round-trips compatibility settings` using the same `settings.compatibility` object.

- [x] **Step 2: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "compatibility settings"`

Expected: FAIL because the reader drops `w:compat`.

- [x] **Step 3: Implement minimal reader support**

Parse `w:compat/w:compatSetting` entries. Map the entry with `w:name="compatibilityMode"` to `compatibility.compatMode`, and preserve all other entries in `compatibility.settings` with `name`, `uri`, and `value`.

- [x] **Step 4: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "compatibility settings"`

Expected: writer and round-trip tests PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 41 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

Run: `git diff --check`

Expected: no whitespace errors.

- [x] **Step 2: Commit**

```bash
git add docs/superpowers/plans/2026-07-04-phase-41-compatibility-settings.md src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts
git commit -m "feat: add phase 41 compatibility settings"
```

- [x] **Step 3: Push**

```bash
git push -u origin phase-41-compatibility-settings
```

Expected: branch `phase-41-compatibility-settings` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers writing, parsing, and round-trip of Word compatibility settings under `w:compat`.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: JSON uses `compatibility.compatMode` and ordered `compatibility.settings`; OOXML uses `w:compatSetting` attributes.
