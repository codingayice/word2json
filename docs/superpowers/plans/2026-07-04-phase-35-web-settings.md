# Phase 35 Web Settings Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve DOCX `word/webSettings.xml` metadata for web-oriented document settings.

**Architecture:** Extend `DocumentSettings` with optional `web` settings. The writer emits `word/webSettings.xml`, adds a document relationship and content type override; the reader parses the part back into `settings.web`.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add `DocumentWebSettings` and `settings.web`.
- `src/docx-writer.ts`: emit web settings part, relationship, and content type override.
- `src/docx-reader.ts`: parse web settings part.
- `tests/docx-core.test.ts`: add writer and round-trip tests for web settings.

## Task 1: Writer Support

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add a test named `writes web settings` with:

```ts
settings: {
  web: {
    optimizeForBrowser: true,
    allowPng: true,
    doNotSaveAsSingleFile: true,
    pixelsPerInch: 120,
  },
}
```

- [x] **Step 2: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "web settings"`

Expected: FAIL because `word/webSettings.xml` is not supported yet.

- [x] **Step 3: Implement minimal writer support**

Emit:

```xml
<w:webSettings>
  <w:optimizeForBrowser/>
  <w:allowPNG/>
  <w:doNotSaveAsSingleFile/>
  <w:pixelsPerInch w:val="120"/>
</w:webSettings>
```

and add relationship/content type entries.

- [x] **Step 4: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "web settings"`

Expected: writer test PASS.

## Task 2: Reader Round-Trip Support

**Files:**
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing round-trip test**

Add a test named `round-trips web settings` using the same `settings.web` object.

- [x] **Step 2: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "web settings"`

Expected: FAIL because the reader drops web settings.

- [x] **Step 3: Implement minimal reader support**

Parse `word/webSettings.xml` into `settings.web`, preserving boolean flags and `pixelsPerInch`.

- [x] **Step 4: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "web settings"`

Expected: writer and round-trip tests PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 35 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

Run: `git diff --check`

Expected: no whitespace errors.

- [x] **Step 2: Commit**

```bash
git add docs/superpowers/plans/2026-07-04-phase-35-web-settings.md src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts
git commit -m "feat: add phase 35 web settings"
```

- [x] **Step 3: Push**

```bash
git push -u origin phase-35-web-settings
```

Expected: branch `phase-35-web-settings` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers writing, relationships, content types, reading, and round-trip for common web settings.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: JSON uses `allowPng`, while XML uses Word's `w:allowPNG`.
