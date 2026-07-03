# Phase 38 Theme Script Fonts Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve East Asian and complex-script theme font slots for better fidelity in multilingual Word documents.

**Architecture:** Extend `ThemeFonts` with optional major/minor East Asian and complex-script typefaces. The writer emits `a:ea` and `a:cs` inside `a:majorFont` and `a:minorFont`; the reader parses those slots back into JSON.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: extend `ThemeFonts` with optional script-specific font fields.
- `src/docx-writer.ts`: emit optional `a:ea` and `a:cs` font entries in `themeXml`.
- `src/docx-reader.ts`: parse optional `a:ea` and `a:cs` font entries in `parseTheme`.
- `tests/docx-core.test.ts`: add writer and round-trip tests for theme script fonts.

## Task 1: Writer Theme Script Fonts

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add a test named `writes theme script fonts` with:

```ts
theme: {
  name: "Multilingual Theme",
  fonts: {
    major: "Aptos Display",
    minor: "Aptos",
    majorEastAsia: "SimSun",
    majorComplexScript: "Arial",
    minorEastAsia: "Microsoft YaHei",
    minorComplexScript: "Arial",
  },
  colors: { accent1: "4472C4" },
}
```

Assert that `word/theme/theme1.xml` contains:

```xml
<a:majorFont><a:latin typeface="Aptos Display"/><a:ea typeface="SimSun"/><a:cs typeface="Arial"/></a:majorFont>
<a:minorFont><a:latin typeface="Aptos"/><a:ea typeface="Microsoft YaHei"/><a:cs typeface="Arial"/></a:minorFont>
```

- [x] **Step 2: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "theme script fonts"`

Expected: FAIL because only `a:latin` is written for major/minor fonts.

- [x] **Step 3: Implement minimal writer support**

Extend `ThemeFonts` with:

```ts
majorEastAsia?: string;
majorComplexScript?: string;
minorEastAsia?: string;
minorComplexScript?: string;
```

Update `themeXml` so `a:majorFont` and `a:minorFont` include optional `a:ea` and `a:cs` elements when those fields are present.

- [x] **Step 4: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "theme script fonts"`

Expected: writer test PASS.

## Task 2: Reader Round-Trip Support

**Files:**
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing round-trip test**

Add a test named `round-trips theme script fonts` using the same `theme.fonts` object.

- [x] **Step 2: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "theme script fonts"`

Expected: FAIL because the reader only preserves latin major/minor fonts.

- [x] **Step 3: Implement minimal reader support**

Parse `a:majorFont/a:ea`, `a:majorFont/a:cs`, `a:minorFont/a:ea`, and `a:minorFont/a:cs` into the matching optional `ThemeFonts` fields.

- [x] **Step 4: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "theme script fonts"`

Expected: writer and round-trip tests PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 38 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

Run: `git diff --check`

Expected: no whitespace errors.

- [x] **Step 2: Commit**

```bash
git add docs/superpowers/plans/2026-07-04-phase-38-theme-script-fonts.md src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts
git commit -m "feat: add phase 38 theme script fonts"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-38-theme-script-fonts
```

Expected: branch `phase-38-theme-script-fonts` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers writing, parsing, and round-trip for East Asian and complex-script theme font slots.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: JSON names map to OOXML `ea` and `cs` elements inside major/minor font groups.
