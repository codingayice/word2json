# Phase 39 Theme Supplemental Fonts Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve theme supplemental script font mappings so multilingual DOCX templates retain script-specific typeface choices.

**Architecture:** Add `ThemeFonts.supplemental` as an ordered list of `{ script, typeface }` mappings. The writer emits those mappings as `a:font` children of `a:majorFont` and `a:minorFont`; the reader parses both groups back into JSON while preserving order.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add `ThemeSupplementalFont` and optional `ThemeFonts.supplemental`.
- `src/docx-writer.ts`: emit supplemental `a:font` entries inside theme major/minor font groups.
- `src/docx-reader.ts`: parse supplemental `a:font` entries from theme major/minor font groups.
- `tests/docx-core.test.ts`: add writer and round-trip tests for supplemental theme fonts.

## Task 1: Writer Supplemental Fonts

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add a test named `writes theme supplemental fonts` with:

```ts
theme: {
  name: "Multilingual Theme",
  fonts: {
    major: "Aptos Display",
    minor: "Aptos",
    supplemental: [
      { group: "major", script: "Hans", typeface: "SimSun" },
      { group: "major", script: "Jpan", typeface: "Yu Gothic" },
      { group: "minor", script: "Hans", typeface: "Microsoft YaHei" },
      { group: "minor", script: "Hang", typeface: "Malgun Gothic" },
    ],
  },
  colors: { accent1: "4472C4" },
}
```

Assert that `word/theme/theme1.xml` contains:

```xml
<a:font script="Hans" typeface="SimSun"/>
<a:font script="Jpan" typeface="Yu Gothic"/>
<a:font script="Hans" typeface="Microsoft YaHei"/>
<a:font script="Hang" typeface="Malgun Gothic"/>
```

- [x] **Step 2: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "theme supplemental fonts"`

Expected: FAIL because supplemental `a:font` entries are not written yet.

- [x] **Step 3: Implement minimal writer support**

Add:

```ts
export type ThemeSupplementalFont = {
  group: "major" | "minor";
  script: string;
  typeface: string;
};
```

and `supplemental?: ThemeSupplementalFont[]` to `ThemeFonts`. Update `themeXml` so each supplemental entry is emitted inside the matching `a:majorFont` or `a:minorFont` group.

- [x] **Step 4: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "theme supplemental fonts"`

Expected: writer test PASS.

## Task 2: Reader Round-Trip Support

**Files:**
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing round-trip test**

Add a test named `round-trips theme supplemental fonts` using the same `theme.fonts.supplemental` array.

- [x] **Step 2: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "theme supplemental fonts"`

Expected: FAIL because the reader drops supplemental `a:font` entries.

- [x] **Step 3: Implement minimal reader support**

Parse `a:majorFont/a:font` into `{ group: "major", script, typeface }` and `a:minorFont/a:font` into `{ group: "minor", script, typeface }`, preserving major entries before minor entries and preserving each group's XML order.

- [x] **Step 4: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "theme supplemental fonts"`

Expected: writer and round-trip tests PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 39 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

Run: `git diff --check`

Expected: no whitespace errors.

- [x] **Step 2: Commit**

```bash
git add docs/superpowers/plans/2026-07-04-phase-39-theme-supplemental-fonts.md src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts
git commit -m "feat: add phase 39 theme supplemental fonts"
```

- [x] **Step 3: Push**

```bash
git push -u origin phase-39-theme-supplemental-fonts
```

Expected: branch `phase-39-theme-supplemental-fonts` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers writing, parsing, and round-trip of supplemental theme script font mappings.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: JSON uses `group`, `script`, and `typeface`; OOXML uses `a:font` under either major or minor font groups.
