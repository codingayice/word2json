# Phase 37 Theme Color Scheme Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve richer DOCX theme color schemes so generated documents keep theme-driven colors more faithfully.

**Architecture:** Extend `ThemeColors` from only `accent1` to the common OOXML theme color slots. The writer emits each supplied slot inside `word/theme/theme1.xml`; the reader parses the same slots back from `a:clrScheme`.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: extend `ThemeColors` with optional slots for dark/light, accent, hyperlink, and followed hyperlink colors.
- `src/docx-writer.ts`: emit supported theme color slots in `themeXml`.
- `src/docx-reader.ts`: parse supported theme color slots in `parseTheme`.
- `tests/docx-core.test.ts`: add writer and round-trip tests for rich theme colors.

## Task 1: Writer Theme Colors

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add a test named `writes rich theme color scheme` with:

```ts
theme: {
  name: "Contract Theme",
  fonts: { major: "Aptos Display", minor: "Aptos" },
  colors: {
    dark1: "000000",
    light1: "FFFFFF",
    dark2: "1F2937",
    light2: "F8FAFC",
    accent1: "4472C4",
    accent2: "ED7D31",
    accent3: "A5A5A5",
    accent4: "FFC000",
    accent5: "5B9BD5",
    accent6: "70AD47",
    hyperlink: "0563C1",
    followedHyperlink: "954F72",
  },
}
```

Assert that `word/theme/theme1.xml` contains `a:dk1`, `a:lt1`, `a:dk2`, `a:lt2`, `a:accent1` through `a:accent6`, `a:hlink`, and `a:folHlink` with their `a:srgbClr` values.

- [x] **Step 2: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "rich theme color scheme"`

Expected: FAIL because only `accent1` is written.

- [x] **Step 3: Implement minimal writer support**

Extend `ThemeColors` with optional fields and update `themeXml` to emit:

```xml
<a:dk1><a:srgbClr val="000000"/></a:dk1>
<a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="1F2937"/></a:dk2>
<a:lt2><a:srgbClr val="F8FAFC"/></a:lt2>
<a:accent1><a:srgbClr val="4472C4"/></a:accent1>
<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
<a:accent3><a:srgbClr val="A5A5A5"/></a:accent3>
<a:accent4><a:srgbClr val="FFC000"/></a:accent4>
<a:accent5><a:srgbClr val="5B9BD5"/></a:accent5>
<a:accent6><a:srgbClr val="70AD47"/></a:accent6>
<a:hlink><a:srgbClr val="0563C1"/></a:hlink>
<a:folHlink><a:srgbClr val="954F72"/></a:folHlink>
```

- [x] **Step 4: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "rich theme color scheme"`

Expected: writer test PASS.

## Task 2: Reader Round-Trip Support

**Files:**
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing round-trip test**

Add a test named `round-trips rich theme color scheme` using the same `theme.colors` object.

- [x] **Step 2: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "rich theme color scheme"`

Expected: FAIL because the reader only preserves `accent1`.

- [x] **Step 3: Implement minimal reader support**

Parse each supported color slot from `a:clrScheme`, mapping OOXML names to JSON names:

```ts
dk1 -> dark1
lt1 -> light1
dk2 -> dark2
lt2 -> light2
accent1 -> accent1
accent2 -> accent2
accent3 -> accent3
accent4 -> accent4
accent5 -> accent5
accent6 -> accent6
hlink -> hyperlink
folHlink -> followedHyperlink
```

- [x] **Step 4: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "rich theme color scheme"`

Expected: writer and round-trip tests PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 37 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

Run: `git diff --check`

Expected: no whitespace errors.

- [x] **Step 2: Commit**

```bash
git add docs/superpowers/plans/2026-07-04-phase-37-theme-color-scheme.md src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts
git commit -m "feat: add phase 37 theme color scheme"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-37-theme-color-scheme
```

Expected: branch `phase-37-theme-color-scheme` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers writer output, reader parsing, relationships already covered by existing theme support, and round-trip for common theme color slots.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: JSON uses `followedHyperlink`, and OOXML uses `folHlink`.
