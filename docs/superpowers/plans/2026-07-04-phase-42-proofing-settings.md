# Phase 42 Proofing Settings Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve proofing and hyphenation-related document settings that affect Word review state and line-breaking behavior.

**Architecture:** Add optional `DocumentSettings.proofing` with proof-state flags and hyphenation controls. The writer emits these values into `word/settings.xml`; the reader parses the same supported subset back into JSON.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add `DocumentProofingSettings` and reference it from `DocumentSettings.proofing`.
- `src/docx-writer.ts`: emit proofing and hyphenation elements inside `settingsXml`.
- `src/docx-reader.ts`: parse proofing and hyphenation elements from `word/settings.xml`.
- `tests/docx-core.test.ts`: add writer and round-trip tests for proofing settings.

## Task 1: Writer Proofing Settings

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add a test named `writes proofing settings` with:

```ts
settings: {
  proofing: {
    spelling: "clean",
    grammar: "dirty",
    doNotHyphenateCaps: true,
    hyphenationZone: 360,
  },
}
```

Assert that `word/settings.xml` contains:

```xml
<w:proofState w:spelling="clean" w:grammar="dirty"/>
<w:doNotHyphenateCaps/>
<w:hyphenationZone w:val="360"/>
```

- [x] **Step 2: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "proofing settings"`

Expected: FAIL because proofing settings are not written yet.

- [x] **Step 3: Implement minimal writer support**

Add:

```ts
export type DocumentProofingSettings = {
  spelling?: "clean" | "dirty";
  grammar?: "clean" | "dirty";
  doNotHyphenateCaps?: boolean;
  hyphenationZone?: number;
};
```

and `proofing?: DocumentProofingSettings` to `DocumentSettings`. Emit `w:proofState`, `w:doNotHyphenateCaps`, and `w:hyphenationZone` when those fields exist.

- [x] **Step 4: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "proofing settings"`

Expected: writer test PASS.

## Task 2: Reader Round-Trip Support

**Files:**
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing round-trip test**

Add a test named `round-trips proofing settings` using the same `settings.proofing` object.

- [x] **Step 2: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "proofing settings"`

Expected: FAIL because the reader drops proofing settings.

- [x] **Step 3: Implement minimal reader support**

Parse:

```ts
w:proofState w:spelling -> proofing.spelling
w:proofState w:grammar -> proofing.grammar
w:doNotHyphenateCaps -> proofing.doNotHyphenateCaps
w:hyphenationZone w:val -> proofing.hyphenationZone
```

Only accept `clean` and `dirty` for proof state values.

- [x] **Step 4: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "proofing settings"`

Expected: writer and round-trip tests PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 42 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

Run: `git diff --check`

Expected: no whitespace errors.

- [x] **Step 2: Commit**

```bash
git add docs/superpowers/plans/2026-07-04-phase-42-proofing-settings.md src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts
git commit -m "feat: add phase 42 proofing settings"
```

- [x] **Step 3: Push**

```bash
git push -u origin phase-42-proofing-settings
```

Expected: branch `phase-42-proofing-settings` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers writing, parsing, and round-trip of proof state and common hyphenation controls in `word/settings.xml`.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: JSON uses `settings.proofing`; OOXML uses `w:proofState`, `w:doNotHyphenateCaps`, and `w:hyphenationZone`.
