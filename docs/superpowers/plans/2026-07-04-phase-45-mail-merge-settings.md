# Phase 45 Mail Merge Settings Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve common mail merge settings so template DOCX files retain their merge metadata through JSON round-trip.

**Architecture:** Add optional `DocumentSettings.mailMerge` for the supported `w:mailMerge` children inside `word/settings.xml`. The writer emits the configured mail merge elements; the reader parses the same subset back into JSON.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add `DocumentMailMergeSettings` and reference it from `DocumentSettings.mailMerge`.
- `src/docx-writer.ts`: emit `w:mailMerge` inside `settingsXml`.
- `src/docx-reader.ts`: parse `w:mailMerge` from `word/settings.xml`.
- `tests/docx-core.test.ts`: add writer and round-trip tests for mail merge settings.

## Task 1: Writer Mail Merge Settings

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add a test named `writes mail merge settings` with:

```ts
settings: {
  mailMerge: {
    mainDocumentType: "formLetters",
    dataType: "native",
    connectString: "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=contacts.xlsx;",
    query: "SELECT * FROM `Contacts$`",
    viewMergedData: true,
    activeRecord: 3,
    checkErrors: 1,
  },
}
```

Assert that `word/settings.xml` contains:

```xml
<w:mailMerge>
<w:mainDocumentType w:val="formLetters"/>
<w:dataType w:val="native"/>
<w:connectString w:val="Provider=Microsoft.ACE.OLEDB.12.0;Data Source=contacts.xlsx;"/>
<w:query w:val="SELECT * FROM `Contacts$`"/>
<w:viewMergedData/>
<w:activeRecord w:val="3"/>
<w:checkErrors w:val="1"/>
</w:mailMerge>
```

- [x] **Step 2: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "mail merge settings"`

Expected: FAIL because `w:mailMerge` is not written yet.

- [x] **Step 3: Implement minimal writer support**

Add:

```ts
export type DocumentMailMergeSettings = {
  mainDocumentType?: string;
  dataType?: string;
  connectString?: string;
  query?: string;
  viewMergedData?: boolean;
  activeRecord?: number;
  checkErrors?: number;
};
```

and `mailMerge?: DocumentMailMergeSettings` to `DocumentSettings`. Emit `w:mailMerge` when any field is present.

- [x] **Step 4: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "mail merge settings"`

Expected: writer test PASS.

## Task 2: Reader Round-Trip Support

**Files:**
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing round-trip test**

Add a test named `round-trips mail merge settings` using the same `settings.mailMerge` object.

- [x] **Step 2: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "mail merge settings"`

Expected: FAIL because the reader drops `w:mailMerge`.

- [x] **Step 3: Implement minimal reader support**

Parse:

```ts
w:mainDocumentType w:val -> mainDocumentType
w:dataType w:val -> dataType
w:connectString w:val -> connectString
w:query w:val -> query
w:viewMergedData -> viewMergedData
w:activeRecord w:val -> activeRecord
w:checkErrors w:val -> checkErrors
```

- [x] **Step 4: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "mail merge settings"`

Expected: writer and round-trip tests PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 45 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

Run: `git diff --check`

Expected: no whitespace errors.

- [x] **Step 2: Commit**

```bash
git add docs/superpowers/plans/2026-07-04-phase-45-mail-merge-settings.md src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts
git commit -m "feat: add phase 45 mail merge settings"
```

- [x] **Step 3: Push**

```bash
git push -u origin phase-45-mail-merge-settings
```

Expected: branch `phase-45-mail-merge-settings` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers writing, parsing, and round-trip of common `w:mailMerge` metadata in settings.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: JSON uses `settings.mailMerge`; OOXML uses `w:mailMerge` child elements.
