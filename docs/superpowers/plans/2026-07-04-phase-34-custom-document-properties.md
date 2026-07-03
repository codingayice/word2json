# Phase 34 Custom Document Properties Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve DOCX custom document properties in `docProps/custom.xml`.

**Architecture:** Extend `DocumentJson.properties` with a `custom` array. The writer emits `docProps/custom.xml` with stable `pid` values and VT typed values, adds the package relationship and content type override; the reader parses custom property names and VT value types back into JSON.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add custom document property types.
- `src/docx-writer.ts`: emit `docProps/custom.xml`, package relationship, and content type override.
- `src/docx-reader.ts`: parse `docProps/custom.xml`.
- `tests/docx-core.test.ts`: add writer and round-trip tests for custom document properties.

## Task 1: Custom Properties Writer

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add a test named `writes custom document properties` with string, number, boolean, and date properties.

- [x] **Step 2: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "custom document properties"`

Expected: FAIL because `docProps/custom.xml` is not supported yet.

- [x] **Step 3: Implement minimal writer support**

Add:

```ts
export type CustomDocumentProperty = {
  name: string;
  type: "string" | "number" | "boolean" | "date";
  value: string | number | boolean;
};
```

Emit VT nodes:

```xml
<vt:lpwstr>value</vt:lpwstr>
<vt:i4>42</vt:i4>
<vt:bool>true</vt:bool>
<vt:filetime>2026-07-04T00:00:00Z</vt:filetime>
```

- [x] **Step 4: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "custom document properties"`

Expected: writer test PASS.

## Task 2: Custom Properties Reader

**Files:**
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing round-trip test**

Add a test named `round-trips custom document properties` using the same four custom properties.

- [x] **Step 2: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "custom document properties"`

Expected: FAIL because the reader drops custom properties.

- [x] **Step 3: Implement minimal reader support**

Parse `property` entries from `docProps/custom.xml`, identify the VT child, and return the matching JSON type/value.

- [x] **Step 4: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "custom document properties"`

Expected: writer and round-trip tests PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 34 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

Run: `git diff --check`

Expected: no whitespace errors.

- [x] **Step 2: Commit**

```bash
git add docs/superpowers/plans/2026-07-04-phase-34-custom-document-properties.md src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts
git commit -m "feat: add phase 34 custom document properties"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-34-custom-document-properties
```

Expected: branch `phase-34-custom-document-properties` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers string, number, boolean, and date custom document properties plus relationships and content type override.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: `properties.custom` is an array of named typed values.
