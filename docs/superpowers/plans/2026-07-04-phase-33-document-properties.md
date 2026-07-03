# Phase 33 Document Properties Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve DOCX core and extended document properties.

**Architecture:** Add optional `properties` to `DocumentJson` with `core` and `app` groups. The writer emits `docProps/core.xml` and `docProps/app.xml`, registers package relationships and content type overrides; the reader parses those parts back into JSON when present.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add document property types and `DocumentJson.properties`.
- `src/docx-writer.ts`: emit core/app property parts, package relationships, and content type overrides.
- `src/docx-reader.ts`: parse core/app property parts into JSON.
- `tests/docx-core.test.ts`: add writer and round-trip tests for document properties.

## Task 1: Core Properties

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add writer and round-trip tests for:

```ts
properties: {
  core: {
    title: "Quarterly Report",
    subject: "Sales",
    creator: "Ada Lovelace",
    keywords: "sales,quarterly",
    description: "Executive summary",
    lastModifiedBy: "Grace Hopper",
    created: "2026-07-04T00:00:00Z",
    modified: "2026-07-04T01:00:00Z",
  },
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "document core properties"`

Expected: FAIL because docProps/core.xml is not supported yet.

- [x] **Step 3: Implement minimal support**

Emit and parse `docProps/core.xml` using Dublin Core, DCTerms, and cp namespaces.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "document core properties"`

Expected: PASS.

## Task 2: Extended App Properties

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing tests**

Add writer and round-trip tests for:

```ts
properties: {
  app: {
    application: "word2json",
    company: "ACME",
    manager: "Mira",
    pages: 3,
    words: 1200,
    characters: 6400,
    lines: 80,
    paragraphs: 12,
  },
}
```

- [x] **Step 2: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "document app properties"`

Expected: FAIL because docProps/app.xml is not supported yet.

- [x] **Step 3: Implement minimal support**

Emit and parse `docProps/app.xml` with standard extended properties namespace.

- [x] **Step 4: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "document app properties"`

Expected: PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 33 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

Run: `git diff --check`

Expected: no whitespace errors.

- [x] **Step 2: Commit**

```bash
git add docs/superpowers/plans/2026-07-04-phase-33-document-properties.md src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts
git commit -m "feat: add phase 33 document properties"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-33-document-properties
```

Expected: branch `phase-33-document-properties` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers core properties, app properties, relationships, content types, writer output, and reader round-trip.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: JSON properties are grouped under `properties.core` and `properties.app`.
