# Phase 32 Custom XML Schema Refs Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve `ds:schemaRef` entries inside custom XML properties parts.

**Architecture:** Extend `CustomXmlPartProperties` with `schemaRefs: string[]`. The writer emits each schema URI as `ds:schemaRef ds:uri="..."` inside `ds:schemaRefs`; the reader parses single or multiple schema references from the custom XML properties part back into JSON.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add optional `schemaRefs` to `CustomXmlPartProperties`.
- `src/docx-writer.ts`: emit schema references in custom XML properties.
- `src/docx-reader.ts`: parse schema references from custom XML properties.
- `tests/docx-core.test.ts`: add writer and round-trip tests for schema refs.

## Task 1: Writer Support

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add a test named `writes custom xml schema refs`:

```ts
it("writes custom xml schema refs", async () => {
  const document = {
    ...createDocumentJson([]),
    customXmlParts: [
      {
        path: "customXml/item1.xml",
        xml: "<customer><name>Ada Lovelace</name></customer>",
        properties: {
          path: "customXml/itemProps1.xml",
          storeItemId: "{11111111-2222-3333-4444-555555555555}",
          schemaRefs: ["urn:customer", "urn:crm"],
        },
      },
    ],
  };

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const itemProps = await zip.file("customXml/itemProps1.xml")!.async("string");

  expect(itemProps).toContain('<ds:schemaRef ds:uri="urn:customer"/>');
  expect(itemProps).toContain('<ds:schemaRef ds:uri="urn:crm"/>');
});
```

- [x] **Step 2: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "custom xml schema refs"`

Expected: FAIL because schema refs are not written.

- [x] **Step 3: Implement minimal writer support**

Add `schemaRefs?: string[]` to `CustomXmlPartProperties` and emit each item as:

```xml
<ds:schemaRef ds:uri="urn:customer"/>
```

- [x] **Step 4: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "custom xml schema refs"`

Expected: writer test PASS.

## Task 2: Reader Round-Trip Support

**Files:**
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing round-trip test**

Add a test named `round-trips custom xml schema refs`:

```ts
it("round-trips custom xml schema refs", async () => {
  const source = {
    ...createDocumentJson([]),
    customXmlParts: [
      {
        path: "customXml/item1.xml",
        xml: "<customer><name>Ada Lovelace</name></customer>",
        properties: {
          path: "customXml/itemProps1.xml",
          storeItemId: "{11111111-2222-3333-4444-555555555555}",
          schemaRefs: ["urn:customer", "urn:crm"],
        },
      },
    ],
  };

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed.customXmlParts).toEqual(source.customXmlParts);
});
```

- [x] **Step 2: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "custom xml schema refs"`

Expected: FAIL because the reader drops schema refs.

- [x] **Step 3: Implement minimal reader support**

Parse `datastoreItem.schemaRefs.schemaRef` into `schemaRefs` when one or more string `uri` values are present.

- [x] **Step 4: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "custom xml schema refs"`

Expected: writer and round-trip tests PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 32 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

Run: `git diff --check`

Expected: no whitespace errors.

- [x] **Step 2: Commit**

```bash
git add docs/superpowers/plans/2026-07-04-phase-32-custom-xml-schema-refs.md src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts
git commit -m "feat: add phase 32 custom xml schema refs"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-32-custom-xml-schema-refs
```

Expected: branch `phase-32-custom-xml-schema-refs` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan preserves multiple custom XML schema references.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: JSON uses `schemaRefs`, and XML uses `ds:schemaRef ds:uri`.
