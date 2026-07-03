# Phase 29 Content Control Data Binding Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve Word content control XML data binding metadata in JSON to DOCX writing and DOCX to JSON round-tripping.

**Architecture:** Extend the shared `ContentControl` type with a `dataBinding` object because block and run SDTs already use the same content control serialization path. The writer emits `w:dataBinding` inside `w:sdtPr`; the reader parses `storeItemID`, `xpath`, and `prefixMappings` from the same node through `parseContentControl()`.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add `ContentControlDataBinding` and `ContentControl.dataBinding`.
- `src/docx-writer.ts`: emit `w:dataBinding` metadata inside content control properties.
- `src/docx-reader.ts`: parse `w:dataBinding` into JSON.
- `tests/docx-core.test.ts`: add writer and round-trip tests for data binding metadata.

## Task 1: Writer Support

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add a test named `writes content control data bindings`:

```ts
it("writes content control data bindings", async () => {
  const document = createDocumentJson([
    {
      type: "paragraph",
      runs: [
        {
          text: "Ada Lovelace",
          contentControl: {
            alias: "Customer",
            tag: "customer",
            dataBinding: {
              storeItemId: "{11111111-2222-3333-4444-555555555555}",
              xpath: "/customer/name[1]",
              prefixMappings: "xmlns:crm='urn:crm'",
            },
          },
        },
      ],
    },
  ]);

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain('<w:dataBinding w:storeItemID="{11111111-2222-3333-4444-555555555555}" w:xpath="/customer/name[1]" w:prefixMappings="xmlns:crm=&apos;urn:crm&apos;"/>');
});
```

- [x] **Step 2: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "content control data bindings"`

Expected: FAIL because `dataBinding` is not part of the schema and writer yet.

- [x] **Step 3: Implement minimal writer support**

Add:

```ts
export type ContentControlDataBinding = {
  storeItemId?: string;
  xpath?: string;
  prefixMappings?: string;
};
```

and `dataBinding?: ContentControlDataBinding` to `ContentControl`.

Emit:

```xml
<w:dataBinding w:storeItemID="..." w:xpath="..." w:prefixMappings="..."/>
```

with only provided attributes.

- [x] **Step 4: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "content control data bindings"`

Expected: writer test PASS.

## Task 2: Reader Round-Trip Support

**Files:**
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing round-trip test**

Add a test named `round-trips content control data bindings`:

```ts
it("round-trips content control data bindings", async () => {
  const source = createDocumentJson([
    {
      type: "paragraph",
      runs: [
        {
          text: "Ada Lovelace",
          contentControl: {
            alias: "Customer",
            tag: "customer",
            dataBinding: {
              storeItemId: "{11111111-2222-3333-4444-555555555555}",
              xpath: "/customer/name[1]",
              prefixMappings: "xmlns:crm='urn:crm'",
            },
          },
        },
      ],
    },
  ]);

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 2: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "content control data bindings"`

Expected: FAIL because the reader drops data binding metadata.

- [x] **Step 3: Implement minimal reader support**

Parse `properties.dataBinding` into:

```ts
{
  storeItemId: dataBinding.storeItemID,
  xpath: dataBinding.xpath,
  prefixMappings: dataBinding.prefixMappings,
}
```

omitting keys that are not strings.

- [x] **Step 4: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "content control data bindings"`

Expected: writer and round-trip tests PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 29 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

Run: `git diff --check`

Expected: no whitespace errors.

- [x] **Step 2: Commit**

```bash
git add docs/superpowers/plans/2026-07-04-phase-29-content-control-data-binding.md src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts
git commit -m "feat: add phase 29 content control data binding"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-29-content-control-data-binding
```

Expected: branch `phase-29-content-control-data-binding` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers the three key Word `w:dataBinding` attributes: `storeItemID`, `xpath`, and `prefixMappings`.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: JSON uses `storeItemId` while writer/reader map it to the Word XML attribute `storeItemID`.
