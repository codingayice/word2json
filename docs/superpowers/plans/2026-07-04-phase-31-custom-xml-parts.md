# Phase 31 Custom XML Parts Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve DOCX custom XML data parts and their package relationships so content control data bindings can keep their backing data source.

**Architecture:** Add top-level `customXmlParts` to `DocumentJson`. The writer stores custom XML under `customXml/itemN.xml`, adds package relationships from `/_rels/.rels`, optionally writes `customXml/itemPropsN.xml`, and links item to properties through `customXml/_rels/itemN.xml.rels`; the reader follows package customXml relationships and reconstructs the same JSON.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add `CustomXmlPart` and optional `DocumentJson.customXmlParts`.
- `src/docx-writer.ts`: write custom XML parts, custom XML properties, package relationships, item relationships, and content type overrides.
- `src/docx-reader.ts`: parse package custom XML relationships, item XML, item properties, and item-level relationships.
- `tests/docx-core.test.ts`: add writer and round-trip tests for custom XML parts.

## Task 1: Writer Support

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add a test named `writes custom xml parts with relationships`:

```ts
it("writes custom xml parts with relationships", async () => {
  const document = {
    ...createDocumentJson([]),
    customXmlParts: [
      {
        path: "customXml/item1.xml",
        xml: "<customer><name>Ada Lovelace</name></customer>",
        properties: {
          path: "customXml/itemProps1.xml",
          storeItemId: "{11111111-2222-3333-4444-555555555555}",
        },
      },
    ],
  };

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const packageRels = await zip.file("_rels/.rels")!.async("string");
  const item = await zip.file("customXml/item1.xml")!.async("string");
  const itemRels = await zip.file("customXml/_rels/item1.xml.rels")!.async("string");
  const itemProps = await zip.file("customXml/itemProps1.xml")!.async("string");
  const contentTypes = await zip.file("[Content_Types].xml")!.async("string");

  expect(packageRels).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"');
  expect(packageRels).toContain('Target="customXml/item1.xml"');
  expect(item).toBe("<customer><name>Ada Lovelace</name></customer>");
  expect(itemRels).toContain('Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"');
  expect(itemRels).toContain('Target="itemProps1.xml"');
  expect(itemProps).toContain('<ds:datastoreItem ds:itemID="{11111111-2222-3333-4444-555555555555}"');
  expect(contentTypes).toContain('PartName="/customXml/itemProps1.xml"');
  expect(contentTypes).toContain('ContentType="application/vnd.openxmlformats-officedocument.customXmlProperties+xml"');
});
```

- [x] **Step 2: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "custom xml parts"`

Expected: FAIL because custom XML parts are not written.

- [x] **Step 3: Implement minimal writer support**

Add:

```ts
export type CustomXmlPart = {
  path: string;
  xml: string;
  properties?: {
    path?: string;
    storeItemId?: string;
  };
};
```

and `customXmlParts?: CustomXmlPart[]` to `DocumentJson`.

Write package and item relationships using stable IDs:

```xml
<Relationship Id="rIdCustomXml1" Type=".../customXml" Target="customXml/item1.xml"/>
<Relationship Id="rIdCustomXmlProps1" Type=".../customXmlProps" Target="itemProps1.xml"/>
```

- [x] **Step 4: Run targeted writer test**

Run: `npm test -- tests/docx-core.test.ts -t "custom xml parts"`

Expected: writer test PASS.

## Task 2: Reader Round-Trip Support

**Files:**
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing round-trip test**

Add a test named `round-trips custom xml parts`:

```ts
it("round-trips custom xml parts", async () => {
  const source = {
    ...createDocumentJson([]),
    customXmlParts: [
      {
        path: "customXml/item1.xml",
        xml: "<customer><name>Ada Lovelace</name></customer>",
        properties: {
          path: "customXml/itemProps1.xml",
          storeItemId: "{11111111-2222-3333-4444-555555555555}",
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

Run: `npm test -- tests/docx-core.test.ts -t "custom xml parts"`

Expected: FAIL because the reader drops custom XML parts.

- [x] **Step 3: Implement minimal reader support**

Read package relationships whose type is customXml, load each target part, then load any item-level `customXmlProps` relationship and parse `ds:itemID` from the properties part.

- [x] **Step 4: Run targeted round-trip test**

Run: `npm test -- tests/docx-core.test.ts -t "custom xml parts"`

Expected: writer and round-trip tests PASS.

## Task 3: Full Verification and Push

**Files:**
- Modify: all Phase 31 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

Run: `git diff --check`

Expected: no whitespace errors.

- [x] **Step 2: Commit**

```bash
git add docs/superpowers/plans/2026-07-04-phase-31-custom-xml-parts.md src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts
git commit -m "feat: add phase 31 custom xml parts"
```

- [ ] **Step 3: Push**

```bash
git push -u origin phase-31-custom-xml-parts
```

Expected: branch `phase-31-custom-xml-parts` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan preserves custom XML item parts, properties parts, package relationships, item relationships, and properties content type overrides.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: `storeItemId` maps to `ds:itemID` in the custom XML properties part.
