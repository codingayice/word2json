# Phase 16 Image DrawingML Fidelity Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve common DrawingML image properties that affect visual fidelity in image-heavy Word documents.

**Architecture:** Extend `ImageNode` with image crop, rotation, and floating layout metadata. The writer emits these values into the existing DrawingML picture tree and switches from `wp:inline` to `wp:anchor` when floating layout is requested; the reader parses both inline and anchor containers back into JSON.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser.

---

## File Structure

- `src/schema.ts`: add `ImageCrop` and `ImageFloatingLayout`.
- `src/docx-writer.ts`: emit `a:srcRect`, `a:xfrm rot`, and `wp:anchor` with position/wrap metadata.
- `src/docx-reader.ts`: parse image crop, rotation, and floating layout from DrawingML.
- `tests/docx-core.test.ts`: add XML writer and round-trip tests.

## Task 1: Image Crop

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add this test under `describe("DOCX writer", ...)`:

```ts
it("writes image crop", async () => {
  const imageData = Buffer.from("fake-png").toString("base64");
  const document = createDocumentJson([
    {
      type: "image",
      data: imageData,
      contentType: "image/png",
      width: 120,
      height: 80,
      crop: { left: 1000, top: 2000, right: 3000, bottom: 4000 },
    },
  ]);

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain('<a:srcRect l="1000" t="2000" r="3000" b="4000"/>');
});
```

- [x] **Step 2: Write failing round-trip test**

Add this test under `describe("DOCX reader", ...)`:

```ts
it("round-trips image crop", async () => {
  const imageData = Buffer.from("fake-png").toString("base64");
  const source = createDocumentJson([
    {
      type: "image",
      data: imageData,
      contentType: "image/png",
      width: 120,
      height: 80,
      crop: { left: 1000, top: 2000, right: 3000, bottom: 4000 },
    },
  ]);

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "image crop"`

Expected: FAIL because image crop is not supported yet.

- [x] **Step 4: Implement minimal support**

Add `crop?: ImageCrop` to `ImageNode`. Emit `a:srcRect` inside `pic:blipFill` and parse it back from the picture tree.

- [x] **Step 5: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "image crop"`

Expected: PASS.

## Task 2: Image Rotation

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add this test under `describe("DOCX writer", ...)`:

```ts
it("writes image rotation", async () => {
  const imageData = Buffer.from("fake-png").toString("base64");
  const document = createDocumentJson([
    {
      type: "image",
      data: imageData,
      contentType: "image/png",
      width: 120,
      height: 80,
      rotation: 15,
    },
  ]);

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain('<a:xfrm rot="900000">');
});
```

- [x] **Step 2: Write failing round-trip test**

Add this test under `describe("DOCX reader", ...)`:

```ts
it("round-trips image rotation", async () => {
  const imageData = Buffer.from("fake-png").toString("base64");
  const source = createDocumentJson([
    {
      type: "image",
      data: imageData,
      contentType: "image/png",
      width: 120,
      height: 80,
      rotation: 15,
    },
  ]);

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "image rotation"`

Expected: FAIL because image rotation is not supported yet.

- [x] **Step 4: Implement minimal support**

Add `rotation?: number` in degrees to `ImageNode`. Emit DrawingML `rot` as degrees multiplied by 60000 and parse it back as degrees.

- [x] **Step 5: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "image rotation"`

Expected: PASS.

## Task 3: Floating Image Layout

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`
- Test: `tests/docx-core.test.ts`

- [x] **Step 1: Write failing writer test**

Add this test under `describe("DOCX writer", ...)`:

```ts
it("writes floating image layout", async () => {
  const imageData = Buffer.from("fake-png").toString("base64");
  const document = createDocumentJson([
    {
      type: "image",
      data: imageData,
      contentType: "image/png",
      width: 120,
      height: 80,
      floating: { wrap: "square", horizontalOffset: 1440, verticalOffset: 720 },
    },
  ]);

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain("<wp:anchor ");
  expect(xml).toContain("<wp:wrapSquare/>");
  expect(xml).toContain('<wp:positionH relativeFrom="page"><wp:posOffset>1440</wp:posOffset></wp:positionH>');
  expect(xml).toContain('<wp:positionV relativeFrom="page"><wp:posOffset>720</wp:posOffset></wp:positionV>');
});
```

- [x] **Step 2: Write failing round-trip test**

Add this test under `describe("DOCX reader", ...)`:

```ts
it("round-trips floating image layout", async () => {
  const imageData = Buffer.from("fake-png").toString("base64");
  const source = createDocumentJson([
    {
      type: "image",
      data: imageData,
      contentType: "image/png",
      width: 120,
      height: 80,
      floating: { wrap: "square", horizontalOffset: 1440, verticalOffset: 720 },
    },
  ]);

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [x] **Step 3: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "floating image layout"`

Expected: FAIL because floating image layout is not supported yet.

- [x] **Step 4: Implement minimal support**

Add `floating?: ImageFloatingLayout` to `ImageNode`. Emit `wp:anchor` instead of `wp:inline` when present, include square wrapping and page-relative offsets, and parse anchor containers back into `floating`.

- [x] **Step 5: Run targeted tests**

Run: `npm test -- tests/docx-core.test.ts -t "floating image layout"`

Expected: PASS.

## Task 4: Full Verification and Push

**Files:**
- Modify: all Phase 16 files

- [x] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [x] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 16 image drawingml fidelity"
```

- [x] **Step 3: Push**

```bash
git push -u origin phase-16-image-drawingml-fidelity
```

Expected: branch `phase-16-image-drawingml-fidelity` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan covers image crop, rotation, and floating layout/wrap metadata.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: `crop`, `rotation`, and `floating` live on `ImageNode`.
