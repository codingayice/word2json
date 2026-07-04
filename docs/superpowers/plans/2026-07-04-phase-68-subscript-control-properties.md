# Phase 68 Subscript Control Properties Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:sSubPr/m:ctrlPr` support for `subscript` math nodes so subscript object formatting survives JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Reuse the existing `MathControlProperties`, writer `mathControlPropertiesXml`, and reader `parseMathControlProperties` helpers. Extend only the `subscript` node with `controlProperties` and map it to the real OMML property container `m:sSubPr`.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `controlProperties` on `subscript`.
- Modify `src/docx-writer.ts`: emit `sSubPr/ctrlPr` when `subscript.controlProperties` is present.
- Modify `src/docx-reader.ts`: parse `sSubPr.ctrlPr.rPr` into `controlProperties`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add this test near the other writer-side Office Math tests:

```ts
  it("writes subscript control properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "subscript" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 11,
                    color: "00B050",
                    underline: true,
                    highlight: "magenta",
                  },
                  base: [{ type: "text" as const, text: "a" }],
                  subscript: [{ type: "text" as const, text: "i" }],
                },
              ],
            },
          },
        ],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<m:sSub><m:sSubPr><m:ctrlPr><w:rPr><w:b/><w:i/><w:u w:val="single"/><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:sz w:val="22"/><w:color w:val="00B050"/><w:highlight w:val="magenta"/></w:rPr></m:ctrlPr></m:sSubPr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub></m:sSub>');
  });
```

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes subscript control properties"
```

Expected: FAIL because `subscript.controlProperties` is ignored and no `m:sSubPr` XML is emitted.

- [x] **Step 3: Extend schema**

Update the `subscript` node:

```ts
  | { type: "subscript"; controlProperties?: MathControlProperties; base: MathNode[]; subscript: MathNode[] }
```

- [x] **Step 4: Implement writer support**

Update the `subscript` writer branch:

```ts
    const properties = mathControlPropertiesXml(node.controlProperties);
    return `<m:sSub>${properties ? `<m:sSubPr><m:ctrlPr>${properties}</m:ctrlPr></m:sSubPr>` : ""}<m:e>${node.base.map((child) => mathNodeXml(child)).join("")}</m:e><m:sub>${node.subscript.map((child) => mathNodeXml(child)).join("")}</m:sub></m:sSub>`;
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes subscript control properties"
```

Expected: PASS with the expected `m:sSubPr/m:ctrlPr/w:rPr` XML.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add this test near the other reader-side Office Math round-trip tests:

```ts
  it("round-trips subscript control properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "subscript" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 11,
                    color: "00B050",
                    underline: true,
                    highlight: "magenta",
                  },
                  base: [{ type: "text" as const, text: "a" }],
                  subscript: [{ type: "text" as const, text: "i" }],
                },
              ],
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

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips subscript control properties"
```

Expected: FAIL because `parseMathNodes` ignores `sSubPr.ctrlPr.rPr`.

- [x] **Step 3: Implement reader support**

Update the `subscript` reader branch:

```ts
        const controlProperties = parseMathControlProperties(asObject(asObject(subscriptNode.sSubPr).ctrlPr).rPr);
```

and include:

```ts
          ...(controlProperties ? { controlProperties } : {}),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips subscript control properties"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-68-subscript-control-properties.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "subscript control properties"
```

Expected: PASS for both Phase 68 tests.

- [x] **Step 2: Run full test suite**

Run:

```bash
npm test
```

Expected: all tests pass.

- [x] **Step 3: Run build**

Run:

```bash
npm run build
```

Expected: build exits with code 0.

- [x] **Step 4: Run whitespace check**

Run:

```bash
git diff --check
```

Expected: exit code 0.

- [x] **Step 5: Commit implementation**

Run:

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-68-subscript-control-properties.md
git commit -m "feat: add phase 68 subscript control properties"
```

- [x] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-68-subscript-control-properties
```

- [x] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-68-subscript-control-properties.md
git commit -m "docs: mark phase 68 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 68 covers `m:sSubPr/m:ctrlPr/w:rPr` writer and reader round-trip for common math control properties.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: The property is consistently named `controlProperties` and reuses `MathControlProperties`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes subscript control properties"` failed because `m:sSubPr/m:ctrlPr` was missing from generated XML.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes subscript control properties"` passed 1 targeted test.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips subscript control properties"` failed because `controlProperties` was missing from parsed JSON.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips subscript control properties"` passed 1 targeted test.
- Targeted phase verification: `npm test -- tests/docx-core.test.ts -t "subscript control properties"` passed 2 targeted tests.
- Full suite: `npm test` passed 228 tests.
- Build: `npm run build` exited 0.
- Diff check: `git diff --check` exited 0 with LF-to-CRLF warnings only.

### Push Record

- Branch: `phase-68-subscript-control-properties`
- Implementation commit: `58fe74a feat: add phase 68 subscript control properties`
- Remote: `origin/phase-68-subscript-control-properties`
