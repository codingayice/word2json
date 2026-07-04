# Phase 70 Radical Control Properties Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:radPr/m:ctrlPr` support for `radical` math nodes so radical object formatting survives JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Reuse the existing `MathControlProperties`, writer `mathControlPropertiesXml`, and reader `parseMathControlProperties` helpers. Extend only the `radical` node with `controlProperties` and map it to the real OMML property container `m:radPr`.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `controlProperties` on `radical`.
- Modify `src/docx-writer.ts`: emit `radPr/ctrlPr` when `radical.controlProperties` is present.
- Modify `src/docx-reader.ts`: parse `radPr.ctrlPr.rPr` into `controlProperties`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add this test near the other writer-side Office Math tests:

```ts
  it("writes radical control properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "radical" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  degree: [{ type: "text" as const, text: "3" }],
                  content: [{ type: "text" as const, text: "x" }],
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

    expect(xml).toContain('<m:rad><m:radPr><m:ctrlPr><w:rPr><w:b/><w:i/><w:u w:val="single"/><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:sz w:val="32"/><w:color w:val="8064A2"/><w:highlight w:val="darkYellow"/></w:rPr></m:ctrlPr></m:radPr><m:deg><m:r><m:t>3</m:t></m:r></m:deg><m:e><m:r><m:t>x</m:t></m:r></m:e></m:rad>');
  });
```

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes radical control properties"
```

Expected: FAIL because `radical.controlProperties` is ignored and no `m:radPr` XML is emitted.

- [x] **Step 3: Extend schema**

Update the `radical` node:

```ts
  | { type: "radical"; controlProperties?: MathControlProperties; degree?: MathNode[]; content: MathNode[] }
```

- [x] **Step 4: Implement writer support**

Update the `radical` writer branch:

```ts
    const properties = mathControlPropertiesXml(node.controlProperties);
    const degree = node.degree ? `<m:deg>${node.degree.map((child) => mathNodeXml(child)).join("")}</m:deg>` : "";
    return `<m:rad>${properties ? `<m:radPr><m:ctrlPr>${properties}</m:ctrlPr></m:radPr>` : ""}${degree}<m:e>${node.content.map((child) => mathNodeXml(child)).join("")}</m:e></m:rad>`;
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes radical control properties"
```

Expected: PASS with the expected `m:radPr/m:ctrlPr/w:rPr` XML.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add this test near the other reader-side Office Math round-trip tests:

```ts
  it("round-trips radical control properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "radical" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 16,
                    color: "8064A2",
                    underline: true,
                    highlight: "darkYellow",
                  },
                  degree: [{ type: "text" as const, text: "3" }],
                  content: [{ type: "text" as const, text: "x" }],
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
npm test -- tests/docx-core.test.ts -t "round-trips radical control properties"
```

Expected: FAIL because `parseMathNodes` ignores `radPr.ctrlPr.rPr`.

- [x] **Step 3: Implement reader support**

Update the `radical` reader branch:

```ts
        const controlProperties = parseMathControlProperties(asObject(asObject(radicalNode.radPr).ctrlPr).rPr);
```

and include:

```ts
          ...(controlProperties ? { controlProperties } : {}),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips radical control properties"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-70-radical-control-properties.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "radical control properties"
```

Expected: PASS for both Phase 70 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-70-radical-control-properties.md
git commit -m "feat: add phase 70 radical control properties"
```

- [x] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-70-radical-control-properties
```

- [x] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-70-radical-control-properties.md
git commit -m "docs: mark phase 70 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 70 covers `m:radPr/m:ctrlPr/w:rPr` writer and reader round-trip for common math control properties.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: The property is consistently named `controlProperties` and reuses `MathControlProperties`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes radical control properties"` failed because `m:radPr/m:ctrlPr` was not emitted.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes radical control properties"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips radical control properties"` failed because parsed JSON did not include `controlProperties`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips radical control properties"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "radical control properties"` passed 2 tests, 230 skipped.
- Full suite: `npm test` passed 232 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.

### Push Record

- Branch: `phase-70-radical-control-properties`
- Implementation commit: `6cc53b1 feat: add phase 70 radical control properties`
- Remote: `origin/phase-70-radical-control-properties`
