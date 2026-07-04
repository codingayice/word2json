# Phase 65 S Pre Control Properties Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:sPrePr/m:ctrlPr` support for `sPre` math nodes so common formula object run properties survive JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Extend the `sPre` math node with `controlProperties?: MathControlProperties`, a small reusable subset of `TextRun` formatting that maps to `w:rPr` inside `m:ctrlPr`. The writer emits `<m:sPrePr><m:ctrlPr><w:rPr>...</w:rPr></m:ctrlPr></m:sPrePr>` before the `sub/sup/e` children, and the reader parses `sPrePr.ctrlPr.rPr` using the existing run-property parsing helpers.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: add `MathControlProperties` and allow `controlProperties` on `sPre`.
- Modify `src/docx-writer.ts`: emit `sPrePr/ctrlPr` when `sPre.controlProperties` is present.
- Modify `src/docx-reader.ts`: parse `sPrePr.ctrlPr.rPr` into `controlProperties`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add this test near the other writer-side Office Math tests:

```ts
  it("writes s pre control properties", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "sPre" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 14,
                    color: "C00000",
                    underline: true,
                    highlight: "yellow",
                  },
                  base: [{ type: "text" as const, text: "X" }],
                  subscript: [{ type: "text" as const, text: "i" }],
                  superscript: [{ type: "text" as const, text: "j" }],
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

    expect(xml).toContain('<m:sPre><m:sPrePr><m:ctrlPr><w:rPr><w:b/><w:i/><w:u w:val="single"/><w:rFonts w:ascii="Cambria Math" w:hAnsi="Cambria Math"/><w:sz w:val="28"/><w:color w:val="C00000"/><w:highlight w:val="yellow"/></w:rPr></m:ctrlPr></m:sPrePr><m:sub><m:r><m:t>i</m:t></m:r></m:sub><m:sup><m:r><m:t>j</m:t></m:r></m:sup><m:e><m:r><m:t>X</m:t></m:r></m:e></m:sPre>');
  });
```

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes s pre control properties"
```

Expected: FAIL because `sPre.controlProperties` is ignored and no `m:sPrePr` XML is emitted.

- [x] **Step 3: Extend schema**

Add the type and update the `sPre` node:

```ts
export type MathControlProperties = Pick<StyleRunProperties, "bold" | "italic" | "underline" | "fontFamily" | "fontSize" | "color" | "highlight"> & {
  styleId?: string;
};
```

```ts
  | { type: "sPre"; controlProperties?: MathControlProperties; base: MathNode[]; subscript: MathNode[]; superscript: MathNode[] }
```

- [x] **Step 4: Implement writer branch**

Update `sPre` writer branch to include:

```ts
    const properties = mathControlPropertiesXml(node.controlProperties);
    return `<m:sPre>${properties ? `<m:sPrePr><m:ctrlPr>${properties}</m:ctrlPr></m:sPrePr>` : ""}<m:sub>${node.subscript.map((child) => mathNodeXml(child)).join("")}</m:sub><m:sup>${node.superscript.map((child) => mathNodeXml(child)).join("")}</m:sup><m:e>${node.base.map((child) => mathNodeXml(child)).join("")}</m:e></m:sPre>`;
```

Add helper near `mathNodeXml`:

```ts
function mathControlPropertiesXml(properties: MathControlProperties | undefined): string {
  return properties ? runPropertiesXml({ text: "", ...properties }) : "";
}
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes s pre control properties"
```

Expected: PASS with the expected `m:sPrePr/m:ctrlPr/w:rPr` XML.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add this test near the other reader-side Office Math round-trip tests:

```ts
  it("round-trips s pre control properties", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "sPre" as const,
                  controlProperties: {
                    bold: true,
                    italic: true,
                    fontFamily: "Cambria Math",
                    fontSize: 14,
                    color: "C00000",
                    underline: true,
                    highlight: "yellow",
                  },
                  base: [{ type: "text" as const, text: "T" }],
                  subscript: [{ type: "text" as const, text: "i" }],
                  superscript: [{ type: "text" as const, text: "j" }],
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
npm test -- tests/docx-core.test.ts -t "round-trips s pre control properties"
```

Expected: FAIL because `parseMathNodes` ignores `sPrePr.ctrlPr.rPr`.

- [x] **Step 3: Implement reader support**

Add helper near math parsing:

```ts
function parseMathControlProperties(value: unknown): MathControlProperties | undefined {
  const properties = asObject(value);
  const parsed = {
    ...parseRunStyle(properties),
    ...(properties.b !== undefined ? { bold: true } : {}),
    ...(properties.i !== undefined ? { italic: true } : {}),
    ...(properties.u !== undefined ? { underline: true } : {}),
    ...parseRunFont(properties),
  };
  return Object.keys(parsed).length > 0 ? parsed : undefined;
}
```

Update the `sPre` reader branch:

```ts
        const controlProperties = parseMathControlProperties(asObject(asObject(sPreNode.sPrePr).ctrlPr).rPr);
```

and include:

```ts
          ...(controlProperties ? { controlProperties } : {}),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips s pre control properties"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-65-s-pre-control-properties.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "s pre control properties"
```

Expected: PASS for both Phase 65 tests.

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

- [ ] **Step 5: Commit implementation**

Run:

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-65-s-pre-control-properties.md
git commit -m "feat: add phase 65 s pre control properties"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-65-s-pre-control-properties
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-65-s-pre-control-properties.md
git commit -m "docs: mark phase 65 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 65 covers `m:sPrePr/m:ctrlPr/w:rPr` writer and reader round-trip for the most common control properties.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: The property is consistently named `controlProperties` and uses the same subset in schema, writer, reader, and tests.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes s pre control properties"` failed because `m:sPrePr/m:ctrlPr` was missing from the generated XML.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes s pre control properties"` passed 1 targeted test.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips s pre control properties"` failed because `controlProperties` was missing from parsed JSON.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips s pre control properties"` passed 1 targeted test.
- Targeted phase verification: `npm test -- tests/docx-core.test.ts -t "s pre control properties"` passed 2 targeted tests.
- Full suite: `npm test` passed 222 tests.
- Build: `npm run build` exited 0.
- Diff check: `git diff --check` exited 0 with LF-to-CRLF warnings only.
