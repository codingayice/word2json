# Phase 62 Pre Sub Sup Math Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity Office Math `m:preSubSup` support so JSON can represent and round-trip front/left-side subscripts and superscripts.

**Architecture:** Follow the existing OMML math pattern: add one discriminated `MathNode` variant, serialize it in `mathNodeXml`, parse it in `parseMathNodes`, and cover both direct XML output and full DOCX round-trip. The node shape mirrors OOXML child names while keeping JSON property names readable.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: extend `MathNode` with `{ type: "preSubSup"; base; subscript; superscript }`.
- Modify `src/docx-writer.ts`: add `m:preSubSup` serialization branch before the final `nary` fallback.
- Modify `src/docx-reader.ts`: add `m:preSubSup` parsing in `parseMathNodes`.
- Modify `tests/docx-core.test.ts`: add a writer XML test and a DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add this test near the other writer-side Office Math tests:

```ts
  it("writes pre sub sup office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "preSubSup" as const,
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

    expect(xml).toContain('<m:preSubSup><m:e><m:r><m:t>X</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub><m:sup><m:r><m:t>j</m:t></m:r></m:sup></m:preSubSup>');
  });
```

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes pre sub sup office math runs"
```

Expected: FAIL because TypeScript does not yet accept `type: "preSubSup"` or writer falls through to the wrong OMML.

- [x] **Step 3: Extend schema**

Add the variant to `MathNode` in `src/schema.ts`:

```ts
  | { type: "preSubSup"; base: MathNode[]; subscript: MathNode[]; superscript: MathNode[] }
```

- [x] **Step 4: Implement writer branch**

Add this branch in `mathNodeXml` after the `subscript` branch and before `radical`:

```ts
  if (node.type === "preSubSup") {
    return `<m:preSubSup><m:e>${node.base.map((child) => mathNodeXml(child)).join("")}</m:e><m:sub>${node.subscript.map((child) => mathNodeXml(child)).join("")}</m:sub><m:sup>${node.superscript.map((child) => mathNodeXml(child)).join("")}</m:sup></m:preSubSup>`;
  }
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes pre sub sup office math runs"
```

Expected: PASS with the expected `m:preSubSup` XML present in `word/document.xml`.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add this test near the other reader-side Office Math round-trip tests:

```ts
  it("round-trips pre sub sup office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "preSubSup" as const,
                  base: [{ type: "text" as const, text: "T" }],
                  subscript: [
                    {
                      type: "fraction" as const,
                      numerator: [{ type: "text" as const, text: "i" }],
                      denominator: [{ type: "text" as const, text: "n" }],
                    },
                  ],
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
npm test -- tests/docx-core.test.ts -t "round-trips pre sub sup office math runs"
```

Expected: FAIL because `parseMathNodes` does not yet map `m:preSubSup` back into JSON.

- [x] **Step 3: Implement reader branch**

Add this branch in `parseMathNodes` after `container.sSub` and before `container.rad`:

```ts
    ...asArray(container.preSubSup)
      .map((preSubSup) => {
        const preSubSupNode = asObject(preSubSup);
        return {
          type: "preSubSup" as const,
          base: parseMathNodes(asObject(preSubSupNode.e)),
          subscript: parseMathNodes(asObject(preSubSupNode.sub)),
          superscript: parseMathNodes(asObject(preSubSupNode.sup)),
        };
      }),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips pre sub sup office math runs"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-62-pre-sub-sup-math.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "pre sub sup office math runs"
```

Expected: PASS for both Phase 62 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-62-pre-sub-sup-math.md
git commit -m "feat: add phase 62 pre sub sup math"
```

- [x] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-62-pre-sub-sup-math
```

- [x] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-62-pre-sub-sup-math.md
git commit -m "docs: mark phase 62 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 62 adds one focused high-fidelity OMML construct for both JSON-to-DOCX and DOCX-to-JSON.
- Placeholder scan: No placeholders or deferred implementation notes.
- Type consistency: The node type is consistently named `preSubSup`, with `base`, `subscript`, and `superscript` arrays in schema, writer, reader, and tests.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes pre sub sup office math runs"` failed with `Cannot read properties of undefined (reading 'map')` because writer fell through to the `nary` fallback.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes pre sub sup office math runs"` passed 1 targeted test.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips pre sub sup office math runs"` failed because parsed runs were empty.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips pre sub sup office math runs"` passed 1 targeted test.
- Targeted phase verification: `npm test -- tests/docx-core.test.ts -t "pre sub sup office math runs"` passed 2 targeted tests.
- Full suite: `npm test` passed 216 tests.
- Build: `npm run build` exited 0.
- Diff check: `git diff --check` exited 0 with LF-to-CRLF warnings only.

### Push Record

- Branch: `phase-62-pre-sub-sup-math`
- Implementation commit: `ba79175 feat: add phase 62 pre sub sup math`
- Remote: `origin/phase-62-pre-sub-sup-math`
