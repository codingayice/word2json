# Phase 63 Sub Sup Math Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity Office Math `m:sSubSup` support so JSON can represent and round-trip right-side subscripts and superscripts attached to one base.

**Architecture:** Follow the existing OMML math pattern: one `MathNode` discriminated union member, one writer branch in `mathNodeXml`, one reader branch in `parseMathNodes`, and targeted tests for XML output plus DOCX round-trip. The JSON node is named `subSup` to distinguish it from Phase 62 `preSubSup` while preserving familiar `base`, `subscript`, and `superscript` fields.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: extend `MathNode` with `{ type: "subSup"; base; subscript; superscript }`.
- Modify `src/docx-writer.ts`: add `m:sSubSup` serialization branch near `sSup`, `sSub`, and `preSubSup`.
- Modify `src/docx-reader.ts`: add `m:sSubSup` parsing in `parseMathNodes`.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add this test near the other writer-side Office Math tests:

```ts
  it("writes sub sup office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "subSup" as const,
                  base: [{ type: "text" as const, text: "x" }],
                  subscript: [{ type: "text" as const, text: "i" }],
                  superscript: [{ type: "text" as const, text: "2" }],
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

    expect(xml).toContain('<m:sSubSup><m:e><m:r><m:t>x</m:t></m:r></m:e><m:sub><m:r><m:t>i</m:t></m:r></m:sub><m:sup><m:r><m:t>2</m:t></m:r></m:sup></m:sSubSup>');
  });
```

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes sub sup office math runs"
```

Expected: FAIL because `subSup` is not yet recognized by the writer and falls through to the final `nary` path.

- [x] **Step 3: Extend schema**

Add the variant to `MathNode` in `src/schema.ts`:

```ts
  | { type: "subSup"; base: MathNode[]; subscript: MathNode[]; superscript: MathNode[] }
```

- [x] **Step 4: Implement writer branch**

Add this branch in `mathNodeXml` after the `subscript` branch and before `preSubSup`:

```ts
  if (node.type === "subSup") {
    return `<m:sSubSup><m:e>${node.base.map((child) => mathNodeXml(child)).join("")}</m:e><m:sub>${node.subscript.map((child) => mathNodeXml(child)).join("")}</m:sub><m:sup>${node.superscript.map((child) => mathNodeXml(child)).join("")}</m:sup></m:sSubSup>`;
  }
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes sub sup office math runs"
```

Expected: PASS with the expected `m:sSubSup` XML present in `word/document.xml`.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add this test near the other reader-side Office Math round-trip tests:

```ts
  it("round-trips sub sup office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "subSup" as const,
                  base: [{ type: "text" as const, text: "x" }],
                  subscript: [{ type: "text" as const, text: "i" }],
                  superscript: [
                    {
                      type: "fraction" as const,
                      numerator: [{ type: "text" as const, text: "m" }],
                      denominator: [{ type: "text" as const, text: "n" }],
                    },
                  ],
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
npm test -- tests/docx-core.test.ts -t "round-trips sub sup office math runs"
```

Expected: FAIL because `parseMathNodes` does not yet map `m:sSubSup` back into JSON.

- [x] **Step 3: Implement reader branch**

Add this branch in `parseMathNodes` after `container.sSub` and before `container.preSubSup`:

```ts
    ...asArray(container.sSubSup)
      .map((subSup) => {
        const subSupNode = asObject(subSup);
        return {
          type: "subSup" as const,
          base: parseMathNodes(asObject(subSupNode.e)),
          subscript: parseMathNodes(asObject(subSupNode.sub)),
          superscript: parseMathNodes(asObject(subSupNode.sup)),
        };
      }),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips sub sup office math runs"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-63-sub-sup-math.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "sub sup office math runs"
```

Expected: PASS for both Phase 63 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-63-sub-sup-math.md
git commit -m "feat: add phase 63 sub sup math"
```

- [x] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-63-sub-sup-math
```

- [x] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-63-sub-sup-math.md
git commit -m "docs: mark phase 63 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 63 adds both JSON-to-DOCX and DOCX-to-JSON support for the common OMML right-side subscript plus superscript construct.
- Placeholder scan: No placeholder text or deferred work remains.
- Type consistency: The node type is consistently `subSup`, with `base`, `subscript`, and `superscript` arrays in tests, schema, writer, and reader.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes sub sup office math runs"` failed with `Cannot read properties of undefined (reading 'map')` because writer fell through to the `nary` fallback.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes sub sup office math runs"` passed 1 targeted test.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips sub sup office math runs"` failed because parsed runs were empty.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips sub sup office math runs"` passed 1 targeted test.
- Targeted phase verification: `npm test -- tests/docx-core.test.ts -t "sub sup office math runs"` passed 4 related tests.
- Full suite: `npm test` passed 218 tests.
- Build: `npm run build` exited 0.
- Diff check: `git diff --check` exited 0 with LF-to-CRLF warnings only.

### Push Record

- Branch: `phase-63-sub-sup-math`
- Implementation commit: `e74d608 feat: add phase 63 sub sup math`
- Remote: `origin/phase-63-sub-sup-math`
