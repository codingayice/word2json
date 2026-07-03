# Phase 51 Matrix Math Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve Office Math matrices so common vectors, matrices, and equation arrays can round-trip through JSON.

**Architecture:** Extend the recursive `MathNode` union with a `matrix` node whose rows contain cells of nested `MathNode[]`. Serialize the node into OMML `m:m` with `m:mr` rows and `m:e` cells, then parse those elements back while preserving existing math node behavior.

**Tech Stack:** TypeScript, JSZip, fast-xml-parser, Vitest, OOXML OMML.

---

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add this test after `writes radical and n-ary office math runs` in `tests/docx-core.test.ts`:

```ts
  it("writes matrix office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "matrix" as const,
                  rows: [
                    [[{ type: "text" as const, text: "a" }], [{ type: "text" as const, text: "b" }]],
                    [[{ type: "text" as const, text: "c" }], [{ type: "text" as const, text: "d" }]],
                  ],
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

    expect(xml).toContain("<m:m><m:mr><m:e><m:r><m:t>a</m:t></m:r></m:e><m:e><m:r><m:t>b</m:t></m:r></m:e></m:mr><m:mr><m:e><m:r><m:t>c</m:t></m:r></m:e><m:e><m:r><m:t>d</m:t></m:r></m:e></m:mr></m:m>");
  });
```

- [x] **Step 2: Run the writer test to verify RED**

Run: `npm test -- tests/docx-core.test.ts -t "writes matrix office math runs"`

Expected: FAIL because `MathNode` has no `matrix` node and the writer cannot serialize it.

- [x] **Step 3: Add the minimal writer implementation**

In `src/schema.ts`, extend `MathNode` with:

```ts
  | { type: "matrix"; rows: MathNode[][][] };
```

In `src/docx-writer.ts`, add a `mathNodeXml` branch:

```ts
  if (node.type === "matrix") {
    return `<m:m>${node.rows.map((row) => `<m:mr>${row.map((cell) => `<m:e>${cell.map((child) => mathNodeXml(child)).join("")}</m:e>`).join("")}</m:mr>`).join("")}</m:m>`;
  }
```

- [x] **Step 4: Run the writer test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "writes matrix office math runs"`

Expected: PASS.

### Task 2: Reader Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add this test after `round-trips radical and n-ary office math runs` in `tests/docx-core.test.ts`:

```ts
  it("round-trips matrix office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "matrix" as const,
                  rows: [
                    [[{ type: "text" as const, text: "a" }], [{ type: "text" as const, text: "b" }]],
                    [[{ type: "text" as const, text: "c" }], [{ type: "text" as const, text: "d" }]],
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

- [x] **Step 2: Run the round-trip test to verify RED**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips matrix office math runs"`

Expected: FAIL because parser omits `m:m` nodes.

- [x] **Step 3: Add the minimal reader implementation**

In `src/docx-reader.ts`, extend `parseMathNodes` with a `container.m` branch:

```ts
    ...asArray(container.m)
      .map((matrix) => {
        const matrixNode = asObject(matrix);
        return {
          type: "matrix" as const,
          rows: asArray(matrixNode.mr).map((row) =>
            asArray(asObject(row).e).map((cell) => parseMathNodes(asObject(cell))),
          ),
        };
      }),
```

- [x] **Step 4: Run the round-trip test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips matrix office math runs"`

Expected: PASS.

### Task 3: Full Verification and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-51-matrix-math.md`

- [x] **Step 1: Run full verification**

Run:

```powershell
npm test
npm run build
git diff --check
```

Expected: all tests pass, build exits 0, diff check exits 0 with only Windows line-ending warnings if any.

- [ ] **Step 2: Commit implementation**

Run:

```powershell
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-51-matrix-math.md
git commit -m "feat: add phase 51 matrix math"
```

- [ ] **Step 3: Push branch**

Run: `git push -u origin phase-51-matrix-math`

- [ ] **Step 4: Mark the plan pushed and commit docs status**

Append the pushed branch and commit hash to this plan, then run:

```powershell
git add docs/superpowers/plans/2026-07-04-phase-51-matrix-math.md
git commit -m "docs: mark phase 51 pushed"
git push
```
