# Phase 57 Equation Array Math Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve Office Math equation arrays so multi-line formulas and aligned equation groups round-trip through JSON.

**Architecture:** Extend the recursive `MathNode` union with an `equationArray` node containing rows of nested math nodes. Serialize rows to OMML `m:eqArr` with one `m:e` element per row, then parse the same structure back from `m:oMath`.

**Tech Stack:** TypeScript, JSZip, fast-xml-parser, Vitest, OOXML OMML.

---

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add this test after `writes limit office math runs` in `tests/docx-core.test.ts`:

```ts
  it("writes equation array office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "equationArray" as const,
                  rows: [
                    [{ type: "text" as const, text: "x=1" }],
                    [
                      {
                        type: "fraction" as const,
                        numerator: [{ type: "text" as const, text: "a" }],
                        denominator: [{ type: "text" as const, text: "b" }],
                      },
                    ],
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

    expect(xml).toContain("<m:eqArr><m:e><m:r><m:t>x=1</m:t></m:r></m:e><m:e><m:f><m:num><m:r><m:t>a</m:t></m:r></m:num><m:den><m:r><m:t>b</m:t></m:r></m:den></m:f></m:e></m:eqArr>");
  });
```

- [x] **Step 2: Run the writer test to verify RED**

Run: `npm test -- tests/docx-core.test.ts -t "writes equation array office math runs"`

Expected: FAIL because `MathNode` has no `equationArray` node and the writer cannot serialize it.

- [x] **Step 3: Add the minimal writer implementation**

In `src/schema.ts`, extend `MathNode` with:

```ts
  | { type: "equationArray"; rows: MathNode[][] };
```

In `src/docx-writer.ts`, add a `mathNodeXml` branch before the final `nary` branch:

```ts
  if (node.type === "equationArray") {
    return `<m:eqArr>${node.rows.map((row) => `<m:e>${row.map((child) => mathNodeXml(child)).join("")}</m:e>`).join("")}</m:eqArr>`;
  }
```

- [x] **Step 4: Run the writer test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "writes equation array office math runs"`

Expected: PASS.

### Task 2: Reader Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add this test after `round-trips limit office math runs` in `tests/docx-core.test.ts`:

```ts
  it("round-trips equation array office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "equationArray" as const,
                  rows: [
                    [{ type: "text" as const, text: "x=1" }],
                    [
                      {
                        type: "superscript" as const,
                        base: [{ type: "text" as const, text: "y" }],
                        superscript: [{ type: "text" as const, text: "2" }],
                      },
                    ],
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

Run: `npm test -- tests/docx-core.test.ts -t "round-trips equation array office math runs"`

Expected: FAIL because parser omits `m:eqArr` nodes.

- [x] **Step 3: Add the minimal reader implementation**

In `src/docx-reader.ts`, extend `parseMathNodes` with a `container.eqArr` branch:

```ts
    ...asArray(container.eqArr)
      .map((equationArray) => {
        const equationArrayNode = asObject(equationArray);
        return {
          type: "equationArray" as const,
          rows: asArray(equationArrayNode.e).map((row) => parseMathNodes(asObject(row))),
        };
      }),
```

- [x] **Step 4: Run the round-trip test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips equation array office math runs"`

Expected: PASS.

### Task 3: Full Verification and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-57-equation-array-math.md`

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-57-equation-array-math.md
git commit -m "feat: add phase 57 equation array math"
```

- [x] **Step 3: Push branch**

Run: `git push -u origin phase-57-equation-array-math`

- [x] **Step 4: Mark the plan pushed and commit docs status**

Append the pushed branch and commit hash to this plan, then run:

```powershell
git add docs/superpowers/plans/2026-07-04-phase-57-equation-array-math.md
git commit -m "docs: mark phase 57 pushed"
git push
```

---

**Pushed Branch:** `phase-57-equation-array-math`

**Implementation Commit:** `c77e035 feat: add phase 57 equation array math`
