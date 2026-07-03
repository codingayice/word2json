# Phase 55 Func Math Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve Office Math function structures so named functions such as `sin(x)` and `log(x)` round-trip through JSON.

**Architecture:** Extend the recursive `MathNode` union with a `function` node containing recursive function-name nodes and recursive argument nodes. Serialize the node to OMML `m:func` with `m:fName` and `m:e`, then parse the same structure back from `m:oMath`.

**Tech Stack:** TypeScript, JSZip, fast-xml-parser, Vitest, OOXML OMML.

---

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add this test after `writes bar office math runs` in `tests/docx-core.test.ts`:

```ts
  it("writes function office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "function" as const,
                  name: [{ type: "text" as const, text: "sin" }],
                  argument: [
                    {
                      type: "delimiter" as const,
                      begin: "(",
                      end: ")",
                      content: [{ type: "text" as const, text: "x" }],
                    },
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

    expect(xml).toContain('<m:func><m:fName><m:r><m:t>sin</m:t></m:r></m:fName><m:e><m:d><m:dPr><m:begChr m:val="("/><m:endChr m:val=")"/></m:dPr><m:e><m:r><m:t>x</m:t></m:r></m:e></m:d></m:e></m:func>');
  });
```

- [x] **Step 2: Run the writer test to verify RED**

Run: `npm test -- tests/docx-core.test.ts -t "writes function office math runs"`

Expected: FAIL because `MathNode` has no `function` node and the writer cannot serialize it.

- [x] **Step 3: Add the minimal writer implementation**

In `src/schema.ts`, extend `MathNode` with:

```ts
  | { type: "function"; name: MathNode[]; argument: MathNode[] };
```

In `src/docx-writer.ts`, add a `mathNodeXml` branch before the final `nary` branch:

```ts
  if (node.type === "function") {
    return `<m:func><m:fName>${node.name.map((child) => mathNodeXml(child)).join("")}</m:fName><m:e>${node.argument.map((child) => mathNodeXml(child)).join("")}</m:e></m:func>`;
  }
```

- [x] **Step 4: Run the writer test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "writes function office math runs"`

Expected: PASS.

### Task 2: Reader Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add this test after `round-trips bar office math runs` in `tests/docx-core.test.ts`:

```ts
  it("round-trips function office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "function" as const,
                  name: [{ type: "text" as const, text: "log" }],
                  argument: [
                    {
                      type: "fraction" as const,
                      numerator: [{ type: "text" as const, text: "a" }],
                      denominator: [{ type: "text" as const, text: "b" }],
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

- [x] **Step 2: Run the round-trip test to verify RED**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips function office math runs"`

Expected: FAIL because parser omits `m:func` nodes.

- [x] **Step 3: Add the minimal reader implementation**

In `src/docx-reader.ts`, extend `parseMathNodes` with a `container.func` branch:

```ts
    ...asArray(container.func)
      .map((func) => {
        const funcNode = asObject(func);
        return {
          type: "function" as const,
          name: parseMathNodes(asObject(funcNode.fName)),
          argument: parseMathNodes(asObject(funcNode.e)),
        };
      }),
```

- [x] **Step 4: Run the round-trip test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips function office math runs"`

Expected: PASS.

### Task 3: Full Verification and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-55-func-math.md`

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-55-func-math.md
git commit -m "feat: add phase 55 func math"
```

- [ ] **Step 3: Push branch**

Run: `git push -u origin phase-55-func-math`

- [ ] **Step 4: Mark the plan pushed and commit docs status**

Append the pushed branch and commit hash to this plan, then run:

```powershell
git add docs/superpowers/plans/2026-07-04-phase-55-func-math.md
git commit -m "docs: mark phase 55 pushed"
git push
```
