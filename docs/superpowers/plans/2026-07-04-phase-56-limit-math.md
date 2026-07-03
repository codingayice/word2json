# Phase 56 Limit Math Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve Office Math lower-limit and upper-limit structures so limits and function annotations round-trip through JSON.

**Architecture:** Extend the recursive `MathNode` union with `limitLower` and `limitUpper` nodes, each containing a base expression and a limit expression. Serialize them to OMML `m:limLow` and `m:limUpp`, then parse those same structures back from `m:oMath`.

**Tech Stack:** TypeScript, JSZip, fast-xml-parser, Vitest, OOXML OMML.

---

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add this test after `writes function office math runs` in `tests/docx-core.test.ts`:

```ts
  it("writes limit office math runs", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "limitLower" as const,
                  base: [{ type: "text" as const, text: "lim" }],
                  limit: [{ type: "text" as const, text: "x→0" }],
                },
                {
                  type: "limitUpper" as const,
                  base: [{ type: "text" as const, text: "max" }],
                  limit: [{ type: "text" as const, text: "n" }],
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

    expect(xml).toContain("<m:limLow><m:e><m:r><m:t>lim</m:t></m:r></m:e><m:lim><m:r><m:t>x→0</m:t></m:r></m:lim></m:limLow>");
    expect(xml).toContain("<m:limUpp><m:e><m:r><m:t>max</m:t></m:r></m:e><m:lim><m:r><m:t>n</m:t></m:r></m:lim></m:limUpp>");
  });
```

- [x] **Step 2: Run the writer test to verify RED**

Run: `npm test -- tests/docx-core.test.ts -t "writes limit office math runs"`

Expected: FAIL because `MathNode` has no `limitLower` or `limitUpper` nodes and the writer cannot serialize them.

- [x] **Step 3: Add the minimal writer implementation**

In `src/schema.ts`, extend `MathNode` with:

```ts
  | { type: "limitLower"; base: MathNode[]; limit: MathNode[] }
  | { type: "limitUpper"; base: MathNode[]; limit: MathNode[] };
```

In `src/docx-writer.ts`, add `mathNodeXml` branches before the final `nary` branch:

```ts
  if (node.type === "limitLower") {
    return `<m:limLow><m:e>${node.base.map((child) => mathNodeXml(child)).join("")}</m:e><m:lim>${node.limit.map((child) => mathNodeXml(child)).join("")}</m:lim></m:limLow>`;
  }

  if (node.type === "limitUpper") {
    return `<m:limUpp><m:e>${node.base.map((child) => mathNodeXml(child)).join("")}</m:e><m:lim>${node.limit.map((child) => mathNodeXml(child)).join("")}</m:lim></m:limUpp>`;
  }
```

- [x] **Step 4: Run the writer test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "writes limit office math runs"`

Expected: PASS.

### Task 2: Reader Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add this test after `round-trips function office math runs` in `tests/docx-core.test.ts`:

```ts
  it("round-trips limit office math runs", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [
          {
            text: "",
            math: {
              nodes: [
                {
                  type: "limitLower" as const,
                  base: [{ type: "text" as const, text: "lim" }],
                  limit: [
                    {
                      type: "subscript" as const,
                      base: [{ type: "text" as const, text: "x" }],
                      subscript: [{ type: "text" as const, text: "0" }],
                    },
                  ],
                },
                {
                  type: "limitUpper" as const,
                  base: [{ type: "text" as const, text: "sup" }],
                  limit: [{ type: "text" as const, text: "n" }],
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

Run: `npm test -- tests/docx-core.test.ts -t "round-trips limit office math runs"`

Expected: FAIL because parser omits `m:limLow` and `m:limUpp` nodes.

- [x] **Step 3: Add the minimal reader implementation**

In `src/docx-reader.ts`, extend `parseMathNodes` with `container.limLow` and `container.limUpp` branches:

```ts
    ...asArray(container.limLow)
      .map((limitLower) => {
        const limitNode = asObject(limitLower);
        return {
          type: "limitLower" as const,
          base: parseMathNodes(asObject(limitNode.e)),
          limit: parseMathNodes(asObject(limitNode.lim)),
        };
      }),
    ...asArray(container.limUpp)
      .map((limitUpper) => {
        const limitNode = asObject(limitUpper);
        return {
          type: "limitUpper" as const,
          base: parseMathNodes(asObject(limitNode.e)),
          limit: parseMathNodes(asObject(limitNode.lim)),
        };
      }),
```

- [x] **Step 4: Run the round-trip test to verify GREEN**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips limit office math runs"`

Expected: PASS.

### Task 3: Full Verification and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-56-limit-math.md`

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-56-limit-math.md
git commit -m "feat: add phase 56 limit math"
```

- [x] **Step 3: Push branch**

Run: `git push -u origin phase-56-limit-math`

- [x] **Step 4: Mark the plan pushed and commit docs status**

Append the pushed branch and commit hash to this plan, then run:

```powershell
git add docs/superpowers/plans/2026-07-04-phase-56-limit-math.md
git commit -m "docs: mark phase 56 pushed"
git push
```

---

**Pushed Branch:** `phase-56-limit-math`

**Implementation Commit:** `12848e6 feat: add phase 56 limit math`
