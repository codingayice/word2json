# Phase 133 Text Run Office Math Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity JSON round-trip support for the WordprocessingML text run Office Math marker `<w:oMath>`.

**Architecture:** The existing `TextRun.math` field represents real inline Office Math content serialized as `<m:oMath>`, so this phase adds a separate boolean run property named `officeMath`. The writer emits `<w:oMath/>` or `<w:oMath w:val="0"/>` inside `<w:rPr>`, and the reader parses the same on/off property through the existing run property parser.

**Tech Stack:** TypeScript, JSZip-based DOCX package tests, existing DOCX reader/writer XML helpers, Vitest.

---

### Task 1: Add RED Tests For Text Run Office Math

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write writer tests before implementation**

Add two tests near the other text run property writer tests:

```ts
  it("writes text run office math on", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Formula", officeMath: true }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:rPr><w:oMath/></w:rPr><w:t>Formula</w:t>");
  });

  it("writes text run office math off", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", officeMath: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:oMath w:val="0"/></w:rPr><w:t>Plain</w:t>');
  });
```

- [x] **Step 2: Write reader round-trip tests before implementation**

Add two tests near the other text run property round-trip tests:

```ts
  it("round-trips text run office math on", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Formula", officeMath: true }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run office math off", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Plain", officeMath: false }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 3: Run writer RED test**

Run: `npm test -- tests/docx-core.test.ts -t "writes text run office math"`

Expected: FAIL because `<w:oMath>` is not emitted yet.

- [x] **Step 4: Run reader RED test**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips text run office math"`

Expected: FAIL because `officeMath` is not preserved through parsing yet.

### Task 2: Implement Office Math Run Property

**Files:**
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Add schema property**

Add the boolean field to `StyleRunProperties`:

```ts
  officeMath?: boolean;
```

- [x] **Step 2: Emit writer XML**

Add the run property in `runPropertiesXml` after `noProof`:

```ts
    run.officeMath !== undefined ? (run.officeMath ? "<w:oMath/>" : '<w:oMath w:val="0"/>') : "",
```

- [x] **Step 3: Parse reader XML**

Add parsing to `parseRunFont` after `noProof`:

```ts
    ...parseOnOffRunProperty(properties.oMath, "officeMath"),
```

- [x] **Step 4: Run targeted GREEN tests**

Run: `npm test -- tests/docx-core.test.ts -t "text run office math"`

Expected: PASS with all four targeted tests passing.

### Task 3: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-133-text-run-office-math.md`

- [x] **Step 1: Run full verification**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run office math"
npm test
npm run build
git diff --check
```

Expected: Targeted tests, full tests, and build pass. `git diff --check` may print existing LF/CRLF warnings but must exit 0.

- [x] **Step 2: Commit implementation**

```bash
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-133-text-run-office-math.md
git commit -m "feat: add phase 133 text run office math"
```

- [x] **Step 3: Push branch**

```bash
git push -u origin phase-133-text-run-math
```

- [x] **Step 4: Record push metadata**

Append the pushed commit hash and remote branch to this plan, then commit the plan update:

```bash
git add docs/superpowers/plans/2026-07-04-phase-133-text-run-office-math.md
git commit -m "docs: mark phase 133 pushed"
git push
```

## Push Record

- Branch: `phase-133-text-run-math`
- Remote: `origin/phase-133-text-run-math`
- Implementation commit: `3f63168 feat: add phase 133 text run office math`
- Pushed to: `https://github.com/codingayice/word2json.git`
