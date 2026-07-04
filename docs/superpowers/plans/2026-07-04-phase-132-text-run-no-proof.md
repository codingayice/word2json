# Phase 132 Text Run No Proof Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve normal text run no-proof formatting, including explicit `noProof: false` to disable inherited no-proof formatting.

**Architecture:** Add `noProof?: boolean` to `StyleRunProperties`, which makes it available on `TextRun`. Writer emits `<w:noProof/>` for true, `<w:noProof w:val="0"/>` for false, and omits it when undefined; reader maps `w:noProof` values back through the existing on/off run property helper.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: add `noProof?: boolean` to reusable run properties.
- Modify `src/docx-writer.ts`: make `runPropertiesXml` write `w:noProof` for normal text runs.
- Modify `src/docx-reader.ts`: parse normal run `w:noProof` values with optional boolean semantics.
- Modify `tests/docx-core.test.ts`: add writer XML tests and round-trip tests for `TextRun.noProof: true` and `TextRun.noProof: false`.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write failing writer tests**

Add writer-side tests named `writes text run no proof on` and `writes text run no proof off`.

```ts
  it("writes text run no proof on", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Unchecked", noProof: true }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain("<w:rPr><w:noProof/></w:rPr><w:t>Unchecked</w:t>");
  });

  it("writes text run no proof off", async () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Checked", noProof: false }],
      },
    ]);

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const xml = await zip.file("word/document.xml")!.async("string");

    expect(xml).toContain('<w:rPr><w:noProof w:val="0"/></w:rPr><w:t>Checked</w:t>');
  });
```

- [x] **Step 2: Run writer tests to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run no proof"
```

Expected: FAIL because `TextRun` does not accept `noProof` and the writer does not emit `w:noProof`.

- [x] **Step 3: Implement writer support**

Add to `StyleRunProperties`:

```ts
noProof?: boolean;
```

Add to `runPropertiesXml`:

```ts
run.noProof !== undefined ? (run.noProof ? "<w:noProof/>" : '<w:noProof w:val="0"/>') : "",
```

- [x] **Step 4: Run writer tests to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes text run no proof"
```

Expected: PASS for both writer tests.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write failing round-trip tests**

Add reader-side tests named `round-trips text run no proof on` and `round-trips text run no proof off`.

```ts
  it("round-trips text run no proof on", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Unchecked", noProof: true }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });

  it("round-trips text run no proof off", async () => {
    const source = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Checked", noProof: false }],
      },
    ]);

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 2: Run round-trip tests to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run no proof"
```

Expected: FAIL because the reader currently ignores `w:noProof`.

- [x] **Step 3: Implement reader support**

Use the existing on/off run property helper in normal run property parsing:

```ts
...parseOnOffRunProperty(properties.noProof, "noProof"),
```

- [x] **Step 4: Run round-trip tests to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips text run no proof"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-132-text-run-no-proof.md`

- [x] **Step 1: Run targeted tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "text run no proof"
```

Expected: PASS for all Phase 132 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-132-text-run-no-proof.md
git commit -m "feat: add phase 132 text run no proof"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-132-text-run-no-proof
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-132-text-run-no-proof.md
git commit -m "docs: mark phase 132 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 132 covers schema, writer XML, and reader round-trip preservation for normal text run no-proof on/off.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: New `noProof?: boolean` maps to WordprocessingML `w:noProof`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes text run no proof"` failed because generated XML omitted `<w:noProof/>` and `<w:noProof w:val="0"/>`.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run no proof"` failed because parsed JSON dropped both `noProof: true` and `noProof: false`.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes text run no proof"` passed 2 tests after writer emitted `<w:noProof/>` and `<w:noProof w:val="0"/>`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips text run no proof"` passed 2 tests after reader parsed `w:noProof`.
- Targeted verification: `npm test -- tests/docx-core.test.ts -t "text run no proof"` passed 4 tests with 378 skipped.
- Full suite: `npm test` passed 382 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
