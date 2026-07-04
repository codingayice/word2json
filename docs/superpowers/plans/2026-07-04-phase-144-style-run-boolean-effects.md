# Phase 144 Style Run Boolean Effects Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Preserve style run boolean effect and visibility properties in Word style definitions.

**Architecture:** `StyleRunProperties` already models the boolean run properties that inline text runs can write and parse. This phase extends style `<w:rPr>` writing and parsing for `shadow`, `outline`, `emboss`, `imprint`, `rtl`, `complexScript`, `specVanish`, `hidden`, `webHidden`, `snapToGrid`, `noProof`, and `officeMath`.

**Tech Stack:** TypeScript, JSZip-based DOCX package tests, existing DOCX reader/writer XML helpers, Vitest.

---

### Task 1: Add RED Tests For Style Run Boolean Effects

**Files:**
- Modify: `tests/docx-core.test.ts`

- [x] **Step 1: Write writer test before implementation**

Add a writer test near the existing style run tests:

```ts
  it("writes style run boolean effect properties", async () => {
    const document = {
      version: "1.0" as const,
      styles: {
        character: [{
          id: "HiddenEffectText",
          name: "Hidden Effect Text",
          run: {
            shadow: true,
            outline: true,
            emboss: true,
            imprint: true,
            rtl: true,
            complexScript: true,
            specVanish: true,
            hidden: true,
            webHidden: true,
            snapToGrid: true,
            noProof: true,
            officeMath: true,
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Hidden", styleId: "HiddenEffectText" }] }] }],
    };

    const buffer = await buildDocx(document);
    const zip = await JSZip.loadAsync(buffer);
    const styles = await zip.file("word/styles.xml")!.async("string");

    expect(styles).toContain('<w:rPr><w:shadow/><w:outline/><w:emboss/><w:imprint/><w:rtl/><w:cs/><w:specVanish/><w:vanish/><w:webHidden/><w:snapToGrid/><w:noProof/><w:oMath/></w:rPr>');
  });
```

- [x] **Step 2: Write reader round-trip test before implementation**

Add a reader test near the existing style run round-trip tests:

```ts
  it("round-trips style run boolean effect properties", async () => {
    const source = {
      version: "1.0" as const,
      styles: {
        character: [{
          id: "HiddenEffectText",
          name: "Hidden Effect Text",
          run: {
            shadow: true,
            outline: true,
            emboss: true,
            imprint: true,
            rtl: true,
            complexScript: true,
            specVanish: true,
            hidden: true,
            webHidden: true,
            snapToGrid: true,
            noProof: true,
            officeMath: true,
          },
        }],
      },
      sections: [{ blocks: [{ type: "paragraph" as const, runs: [{ text: "Hidden", styleId: "HiddenEffectText" }] }] }],
    };

    const docx = await buildDocx(source);
    const parsed = await parseDocx(docx);

    expect(parsed).toEqual(source);
  });
```

- [x] **Step 3: Run writer RED test**

Run: `npm test -- tests/docx-core.test.ts -t "writes style run boolean effect"`

Expected: FAIL because style `<w:rPr>` does not emit these boolean properties yet.

- [x] **Step 4: Run reader RED test**

Run: `npm test -- tests/docx-core.test.ts -t "round-trips style run boolean effect"`

Expected: FAIL because style parsing drops these boolean properties.

### Task 2: Implement Style Run Boolean Effects

**Files:**
- Modify: `src/docx-writer.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Emit style run boolean XML**

Extend `styleRunPropertiesXml` with:

```ts
    run.shadow ? "<w:shadow/>" : "",
    run.outline ? "<w:outline/>" : "",
    run.emboss ? "<w:emboss/>" : "",
    run.imprint ? "<w:imprint/>" : "",
    run.rtl ? "<w:rtl/>" : "",
    run.complexScript ? "<w:cs/>" : "",
    run.specVanish ? "<w:specVanish/>" : "",
    run.hidden ? "<w:vanish/>" : "",
    run.webHidden ? "<w:webHidden/>" : "",
    run.snapToGrid ? "<w:snapToGrid/>" : "",
    run.noProof ? "<w:noProof/>" : "",
    run.officeMath ? "<w:oMath/>" : "",
```

- [x] **Step 2: Parse style run boolean XML**

Extend `parseStyleRunProperties` with:

```ts
    ...(properties.shadow !== undefined ? { shadow: true } : {}),
    ...(properties.outline !== undefined ? { outline: true } : {}),
    ...(properties.emboss !== undefined ? { emboss: true } : {}),
    ...(properties.imprint !== undefined ? { imprint: true } : {}),
    ...(properties.rtl !== undefined ? { rtl: true } : {}),
    ...(properties.cs !== undefined ? { complexScript: true } : {}),
    ...(properties.specVanish !== undefined ? { specVanish: true } : {}),
    ...(properties.vanish !== undefined ? { hidden: true } : {}),
    ...(properties.webHidden !== undefined ? { webHidden: true } : {}),
    ...(properties.snapToGrid !== undefined ? { snapToGrid: true } : {}),
    ...(properties.noProof !== undefined ? { noProof: true } : {}),
    ...(properties.oMath !== undefined ? { officeMath: true } : {}),
```

- [x] **Step 3: Run targeted GREEN tests**

Run: `npm test -- tests/docx-core.test.ts -t "style run boolean effect"`

Expected: PASS with both targeted tests passing.

### Task 3: Verify, Commit, Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-144-style-run-boolean-effects.md`

- [x] **Step 1: Run full verification**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "style run boolean effect"
npm test
npm run build
git diff --check
```

Expected: Targeted tests, full tests, and build pass. `git diff --check` may print existing LF/CRLF warnings but must exit 0.

- [ ] **Step 2: Commit implementation**

```bash
git add src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-144-style-run-boolean-effects.md
git commit -m "feat: add phase 144 style run boolean effects"
```

- [ ] **Step 3: Push branch**

```bash
git push -u origin phase-144-style-run-boolean-effects
```

- [ ] **Step 4: Record push metadata**

Append the pushed commit hash and remote branch to this plan, then commit the plan update:

```bash
git add docs/superpowers/plans/2026-07-04-phase-144-style-run-boolean-effects.md
git commit -m "docs: mark phase 144 pushed"
git push
```

## Push Record

- Pending.
