# Phase 1 DOCX Core Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build the first high-fidelity foundation: a typed Word JSON schema, JSON-to-DOCX generation for common text documents, and DOCX-to-JSON parsing that can round-trip the supported subset.

**Architecture:** The project starts as a TypeScript library plus CLI. `src/schema.ts` owns the public JSON model, `src/docx-writer.ts` writes OOXML package parts, `src/docx-reader.ts` parses the supported OOXML subset, and tests verify round-trip behavior against the generated `.docx` zip package.

**Tech Stack:** TypeScript, Vitest, JSZip, fast-xml-parser, tsx.

---

## File Structure

- `package.json`: npm scripts and runtime/dev dependencies.
- `tsconfig.json`: strict TypeScript configuration.
- `src/schema.ts`: public Word JSON types and validation helpers.
- `src/docx-writer.ts`: JSON-to-DOCX OOXML package writer.
- `src/docx-reader.ts`: DOCX package reader for the supported subset.
- `src/index.ts`: public exports.
- `src/cli.ts`: command-line entry point for `json2docx` and `docx2json`.
- `tests/docx-core.test.ts`: behavior tests for writer, reader, and round-trip.

## Task 1: Project Scaffold and Public API

**Files:**
- Create: `package.json`
- Create: `tsconfig.json`
- Create: `src/schema.ts`
- Create: `src/index.ts`
- Test: `tests/docx-core.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
import { describe, expect, it } from "vitest";
import { createDocumentJson } from "../src/index";

describe("DocumentJson schema", () => {
  it("creates a versioned document with sections and paragraphs", () => {
    const document = createDocumentJson([
      {
        type: "paragraph",
        runs: [{ text: "Hello Word", bold: true }],
      },
    ]);

    expect(document.version).toBe("1.0");
    expect(document.sections).toHaveLength(1);
    expect(document.sections[0].blocks[0]).toMatchObject({
      type: "paragraph",
      runs: [{ text: "Hello Word", bold: true }],
    });
  });
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- tests/docx-core.test.ts -t "creates a versioned document"`

Expected: FAIL because the project has no package or source files yet.

- [ ] **Step 3: Write minimal implementation**

Create `package.json`, `tsconfig.json`, `src/schema.ts`, and `src/index.ts`. `createDocumentJson(blocks)` returns `{ version: "1.0", sections: [{ blocks }] }`.

- [ ] **Step 4: Run test to verify it passes**

Run: `npm test -- tests/docx-core.test.ts -t "creates a versioned document"`

Expected: PASS.

## Task 2: JSON to DOCX Writer

**Files:**
- Modify: `src/schema.ts`
- Create: `src/docx-writer.ts`
- Modify: `src/index.ts`
- Test: `tests/docx-core.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
import JSZip from "jszip";
import { buildDocx, createDocumentJson } from "../src/index";

it("writes paragraphs and text run formatting into document.xml", async () => {
  const document = createDocumentJson([
    {
      type: "paragraph",
      alignment: "center",
      runs: [
        { text: "Hello ", bold: true },
        { text: "Word", italic: true, underline: true },
      ],
    },
  ]);

  const buffer = await buildDocx(document);
  const zip = await JSZip.loadAsync(buffer);
  const xml = await zip.file("word/document.xml")!.async("string");

  expect(xml).toContain("<w:jc w:val=\"center\"/>");
  expect(xml).toContain("<w:b/>");
  expect(xml).toContain("<w:i/>");
  expect(xml).toContain("<w:u w:val=\"single\"/>");
  expect(xml).toContain("<w:t>Hello </w:t>");
  expect(xml).toContain("<w:t>Word</w:t>");
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- tests/docx-core.test.ts -t "writes paragraphs"`

Expected: FAIL because `buildDocx` is not implemented.

- [ ] **Step 3: Write minimal implementation**

Implement a DOCX zip package with `[Content_Types].xml`, `_rels/.rels`, `word/_rels/document.xml.rels`, `word/styles.xml`, and `word/document.xml`. Support paragraph alignment and run flags: bold, italic, underline.

- [ ] **Step 4: Run test to verify it passes**

Run: `npm test -- tests/docx-core.test.ts -t "writes paragraphs"`

Expected: PASS.

## Task 3: DOCX to JSON Reader and Round-Trip

**Files:**
- Create: `src/docx-reader.ts`
- Modify: `src/index.ts`
- Test: `tests/docx-core.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
import { buildDocx, createDocumentJson, parseDocx } from "../src/index";

it("parses supported paragraphs from generated docx back to json", async () => {
  const source = createDocumentJson([
    {
      type: "paragraph",
      alignment: "right",
      runs: [
        { text: "A", bold: true },
        { text: "B", italic: true, underline: true },
      ],
    },
  ]);

  const docx = await buildDocx(source);
  const parsed = await parseDocx(docx);

  expect(parsed).toEqual(source);
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- tests/docx-core.test.ts -t "parses supported paragraphs"`

Expected: FAIL because `parseDocx` is not implemented.

- [ ] **Step 3: Write minimal implementation**

Read `word/document.xml` from the zip, parse XML with `fast-xml-parser`, collect `w:p` paragraph nodes, map `w:pPr/w:jc` to `alignment`, and map `w:rPr` formatting plus `w:t` text back to JSON runs.

- [ ] **Step 4: Run test to verify it passes**

Run: `npm test -- tests/docx-core.test.ts -t "parses supported paragraphs"`

Expected: PASS.

## Task 4: CLI Smoke Path

**Files:**
- Create: `src/cli.ts`
- Modify: `package.json`
- Test: `tests/docx-core.test.ts`

- [ ] **Step 1: Write the failing test**

```ts
import { mkdtemp, readFile, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import path from "node:path";
import { execa } from "execa";

it("converts json to docx and back through the CLI", async () => {
  const dir = await mkdtemp(path.join(tmpdir(), "word2json-"));
  const jsonPath = path.join(dir, "input.json");
  const docxPath = path.join(dir, "output.docx");
  const roundTripPath = path.join(dir, "roundtrip.json");

  await writeFile(jsonPath, JSON.stringify(createDocumentJson([
    { type: "paragraph", runs: [{ text: "CLI" }] },
  ])));

  await execa("npx", ["tsx", "src/cli.ts", "json2docx", jsonPath, docxPath]);
  await execa("npx", ["tsx", "src/cli.ts", "docx2json", docxPath, roundTripPath]);

  expect(JSON.parse(await readFile(roundTripPath, "utf8"))).toEqual(createDocumentJson([
    { type: "paragraph", runs: [{ text: "CLI" }] },
  ]));
});
```

- [ ] **Step 2: Run test to verify it fails**

Run: `npm test -- tests/docx-core.test.ts -t "CLI"`

Expected: FAIL because `src/cli.ts` does not exist.

- [ ] **Step 3: Write minimal implementation**

Implement commands:

```text
word2json json2docx input.json output.docx
word2json docx2json input.docx output.json
```

- [ ] **Step 4: Run test to verify it passes**

Run: `npm test -- tests/docx-core.test.ts -t "CLI"`

Expected: PASS.

## Task 5: Verify, Commit, and Push Phase 1

**Files:**
- Modify: all created files

- [ ] **Step 1: Run full verification**

Run: `npm test`

Expected: all tests pass.

Run: `npm run build`

Expected: TypeScript build exits with code 0.

- [ ] **Step 2: Commit**

```bash
git add .
git commit -m "feat: add phase 1 docx core"
```

- [ ] **Step 3: Push**

```bash
git push -u origin main
```

Expected: branch `main` is pushed to `https://github.com/codingayice/word2json.git`.

## Self-Review

- Spec coverage: This plan builds a minimal high-fidelity foundation for the supported subset: typed JSON, DOCX generation, DOCX parsing, round-trip tests, and CLI.
- Placeholder scan: No unresolved placeholders remain.
- Type consistency: Public functions are `createDocumentJson`, `buildDocx`, and `parseDocx`; tests and implementation tasks use the same names.
