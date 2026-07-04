# Phase 107 N-Ary Operator Character Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:naryPr/m:chr` custom character preservation for `nary` math nodes so non-enum large operators survive JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Keep the existing semantic `operator` enum for common operators, and add optional `operatorCharacter?: string` as an exact OMML character override. Writer emits `operatorCharacter` when present; reader maps known characters to the enum while also preserving the raw character when it differs from the enum default.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `operatorCharacter?: string` on `nary`.
- Modify `src/docx-writer.ts`: emit `m:naryPr/m:chr` from `operatorCharacter` when present, otherwise use the existing `operator` mapping.
- Modify `src/docx-reader.ts`: preserve raw `m:chr` as `operatorCharacter` when it is present and not the default character for the parsed enum operator.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes nary custom operator character` that builds an n-ary node with `operator: "sum"` and `operatorCharacter: "⊕"`, then expects `m:chr m:val="⊕"` inside `m:naryPr`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes nary custom operator character"
```

Expected: FAIL because writer emits the enum default `∑` instead of `⊕`.

- [x] **Step 3: Extend schema**

Update the `nary` node:

```ts
  | { type: "nary"; controlProperties?: MathControlProperties; operator: "sum" | "integral" | "product" | "coproduct" | "intersection" | "union"; operatorCharacter?: string; limitLocation?: "underOver" | "subSup"; grow?: boolean; hideLowerLimit?: boolean; hideUpperLimit?: boolean; lowerLimit?: MathNode[]; upperLimit?: MathNode[]; body: MathNode[] }
```

- [x] **Step 4: Implement writer support**

Update the n-ary property list:

```ts
naryOperatorXml(node.operator, node.operatorCharacter),
```

Update the helper:

```ts
function naryOperatorXml(operator: Extract<MathNode, { type: "nary" }>["operator"], character?: string): string {
  const value = character ?? naryOperatorCharacter(operator);
  return `<m:chr m:val="${escapeAttribute(value)}"/>`;
}
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes nary custom operator character"
```

Expected: PASS with `m:chr m:val="⊕"` while preserving limits and body.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips nary custom operator character` that builds, parses, and compares an n-ary node with `operator: "sum"` and `operatorCharacter: "⊕"`.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips nary custom operator character"
```

Expected: FAIL because `parseMathNodes` does not preserve raw `m:chr` for custom characters.

- [x] **Step 3: Implement reader support**

Read the raw character and preserve it when it is not the default for the parsed operator:

```ts
const operatorCharacter = asObject(naryProperties.chr).val;
const operator = naryOperatorValue(operatorCharacter);
```

```ts
operator,
...(typeof operatorCharacter === "string" && operatorCharacter !== naryOperatorCharacter(operator) ? { operatorCharacter } : {}),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips nary custom operator character"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-107-nary-operator-character.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "nary custom operator character"
```

Expected: PASS for both Phase 107 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-107-nary-operator-character.md
git commit -m "feat: add phase 107 nary operator character"
```

- [x] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-107-nary-operator-character
```

- [x] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-107-nary-operator-character.md
git commit -m "docs: mark phase 107 pushed"
git push
```

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes nary custom operator character"` failed because the writer emitted the enum default `∑` instead of custom `⊕`.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes nary custom operator character"` passed after `operatorCharacter` writer support.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips nary custom operator character"` failed because parsed JSON did not preserve `operatorCharacter: "⊕"`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips nary custom operator character"` passed after raw `m:chr` preservation.
- Targeted verification: `npm test -- tests/docx-core.test.ts -t "nary custom operator character"` passed 2 tests with 304 skipped.
- Full suite: `npm test` passed 306 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.

### Push Record

- Branch: `phase-107-nary-operator-character`
- Implementation commit: `b9e11c1 feat: add phase 107 nary operator character`
- Remote: `origin/phase-107-nary-operator-character`

### Self-Review

- Spec coverage: Phase 107 covers writer and reader round-trip for custom n-ary operator characters while preserving current enum behavior for common operators.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: JSON property is consistently named `operatorCharacter` and maps to OMML `m:naryPr/m:chr`.
