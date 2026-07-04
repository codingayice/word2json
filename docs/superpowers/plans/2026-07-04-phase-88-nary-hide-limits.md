# Phase 88 N-Ary Hide Limits Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add high-fidelity `m:naryPr/m:subHide` and `m:naryPr/m:supHide` support for `nary` math nodes so hidden lower and upper limits survive JSON-to-DOCX and DOCX-to-JSON.

**Architecture:** Extend the existing `nary` node with optional `hideLowerLimit` and `hideUpperLimit` booleans. Reuse the current `m:naryPr` container so hide flags can coexist with operator `chr`, `limLoc`, and `controlProperties`.

**Tech Stack:** TypeScript, Vitest, JSZip, existing DOCX reader/writer utilities.

---

### File Structure

- Modify `src/schema.ts`: allow `hideLowerLimit?: boolean` and `hideUpperLimit?: boolean` on `nary`.
- Modify `src/docx-writer.ts`: emit `m:naryPr/m:subHide` and `m:supHide` when the corresponding JSON booleans are true.
- Modify `src/docx-reader.ts`: parse `naryPr.subHide` and `naryPr.supHide` into JSON booleans.
- Modify `tests/docx-core.test.ts`: add one writer XML test and one DOCX round-trip test.

### Task 1: Writer Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/schema.ts`
- Modify: `src/docx-writer.ts`

- [x] **Step 1: Write the failing writer test**

Add a writer-side Office Math test named `writes nary hide limits` that builds a sum with `hideLowerLimit: true` and `hideUpperLimit: true`.

- [x] **Step 2: Run writer test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes nary hide limits"
```

Expected: FAIL because writer does not emit `m:subHide` or `m:supHide`.

- [x] **Step 3: Extend schema**

Update the `nary` node:

```ts
  | { type: "nary"; controlProperties?: MathControlProperties; operator: "sum" | "integral" | "product" | "coproduct" | "intersection" | "union"; limitLocation?: "underOver" | "subSup"; hideLowerLimit?: boolean; hideUpperLimit?: boolean; lowerLimit?: MathNode[]; upperLimit?: MathNode[]; body: MathNode[] }
```

- [x] **Step 4: Implement writer support**

Emit these properties in `m:naryPr`:

```ts
node.hideLowerLimit ? '<m:subHide m:val="1"/>' : ""
node.hideUpperLimit ? '<m:supHide m:val="1"/>' : ""
```

- [x] **Step 5: Run writer test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "writes nary hide limits"
```

Expected: PASS with both `m:subHide` and `m:supHide` while preserving operator, limits, and body.

### Task 2: Reader Round-Trip Support

**Files:**
- Modify: `tests/docx-core.test.ts`
- Modify: `src/docx-reader.ts`

- [x] **Step 1: Write the failing round-trip test**

Add a reader-side Office Math test named `round-trips nary hide limits` that builds, parses, and compares an n-ary node with both hide flags.

- [x] **Step 2: Run round-trip test to verify RED**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips nary hide limits"
```

Expected: FAIL because `parseMathNodes` ignores `naryPr.subHide` and `naryPr.supHide`.

- [x] **Step 3: Implement reader support**

Reuse `mathBooleanProperty`:

```ts
          ...mathBooleanProperty(naryProperties.subHide, "hideLowerLimit"),
          ...mathBooleanProperty(naryProperties.supHide, "hideUpperLimit"),
```

- [x] **Step 4: Run round-trip test to verify GREEN**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "round-trips nary hide limits"
```

Expected: PASS and parsed JSON equals the source JSON.

### Task 3: Full Verification, Commit, and Push

**Files:**
- Modify: `docs/superpowers/plans/2026-07-04-phase-88-nary-hide-limits.md`

- [x] **Step 1: Run targeted Office Math tests**

Run:

```bash
npm test -- tests/docx-core.test.ts -t "nary hide limits"
```

Expected: PASS for both Phase 88 tests.

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
git add src/schema.ts src/docx-writer.ts src/docx-reader.ts tests/docx-core.test.ts docs/superpowers/plans/2026-07-04-phase-88-nary-hide-limits.md
git commit -m "feat: add phase 88 nary hide limits"
```

- [ ] **Step 6: Push implementation branch**

Run:

```bash
git push -u origin phase-88-nary-hide-limits
```

- [ ] **Step 7: Mark this plan pushed and commit docs**

Update this plan with the pushed branch name and implementation commit hash, then run:

```bash
git add docs/superpowers/plans/2026-07-04-phase-88-nary-hide-limits.md
git commit -m "docs: mark phase 88 pushed"
git push
```

### Self-Review

- Spec coverage: Phase 88 covers writer and reader round-trip for hidden lower and upper n-ary limits while preserving existing operator, limit location, limits, body, and control properties behavior.
- Placeholder scan: No placeholder text or deferred implementation notes.
- Type consistency: JSON properties are consistently named `hideLowerLimit` and `hideUpperLimit`, mapped to OMML `subHide` and `supHide`.

### Verification Evidence

- RED writer: `npm test -- tests/docx-core.test.ts -t "writes nary hide limits"` failed because `m:naryPr/m:subHide` and `m:supHide` were not emitted.
- GREEN writer: `npm test -- tests/docx-core.test.ts -t "writes nary hide limits"` passed.
- RED reader: `npm test -- tests/docx-core.test.ts -t "round-trips nary hide limits"` failed because parsed JSON did not include `hideLowerLimit` or `hideUpperLimit`.
- GREEN reader: `npm test -- tests/docx-core.test.ts -t "round-trips nary hide limits"` passed.
- Targeted phase check: `npm test -- tests/docx-core.test.ts -t "nary hide limits"` passed 2 tests, 266 skipped.
- Full suite: `npm test` passed 268 tests.
- Build: `npm run build` exited 0.
- Whitespace: `git diff --check` exited 0 with LF/CRLF warnings only.
