# Engineering, not vibe coding

The default in this repo is **real engineering**: understand the problem, design
before writing, write tests, then implement in small verifiable steps, and review
the result. Don't jump straight to code for anything non-trivial.

A skills library lives in `.claude/skills/` (from
[mattpocock/skills](https://github.com/mattpocock/skills)). Reach for the right
skill instead of improvising. When unsure which one fits, use **`ask-matt`** — it
routes you to the right skill or flow.

## The default flow

For any feature, bug, or refactor beyond a trivial one-liner, work in this order.
Skip a step only when it genuinely doesn't apply, and say so.

1. **Understand** — `research` to gather facts against primary sources;
   `diagnosing-bugs` when something is broken, throwing, failing, or slow.
2. **Sharpen the plan** — `grill-me` (or `grill-with-docs`) to stress-test the
   design *before* building. Don't skip this on anything ambiguous.
3. **Design the shape** — `codebase-design` / `design-an-interface` for module
   boundaries and interfaces; `domain-modeling` to pin down terminology;
   `prototype` for a throwaway to sanity-check an approach.
4. **Write it down** — `to-prd` to turn the discussion into a PRD;
   `to-issues` to slice it into independently-shippable vertical slices.
5. **Build test-first** — `tdd` (red → green → refactor). Tests come before the
   implementation, not after.
6. **Implement** — `implement` to execute against a PRD or set of issues.
7. **Review** — `code-review` before considering the work done (checks against
   both repo standards and the original spec).

## Rules of engagement

- **No vibe coding.** Don't guess at behavior or write speculative code hoping it
  works. If you don't know how something behaves, `research` or `prototype` it.
- **Plan before code.** For non-trivial work, get the design grilled first.
- **Test-first.** New behavior and bug fixes start with a failing test (`tdd`).
- **Small, verifiable steps.** Vertical slices (`to-issues`), tiny commits, each
  independently reviewable.
- **Review before done.** Run `code-review` on the diff before calling it finished.

## Skills reference

**Engineering**
- `ask-matt` — router; asks which skill or flow fits your situation
- `research` — investigate a question against high-trust primary sources
- `diagnosing-bugs` — diagnosis loop for hard bugs and performance regressions
- `grill-me` / `grill-with-docs` — relentless interview to sharpen a plan or design
- `codebase-design` / `design-an-interface` — design deep modules and interfaces
- `domain-modeling` — build and sharpen the domain model / ubiquitous language
- `prototype` — throwaway prototype to answer a design question
- `to-prd` — turn the conversation into a PRD on the issue tracker
- `to-issues` — break a plan/PRD into independently-grabbable vertical slices
- `tdd` — test-driven development (red-green-refactor)
- `implement` — implement work from a PRD or set of issues
- `code-review` — review changes against repo standards and the original spec
- `improve-codebase-architecture` — scan for deepening opportunities and grill them
- `triage` — move issues/PRs through a triage state machine
- `setup-matt-pocock-skills` — one-time setup of tracker, labels, and doc layout

**Productivity**
- `handoff` — compact the conversation into a handoff doc for another agent
- `teach` — teach a new skill or concept within this workspace
- `writing-great-skills` — reference for writing and editing skills well
- `grilling` — grill the user relentlessly about a plan or design

Additional skills (`misc/`, `personal/`, `in-progress/`, `deprecated/` from the
source repo) are also installed under `.claude/skills/` and available by name.
