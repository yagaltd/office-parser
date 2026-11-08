# Role registry — dispatch policy (single source of truth)

Dispatch policy for **office-parser** under pi-workflows. Toolset, model, thinking,
verification tier, and the binding for the live plan `20261108-001`.

Role **prompts** are not duplicated here: they are the verbatim subagent system prompts
in the pi-workflows package
(`/home/aurel/.pi/agent/git/github.com/yagaltd/pi-workflows/agents/<role>.md`), referenced
as `@role:<name>` and substituted by the extension at dispatch time. If this table and a
prompt disagree, **this file wins** and the prompt must be fixed.

## Model + thinking policy

Both tiers resolve to the same model — `deepseek/deepseek-flash` (DeepSeek V4.1 Flash,
1,000,000 context) via the DeepSeek API. Thinking level is what separates them.

| Bottleneck | Tier | model | thinking | extra |
|---|---|---|---|---|
| 🔴 BLOCKING | B or C | `deepseek/deepseek-flash` | **max** | human review after |
| 🟡 RISKY | B | `deepseek/deepseek-flash` | **max** | prototype first, then build |
| 🔵 TIME_CONSUMING | B | `deepseek/deepseek-flash` | high | break into smaller steps |
| 🟠 VERIFICATION_HEAVY | B | `deepseek/deepseek-flash` | high | budget extra verification time |
| ⚪ STANDARD | A or B | `deepseek/deepseek-flash` | **high** | default flow |

`@model:strong` → `deepseek/deepseek-flash` + **max**.
`@model:standard` → `deepseek/deepseek-flash` + **high**.

### Verification (human instruction, 2026-11-08)

`thinking: max` was **empirically probed** on `deepseek/deepseek-flash` and accepted
(1/1 succeeded). `high` therefore also holds — it is lower on the same ladder. Do not
change these levels without re-probing; pi validates loudly but silently-wrong levels
cost a full dispatch round.

### Tier A note

Mechanical tasks (≤20 lines, single file, grep-able done) inherit `high`. The
pi-workflows doctrine default would be `low` for Tier A; that default is **not** applied
here because the human specified `high` as the base level. Flip Tier A to `low` only on
explicit instruction.

## Roles

| id | prompt file | toolset | model | thinking |
|---|---|---|---|---|
| `@role:scout` | `agents/scout.md` | read-only | `deepseek/deepseek-flash` | high |
| `@role:worker` | `agents/worker.md` | write | per task (table above) | per task |
| `@role:reviewer` | `agents/reviewer.md` | read-only + bash | `deepseek/deepseek-flash` | high |
| `@role:quality-reviewer` | `agents/quality-reviewer.md` | read-only + bash | `deepseek/deepseek-flash` | max |
| `@role:spec-drafter` | `agents/spec-drafter.md` | write, `.workflows/specs/<file>.spec` only | `deepseek/deepseek-flash` | high |
| `@role:tester` | `agents/tester.md` | write, tests only | `deepseek/deepseek-flash` | high |

Reviewers verify, never fix. A task's ✅ comes only from a reviewer verdict `ok:true`,
after the mechanical gate passes. No model marks its own work done.

## Verification tiers

| Tier | Size | Routing | Review | Parallelism |
|---|---|---|---|---|
| **A — mechanical** | ≤20 lines changed, single file, grep-able done | `standard` + high (see Tier A note) | grep-only, or the task's `verify:` line; orchestrator-inline allowed | any group |
| **B — bounded** | module + tests under contract, one concern | per bottleneck tag | full mechanical pipeline | any group |
| **C — spike / architecture** | verdict artifact required | **max**, serial | full + human review after | **never** in a parallel group |

Tier is decided at plan time and written next to the task, never inferred by the worker.
Prototype tasks (`TASK <N>p (prototype)`) carry Tier C routing and their own spec.

## Live plan `20261108-001` — task routing

| Task | Bottleneck | Tier | model | thinking | Wave |
|---|---|---|---|---|---|
| TASK 1 scout | ⚪ STANDARD | A | flash | high | 0 |
| TASK 2p prototype | 🟡 RISKY | C | flash | **max** | 1 |
| TASK 3p prototype | 🟡 RISKY | C | flash | **max** | 1 |
| TASK 4 page-geometry | 🔴 BLOCKING | B | flash | **max** | 2 (group A) |
| TASK 5 quality-types | 🟡 RISKY | B | flash | **max** | 2 (group A) |
| TASK 6 xy-cut | 🔴 BLOCKING | B | flash | **max** | 3 (group B) |
| TASK 7 complexity-score | 🟡 RISKY | B | flash | **max** | 3 (group B) |
| TASK 8 page-verdicts | 🔴 BLOCKING | B | flash | **max** | 4 |
| TASK 9 no-ink-check | 🟡 RISKY | B | flash | **max** | 5 |
| TASK 10 ocr-page-contract | 🟡 RISKY | B | flash | **max** | 6 |
| TASK 11 rusto-core-tables | 🟡 RISKY | B | flash | **max** | 7 |
| TASK 12 docs | ⚪ STANDARD | A | flash | high | 8 |

**Cost note:** 10 of 12 tasks land on `max` because 🔴/🟡 dominate this plan. If spend
matters more than first-pass accuracy, downgrade the 🟡 build tasks (5, 7, 9, 10, 11) to
`high` and keep `max` for the 🔴 critical path (4, 6, 8) and both prototypes, which are
the two tasks whose verdicts re-shape the plan.

## Parallel groups — disjointness (verified at plan time)

Waves run concurrently only when `eligible` in the `dispatch_next` result has >1 entry,
and only after human confirmation.

| Group | Members | Disjoint boundaries |
|---|---|---|
| A | TASK 4 ‖ TASK 5 | `src/formats/pdf.rs` ‖ `src/quality.rs` + `src/document.rs` + `src/lib.rs` |
| B | TASK 6 ‖ TASK 7 | `src/formats/pdf.rs` ‖ `src/quality.rs` |

Every other task edits `src/formats/pdf.rs` and runs **serially**. Parallel subagents
share one filesystem (no worktrees) — overlapping boundaries must be resequenced, never
run together.

## Budget

- Default: 60 tool calls per task (role doctrine).
- Prototype tasks: 40 tool calls, unbounded cheap greps, **no** production writes.
- Quality-reviewer at SHIP: 80 tool calls.
- Spend is journaled to `.workflows/journal.jsonl` and checked against
  `.workflows/budget.json` before every re-dispatch. Exhausted ⇒ **BUDGET STOP**, not a
  failure: print done vs remaining and wait for the human.

## Blocked states (do not work around these)

`no-pending` / `no-tasks` → plan complete, stop. `deps-unsatisfied` → surface and stop.
`stale-status` → a ✅ without its `ok:true` verdict on file; repair the record, never
re-mark. `missing-spec` → the contract was not generated; stop, do not improvise one.
`no-policy` → this file is missing the entry; add it here before dispatching.
`no-plan` → wrong project; the plan lives in this repo, so re-run from this repo's cwd.

## Standing risks for this plan

- TASK 2p NO-GO re-tiers TASK 4 as `refactor-first` and stops for human review. That is
  designed behaviour, not a failure.
- TASK 6 touching the heading `level` walk is the likeliest surprise. Escalate; never
  patch locally.
- TASK 11 must never add `rusto-rs`, `rusto-nn`, or any crate other than `rusto-core`,
  and must never enter `default` features (`.workflows/CONTEXT.md` §Rulings).
