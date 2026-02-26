# Project Retrospective

*A living document updated after each milestone. Lessons feed forward into future planning.*

## Milestone: v1.0 — Reliability Release

**Shipped:** 2026-02-27
**Phases:** 6 | **Plans:** 11 | **Sessions:** ~4

### What Was Built
- Async GoPhish HTTP via static HttpClient singleton + Polly exponential backoff retry
- Immutable EmailReport DTO at async boundary eliminating all COMException risk
- ReportOrchestrator separating GoPhish async branch from standard sync branch
- UrlExtractor, AttachmentHasher, GoPhish enum — extracted from 442-line monolithic Ribbon.cs
- NLog structured logging to %AppData% with 7-day rolling retention
- Deferred startup via Application.Startup + CRL bypass + CreateRibbonExtensibilityObject override
- MSI InstallerActions custom action: LoadBehavior reset, DisabledItems/CrashingAddinList cleanup, AddinList "always enabled" key

### What Worked
- Phase ordering was correct: foundation → extraction → async → orchestration → startup → deployment built naturally on each other
- Research phase before planning caught critical pitfalls (DisabledItems-before-LoadBehavior ordering, REG_SZ vs DWORD for AddinList, HKCU architecture neutrality)
- Immutable DTO pattern at the async boundary completely eliminated COM threading concerns — zero need for STA marshaling
- Plan verification loop caught issues before execution (requirement coverage gaps, missing key links)
- Checkpoint:human-action for VS IDE custom action wiring was the right call — GUIDs can't be predicted

### What Was Inefficient
- Phase 4 (Orchestration) was planned and executed in a previous session but roadmap wasn't updated — required manual fix during Phase 6 planning
- Some early phase plans (Phase 1, 2) had plan checkboxes not updated in ROADMAP.md — accumulated tracking debt
- SUMMARY.md files lacked `one_liner` frontmatter field — had to grep accomplishments sections during milestone audit

### Patterns Established
- **Immutable DTO at async boundary**: Extract all COM/OOM data into sealed class before any `await` — prevents COMException on background threads
- **Silent best-effort custom actions**: MSI custom actions must never fail the install; wrap all logic in try/catch
- **Isolated LogFactory**: Use NLog LogFactory (not global LogManager) in VSTO add-ins to avoid config conflicts
- **HKPU per-user registry**: Match existing add-in registration pattern (HKCU, not HKLM)
- **Defensive Resiliency cleanup**: Clear both DisabledItems AND CrashingAddinList — add-in may have been in either list

### Key Lessons
1. Phase ordering matters enormously — doing async conversion before code extraction would have been painful (OOM calls scattered across monolith)
2. DisabledItems must be cleared BEFORE LoadBehavior reset — Outlook checks DisabledItems on startup and re-disables if the add-in is still listed
3. `async void` is correct for COM ribbon callbacks because you can't return `Task` — but always wrap in try/catch
4. `.vdproj` custom action wiring generates GUIDs that can't be predicted — must be done in VS IDE, not programmatically
5. HKCU\Software is shared across 32-bit and 64-bit processes — no WOW6432Node concern for per-user registry keys

### Cost Observations
- Model mix: ~70% opus (planning, execution), ~30% sonnet (verification, checking)
- Sessions: ~4 sessions across 2 days
- Notable: Single-plan phases (Phase 5, 6) executed very quickly; multi-plan phases (Phase 1 with 3 plans) took longer but parallelized well within waves

---

## Cross-Milestone Trends

### Process Evolution

| Milestone | Sessions | Phases | Key Change |
|-----------|----------|--------|------------|
| v1.0 | ~4 | 6 | First milestone — established research→plan→verify→execute pipeline |

### Cumulative Quality

| Milestone | Tests | Coverage | Zero-Dep Additions |
|-----------|-------|----------|-------------------|
| v1.0 | 0 | 0% | 2 (Polly 8.4.2, NLog 5.4.0) |

### Top Lessons (Verified Across Milestones)

1. Research before planning catches critical pitfalls that would cause rework during execution
2. Immutable DTOs at async boundaries eliminate entire categories of threading bugs
