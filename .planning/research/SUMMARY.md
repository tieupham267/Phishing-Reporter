# Project Research Summary

**Project:** Phishing-Reporter VSTO Outlook Add-in — Reliability Fixes
**Domain:** Enterprise VSTO Outlook Add-in (C# / .NET Framework)
**Researched:** 2026-02-25
**Confidence:** HIGH

## Executive Summary

The Phishing-Reporter add-in suffers from two compounding root causes: Outlook's hard-coded 1,000 ms resiliency threshold auto-disabling the add-in on approximately 50% of enterprise machines during cold-start, and a synchronous HTTP call to the GoPhish server on the UI thread that freezes Outlook for 1–5 seconds on every simulated campaign report. These are not configuration issues — they require code changes. The resiliency disable is caused by .NET Framework JIT overhead on cold-start combined with any work in the startup path; the UI freeze is caused by `HttpWebRequest.GetResponse()` blocking the STA (UI) thread during the GoPhish notification. Both problems are solvable with well-documented .NET Framework and VSTO patterns.

The recommended approach is a phased refactoring: first establish safety nets (registry resiliency keys in the MSI and entry-point exception handling), then convert the GoPhish HTTP call to `async`/`await` using a static `HttpClient` singleton, then move initialization out of `ThisAddIn_Startup` into the `Application.Startup` event which fires outside Outlook's resiliency measurement window. Alongside this, decompose the 442-line `Ribbon.cs` monolith into focused components (orchestrator, extractor, detector, HTTP client) to make the async conversion safe and the code reviewable. Add NLog file-based logging so that post-fix failures are diagnosable in production without remote access.

The primary risks are async conversion pitfalls specific to VSTO: calling Outlook Object Model (OOM) from background threads (causes `COMException E_RPC_WRONG_THREAD`), using `.Result`/`.Wait()` on a `Task` from the STA thread (causes permanent deadlock worse than the original sync call), and leaving `async void` event handlers with unhandled exceptions that crash Outlook silently. All three risks have clear prevention patterns that must be enforced during implementation. The enterprise deployment also carries a non-obvious risk: Outlook writes a disabled `LoadBehavior` to `HKCU` which survives MSI upgrades; the installer must explicitly reset this key as a Custom Action or the fix will appear to deploy successfully while 50% of users remain broken.

---

## Key Findings

### Recommended Stack

The project must stay on .NET Framework — VSTO cannot run on modern .NET (5+/Core), and Microsoft has confirmed VSTO will not move beyond .NET Framework 4.8. Upgrading the target framework from 4.6.1 to 4.8 is the safest move: it is an in-place upgrade (no PIA recompilation), is pre-installed on Windows 10/11 enterprise machines (reducing cold-start JIT overhead), and is the final supported VSTO runtime version. Three new packages are required and all have confirmed .NET Framework 4.8 support: NLog 6.1.0 for file-based logging (XML config allows log-level changes post-deployment without a rebuild — critical for enterprise support), Polly 8.6.5 for HTTP timeout and retry policies wrapping the GoPhish call, and an update to HtmlAgilityPack 1.12.4 (safe drop-in, no breaking changes).

**Core technologies:**
- .NET Framework 4.8: Runtime — in-place upgrade from 4.6.1, final VSTO-supported version, reduces JIT cold-start overhead
- VSTO 4.0 (unchanged): Outlook add-in host — no alternative for classic desktop Outlook
- C# with `async`/`await` (via VS 2022): Implementation — required to move GoPhish HTTP call off UI thread
- `System.Net.Http.HttpClient` (static singleton): HTTP — replaces `HttpWebRequest`; has native async API; must be singleton to avoid socket exhaustion
- NLog 6.1.0: Logging — XML config adjustable post-deployment; supports .NET Framework 3.5–4.8
- Polly 8.6.5: Resilience — adds 10-second timeout and exponential backoff retry to GoPhish call; fixes current infinite-hang risk
- HtmlAgilityPack 1.12.4: HTML parsing — update from 1.11.23; no breaking changes to existing XPath queries

**Do NOT use:** `IHttpClientFactory` (unsupported on .NET Framework, requires DI container), `HttpWebRequest` (no async API, active disposal bug), `Microsoft.Extensions.Logging` (targets .NET 8, DLL conflicts on Framework), Serilog (code-first config, not adjustable post-deployment), modern .NET targets (breaks VSTO entirely).

### Expected Features

This is a reliability milestone, not a feature milestone. "Features" are reliability properties the add-in must have to survive Outlook's resiliency system and enterprise deployment.

**Must have (P1 — table stakes to avoid auto-disable and UI freeze):**
- Startup time under 1,000 ms median — Outlook hard-codes this threshold; the only fix is code optimization
- Override `CreateRibbonExtensibilityObject` explicitly — eliminates reflection scan from startup path (low complexity, high impact)
- Async GoPhish HTTP call — eliminates 1–5 second UI freeze; converts `HttpWebRequest.GetResponse()` to `await HttpClient.GetAsync()`
- Try/catch on every event handler entry point — prevents unhandled exceptions from triggering Outlook's soft-disable mechanism
- GPO resiliency registry keys in MSI — immediate relief for 50% failure rate during the transition before code fixes are deployed

**Should have (P2 — prevent the next class of failures):**
- Structured file logging (NLog to `%AppData%\PhishingReporter\logs\`) — diagnoses remaining failures post-P1 without remote access
- COM object cleanup in attachment/URL processing loops — prevents long-term Exchange object limit failures; `foreach` on COM collections leaks enumerators
- HttpClient singleton initialization — prevents socket exhaustion in high-volume simulation environments
- `Properties.Settings.Default.Save()` call after counter increments — one-line fix for counter reset bug

**Defer (P3 — not this milestone):**
- New user-facing features — explicitly out of scope
- Test framework — desirable but not immediate goal
- Office.js migration — separate project, different architecture

**Anti-features (do not build):**
- Blocking retry logic — extends UI freeze or adds complexity with no user benefit for a fire-and-forget notification
- Polling-based health check — Microsoft explicitly identifies polling as an expensive pattern that contributes to resiliency failures
- In-process crash recovery / self-restart — VSTO does not provide isolation for this; crashing add-ins get hard-disabled regardless
- Configuration UI for end users — enterprise add-ins are IT-managed; user-facing config creates support burden

### Architecture Approach

The recommended architecture decomposes the current 442-line `Ribbon.cs` monolith into focused components organized around the single most important constraint: Outlook OOM must be called on the UI thread, but network I/O must never block it. A `ReportOrchestrator` acts as the async coordinator — it reads all email data from Outlook OOM on the UI thread, packages it into an immutable `EmailReport` POCO, then dispatches the GoPhish HTTP notification to a background thread via `await`. The `Ribbon.cs` becomes a thin UI callback layer that delegates to the orchestrator. Startup initialization moves from `ThisAddIn_Startup` (measured by Outlook's resiliency clock) to `Application.Startup` (fires after all add-ins are loaded, outside the measurement window).

**Major components:**
1. `ThisAddIn` — VSTO lifecycle only; hooks `Application.Startup` event; returns in under 5 ms
2. `Ribbon` — UI callbacks only; `async void` handlers that delegate to `ReportOrchestrator` and wrap awaits in try/catch
3. `ReportOrchestrator` — async workflow coordinator; reads Outlook OOM on UI thread first; dispatches network to background
4. `EmailExtractor` — extracts URLs, attachments, hashes from `MailItem`; pure OOM access; returns immutable `EmailReport`
5. `GoPhishDetector` — pure string logic; parses custom header; constructs report URL; no I/O; fully unit-testable
6. `GoPhishHttpClient` — static `HttpClient` singleton; `async Task<bool> NotifyAsync()`; 10-second timeout; never touches OOM
7. `SettingsValidator` — validates required settings in `Application.Startup` handler; logs warnings; does not crash on missing config

**Suggested build order (from architecture research):**
- Phase 1: Foundation — `MailItemExtensions`, `EmailReport`, `SettingsValidator` (no behavior change, enables everything)
- Phase 2: Extraction layer — `EmailExtractor`, `GoPhishDetector` (pure refactoring)
- Phase 3: Async network layer — `GoPhishHttpClient` (highest impact, isolated change)
- Phase 4: Orchestration — `ReportOrchestrator`, thin `Ribbon.cs` (wires everything together)
- Phase 5: Startup fix — `ThisAddIn.cs` deferred initialization (full async stack must exist first)
- Phase 6: Deployment — MSI installer registry key updates

### Critical Pitfalls

1. **Registry override without fixing startup speed** — GPO `AddinList=1` hides the problem but does not fix it. Event ID 59 still appears; crashes still cause hard-disable that the policy cannot suppress. Fix: measure actual load time via Event ID 45, optimize code first, add registry key as defense-in-depth only after startup is genuinely fast.

2. **Calling Outlook OOM from background threads** — `Task.Run()` lambdas that access `mailItem.HTMLBody`, `Attachments`, or `ExchangeUser` throw `COMException E_RPC_WRONG_THREAD` (HRESULT 0x8001010E). Fix: extract ALL needed data from OOM on the UI thread into an immutable `EmailReport` record BEFORE any `await` or `Task.Run()` boundary.

3. **`async void` with unhandled exceptions** — exceptions thrown after an `await` in an `async void` handler bypass the surrounding try/catch and crash Outlook. Fix: `async void` handler wraps the await of an `async Task` method in a try/catch; never put async logic directly in the `async void` body without a wrapping try/catch.

4. **`.Result`/`.Wait()` on UI thread deadlock** — partial async conversions where the caller still expects a sync return call `.Result` on a Task from the STA thread, causing permanent deadlock (worse than the original sync call). Fix: async conversion must be end-to-end ("async all the way up"); there is no safe way to call `.Result` from a VSTO STA event handler.

5. **HKCU `LoadBehavior` surviving MSI upgrades** — Outlook writes `LoadBehavior=2` (disabled) to `HKCU` when it auto-disables the add-in. MSI upgrades write to `HKLM` only. The `HKCU` key wins, so the fix deploys successfully but 50% of users remain broken. Fix: MSI Custom Action must explicitly reset `HKCU\...\LoadBehavior=3` and clear `DisabledItems` entries during upgrade; test on machines with a previously-disabled state, not just clean installs.

---

## Implications for Roadmap

Based on combined research, the suggested phase structure follows the architecture's build order with deployment considerations grouped at the end.

### Phase 1: Foundation and Safety Net

**Rationale:** Deliver immediate relief to the 50% failure rate through the registry key (deployable without any code changes) while establishing the structural foundation that makes subsequent code changes safe. No behavior changes to existing code paths.

**Delivers:** MSI with GPO resiliency registry key (immediate production fix); immutable `EmailReport` data record; `MailItemExtensions` moved to its own file; `SettingsValidator` for early startup failure detection; NLog initialization.

**Addresses features:** GPO resiliency registry keys in MSI (P1), HttpClient singleton initialization (P2), structured file logging setup (P2), startup configuration validation (P2).

**Avoids pitfall:** HKCU `LoadBehavior` surviving MSI upgrades (must test upgrade path from disabled-state machine, not clean install).

**Research flag:** Standard patterns — MSI Custom Actions are well-documented; no research phase needed.

---

### Phase 2: Extraction Layer (Pure Refactoring)

**Rationale:** Extract `EmailExtractor` and `GoPhishDetector` from `Ribbon.cs` with zero behavior change. This creates the testable, OOM-isolated components that Phase 3's async conversion depends on. Doing this before the async work means the async conversion is a narrowly scoped change to `GoPhishHttpClient` only.

**Delivers:** `EmailExtractor.cs` (URL/attachment/hash extraction from MailItem), `GoPhishDetector.cs` (pure string header parsing; fully unit-testable with no Outlook dependency), `EmailReport.cs` (immutable POCO carrying all extracted data across thread boundaries).

**Addresses features:** COM object cleanup in attachment/URL processing loops (fix `foreach` leaks during extraction refactor).

**Avoids pitfall:** Calling Outlook OOM from background threads — by isolating all OOM access in `EmailExtractor` (UI-thread-only component), the async layer has no temptation to reach back into Outlook objects.

**Research flag:** Standard patterns — pure C# refactoring; no research phase needed.

---

### Phase 3: Async Network Layer

**Rationale:** This is the highest-impact single change — it eliminates the 1–5 second UI freeze on every GoPhish campaign report. Doing it in Phase 3 (not Phase 4) means the `GoPhishHttpClient` component can be reviewed and tested in isolation before it is wired into the full orchestration layer.

**Delivers:** `GoPhishHttpClient.cs` with static `HttpClient` singleton, `async Task<bool> NotifyAsync()`, 10-second timeout via Polly, graceful failure handling (network errors return `false`, not exceptions). Replaces `HttpWebRequest.GetResponse()` entirely.

**Uses stack:** `System.Net.Http.HttpClient`, Polly 8.6.5, .NET Framework 4.8 async/await.

**Avoids pitfalls:** `.Result`/`.Wait()` deadlock (the method is purely async, no sync callers exist yet); async void unhandled exceptions (this is a Task-returning method, not a void event handler); socket exhaustion (singleton pattern).

**Research flag:** Standard patterns — HttpClient singleton and Polly v8 API are well-documented; no research phase needed. Note: Polly 8.x uses `ResiliencePipeline` builder API, not the old `Policy.Handle()` chain — do not mix v7 and v8 patterns.

---

### Phase 4: Orchestration and Ribbon Thinning

**Rationale:** Wire Phase 2 and Phase 3 components together through `ReportOrchestrator`. Reduce `Ribbon.cs` from 442 lines to a thin UI callback layer. This is where the `async void` / `async Task` split is implemented for all ribbon handlers.

**Delivers:** `ReportOrchestrator.cs` (async coordinator that reads OOM on UI thread, then awaits GoPhish notification on background thread), thin `Ribbon.cs` (UI callbacks with try/catch-wrapped awaits), `Properties.Settings.Default.Save()` fix for counter persistence.

**Implements architecture:** Complete async data flow — UI thread reads OOM into `EmailReport`, background thread handles HTTP, no OOM objects cross thread boundaries.

**Avoids pitfalls:** `async void` unhandled exceptions (each ribbon handler wraps `await _orchestrator.ReportAsync()` in try/catch); Outlook OOM from background threads (OOM extraction is complete before the first `await`).

**Research flag:** The `async void` + `async Task` split for VSTO event handlers has documented patterns — standard; no research phase needed.

---

### Phase 5: Startup Reliability Fix

**Rationale:** Move all non-trivial initialization out of `ThisAddIn_Startup` (measured by Outlook's resiliency clock) into the `Application.Startup` event handler (fires after all add-ins load, outside the measurement window). Also override `CreateRibbonExtensibilityObject` to eliminate reflection scan from startup. This phase requires Phase 3 and 4 to be complete because the `GoPhishHttpClient` singleton pre-warm and settings validation must exist to call from `Application.Startup`.

**Delivers:** `ThisAddIn.cs` refactored to near-empty startup handler; `Application.Startup` event handler performing settings validation, NLog initialization, and HttpClient pre-warming; `CreateRibbonExtensibilityObject` override; `ngen.exe` post-install MSI action for native image pre-compilation.

**Addresses features:** Startup time under 1,000 ms (P1), override `CreateRibbonExtensibilityObject` (P1), lazy initialization pattern (P2).

**Avoids pitfall:** Heavy work in `ThisAddIn_Startup` — measure via Event ID 45 after deployment to confirm median stays under 1,000 ms across 5 consecutive Outlook starts on representative hardware.

**Research flag:** Standard patterns — `Application.Startup` deferral is directly documented in Microsoft Q&A for this exact symptom; no research phase needed.

---

### Phase 6: Enterprise Deployment Hardening

**Rationale:** Complete the MSI with all registry Custom Actions, verify 32-bit/64-bit compatibility, and validate the full upgrade path from machines with a previously-disabled add-in. This phase is last because it validates against the fully-working add-in from Phases 1–5.

**Delivers:** MSI Custom Action resetting `HKCU LoadBehavior=3` on upgrade; `DisabledItems` registry cleanup on upgrade; registry keys for both Office 15.0 and 16.0 version paths; bitness verification (32-bit and 64-bit Office); upgrade path test from disabled-state machine.

**Addresses features:** MSI deployment with HKLM registration (P1), GPO resiliency configuration (P1 — started in Phase 1, finalized here).

**Avoids pitfalls:** HKCU `LoadBehavior` surviving MSI upgrades; 32-bit vs. 64-bit MSI mismatch; assembly name/ProgID change (do not rename ProgID — it is deployment-permanent).

**Research flag:** The HKCU override behavior and Custom Action remediation pattern are documented; standard. However, the team should verify the actual Office bitness distribution in the enterprise environment before assuming 64-bit only.

---

### Phase Ordering Rationale

- **Registry key first (Phase 1):** The 50% failure rate is causing immediate support burden. The GPO registry key can be deployed as an MSI update with no code changes, providing immediate relief while code fixes are developed.
- **Extraction before async (Phase 2 before 3):** The async conversion is safe only when OOM access is cleanly isolated in UI-thread-only components. Extracting first means the async layer cannot accidentally touch Outlook objects.
- **Network layer before orchestration (Phase 3 before 4):** `GoPhishHttpClient` in isolation is reviewable and testable. Wiring it into orchestration before it is correct creates a harder-to-debug integrated system.
- **Startup fix after async stack (Phase 5 after 3/4):** The `Application.Startup` handler calls `GoPhishHttpClient.Initialize()` — that singleton must exist. Moving startup work before the HTTP layer is ready creates an incomplete initialization.
- **Deployment hardening last (Phase 6):** The HKCU reset Custom Action should be validated against the full working add-in. Testing upgrade paths during active code changes wastes cycles.

### Research Flags

**Standard patterns — skip research phase:**
- Phase 1: Foundation and safety net — MSI Custom Actions, NLog setup, immutable data records are all standard .NET patterns
- Phase 2: Extraction layer — pure C# refactoring with no novel patterns
- Phase 3: Async network layer — `HttpClient` singleton and Polly v8 are well-documented
- Phase 4: Orchestration — `async void` / `async Task` split for VSTO is documented by Microsoft Q&A
- Phase 5: Startup fix — `Application.Startup` deferral pattern is directly documented for this symptom
- Phase 6: Deployment hardening — HKCU behavior is documented; bitness verification is standard

**No phases require a research phase** — all patterns are verified against Microsoft official documentation with HIGH confidence. The primary implementation risk is correctness of async/OOM thread boundaries, which is a code-review concern rather than a research gap.

---

## Confidence Assessment

| Area | Confidence | Notes |
|------|------------|-------|
| Stack | HIGH | All package versions verified on NuGet; .NET Framework 4.8 as final VSTO runtime confirmed by Microsoft; HttpClient singleton and NLog recommendations from official Microsoft docs |
| Features | HIGH | Resiliency threshold (1,000 ms, 5-iteration median) confirmed from multiple official Microsoft sources; feature prioritization derived from documented Outlook disable reason codes |
| Architecture | HIGH | Component decomposition and async patterns confirmed by Microsoft Q&A, official VSTO threading docs, and Add-in Express community with VSTO expertise; SynchronizationContext rules confirmed |
| Pitfalls | HIGH | All critical pitfalls sourced from official Microsoft documentation (threading support, registry entries, VSTO deploy docs); async deadlock and HKCU override pitfalls are well-established |

**Overall confidence: HIGH**

### Gaps to Address

- **Startup measurement baseline:** Actual current median startup time is unknown. Event ID 45 logging will reveal this post-Phase 1 deployment. If the baseline is already under 1,000 ms and the failure is crash-based (reason code 0x3) rather than performance-based (0x1), the Phase 5 optimizations are still correct but less urgent. The architecture research's `Application.Startup` deferred initialization pattern applies to either root cause.

- **Office bitness distribution in the target enterprise:** Research flags this as environment-specific. Before finalizing Phase 6, confirm whether the environment has mixed 32-bit/64-bit Office installs. Modern M365 is predominantly 64-bit; legacy Outlook 2016 licenses may have 32-bit installs. A wrong assumption here results in silent deployment failures for one Office bitness.

- **Polly 8.x vs. 7.x API breakage:** Polly 8.6.5 uses a different API (`ResiliencePipeline` builder) than Polly 7.x (`Policy.Handle()` chain). If any existing code or documentation references Polly 7 patterns, they must not be mixed. This is a verification item for Phase 3, not a research gap.

- **`ngen.exe` per-machine benefit variance:** STACK.md rates `ngen.exe` as MEDIUM confidence because the startup time benefit varies by hardware and .NET assembly composition. It is recommended as a MSI post-install action but should not be the sole strategy for meeting the 1,000 ms threshold. The `Application.Startup` deferral pattern is the primary fix; `ngen.exe` is supplemental.

---

## Sources

### Primary (HIGH confidence)

- [Microsoft Docs — Support for keeping add-ins enabled](https://learn.microsoft.com/en-us/office/vba/outlook/concepts/getting-started/support-for-keeping-add-ins-enabled) — resiliency registry keys, disable reason codes, `DoNotDisableAddinList`/`AddinList`
- [Microsoft Docs — HttpClient Guidelines for .NET](https://learn.microsoft.com/en-us/dotnet/fundamentals/networking/http/httpclient-guidelines) — singleton pattern, IHttpClientFactory framework support boundaries
- [Microsoft Docs — VSTO Runtime Lifecycle Policy](https://learn.microsoft.com/en-us/visualstudio/vsto/visual-studio-tools-for-office-runtime?view=visualstudio) — .NET Framework 4.8 as final VSTO version
- [Microsoft Docs — Threading support in Office (VSTO)](https://learn.microsoft.com/en-us/visualstudio/vsto/threading-support-in-office?view=vs-2022) — STA model, OOM thread rules, `E_RPC_WRONG_THREAD`
- [Microsoft Docs — Registry entries for VSTO Add-ins](https://learn.microsoft.com/en-us/visualstudio/vsto/registry-entries-for-vsto-add-ins?view=vs-2022) — `LoadBehavior` values, HKLM/HKCU precedence
- [Microsoft Docs — Improve VSTO Add-in Performance](https://learn.microsoft.com/en-us/visualstudio/vsto/improving-the-performance-of-a-vsto-add-in?view=vs-2022) — `CreateRibbonExtensibilityObject`, deferred loading patterns
- [Microsoft Docs — Deploy a VSTO Solution with Windows Installer](https://learn.microsoft.com/en-us/visualstudio/vsto/deploying-a-vsto-solution-by-using-windows-installer?view=vs-2022) — assembly name/ProgID permanence, bitness requirements
- [Microsoft Q&A — VSTO Outlook improve and accelerate Add-in startup](https://learn.microsoft.com/en-us/answers/questions/1056423/vsto-outlook-improve-and-accelerate-add-in-startup) — `Application.Startup` deferral pattern
- [Microsoft Q&A — Default VSTO project causes Outlook to start slowly](https://learn.microsoft.com/en-us/answers/questions/1377543/the-default-project-for-my-visual-studio-outlook-v) — 1,000 ms threshold confirmed even on blank VSTO template (1.62 s)
- [NuGet — NLog 6.1.0](https://www.nuget.org/packages/NLog/) — .NET Framework 3.5–4.8 support confirmed
- [NuGet — Polly 8.6.5](https://www.nuget.org/packages/Polly) — .NET Framework 4.6.2+ support confirmed
- [NuGet — HtmlAgilityPack 1.12.4](https://www.nuget.org/packages/htmlagilitypack/) — .NET Framework compatibility confirmed

### Secondary (MEDIUM confidence)

- [Developer Messaging Blog — Outlook slow add-ins resiliency logic](https://developermessaging.azurewebsites.net/2017/08/02/outlooks-slow-add-ins-resiliency-logic-and-how-to-always-enable-slow-add-ins/) — 5-iteration median measurement explained
- [Add-in Express — Releasing COM objects in Outlook](https://www.add-in-express.com/creating-addins-blog/2008/10/30/releasing-office-objects-net/) — COM object release patterns
- [Add-in Express — Threading in managed Office extensions](https://www.add-in-express.com/creating-addins-blog/2010/11/04/threads-managed-office-extensions/) — OOM thread safety rules
- [NLog GitHub Issue #740](https://github.com/NLog/NLog.Extensions.Logging/issues/740) — DLL conflicts between `NLog.Extensions.Logging` and `Microsoft.Extensions.Logging` on .NET Framework 4.8
- [Microsoft Q&A — VSTO Outlook addin update UI from async process](https://learn.microsoft.com/en-us/answers/questions/78894/vsto-outlook-addin-how-to-update-ui-from-asynchron) — SynchronizationContext async patterns

### Tertiary (LOW confidence)

- [TechHit — How to prevent Outlook from disabling add-ins](https://www.techhit.com/how-to/prevent-outlook-from-disabling-add-in/) — third-party; consistent with official docs but not independently authoritative

---
*Research completed: 2026-02-25*
*Ready for roadmap: yes*
