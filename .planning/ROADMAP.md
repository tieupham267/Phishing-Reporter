# Roadmap: Phishing Reporter — Reliability Release

## Overview

This roadmap transforms a fragile VSTO add-in with a 50% enterprise load failure rate and UI-freezing
synchronous HTTP calls into a reliable, production-hardened tool. The work proceeds in a deliberate
order: infrastructure and logging first (so subsequent changes are diagnosable), pure code extraction
second (so the async conversion has clean OOM isolation), async HTTP conversion third (eliminating
UI freeze), orchestration wiring fourth (making the full async pipeline coherent), startup optimization
fifth (eliminating the resiliency auto-disable root cause), and deployment hardening last (validating
the full upgrade path against working code). Each phase delivers a coherent, verifiable capability
and builds on the previous without leaving partial conversions in the codebase.

## Phases

**Phase Numbering:**
- Integer phases (1, 2, 3): Planned milestone work
- Decimal phases (2.1, 2.2): Urgent insertions (marked with INSERTED)

Decimal phases appear between their surrounding integers in numeric order.

- [ ] **Phase 1: Foundation** - Upgrade framework, establish logging, and deploy immediate registry relief via MSI
- [ ] **Phase 2: Code Extraction** - Decompose Ribbon.cs monolith into pure, single-responsibility components
- [ ] **Phase 3: Async Network Layer** - Replace synchronous GoPhish HTTP call with async HttpClient singleton
- [ ] **Phase 4: Orchestration** - Wire extracted components into async ReportOrchestrator; thin Ribbon.cs to callbacks
- [ ] **Phase 5: Startup Reliability** - Move initialization outside Outlook's resiliency measurement window
- [ ] **Phase 6: Enterprise Deployment** - Complete MSI hardening; validate upgrade path from disabled-state machines

## Phase Details

### Phase 1: Foundation
**Goal**: The project builds against .NET Framework 4.8, logs all workflow steps to disk, and the MSI
already contains the resiliency registry key so IT can deploy immediate relief before code changes ship.
**Depends on**: Nothing (first phase)
**Requirements**: INFR-01, INFR-02, INFR-03, INFR-04, STRT-04
**Success Criteria** (what must be TRUE):
  1. The project compiles and all existing behavior works after retargeting to .NET Framework 4.8
  2. On a developer machine, a log file appears at %AppData%\PhishingReporter\logs\ after clicking the report button, containing timestamped entries for each workflow step
  3. The MSI contains the DoNotDisableAddinList registry key, verifiable by inspecting the installer in Orca or running it on a test machine and checking HKLM\Software\Microsoft\Office\Outlook\Addins
  4. HtmlAgilityPack is updated to 1.12.x with no changes to URL extraction behavior
**Plans:** 2/3 plans executed
Plans:
- [ ] 01-01-PLAN.md — Retarget to .NET Framework 4.8 and update HtmlAgilityPack to 1.12.4
- [ ] 01-02-PLAN.md — Add DoNotDisableAddinList registry key to MSI installer
- [ ] 01-03-PLAN.md — Install NLog 5.4.0 and add structured logging to all workflow steps

### Phase 2: Code Extraction
**Goal**: Ribbon.cs is reduced to a thin coordination layer; URL extraction, attachment hashing, and
GoPhish detection live in dedicated single-responsibility classes; all known bugs in these code paths
are fixed as part of the extraction; all entry points are exception-safe.
**Depends on**: Phase 1
**Requirements**: QUAL-03, QUAL-04, QUAL-05, BUGF-01, BUGF-02, BUGF-03, BUGF-05, STRT-05
**Success Criteria** (what must be TRUE):
  1. A reported email produces a report that includes all URLs in the email body (including those previously missed by the broken Contains("a") filter)
  2. The report counter increments and the new value is visible when Outlook is restarted (counter persists across sessions)
  3. An unhandled exception thrown inside any ribbon event handler does not silently disable the add-in (Outlook's COM Add-ins dialog still shows the add-in as loaded)
  4. After reporting an email with attachments, no temporary files remain in the user's temp directory from that operation
  5. Code inspection shows GoPhish result is an enum or bool type, not the string literals "OK", "ERROR", or "NaN"
**Plans:** 2 plans
Plans:
- [x] 02-01-PLAN.md — Exception-safe ribbon callbacks, Settings persistence fix, GoPhish enum refactor
- [x] 02-02-PLAN.md — UrlExtractor and AttachmentHasher class extraction, COM object cleanup

### Phase 3: Async Network Layer
**Goal**: The GoPhish HTTP notification executes on a background thread via a static HttpClient
singleton with a configured timeout and exponential backoff retry, and the Outlook UI remains
responsive throughout — including during network failures.
**Depends on**: Phase 2
**Requirements**: NETW-01, NETW-02, NETW-03, NETW-04, NETW-05, BUGF-04
**Success Criteria** (what must be TRUE):
  1. When the report button is clicked on an email matching a GoPhish campaign, Outlook's UI (ribbon, email list, inspector) remains responsive while the GoPhish notification is in flight
  2. When the GoPhish server is unreachable or times out after 10 seconds, the standard email report still sends successfully and the user sees a confirmation (not a freeze or error dialog)
  3. After reporting 10 consecutive simulated phishing emails, no "address already in use" or socket exhaustion errors appear in the log file
  4. The log shows GoPhish HTTP call attempt, result (success/failure/timeout), and retry count for each report submission
**Plans:** 2 plans
Plans:
- [x] 03-01-PLAN.md -- Install Polly 8.4.2, add System.Net.Http, rewrite GoPhishIntegration with async HttpClient singleton and Polly resilience pipeline
- [ ] 03-02-PLAN.md -- Wire async method into ribbon callback

### Phase 4: Orchestration
**Goal**: All email data is extracted from Outlook OOM on the UI thread into an immutable EmailReport
record, then the async GoPhish notification and SMTP report dispatch occur without any further OOM
access — the full report workflow is async-safe end-to-end and Ribbon.cs contains only UI callbacks.
**Depends on**: Phase 3
**Requirements**: QUAL-01, QUAL-02
**Success Criteria** (what must be TRUE):
  1. Reporting a phishing email completes successfully with all extracted metadata (URLs, hashes, sender, GoPhish detection result) present in the forwarded report
  2. No COMException with HRESULT 0x8001010E appears in the log after reporting any email (confirming OOM is not accessed from background threads)
  3. Ribbon.cs contains no email parsing, URL extraction, hash calculation, or HTTP logic — code inspection shows it delegates entirely to ReportOrchestrator
**Plans**: TBD

### Phase 5: Startup Reliability
**Goal**: The add-in's measured startup time stays under Outlook's 1,000 ms resiliency threshold on
representative enterprise hardware by deferring all non-trivial initialization to the Application.Startup
event and eliminating the VSTO reflection scan from the startup path.
**Depends on**: Phase 4
**Requirements**: STRT-01, STRT-02, STRT-03
**Success Criteria** (what must be TRUE):
  1. On a representative enterprise machine (not a developer workstation), Outlook Event ID 45 for this add-in shows a median load time under 1,000 ms across 5 consecutive cold starts after a reboot
  2. ThisAddIn_Startup returns in under 5 ms (verifiable via log timestamp between "startup begin" and "startup end" entries)
  3. The add-in remains enabled in Outlook's COM Add-ins dialog across 5 consecutive Outlook restarts on a machine that previously had the add-in auto-disabled
**Plans**: TBD

### Phase 6: Enterprise Deployment
**Goal**: The MSI upgrade path correctly remediates previously-disabled machines by resetting HKCU
LoadBehavior and clearing DisabledItems entries, the installer is validated for both 32-bit and 64-bit
Office deployments, and all required resiliency registry keys are in place for Outlook 16.0.
**Depends on**: Phase 5
**Requirements**: DEPL-01, DEPL-02, DEPL-03, DEPL-04
**Success Criteria** (what must be TRUE):
  1. On a test machine where the add-in was previously auto-disabled (HKCU LoadBehavior=2), running the MSI upgrade causes the add-in to load on the next Outlook start without the user taking any action
  2. After MSI upgrade on a previously-disabled machine, the add-in does not reappear in Outlook's Disabled Items list
  3. The MSI installs and the add-in loads correctly in both 32-bit and 64-bit Office environments (verified on two test machines)
  4. The resiliency AddinList registry key for Outlook 16.0 is present after installation, verifiable in HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\AddinList
**Plans**: TBD

## Progress

**Execution Order:**
Phases execute in numeric order: 1 -> 2 -> 3 -> 4 -> 5 -> 6

| Phase | Plans Complete | Status | Completed |
|-------|----------------|--------|-----------|
| 1. Foundation | 3/3 | Complete | 2026-02-26 |
| 2. Code Extraction | 2/2 | Complete | 2026-02-26 |
| 3. Async Network Layer | 1/2 | In Progress | - |
| 4. Orchestration | 0/TBD | Not started | - |
| 5. Startup Reliability | 0/TBD | Not started | - |
| 6. Enterprise Deployment | 0/TBD | Not started | - |
