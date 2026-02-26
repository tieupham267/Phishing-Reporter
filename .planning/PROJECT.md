# Phishing Reporter

## What This Is

A VSTO Outlook add-in that enables enterprise users to report suspected phishing emails to their security team with one click. It extracts email metadata (URLs, attachments, hashes, headers) into a structured report, detects GoPhish simulated campaigns, and forwards everything to the infosec team. Deployed enterprise-wide via MSI/GPO/SCCM across mixed Outlook 2016/2019/365 environments. Now with async HTTP, structured logging, deferred startup, and MSI auto-remediation for previously-disabled machines.

## Core Value

The add-in must load reliably in Outlook and let users report phishing emails without disrupting their workflow.

## Requirements

### Validated

- ✓ User can report a phishing email via ribbon button or right-click context menu — existing
- ✓ Reported email is forwarded as attachment to configured infosec email — existing
- ✓ Report includes extracted metadata: sender, subject, received date — existing
- ✓ Report includes extracted URLs and domains from email body — existing
- ✓ Report includes attachment file hashes (MD5, SHA256) — existing
- ✓ GoPhish simulated campaign detection via custom email headers — existing
- ✓ GoPhish server notification when simulated phishing is reported — existing
- ✓ User information from Exchange/AD included in reports — existing
- ✓ Confirmation dialog before reporting — existing
- ✓ Error auto-report to support email on failure — existing
- ✓ MSI installer for enterprise deployment — existing
- ✓ Add-in loads reliably across all Outlook versions without being auto-disabled — v1.0
- ✓ Outlook UI does not freeze during phishing report submission — v1.0
- ✓ GoPhish HTTP calls execute asynchronously without blocking the UI thread — v1.0
- ✓ Add-in startup time stays under Outlook's resiliency threshold (~1 second) — v1.0
- ✓ Report counters persist across Outlook sessions — v1.0
- ✓ URL detection correctly captures all links in email body — v1.0

### Active

(None — define in next milestone via `/gsd:new-milestone`)

### Out of Scope

- New features (batch reporting, configuration UI, offline mode) — focus was reliability only
- Migration to new Outlook (modern app) — different architecture, separate project
- Migration to Office.js web add-in — major rewrite, not in scope
- .NET Framework upgrade beyond 4.8 — VSTO requires .NET Framework
- Adding a test framework — desirable for future milestone
- OAuth/API key auth for GoPhish — current unauthenticated HTTP sufficient for intranet use

## Context

**Current state (post v1.0):**
- 16 C# source files, 1,846 LOC
- Tech stack: C# / .NET Framework 4.8 / VSTO 4.0 / HtmlAgilityPack 1.12.4 / NLog 5.4.0 / Polly 8.4.2
- Architecture: Ribbon.cs (thin UI callbacks) → ReportOrchestrator → EmailReport (immutable DTO) → GoPhishIntegration (async HttpClient + Polly) / SMTP report
- Logging: NLog file target at %AppData%\PhishingReporter\logs\ with 7-day retention
- MSI: DoNotDisableAddinList + AddinList registry keys, InstallerActions custom action for LoadBehavior reset and DisabledItems cleanup
- Enterprise deployment: hundreds+ users across mixed Outlook 2016/2019/Microsoft 365
- Deployed via MSI through GPO/SCCM

**Known issues / tech debt:**
- GoPhishResult enum value not branched on in ReportOrchestrator (graceful but no differentiated feedback)
- STRT-03 startup timing requires runtime measurement on enterprise hardware (Event ID 45)
- DEPL-04 32/64-bit validation requires physical test machines
- No test framework (unit, integration, or E2E)

## Constraints

- **Tech stack**: Must remain VSTO add-in — users are on desktop Outlook, not new Outlook
- **Compatibility**: Must work with Outlook 2016, 2019, and Microsoft 365 desktop
- **Deployment**: Must continue to deploy via MSI installer
- **Disruption**: Changes must not alter the user-facing workflow (report button, confirmation dialog, report format)
- **Enterprise**: Changes deployed to all users — must be thoroughly validated before rollout

## Key Decisions

| Decision | Rationale | Outcome |
|----------|-----------|---------|
| Focus on reliability over new features | 50% load failure rate is critical for enterprise deployment | ✓ Good — all reliability goals met |
| Keep VSTO architecture | Users on desktop Outlook, no migration path to new Outlook yet | ✓ Good — stable platform for reliability work |
| Async GoPhish calls to fix freezing | Synchronous HTTP on UI thread is root cause of freezes | ✓ Good — UI stays responsive during HTTP calls |
| .NET Framework 4.8 target | Pre-installed on Win10/11, extended support, final .NET Framework version | ✓ Good — no deployment prereqs |
| Isolated NLog LogFactory | Prevents config conflicts between Outlook add-ins | ✓ Good — no interference with other add-ins |
| Immutable EmailReport DTO at async boundary | Prevents COMException from OOM access on background threads | ✓ Good — zero threading issues |
| Polly 8.x for retry with exponential backoff | Modern resilience library, jittered backoff prevents thundering herd | ✓ Good — graceful degradation on GoPhish failures |
| InstallerActions custom action for remediation | Static .vdproj entries can't overwrite Outlook-modified HKCU values | ✓ Good — zero-touch remediation of disabled machines |
| AddinList REG_SZ "1" (not DWORD) | Microsoft managed add-in policy requires string type for "always enabled" | ✓ Good — correct per official docs |
| Deferred init via Application.Startup | Moves heavy work outside Outlook's resiliency measurement window | ✓ Good — minimal startup body |

---
*Last updated: 2026-02-27 after v1.0 milestone*
