# Requirements: Phishing Reporter

**Defined:** 2026-02-25
**Core Value:** The add-in must load reliably in Outlook and let users report phishing emails without disrupting their workflow.

## v1 Requirements

Requirements for reliability release. Each maps to roadmap phases.

### Startup Reliability

- [ ] **STRT-01**: Add-in initialization deferred to Application.Startup event (outside Outlook's resiliency measurement window)
- [ ] **STRT-02**: CreateRibbonExtensibilityObject overridden in ThisAddIn to eliminate VSTO reflection scan from startup path
- [ ] **STRT-03**: Add-in startup time stays under Outlook's 1,000ms resiliency threshold on typical enterprise hardware
- [x] **STRT-04**: Registry DoNotDisableAddinList key deployed via MSI to prevent auto-disabling
- [x] **STRT-05**: All ribbon event handler entry points wrapped in try/catch to prevent unhandled exception soft-disable

### Network Reliability

- [ ] **NETW-01**: GoPhish HTTP call executes asynchronously via HttpClient without blocking the Outlook UI thread
- [ ] **NETW-02**: HttpClient instantiated as static singleton to prevent socket exhaustion
- [ ] **NETW-03**: GoPhish HTTP call has configurable timeout (default 10 seconds)
- [ ] **NETW-04**: GoPhish HTTP call retries with exponential backoff on transient failures (via Polly)
- [ ] **NETW-05**: Report submission completes gracefully if GoPhish server is unreachable (falls back to standard email report)

### Infrastructure

- [x] **INFR-01**: Project upgraded to .NET Framework 4.8 (pre-installed on Windows 10/11, extended support)
- [x] **INFR-02**: NLog file logging to %AppData%\PhishingReporter\logs\ with 7-day retention
- [x] **INFR-03**: Structured log entries for all report workflow steps (start, GoPhish check, email sent, errors)
- [x] **INFR-04**: HtmlAgilityPack updated to latest stable version (1.12.x)

### Bug Fixes

- [x] **BUGF-01**: URL detection correctly captures all links (remove broken Contains("a") filter)
- [x] **BUGF-02**: Report counters persist across Outlook sessions (call Settings.Save() after increment)
- [x] **BUGF-03**: GoPhish integration returns enum/bool instead of magic strings ("OK", "ERROR", "NaN")
- [ ] **BUGF-04**: HttpWebResponse and StreamReader properly disposed (replaced by HttpClient)
- [x] **BUGF-05**: Temporary attachment files cleaned up in finally block

### Code Quality

- [ ] **QUAL-01**: Email processing logic extracted from Ribbon.cs into dedicated EmailProcessor class
- [ ] **QUAL-02**: GoPhish integration refactored with async HttpClient and proper result types
- [x] **QUAL-03**: URL extraction logic extracted into URLExtractor class
- [x] **QUAL-04**: Hash calculation logic extracted into AttachmentHasher class
- [x] **QUAL-05**: COM objects properly released via Marshal.ReleaseComObject in all processing loops

### Enterprise Deployment

- [ ] **DEPL-01**: MSI Custom Action resets HKCU LoadBehavior on upgrade (fixes previously-disabled users)
- [ ] **DEPL-02**: MSI Custom Action clears HKCU DisabledItems for this add-in on upgrade
- [ ] **DEPL-03**: MSI writes resiliency AddinList registry key for Outlook 16.0
- [ ] **DEPL-04**: Installer validated for both 32-bit and 64-bit Office deployments

## v2 Requirements

Deferred to future release. Tracked but not in current roadmap.

### Observability

- **OBSV-01**: Configuration UI for editing settings at runtime
- **OBSV-02**: Health check indicator in ribbon showing add-in status
- **OBSV-03**: Diagnostic export button for support team

### Features

- **FEAT-01**: Batch reporting (report multiple selected emails)
- **FEAT-02**: Offline queue with sync-when-online
- **FEAT-03**: Customizable report email template

## Out of Scope

| Feature | Reason |
|---------|--------|
| Migration to Office.js web add-in | Major rewrite, different architecture, separate project |
| New Outlook (modern app) support | COM/VSTO add-ins cannot load in new Outlook |
| .NET 5+ migration | VSTO is .NET Framework-only, cannot target modern .NET |
| Test framework setup | Desirable but not in this reliability-focused milestone |
| OAuth/API key auth for GoPhish | Current unauthenticated HTTP sufficient for intranet use |

## Traceability

Which phases cover which requirements. Updated during roadmap creation.

| Requirement | Phase | Status |
|-------------|-------|--------|
| STRT-01 | Phase 5 | Pending |
| STRT-02 | Phase 5 | Pending |
| STRT-03 | Phase 5 | Pending |
| STRT-04 | Phase 1 | Complete |
| STRT-05 | Phase 2 | Complete |
| NETW-01 | Phase 3 | Pending |
| NETW-02 | Phase 3 | Pending |
| NETW-03 | Phase 3 | Pending |
| NETW-04 | Phase 3 | Pending |
| NETW-05 | Phase 3 | Pending |
| INFR-01 | Phase 1 | Complete |
| INFR-02 | Phase 1 | Complete |
| INFR-03 | Phase 1 | Complete |
| INFR-04 | Phase 1 | Complete |
| BUGF-01 | Phase 2 | Complete |
| BUGF-02 | Phase 2 | Complete |
| BUGF-03 | Phase 2 | Complete |
| BUGF-04 | Phase 3 | Pending |
| BUGF-05 | Phase 2 | Complete |
| QUAL-01 | Phase 4 | Pending |
| QUAL-02 | Phase 4 | Pending |
| QUAL-03 | Phase 2 | Complete |
| QUAL-04 | Phase 2 | Complete |
| QUAL-05 | Phase 2 | Complete |
| DEPL-01 | Phase 6 | Pending |
| DEPL-02 | Phase 6 | Pending |
| DEPL-03 | Phase 6 | Pending |
| DEPL-04 | Phase 6 | Pending |

**Coverage:**
- v1 requirements: 28 total
- Mapped to phases: 28
- Unmapped: 0

---
*Requirements defined: 2026-02-25*
*Last updated: 2026-02-26 after 02-01 plan completion*
