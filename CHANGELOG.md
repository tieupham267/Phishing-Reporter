# Changelog

All notable changes to the Phishing Reporter Outlook add-in.

## [1.0.0] - 2026-02-27

### Added
- NLog 5.4.0 structured logging to `%AppData%\PhishingReporter\logs\` with 7-day rolling retention
- `UrlExtractor` class for HTML link extraction via HtmlAgilityPack
- `AttachmentHasher` class for MD5/SHA256 file hashing with temp file cleanup
- `EmailReport` immutable DTO capturing all email metadata on the UI thread before async boundary
- `ReportOrchestrator` separating GoPhish async branch from standard sync SMTP branch
- Polly 8.4.2 exponential backoff retry (3 attempts, 10s timeout, jittered delays) for GoPhish HTTP calls
- `InstallerActions` class library with `RegistryRemediation` MSI custom action
- MSI custom action resets `LoadBehavior` from 2 to 3 on upgrade (remediates previously-disabled machines)
- MSI custom action clears `DisabledItems` and `CrashingAddinList` registry subkeys on upgrade
- `AddinList` registry key (REG_SZ "1") for Outlook 16.0 resiliency "always enabled" policy
- `DoNotDisableAddinList` registry key deployed via MSI
- Deferred initialization via `Application.Startup` event (outside Outlook's resiliency measurement window)
- `CreateRibbonExtensibilityObject` override eliminating VSTO reflection scan from startup path
- Stopwatch instrumentation for startup timing diagnostics
- `generatePublisherEvidence` CRL bypass in app.config eliminating Authenticode delays

### Changed
- Retargeted from .NET Framework 4.6.1 to 4.8
- Updated HtmlAgilityPack from 1.11.23 to 1.12.4
- GoPhish HTTP calls converted from synchronous `HttpWebRequest` to async `HttpClient` static singleton
- GoPhish integration returns `GoPhishResult` enum instead of magic strings ("OK", "ERROR", "NaN")
- Ribbon `reportPhishing` callback converted to `async void` with full try/catch safety net
- Ribbon.cs reduced from 442-line monolith to thin UI callback layer delegating to `ReportOrchestrator`
- Report counters now persist across Outlook sessions (`Settings.Save()` called after increment)

### Fixed
- URL detection captures all links (removed broken `Contains("a")` filter that silently dropped URLs)
- `HttpWebResponse` and `StreamReader` resource leaks (replaced by `HttpClient` with proper disposal)
- Temporary attachment files cleaned up in `finally` block
- All ribbon event handler entry points wrapped in try/catch preventing unhandled exception soft-disable
- COM objects properly released via `Marshal.ReleaseComObject` in all processing loops
