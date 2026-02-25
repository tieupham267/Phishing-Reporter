# Phishing Reporter

## What This Is

A VSTO Outlook add-in that enables enterprise users to report suspected phishing emails to their security team with one click. It extracts email metadata (URLs, attachments, hashes, headers) into a structured report, detects GoPhish simulated campaigns, and forwards everything to the infosec team. Deployed enterprise-wide via MSI/GPO/SCCM across mixed Outlook 2016/2019/365 environments.

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

### Active

- [ ] Add-in loads reliably across all Outlook versions without being auto-disabled by Outlook resiliency
- [ ] Outlook UI does not freeze during phishing report submission
- [ ] GoPhish HTTP calls execute asynchronously without blocking the UI thread
- [ ] Add-in startup time stays under Outlook's resiliency threshold (~1 second)
- [ ] Report counters persist across Outlook sessions
- [ ] URL detection correctly captures all links in email body (fix broken filter logic)

### Out of Scope

- New features (batch reporting, configuration UI, offline mode) — focus is reliability only
- Migration to new Outlook (modern app) — different architecture, separate project
- Migration to Office.js web add-in — major rewrite, not in scope for this milestone
- .NET Framework upgrade beyond what's needed for fixes — minimize breaking changes
- Adding a test framework — desirable but not the immediate goal

## Context

- Enterprise deployment: hundreds+ users across mixed Outlook 2016/2019/Microsoft 365
- Deployed via MSI through GPO/SCCM
- ~50% of users experience add-in loading failures where Outlook auto-disables the add-in
- Re-enabling the add-in via COM Add-ins doesn't stick — Outlook disables it again
- Outlook freezes during report submission due to synchronous GoPhish HTTP call on UI thread
- Current stack: C# / .NET Framework 4.6.1 / VSTO 4.0 / HtmlAgilityPack 1.11.23
- Codebase is a monolithic Ribbon.cs (442 lines) mixing UI, email processing, and integration logic
- No logging, no tests, no configuration UI
- Known bugs: URL detection filter, counter persistence, resource disposal in GoPhish integration

## Constraints

- **Tech stack**: Must remain VSTO add-in — users are on desktop Outlook, not new Outlook
- **Compatibility**: Must work with Outlook 2016, 2019, and Microsoft 365 desktop
- **Deployment**: Must continue to deploy via MSI installer
- **Disruption**: Changes must not alter the user-facing workflow (report button, confirmation dialog, report format)
- **Enterprise**: Changes deployed to all users — must be thoroughly validated before rollout

## Key Decisions

| Decision | Rationale | Outcome |
|----------|-----------|---------|
| Focus on reliability over new features | 50% load failure rate is critical for enterprise deployment | — Pending |
| Keep VSTO architecture | Users on desktop Outlook, no migration path to new Outlook yet | — Pending |
| Async GoPhish calls to fix freezing | Synchronous HTTP on UI thread is root cause of freezes | — Pending |

---
*Last updated: 2026-02-25 after initialization*
