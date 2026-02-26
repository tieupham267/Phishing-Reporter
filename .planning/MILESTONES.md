# Milestones

## v1.0 Reliability Release (Shipped: 2026-02-27)

**Phases completed:** 6 phases, 11 plans
**Requirements:** 28/28 satisfied
**Git range:** `feat(01-02)` → `feat(06-01)` (14 feature commits, 73 total)
**Code:** 16 C# files (1,846 LOC), 20 files changed, 4,366 insertions

**Key accomplishments:**
1. Upgraded to .NET Framework 4.8 with NLog 5.4 structured logging to %AppData%
2. All ribbon callbacks exception-safe; UrlExtractor, AttachmentHasher, GoPhish enum extracted from monolithic Ribbon.cs
3. Async GoPhish HTTP via static HttpClient singleton + Polly 8.4 exponential backoff retry
4. Immutable EmailReport DTO extracts all COM data on UI thread before async boundary — zero COMException risk
5. Deferred startup via Application.Startup + CRL bypass for sub-1s load time
6. MSI custom action remediates previously-disabled machines (LoadBehavior reset, DisabledItems/CrashingAddinList cleanup, AddinList "always enabled" key)

**Tech debt accepted:**
- Pre-existing VSTO template TODO comment in Ribbon.cs (Info)
- GoPhishResult enum not branched on in ReportOrchestrator (Low)
- STRT-03/DEPL-04 require runtime measurement on enterprise hardware (Low)

---

