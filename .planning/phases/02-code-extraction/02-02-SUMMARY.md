---
phase: 02-code-extraction
plan: 02
subsystem: core
tags: [vsto, com-interop, url-extraction, hash-computation, html-agility-pack, marshal-release]

# Dependency graph
requires:
  - phase: 02-code-extraction
    provides: Exception-safe ribbon callbacks and GoPhishResult typed enum (from 02-01)
provides:
  - UrlExtractor static class for HTML URL and domain extraction (BUGF-01 fix)
  - AttachmentHasher static class with guaranteed temp file cleanup (BUGF-05 fix)
  - COM object cleanup via Marshal.ReleaseComObject in all processing methods (QUAL-05)
  - Ribbon.cs as thin coordinator delegating to extracted classes (QUAL-03, QUAL-04)
affects: [03-resilience, 04-async]

# Tech tracking
tech-stack:
  added: []
  patterns: [single-responsibility extraction, immutable result DTOs, COM cleanup in finally blocks, GUID-based temp file naming]

key-files:
  created:
    - PhishingReporter/UrlExtractor.cs
    - PhishingReporter/AttachmentHasher.cs
  modified:
    - PhishingReporter/Ribbon.cs
    - PhishingReporter/PhishingReporter.csproj

key-decisions:
  - "Use for-loop with index instead of foreach for attachment iteration to enable per-attachment COM release in try/finally"
  - "Release Attachments collection COM object separately from individual Attachment items"
  - "Inner try/catch around each Marshal.ReleaseComObject call to prevent cleanup exceptions from propagating"

patterns-established:
  - "COM cleanup pattern: declare COM locals as null at top, assign in try, release in reverse order in finally with inner try/catch"
  - "Immutable result DTO pattern: sealed class with get-only properties and constructor injection (UrlExtractionResult, AttachmentHashResult)"
  - "Static extractor pattern: pure static class accepting minimal data (string HTML, Attachment COM object), returning immutable result"

requirements-completed: [QUAL-03, QUAL-04, QUAL-05, BUGF-01, BUGF-05]

# Metrics
duration: 5min
completed: 2026-02-26
---

# Phase 02 Plan 02: Code Extraction Summary

**UrlExtractor and AttachmentHasher extracted from Ribbon.cs with BUGF-01 (all URLs captured), BUGF-05 (temp file cleanup), and QUAL-05 (COM object release in all processing methods)**

## Performance

- **Duration:** 5 min
- **Started:** 2026-02-25T18:36:05Z
- **Completed:** 2026-02-25T18:40:58Z
- **Tasks:** 3
- **Files modified:** 4

## Accomplishments
- UrlExtractor class captures ALL anchor href values from email HTML without any character-based filtering, fixing the BUGF-01 Contains("a") bug that silently dropped URLs
- AttachmentHasher class uses GUID-based temp file naming and finally-block cleanup, fixing BUGF-05 temp file leak; SHA256.Create() replaces obsolete SHA256Managed
- Marshal.ReleaseComObject added in finally blocks across all Ribbon.cs processing methods (reportPhishingEmailToSecurityTeam, GetBasicInfo, GetCurrentUserInfos, GetURLsAndAttachmentsInfo) with 11 total release points
- Ribbon.cs reduced from monolith to thin coordinator: no inline HTML parsing, no hash computation, no cryptography imports

## Task Commits

Each task was committed atomically:

1. **Task 1: Create UrlExtractor class with BUGF-01 fix** - `b3f5221` (feat)
2. **Task 2: Create AttachmentHasher class with BUGF-05 fix** - `2ea4e64` (feat)
3. **Task 3: Wire extracted classes into Ribbon.cs, add COM cleanup, update csproj** - `c3302cd` (refactor)

## Files Created/Modified
- `PhishingReporter/UrlExtractor.cs` - URL and domain extraction from HTML email body; UrlExtractionResult immutable DTO
- `PhishingReporter/AttachmentHasher.cs` - MD5/SHA256 hash computation with guaranteed temp file cleanup; AttachmentHashResult immutable DTO
- `PhishingReporter/Ribbon.cs` - Thin coordination layer delegating to extracted classes; COM object cleanup in all processing methods
- `PhishingReporter/PhishingReporter.csproj` - Compile entries for UrlExtractor.cs and AttachmentHasher.cs

## Decisions Made
- Used for-loop with index (1-based, Outlook convention) instead of foreach for attachment iteration to enable per-attachment COM release in individual try/finally blocks -- avoids InvalidComObjectException risk from releasing during enumeration
- Released Attachments collection COM object in an outer finally block, separate from individual Attachment item release in inner finally blocks -- follows reverse-order-of-acquisition rule
- Wrapped each Marshal.ReleaseComObject call in its own inner try/catch to prevent cleanup exceptions from propagating and masking the original exception
- GetCurrentUserInfos: broke up property chaining (Application.Session.CurrentUser.AddressEntry) into named locals (session, currentUserRecipient, addrEntry) to enable proper COM release of each intermediate RCW

## Deviations from Plan

### Auto-fixed Issues

**1. [Rule 2 - Missing Critical] Added COM cleanup for errorEmail in catch block**
- **Found during:** Task 3 (reportPhishingEmailToSecurityTeam refactoring)
- **Issue:** The error-handling catch block creates an errorEmail MailItem COM object that was never released
- **Fix:** Wrapped errorEmail creation and send in try/finally with Marshal.ReleaseComObject
- **Files modified:** PhishingReporter/Ribbon.cs
- **Verification:** Marshal.ReleaseComObject(errorEmail) visible in line 232
- **Committed in:** c3302cd (Task 3 commit)

**2. [Rule 2 - Missing Critical] Added COM cleanup for Attachments collection object**
- **Found during:** Task 3 (GetURLsAndAttachmentsInfo refactoring)
- **Issue:** mailItem.Attachments returns a COM collection object that needs release; plan mentioned individual attachments but the collection itself also needs cleanup
- **Fix:** Store mailItem.Attachments in named local, release in outer finally block
- **Files modified:** PhishingReporter/Ribbon.cs
- **Verification:** Marshal.ReleaseComObject(attachments) visible in line 364
- **Committed in:** c3302cd (Task 3 commit)

---

**Total deviations:** 2 auto-fixed (2 missing critical)
**Impact on plan:** Both auto-fixes necessary for COM resource correctness. No scope creep.

## Issues Encountered
None

## User Setup Required
None - no external service configuration required.

## Next Phase Readiness
- Ribbon.cs is now a thin coordinator ready for Phase 3 async conversion
- UrlExtractor accepts string input (no COM dependency), making it trivially testable and async-safe
- AttachmentHasher accepts Attachment (minimum COM dependency), ready for Phase 4 further decoupling into immutable EmailReport DTO
- COM cleanup pattern established for reuse in any future methods that obtain Outlook OOM objects

## Self-Check: PASSED

All files exist. All commits verified:
- `b3f5221`: feat(02-02) - UrlExtractor class
- `2ea4e64`: feat(02-02) - AttachmentHasher class
- `c3302cd`: refactor(02-02) - Ribbon.cs wiring + COM cleanup + csproj

---
*Phase: 02-code-extraction*
*Completed: 2026-02-26*
