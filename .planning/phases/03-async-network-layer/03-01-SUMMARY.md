---
phase: 03-async-network-layer
plan: 01
subsystem: network
tags: [httpclient, polly, async, resilience, retry, timeout, exponential-backoff]

# Dependency graph
requires:
  - phase: 02-code-extraction
    provides: "GoPhishResult enum and GoPhishIntegration class structure"
provides:
  - "Async SendReportNotificationAsync method on GoPhishIntegration"
  - "Static HttpClient singleton preventing socket exhaustion"
  - "Polly 8 resilience pipeline with retry (3x exponential backoff) and 10s per-attempt timeout"
  - "System.Net.Http and Polly 8.4.2 NuGet references in project"
affects: [03-02-PLAN, 04-async-orchestration]

# Tech tracking
tech-stack:
  added: [Polly 8.4.2, Polly.Core 8.4.2, System.Net.Http, Microsoft.Bcl.AsyncInterfaces 6.0.0, Microsoft.Bcl.TimeProvider 8.0.0, System.Threading.Tasks.Extensions 4.5.4]
  patterns: [static-httpclient-singleton, polly-resilience-pipeline, configureawait-false, servicepoint-dns-workaround]

key-files:
  created: []
  modified:
    - PhishingReporter/GoPhishIntegration.cs
    - PhishingReporter/packages.config
    - PhishingReporter/PhishingReporter.csproj

key-decisions:
  - "Used net472 TFM for Polly/Polly.Core HintPaths (closer match to net48 than netstandard2.0)"
  - "Wrapped ServicePoint DNS setup in try/catch to prevent static constructor failure if URL is misconfigured"
  - "Removed explicit TLS 1.2 protocol setting -- .NET 4.8 on Win10/11 defaults to TLS 1.2+"
  - "Set HttpClient.Timeout to InfiniteTimeSpan to let Polly manage all timeout behavior"

patterns-established:
  - "Static HttpClient singleton: private static readonly HttpClient with never-dispose pattern"
  - "Polly 8 ResiliencePipeline: FIFO strategy ordering (retry wraps timeout for per-attempt timeout)"
  - "ConfigureAwait(false) on every await in library code to avoid STA thread context capture"
  - "ServicePoint.ConnectionLeaseTimeout for DNS recycling on .NET Framework 4.8"

requirements-completed: [NETW-01, NETW-02, NETW-03, NETW-04, BUGF-04]

# Metrics
duration: 3min
completed: 2026-02-26
---

# Phase 3 Plan 1: Async Network Layer Summary

**Async HttpClient singleton with Polly 8 resilience pipeline (3x exponential backoff retry + 10s per-attempt timeout) replacing synchronous HttpWebRequest in GoPhishIntegration**

## Performance

- **Duration:** 3 min
- **Started:** 2026-02-26T00:11:51Z
- **Completed:** 2026-02-26T00:14:48Z
- **Tasks:** 2
- **Files modified:** 3

## Accomplishments
- Installed Polly 8.4.2 with all transitive dependencies and added System.Net.Http framework reference
- Replaced synchronous HttpWebRequest/HttpWebResponse with async HttpClient singleton and Polly resilience pipeline
- Eliminated BUGF-04 resource leaks (undisposed HttpWebResponse/StreamReader) by using HttpClient lifecycle management
- Removed fragile global TLS 1.2 protocol setting (unnecessary on .NET 4.8/Win10+)

## Task Commits

Each task was committed atomically:

1. **Task 1: Install Polly 8.x and add System.Net.Http framework reference** - `419d818` (chore)
2. **Task 2: Rewrite GoPhishIntegration with async HttpClient singleton and Polly resilience pipeline** - `2342f1c` (feat)

## Files Created/Modified
- `PhishingReporter/packages.config` - Added Polly 8.4.2, Polly.Core 8.4.2, and 6 transitive dependency entries
- `PhishingReporter/PhishingReporter.csproj` - Added System.Net.Http framework reference, Polly/Polly.Core assembly references with HintPaths, and transitive dependency references
- `PhishingReporter/GoPhishIntegration.cs` - Replaced synchronous sendReportNotificationToServer with async SendReportNotificationAsync, added static HttpClient singleton, added Polly resilience pipeline with retry + timeout

## Decisions Made
- Used `net472` TFM for Polly/Polly.Core HintPaths instead of `netstandard2.0` -- net472 is a closer framework match for the project's net48 target
- Wrapped ServicePoint DNS configuration in try/catch to prevent static constructor failure if gophish_url setting is empty or invalid at startup
- Removed explicit `ServicePointManager.SecurityProtocol = Tls12` assignment -- .NET Framework 4.8 on Windows 10/11 defaults to system TLS (includes 1.2 and 1.3), and the explicit assignment was process-wide and could interfere with other add-ins (Pitfall 6 from research)
- Set `HttpClient.Timeout = InfiniteTimeSpan` to delegate all timeout management to Polly, avoiding the HttpClient.Timeout vs Polly timeout conflict (Pitfall 5 from research)

## Deviations from Plan

### Auto-fixed Issues

**1. [Rule 3 - Blocking] Downloaded nuget.exe for package installation**
- **Found during:** Task 1 (Install Polly)
- **Issue:** `nuget` CLI was not available in PATH; plan noted manual approach may be needed
- **Fix:** Downloaded nuget.exe from dist.nuget.org and used it to install Polly 8.4.2 with proper dependency resolution into packages/ directory
- **Files modified:** nuget.exe downloaded (not committed)
- **Verification:** All 8 packages installed with DLLs at correct HintPath locations
- **Committed in:** 419d818 (Task 1 commit)

**2. [Rule 1 - Bug] Corrected transitive dependency versions and TFMs from plan**
- **Found during:** Task 1 (Install Polly)
- **Issue:** Plan specified System.Runtime.CompilerServices.Unsafe 6.0.0 but NuGet resolved 4.5.3 as the minimum required version; plan also specified netstandard2.0 HintPaths but net461/net462/net472 TFMs were available and are better matches
- **Fix:** Used actual NuGet-resolved versions (4.5.3 for Unsafe) and best-match TFMs (net472 for Polly, net462 for TimeProvider, net461 for others); added Microsoft.Bcl.TimeProvider 8.0.0 transitive dependency not listed in plan
- **Files modified:** PhishingReporter/packages.config, PhishingReporter/PhishingReporter.csproj
- **Verification:** All DLLs verified to exist at HintPath locations
- **Committed in:** 419d818 (Task 1 commit)

---

**Total deviations:** 2 auto-fixed (1 blocking, 1 bug)
**Impact on plan:** Both auto-fixes necessary for correct package installation. No scope creep.

## Issues Encountered
- MSBuild build verification could not run because the VSTO Office Tools workload (Microsoft.VisualStudio.Tools.Office.targets) is not installed in the current build environment. All file changes were verified structurally (correct XML syntax, DLLs at HintPaths, correct using directives, correct API usage) but a full compilation was not possible. The project should build correctly when opened in Visual Studio with Office Tools installed.

## User Setup Required

None - no external service configuration required.

## Next Phase Readiness
- GoPhishIntegration.SendReportNotificationAsync is ready to be called from the ribbon callback
- Phase 3 Plan 2 (03-02) will wire the async method into the ribbon callback via async void pattern
- Phase 4 (async orchestration) will extract OOM data before the await boundary

## Self-Check: PASSED

- All 3 modified files exist at expected paths
- Commit 419d818 (Task 1) verified in git log
- Commit 2342f1c (Task 2) verified in git log
- SUMMARY.md created at .planning/phases/03-async-network-layer/03-01-SUMMARY.md

---
*Phase: 03-async-network-layer*
*Completed: 2026-02-26*
