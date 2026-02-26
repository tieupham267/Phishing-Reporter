---
phase: 06-enterprise-deployment
plan: 01
subsystem: infra
tags: [msi, installer, registry, custom-action, vsto, vdproj]

# Dependency graph
requires:
  - phase: 01-foundation
    provides: DoNotDisableAddinList registry key in MSI installer (.vdproj structure)
  - phase: 05-startup-reliability
    provides: Deferred initialization and CRL bypass (ensures add-in passes resiliency window)
provides:
  - InstallerActions class library with RegistryRemediation custom action
  - MSI custom action that resets LoadBehavior from 2 to 3 on upgrade
  - MSI custom action that clears DisabledItems and CrashingAddinList under Resiliency
  - AddinList static registry key (REG_SZ "1" = always enabled) in MSI
affects: []

# Tech tracking
tech-stack:
  added: [System.Configuration.Install]
  patterns: [MSI custom action via Installer class, registry remediation on upgrade, best-effort error swallowing in custom actions]

key-files:
  created:
    - InstallerActions/InstallerActions.csproj
    - InstallerActions/RegistryRemediation.cs
  modified:
    - PhishingReporter.sln
    - Installer/Installer.vdproj

key-decisions:
  - "Remediation runs only in Install override (not Uninstall) per research anti-pattern guidance"
  - "DisabledItems cleared BEFORE LoadBehavior reset per Pitfall 2 (prevents Outlook re-disabling on next launch)"
  - "CrashingAddinList also cleared as defensive measure (add-in may have crashed pre-Phase 5 fixes)"
  - "All registry operations wrapped in try/catch -- custom action must never fail the install"
  - "AddinList uses REG_SZ type (ValueTypes 3:1) not DWORD -- per Microsoft managed add-in policy specification"

patterns-established:
  - "Best-effort custom action: all registry remediation in try/catch, fallback to static registry keys"
  - "Order-dependent registry cleanup: clear DisabledItems before resetting LoadBehavior"
  - "InstallerActions as standalone .NET 4.8 class library with zero Office/VSTO dependencies"

requirements-completed: [DEPL-01, DEPL-02, DEPL-03, DEPL-04]

# Metrics
duration: 18min
completed: 2026-02-26
---

# Phase 6 Plan 01: Enterprise Deployment Summary

**MSI custom action with RegistryRemediation installer class that resets LoadBehavior, clears DisabledItems/CrashingAddinList, and deploys AddinList "always enabled" registry key**

## Performance

- **Duration:** 18 min (includes checkpoint pause for VS IDE custom action wiring)
- **Started:** 2026-02-26T16:32:07Z
- **Completed:** 2026-02-26T16:50:00Z
- **Tasks:** 3
- **Files modified:** 4

## Accomplishments
- Created InstallerActions class library with RegistryRemediation installer class targeting .NET 4.8 (zero Office/VSTO dependencies)
- Added AddinList static registry entry to MSI with correct REG_SZ type (ValueTypes 3:1, Value "1") for Outlook 16.0 managed add-in policy
- Wired InstallerActions custom action in Visual Studio IDE with InstallerClass=TRUE on Install node
- Registry remediation handles LoadBehavior reset (2->3), DisabledItems deletion, and CrashingAddinList deletion with proper ordering and error handling

## Task Commits

Each task was committed atomically:

1. **Task 1: Create InstallerActions class library project with RegistryRemediation Installer class** - `e89056a` (feat)
2. **Task 2: Add AddinList static registry entry to Installer.vdproj** - `160e7f6` (feat)
3. **Task 3: Wire InstallerActions custom action in Visual Studio and validate deployment** - `7b0a313` (feat)

## Files Created/Modified
- `InstallerActions/InstallerActions.csproj` - .NET Framework 4.8 class library with System.Configuration.Install reference
- `InstallerActions/RegistryRemediation.cs` - [RunInstaller(true)] class with LoadBehavior reset, DisabledItems cleanup, CrashingAddinList cleanup
- `PhishingReporter.sln` - Updated to include InstallerActions project
- `Installer/Installer.vdproj` - Added AddinList registry key and InstallerActions custom action wiring

## Decisions Made
- Remediation runs only in Install override (not Uninstall) per research anti-pattern guidance -- uninstall should not modify registry state
- DisabledItems cleared BEFORE LoadBehavior reset per Pitfall 2 -- if DisabledItems is not cleared first, Outlook will re-disable the add-in and reset LoadBehavior to 2 on next launch
- CrashingAddinList also cleared as defensive measure -- the add-in may have crashed on machines running pre-Phase 5 code
- All registry operations wrapped in try/catch -- custom action must never fail the MSI install; DoNotDisableAddinList and AddinList static keys provide fallback protection
- AddinList uses REG_SZ type (ValueTypes 3:1) not DWORD (3:3) -- per Microsoft managed add-in policy; DoNotDisableAddinList correctly uses DWORD for its different purpose

## Deviations from Plan

None - plan executed exactly as written.

## Issues Encountered
- Task 3 required a checkpoint pause for Visual Studio IDE interaction because .vdproj custom action wiring generates GUIDs referencing project output that cannot be predicted or set programmatically. User completed the VS steps successfully and InstallerClass=TRUE was confirmed in the .vdproj.

## User Setup Required

None - no external service configuration required.

## Next Phase Readiness
- All 6 phases of the reliability release are now complete
- The MSI upgrade path is fully hardened: custom action remediates previously-disabled machines, static registry keys provide fallback protection
- Ready for final integration testing on machines with 32-bit and 64-bit Office deployments
- Ready for production deployment to enterprise environment

## Self-Check: PASSED

All files verified present:
- InstallerActions/InstallerActions.csproj
- InstallerActions/RegistryRemediation.cs
- .planning/phases/06-enterprise-deployment/06-01-SUMMARY.md

All commits verified:
- e89056a (Task 1)
- 160e7f6 (Task 2)
- 7b0a313 (Task 3)

---
*Phase: 06-enterprise-deployment*
*Completed: 2026-02-26*
