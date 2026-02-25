---
phase: 01-foundation
plan: 02
subsystem: infra
tags: [vdproj, registry, outlook-resiliency, msi-installer]

# Dependency graph
requires:
  - phase: 01-01
    provides: ".NET Framework 4.8 target for all downstream NuGet packages"
provides:
  - "DoNotDisableAddinList HKPU registry key in MSI installer preventing Outlook 16.0 auto-disable"
affects: [06-enterprise-deployment]

# Tech tracking
tech-stack:
  added: []
  patterns: ["vdproj nested registry key hierarchy under HKPU for per-user Outlook resiliency keys"]

key-files:
  created: []
  modified:
    - "Installer/Installer.vdproj"

key-decisions:
  - "Used HKPU (per-user hive) instead of HKLM to match existing add-in registration pattern and avoid requiring admin elevation"
  - "Set DeleteAtUninstall=FALSE and AlwaysCreate=FALSE on intermediate registry keys to avoid removing shared Office paths on uninstall"

patterns-established:
  - "vdproj registry keys use GUID-prefixed identifiers with {60EA8692-D2D5-43EB-80DC-7906BF13D6EF} for keys and {ADCFDA98-8FDD-45E4-90BC-E3D20B029870} for values"

requirements-completed: [STRT-04]

# Metrics
duration: 5min
completed: 2026-02-26
---

# Phase 1 Plan 2: DoNotDisableAddinList Registry Key Summary

**MSI installer now deploys DoNotDisableAddinList DWORD registry key under HKPU\Software\Microsoft\Office\16.0\Outlook\Resiliency to prevent Outlook auto-disabling the add-in**

## Performance

- **Duration:** 5 min
- **Started:** 2026-02-25T17:54:32Z
- **Completed:** 2026-02-26T17:54:32Z
- **Tasks:** 2 (1 auto + 1 human-verify checkpoint)
- **Files modified:** 1

## Accomplishments
- Added 4-level registry key hierarchy (16.0 > Outlook > Resiliency > DoNotDisableAddinList) under existing HKPU\Software\Microsoft\Office path in Installer.vdproj
- Created DWORD value "ZeroD.PhishReporter" = 1 under DoNotDisableAddinList key, matching the ProgID used in existing Addins registration
- Existing registry entries (LoadBehavior, FriendlyName, Description, Manifest) preserved unchanged
- Human verification checkpoint approved confirming correct key structure

## Task Commits

Each task was committed atomically:

1. **Task 1: Add DoNotDisableAddinList registry key hierarchy to Installer.vdproj** - `bd9c95a` (feat)
2. **Task 2: Verify DoNotDisableAddinList registry key in MSI installer** - checkpoint:human-verify (approved, no commit needed)

## Files Created/Modified
- `Installer/Installer.vdproj` - Added DoNotDisableAddinList registry key hierarchy with ZeroD.PhishReporter DWORD value under HKPU Office 16.0 Outlook Resiliency path

## Decisions Made
- Used HKPU (per-user hive) consistently with existing add-in registration pattern, avoiding HKLM which would require admin elevation during install
- Set DeleteAtUninstall=FALSE and AlwaysCreate=FALSE on intermediate registry keys (16.0, Outlook, Resiliency) to avoid removing shared Office registry paths when the MSI is uninstalled

## Deviations from Plan

None - plan executed exactly as written.

## Issues Encountered

None.

## User Setup Required

None - no external service configuration required.

## Next Phase Readiness
- DoNotDisableAddinList registry key is in place; IT can deploy MSI for immediate relief before code changes ship
- Ready for Plan 03 (NLog 5.4.0 structured logging)
- No blockers for next plan

## Self-Check: PASSED

All artifacts verified:
- Installer/Installer.vdproj: FOUND
- 01-02-SUMMARY.md: FOUND
- Commit bd9c95a: FOUND
- DoNotDisableAddinList in vdproj: FOUND (1 match)

---
*Phase: 01-foundation*
*Completed: 2026-02-26*
