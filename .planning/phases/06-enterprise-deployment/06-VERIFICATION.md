---
phase: 06-enterprise-deployment
verified: 2026-02-26T17:30:00Z
status: human_needed
score: 4/5 must-haves verified
re_verification: false
human_verification:
  - test: "Install MSI on a machine where LoadBehavior was previously set to 2"
    expected: "After install, HKCU\\Software\\Microsoft\\Office\\Outlook\\Addins\\ZeroD.PhishReporter\\LoadBehavior = 3. DisabledItems and CrashingAddinList subkeys under HKCU\\Software\\Microsoft\\Office\\16.0\\Outlook\\Resiliency are deleted. Outlook loads the add-in on next launch without user intervention."
    why_human: "Registry remediation is runtime behavior that cannot be verified by inspecting source code alone. The custom action fires during MSI install — cannot be simulated statically."
  - test: "Validate MSI on a machine with 64-bit Office"
    expected: "MSI installs, add-in registers, and add-in loads in Outlook without errors. All HKCU registry keys (LoadBehavior, AddinList, DoNotDisableAddinList) are present and correct."
    why_human: "DEPL-04 requires manual end-to-end test on each Office bitness. Cannot be verified statically."
  - test: "Validate MSI on a machine with 32-bit Office on 64-bit Windows"
    expected: "MSI installs, add-in registers, and add-in loads in Outlook without errors. HKCU path is architecture-neutral — no WOW6432Node issues expected. All registry keys present and correct."
    why_human: "DEPL-04 requires manual end-to-end test on each Office bitness. Cannot be verified statically."
---

# Phase 6: Enterprise Deployment Verification Report

**Phase Goal:** The MSI upgrade path correctly remediates previously-disabled machines by resetting HKCU LoadBehavior and clearing DisabledItems entries, the installer is validated for both 32-bit and 64-bit Office deployments, and all required resiliency registry keys are in place for Outlook 16.0.
**Verified:** 2026-02-26T17:30:00Z
**Status:** human_needed
**Re-verification:** No — initial verification

## Goal Achievement

### Observable Truths

| # | Truth | Status | Evidence |
|---|-------|--------|----------|
| 1 | MSI upgrade resets HKCU LoadBehavior from 2 to 3 on machines where the add-in was previously disabled | VERIFIED | `RegistryRemediation.cs` line 68-77: `OpenSubKey(AddinsKeyPath, writable: true)` + `SetValue("LoadBehavior", 3, RegistryValueKind.DWord)` when `intValue != 3`. Called from `Install` override via `RemediateDisabledState()`. |
| 2 | MSI upgrade deletes DisabledItems and CrashingAddinList subkeys under HKCU Resiliency | VERIFIED | `RegistryRemediation.cs` lines 82-88 + 90-113: `ClearDisabledItems()` calls `DeleteResiliencySubKey("DisabledItems")` and `ClearCrashingAddinList()` calls `DeleteResiliencySubKey("CrashingAddinList")`. Both use `resiliencyKey.DeleteSubKey(subKeyName, throwOnMissingSubKey: false)`. Order is correct (DisabledItems cleared before LoadBehavior reset, per Pitfall 2). |
| 3 | HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\AddinList contains REG_SZ value ZeroD.PhishReporter = 1 | VERIFIED | `Installer.vdproj` lines 1431-1452: AddinList key is present under HKPU > Software > Microsoft > Office > 16.0 > Outlook > Resiliency. Value entry has `"ValueTypes" = "3:1"` (REG_SZ) and `"Value" = "8:1"` (string "1"). ProgID `ZeroD.PhishReporter` is correct. Distinguished correctly from DoNotDisableAddinList which uses `ValueTypes = 3:3` (DWORD). |
| 4 | The InstallerActions project builds as a plain .NET Framework 4.8 class library with no VSTO or Office dependencies | VERIFIED | `InstallerActions.csproj`: `<TargetFrameworkVersion>v4.8</TargetFrameworkVersion>`, `<OutputType>Library</OutputType>`. References only `System` and `System.Configuration.Install`. No Office, VSTO, or NuGet package references present. `RegistryRemediation.cs` uses only `System.Configuration.Install.Installer` and `Microsoft.Win32` (BCL). |
| 5 | The MSI installs and add-in loads on both 32-bit and 64-bit Office deployments | NEEDS HUMAN | Code analysis confirms HKCU registry is architecture-neutral (no WOW6432Node issue), assembly is AnyCPU, and .vdproj uses HKPU (per-user). Manual end-to-end validation on both Office bitnesses is required to satisfy DEPL-04. |

**Score:** 4/5 truths verified (1 requires human testing)

### Required Artifacts

| Artifact | Expected | Status | Details |
|----------|----------|--------|---------|
| `InstallerActions/InstallerActions.csproj` | Class library targeting .NET Framework 4.8 with System.Configuration.Install reference | VERIFIED | File exists, 41 lines, targets v4.8, OutputType=Library, references System and System.Configuration.Install only, no Office/VSTO dependencies |
| `InstallerActions/RegistryRemediation.cs` | Installer class with LoadBehavior reset, DisabledItems cleanup, CrashingAddinList cleanup | VERIFIED | File exists, 115 lines, `[RunInstaller(true)]` attribute confirmed, all three registry operations implemented and substantive, null-checks and try/catch present |
| `PhishingReporter.sln` | Solution file including InstallerActions project | VERIFIED | Line 8: `Project("{FAE04EC0-301F-11D3-BF4B-00C04F79EFBC}") = "InstallerActions", "InstallerActions\InstallerActions.csproj", "{B7A4E4F1-3D2C-4E8A-9F5B-1A2B3C4D5E6F}"` |
| `Installer/Installer.vdproj` | MSI installer with AddinList static registry entry and InstallerActions custom action | VERIFIED | AddinList key confirmed at lines 1431-1452 (REG_SZ, value "1"). Custom action wired at lines 476-489 with `"InstallerClass" = "11:TRUE"` and `"InstallAction" = "3:1"` (Install node). Object reference `_9801A437AA5C4CF5AEBC63F46636B391` traces to InstallerActions.dll at line 2025. |

### Key Link Verification

| From | To | Via | Status | Details |
|------|----|-----|--------|---------|
| `InstallerActions/RegistryRemediation.cs` | `HKCU\Software\Microsoft\Office\Outlook\Addins\ZeroD.PhishReporter\LoadBehavior` | `Registry.CurrentUser.OpenSubKey(AddinsKeyPath, writable: true)` | WIRED | Line 68: `OpenSubKey(AddinsKeyPath, writable: true)`. Line 75: `key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord)`. Pattern confirmed. |
| `InstallerActions/RegistryRemediation.cs` | `HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems` | `Registry.CurrentUser DeleteSubKey` | WIRED | Line 94-95: `OpenSubKey(ResiliencyBasePath, writable: true)`. Line 105: `resiliencyKey.DeleteSubKey(subKeyName, throwOnMissingSubKey: false)`. Called with `"DisabledItems"` via `ClearDisabledItems()`. |
| `Installer/Installer.vdproj` | InstallerActions primary output | Custom action with InstallerClass=TRUE on Install node | WIRED | Line 480: `"Object" = "8:_9801A437AA5C4CF5AEBC63F46636B391"`. Line 482: `"InstallAction" = "3:1"`. Line 487: `"InstallerClass" = "11:TRUE"`. Object key traces to `InstallerActions\obj\Release\InstallerActions.dll` at line 2025. |

### Requirements Coverage

| Requirement | Source Plan | Description | Status | Evidence |
|-------------|------------|-------------|--------|----------|
| DEPL-01 | 06-01-PLAN.md | MSI Custom Action resets HKCU LoadBehavior on upgrade | SATISFIED | `RegistryRemediation.cs` `ResetLoadBehavior()`: opens HKCU addins key writable, checks if value != 3, sets to 3 via `SetValue`. Runs in `Install` override. |
| DEPL-02 | 06-01-PLAN.md | MSI Custom Action clears HKCU DisabledItems for this add-in on upgrade | SATISFIED | `RegistryRemediation.cs` `ClearDisabledItems()` -> `DeleteResiliencySubKey("DisabledItems")`: opens Resiliency key writable, verifies subkey exists, deletes it. Executed before LoadBehavior reset (per Pitfall 2 ordering requirement). |
| DEPL-03 | 06-01-PLAN.md | MSI writes resiliency AddinList registry key for Outlook 16.0 | SATISFIED | `Installer.vdproj` lines 1431-1452: AddinList key under HKPU > Software > Microsoft > Office > 16.0 > Outlook > Resiliency. ValueTypes = 3:1 (REG_SZ), Value = 8:1 (string "1"), Name = ZeroD.PhishReporter. |
| DEPL-04 | 06-01-PLAN.md | Installer validated for both 32-bit and 64-bit Office deployments | NEEDS HUMAN | Code analysis: HKCU is architecture-neutral, assembly is AnyCPU, .vdproj uses HKPU. Static analysis complete. Manual install and add-in load test required on both 32-bit and 64-bit Office machines. |

No orphaned requirements found. All four DEPL-0x IDs declared in the plan are accounted for in REQUIREMENTS.md and all are attributed to Phase 6.

### Anti-Patterns Found

| File | Line | Pattern | Severity | Impact |
|------|------|---------|----------|--------|
| — | — | No anti-patterns detected | — | — |

Scanned `InstallerActions/RegistryRemediation.cs` for: TODO/FIXME/XXX/HACK, placeholder comments, empty returns, console.log-only implementations. None found. All method bodies are substantive with real registry logic, null-checks, and defensive error handling.

### Human Verification Required

#### 1. Upgrade remediation on previously-disabled machine

**Test:** On a test machine, manually set `HKCU\Software\Microsoft\Office\Outlook\Addins\ZeroD.PhishReporter\LoadBehavior = 2`. Optionally add a binary entry to `HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems`. Install the new MSI upgrade over the existing version.

**Expected:** After install, `LoadBehavior` is 3. The `DisabledItems` subkey is deleted. The `CrashingAddinList` subkey is deleted (if it existed). `HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\AddinList\ZeroD.PhishReporter` is present with string value "1". Outlook launches with the add-in active — no "add-in was disabled" notification.

**Why human:** The custom action runs at MSI install time. Verifying the registry state before and after requires executing the MSI on a real machine. Static code analysis confirms the logic is correct but cannot confirm the custom action fires and modifies the correct hive at runtime.

#### 2. 64-bit Office deployment (DEPL-04)

**Test:** Install the MSI on a machine with Microsoft 365 or Office 2016/2019 64-bit. Open Outlook. Open File > Options > Add-ins.

**Expected:** ZeroD.PhishReporter appears in the active COM add-ins list. The PhishingReporter ribbon button is visible. All HKCU registry keys are present and correct values confirmed via `regedit`.

**Why human:** DEPL-04 is explicitly a validation requirement that requires physical install and Outlook launch on a 64-bit Office machine.

#### 3. 32-bit Office deployment (DEPL-04)

**Test:** Install the MSI on a machine with 32-bit Office on 64-bit Windows. Open Outlook. Open File > Options > Add-ins.

**Expected:** ZeroD.PhishReporter appears in the active COM add-ins list. The PhishingReporter ribbon button is visible. Registry keys confirmed under HKCU (not WOW6432Node, since HKCU is architecture-neutral). `LoadBehavior = 3`, `AddinList\ZeroD.PhishReporter = "1"` (REG_SZ).

**Why human:** DEPL-04 is explicitly a validation requirement. 32-bit/64-bit Office coexistence with the x64-targeted .vdproj is correct per architecture analysis, but must be confirmed on real hardware.

### Gaps Summary

No automated gaps found. All four artifacts exist, are substantive (not stubs), and are wired correctly:

- `RegistryRemediation.cs` is a complete, non-stub implementation with real registry logic, proper ordering (DisabledItems before LoadBehavior per Pitfall 2), null-checks on every OpenSubKey call, and try/catch wrapping per the "best-effort custom action" pattern.
- `InstallerActions.csproj` correctly targets .NET Framework 4.8 with zero Office/VSTO dependencies.
- `PhishingReporter.sln` includes the InstallerActions project.
- `Installer.vdproj` has both the AddinList static registry key (correct REG_SZ type, not DWORD — distinguishing it from the sibling DoNotDisableAddinList which uses DWORD) and the custom action wired on the Install node with `InstallerClass = TRUE`.

The only outstanding item is DEPL-04 (32/64-bit Office validation), which requires manual testing by design — it cannot be verified statically and was noted in the plan as a `checkpoint:human-action` task (Task 3).

All three task commits are confirmed in git history:
- `e89056a` — InstallerActions class library and RegistryRemediation.cs
- `160e7f6` — AddinList static registry entry in Installer.vdproj
- `7b0a313` — InstallerActions custom action wired via VS IDE

---

_Verified: 2026-02-26T17:30:00Z_
_Verifier: Claude (gsd-verifier)_
