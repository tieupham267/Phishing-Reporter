# Phase 6: Enterprise Deployment - Research

**Researched:** 2026-02-26
**Domain:** MSI custom actions, VSTO registry remediation, 32/64-bit Office deployment, Outlook resiliency keys
**Confidence:** HIGH

## Summary

Phase 6 completes the MSI hardening story. The add-in's startup reliability was fixed in Phases 1-5 (deferred init, CRL bypass, exception-safe callbacks, DoNotDisableAddinList key). Now the installer must remediate machines where the add-in was ALREADY disabled before those fixes shipped. This means: (1) resetting HKCU LoadBehavior from 2 back to 3, (2) deleting DisabledItems entries for this add-in, (3) writing the Resiliency\AddinList managed add-in key, and (4) validating the MSI works on both 32-bit and 64-bit Office.

The existing installer is a Visual Studio Deployment Project (.vdproj) that uses `HKPU` (per-user hive) registry entries. It already writes LoadBehavior=3 under the Addins key and DoNotDisableAddinList under Resiliency. However, static .vdproj registry entries cannot conditionally overwrite HKCU values that Outlook has modified post-install. A C# Installer Class custom action is required to programmatically reset LoadBehavior, clear DisabledItems, and write the AddinList key during MSI upgrade.

**Primary recommendation:** Create a small C# class library project with a `[RunInstaller(true)]` Installer class that performs all four registry operations on Install/Commit. Wire it into the existing .vdproj as a custom action on the Install node. The static AddinList registry key can alternatively be added directly to the .vdproj HKPU section (same pattern as DoNotDisableAddinList from Phase 1). For 32/64-bit validation, the current installer already targets x64 (`TargetPlatform=1`, `ProgramFiles64Folder`); HKCU registry is architecture-neutral so no WOW6432Node concern exists for per-user keys.

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|-----------------|
| DEPL-01 | MSI Custom Action resets HKCU LoadBehavior on upgrade (fixes previously-disabled users) | Section: Architecture Patterns -- Pattern 1 (Installer Class), Code Examples -- Example 1 (LoadBehavior reset). LoadBehavior=2 means "unloaded, load-at-startup" but Outlook interprets it as "was disabled"; resetting to 3 re-enables. |
| DEPL-02 | MSI Custom Action clears HKCU DisabledItems for this add-in on upgrade | Section: Architecture Patterns -- Pattern 2 (DisabledItems cleanup), Code Examples -- Example 2. Delete entire DisabledItems subkey under Resiliency since binary format is undocumented; safe because Outlook recreates it as needed. |
| DEPL-03 | MSI writes resiliency AddinList registry key for Outlook 16.0 | Section: Architecture Patterns -- Pattern 3 (AddinList key), Code Examples -- Example 3. REG_SZ value "1" = always enabled. Path: HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\AddinList. Can be static .vdproj entry. |
| DEPL-04 | Installer validated for both 32-bit and 64-bit Office deployments | Section: Architecture Patterns -- Pattern 4 (Bitness validation). Current .vdproj targets x64 (TargetPlatform=1). HKCU registry is shared across bitness. VSTO add-in is AnyCPU. Validation is manual test on both Office bitnesses. |
</phase_requirements>

## Standard Stack

### Core

| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| System.Configuration.Install | .NET Framework 4.8 BCL | Installer class base for MSI custom actions | Built into .NET Framework; the standard mechanism for VS Setup Project custom actions |
| Microsoft.Win32.Registry | .NET Framework 4.8 BCL | Registry read/write/delete operations | Built into .NET Framework; no external dependencies; direct HKCU access |
| Visual Studio Installer Projects | VS 2022 Extension | .vdproj MSI builder with Custom Actions support | Already in use for the project; supports Installer Class custom actions natively |

### Supporting

| Library | Version | Purpose | When to Use |
|---------|---------|---------|-------------|
| Orca (MSI Editor) | Windows SDK | Inspect MSI registry tables post-build | Verification only -- to confirm registry entries in built MSI without installing |

### Alternatives Considered

| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| C# Installer Class | VBScript MSI custom action | VBScript is harder to debug, no type safety, no IDE support; C# Installer Class integrates directly with VS Setup Project |
| C# Installer Class | WiX Toolset with custom actions | Would require migrating from .vdproj to WiX -- massive scope change for a single custom action; not justified |
| Static .vdproj AddinList entry | Custom action for AddinList too | AddinList is a static value that does not need conditional logic; .vdproj native registry entry is simpler and more reliable |
| Deleting entire DisabledItems key | Parsing REG_BINARY to selectively remove entries | Binary format is undocumented by Microsoft; parsing is fragile; deleting entire key is safe because Outlook recreates it |

**No new NuGet packages required.** All registry operations use .NET Framework BCL classes.

## Architecture Patterns

### Recommended Project Structure

```
PhishingReporter.sln
+-- PhishingReporter/           # Existing VSTO add-in project
+-- Installer/                  # Existing .vdproj Setup Project
|   +-- Installer.vdproj       # Modified: custom action + AddinList registry key
+-- InstallerActions/           # NEW: Class library for custom action
    +-- InstallerActions.csproj # Targets .NET Framework 4.8
    +-- RegistryRemediation.cs  # [RunInstaller(true)] Installer class
```

### Pattern 1: C# Installer Class for Registry Remediation

**What:** A class library containing a class derived from `System.Configuration.Install.Installer` with the `[RunInstaller(true)]` attribute. The `Install` method performs registry operations.

**When to use:** When MSI needs to execute programmatic logic (read/write/delete registry values) that cannot be expressed as static registry entries in the .vdproj Registry Editor.

**How it works in the MSI lifecycle:**
1. MSI upgrade triggers `RemovePreviousVersions` (existing setting is TRUE in .vdproj)
2. New MSI installs files to disk
3. Custom actions execute (Install phase) -- this is where registry remediation runs
4. Static registry entries from .vdproj are written

**Critical constraint:** Custom actions in VS Setup Projects run as "deferred" actions AFTER files are installed. They have access to `Context.Parameters` (passed via CustomActionData) but NOT to MSI properties directly. The Installer class runs in the installing user's context, so HKCU operations target the correct user.

**Wiring into .vdproj:**
1. Add InstallerActions project output to the Setup Project's Application Folder
2. Open Custom Actions Editor (View > Custom Actions)
3. Add the primary output from InstallerActions to the "Install" node
4. Set `InstallerClass = True` in the custom action properties
5. Optionally set CustomActionData to pass parameters

### Pattern 2: DisabledItems Key Cleanup Strategy

**What:** Delete the entire `HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\DisabledItems` subkey rather than parsing individual REG_BINARY entries.

**Why delete the whole key:**
- The DisabledItems binary format is NOT officially documented by Microsoft (confirmed via Microsoft Q&A and official docs search)
- Each disabled add-in is stored as a REG_BINARY value with an undocumented encoding
- Parsing the binary to selectively remove only our add-in is fragile and error-prone
- Deleting the entire key is SAFE: Outlook recreates the DisabledItems key automatically when it needs to disable an add-in
- This is the standard remediation approach used by enterprise IT (confirmed via multiple community sources)

**Risk mitigation:** If other add-ins were also disabled, deleting the whole key gives them a fresh chance to load too. This is generally a positive side effect. If an add-in was disabled for good reason (crashes), Outlook will re-disable it on the next occurrence.

### Pattern 3: AddinList Managed Add-in Key

**What:** Write a REG_SZ value under `HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\AddinList` to mark the add-in as "always enabled" in the managed add-in policy.

**Registry specification (from Microsoft Docs):**
```
Key:   HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\AddinList
Name:  ZeroD.PhishReporter  (REG_SZ, the ProgID)
Value: "1"  (String "1" = always enabled; "0" = always disabled; "2" = configurable by user)
```

**Important distinction from DoNotDisableAddinList:**
- `DoNotDisableAddinList` (DWORD) = Prevents the add-in disabling FEATURE from triggering. The DWORD value indicates which disable reason to exempt (0x01 = boot load).
- `AddinList` (REG_SZ) = Managed add-in list from group policy. Value "1" = always enabled regardless of user actions or disabling feature. This is the stronger protection.

**Implementation choice:** This can be a static .vdproj registry entry (same pattern as DoNotDisableAddinList from Plan 01-02) because it is a fixed value that does not depend on existing machine state. No custom action needed for this specific key.

### Pattern 4: 32-bit / 64-bit Office Validation

**What:** Verify the MSI and add-in work on both 32-bit and 64-bit Office installations.

**Current state:**
- The .vdproj has `TargetPlatform = 3:1` (x64) and `DefaultLocation = [ProgramFiles64Folder]`
- The VSTO project compiles as `AnyCPU` -- the managed assembly loads in both bitnesses
- Office Interop references use `EmbedInteropTypes = true` -- no PIA deployment needed
- HKCU registry keys are NOT affected by WOW6432Node redirection (only HKLM is redirected)

**Key fact from Microsoft Docs:** "If the installer is targeting the current user, it does NOT need to install to the WOW6432Node because the HKEY_CURRENT_USER\Software path is shared." This means all HKCU registry operations (LoadBehavior, DisabledItems, DoNotDisableAddinList, AddinList) work identically on both Office bitnesses.

**What to validate:**
1. MSI installs on a machine with 64-bit Office -- add-in loads in Outlook
2. MSI installs on a machine with 32-bit Office -- add-in loads in Outlook
3. Registry keys are present and correct on both configurations
4. The VSTO Runtime prerequisite check passes on both

**Potential issue:** The installer uses `ProgramFiles64Folder` as the install target. On a machine with 32-bit Office running on 64-bit Windows, the add-in files will be in `C:\Program Files\...` but the VSTO Manifest registry value points to `[TARGETDIR]` which resolves correctly regardless. The VSTO runtime loads assemblies by the Manifest path, not by the Office bitness.

### Anti-Patterns to Avoid

- **Do NOT parse DisabledItems REG_BINARY values** -- the format is undocumented; delete the entire key instead
- **Do NOT use HKLM for remediation keys** -- the existing installer writes per-user (HKPU/HKCU); mixing hives creates inconsistency
- **Do NOT set LoadBehavior during Uninstall** -- only during Install/Commit; otherwise the old add-in version's uninstall step may conflict
- **Do NOT skip the AddinList key because DoNotDisableAddinList exists** -- they serve different purposes; belt-and-suspenders approach per project research
- **Do NOT create two separate MSIs for 32/64-bit** -- unnecessary when registry is HKCU (shared) and assembly is AnyCPU
- **Do NOT add the custom action to the Uninstall node** -- registry remediation only makes sense on install/upgrade

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| DisabledItems binary parsing | Custom REG_BINARY decoder | Delete entire DisabledItems subkey | Format is undocumented by Microsoft; Outlook recreates the key automatically |
| MSI custom action framework | Custom DLL with DllExport | System.Configuration.Install.Installer class | .vdproj natively supports Installer classes; no custom interop needed |
| Registry path discovery | Hardcoded version-specific paths | Constants with version string | Version 16.0 covers Office 2016/2019/365; single constant suffices for current targets |
| Conditional upgrade detection | Custom registry flag to track "already remediated" | MSI RemovePreviousVersions=TRUE + unconditional remediation | Setting LoadBehavior=3 is idempotent; running on clean installs is harmless |

**Key insight:** The remediation operations (set LoadBehavior=3, delete DisabledItems, write AddinList) are all IDEMPOTENT. Running them on a clean machine or a machine that was never disabled causes no harm. This eliminates the need for complex "is this an upgrade of a disabled machine?" detection logic.

## Common Pitfalls

### Pitfall 1: Custom Action Runs But Registry Not Changed

**What goes wrong:** The Installer Class executes without errors, but HKCU values are unchanged after MSI install.

**Why it happens:** The `InstallerClass` property on the custom action in the .vdproj is set to `False` (default). When False, the setup project tries to run the DLL as a native MSI custom action, not as a .NET Installer class. It silently does nothing.

**How to avoid:** After adding the custom action to the Install node, select it in the Custom Actions editor and set `InstallerClass = True` in the Properties window.

**Warning signs:** No exceptions, no errors, but registry values unchanged after install.

### Pitfall 2: LoadBehavior Reset to 2 Again After First Outlook Launch

**What goes wrong:** The custom action successfully sets LoadBehavior=3 during install, but after Outlook starts, it reverts to 2.

**Why it happens:** The DisabledItems key was not cleared. Outlook reads DisabledItems on startup and re-disables the add-in (setting LoadBehavior back to 2) if it finds the add-in in the disabled list.

**How to avoid:** Always clear DisabledItems IN ADDITION to setting LoadBehavior. The two operations must be performed together. Order: (1) delete DisabledItems, (2) set LoadBehavior=3.

**Warning signs:** LoadBehavior=3 in registry immediately after install, but LoadBehavior=2 after first Outlook launch.

### Pitfall 3: Custom Action Fails on Clean Install (No Resiliency Key)

**What goes wrong:** The custom action throws an exception because `HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency` does not exist on a fresh machine.

**Why it happens:** The Resiliency key and its subkeys (DisabledItems, etc.) are created by Outlook on first use, not during Office installation. On a machine that never had an add-in disabled, these keys may not exist.

**How to avoid:** Check for key existence before attempting delete operations. Use `Registry.CurrentUser.OpenSubKey(path, writable: true)` and check for null return. Wrap all operations in try/catch with logging.

**Warning signs:** MSI install fails on clean test machines; works on machines that previously had the add-in.

### Pitfall 4: vdproj AddinList Uses Wrong Value Type

**What goes wrong:** The AddinList key is deployed but Outlook does not honor it.

**Why it happens:** AddinList uses REG_SZ (string) values, NOT DWORD. The value must be the string `"1"`, not the integer `1`. DoNotDisableAddinList uses DWORD. Mixing them up is easy because they are sibling keys under Resiliency.

**How to avoid:** In the .vdproj Registry Editor: AddinList value must use `ValueTypes = 3:1` (REG_SZ) and `Value = 8:1` (string "1"). DoNotDisableAddinList uses `ValueTypes = 3:3` (DWORD) and `Value = 3:1` (integer 1).

**Warning signs:** AddinList key exists in RegEdit but shows as DWORD instead of REG_SZ; Outlook ignores it.

### Pitfall 5: MSI TargetPlatform Mismatch With 32-bit Office

**What goes wrong:** The MSI installs but the add-in does not load in 32-bit Office.

**Why it happens:** The .vdproj `TargetPlatform = 3:1` (x64) means the MSI is a 64-bit package. On a machine with 32-bit Office on 64-bit Windows, if the add-in files were registered under HKLM, the 32-bit Office would look in WOW6432Node. However, this installer uses HKCU (per-user), which is shared and not redirected.

**How to avoid:** Since all registry entries are under HKCU/HKPU (per-user), the HKCU path is architecture-neutral. The current x64 TargetPlatform is acceptable. The VSTO manifest path `file:///[TARGETDIR]PhishingReporter.vsto|vstolocal` resolves correctly because TARGETDIR is an absolute path set during install. Verify on both Office bitnesses as a manual test.

**Warning signs:** Add-in loads on 64-bit Office but not 32-bit Office on the same Windows architecture.

## Code Examples

### Example 1: Installer Class with LoadBehavior Reset (DEPL-01)

```csharp
// Source: Microsoft Docs - Registry entries for VSTO Add-ins + Custom Actions in VS Setup Projects
using System;
using System.Collections;
using System.ComponentModel;
using System.Configuration.Install;
using Microsoft.Win32;

namespace InstallerActions
{
    [RunInstaller(true)]
    public class RegistryRemediation : Installer
    {
        // Add-in ProgID must match the registration in Addins key
        private const string AddInProgId = "ZeroD.PhishReporter";

        // Office 16.0 covers Office 2016, 2019, and Microsoft 365
        private const string OfficeVersion = "16.0";

        private const string AddinsKeyPath =
            @"Software\Microsoft\Office\Outlook\Addins\" + AddInProgId;

        private const string ResiliencyBasePath =
            @"Software\Microsoft\Office\" + OfficeVersion + @"\Outlook\Resiliency";

        public override void Install(IDictionary stateSaver)
        {
            base.Install(stateSaver);
            RemediateDisabledState();
        }

        public override void Commit(IDictionary savedState)
        {
            base.Commit(savedState);
        }

        private void RemediateDisabledState()
        {
            try
            {
                // DEPL-02: Clear DisabledItems FIRST (before LoadBehavior reset)
                ClearDisabledItems();

                // DEPL-01: Reset LoadBehavior to 3 (load at startup)
                ResetLoadBehavior();
            }
            catch (Exception)
            {
                // Custom action must not fail the install
                // Registry remediation is best-effort
            }
        }

        private void ResetLoadBehavior()
        {
            using (var key = Registry.CurrentUser.OpenSubKey(AddinsKeyPath, writable: true))
            {
                if (key == null) return; // Fresh install, no existing key

                var currentValue = key.GetValue("LoadBehavior");
                if (currentValue is int intValue && intValue != 3)
                {
                    key.SetValue("LoadBehavior", 3, RegistryValueKind.DWord);
                }
            }
        }

        private void ClearDisabledItems()
        {
            var disabledItemsPath = ResiliencyBasePath + @"\DisabledItems";

            try
            {
                using (var resiliencyKey = Registry.CurrentUser.OpenSubKey(
                    ResiliencyBasePath, writable: true))
                {
                    if (resiliencyKey == null) return; // No Resiliency key exists

                    // Check if DisabledItems subkey exists before deleting
                    using (var disabledKey = resiliencyKey.OpenSubKey("DisabledItems"))
                    {
                        if (disabledKey == null) return; // No DisabledItems key
                    }

                    resiliencyKey.DeleteSubKey("DisabledItems", throwOnMissingSubKey: false);
                }
            }
            catch (Exception)
            {
                // Best-effort: if we cannot clear DisabledItems,
                // the DoNotDisableAddinList key provides fallback protection
            }
        }
    }
}
```

### Example 2: InstallerActions.csproj (Minimal Class Library)

```xml
<Project ToolsVersion="15.0" DefaultTargets="Build"
         xmlns="http://schemas.microsoft.com/developer/msbuild/2003">
  <PropertyGroup>
    <Configuration Condition=" '$(Configuration)' == '' ">Release</Configuration>
    <Platform Condition=" '$(Platform)' == '' ">AnyCPU</Platform>
    <OutputType>Library</OutputType>
    <RootNamespace>InstallerActions</RootNamespace>
    <AssemblyName>InstallerActions</AssemblyName>
    <TargetFrameworkVersion>v4.8</TargetFrameworkVersion>
  </PropertyGroup>
  <PropertyGroup Condition=" '$(Configuration)|$(Platform)' == 'Release|AnyCPU' ">
    <DebugType>pdbonly</DebugType>
    <Optimize>true</Optimize>
    <OutputPath>bin\Release\</OutputPath>
  </PropertyGroup>
  <ItemGroup>
    <Reference Include="System" />
    <Reference Include="System.Configuration.Install" />
  </ItemGroup>
  <ItemGroup>
    <Compile Include="RegistryRemediation.cs" />
  </ItemGroup>
  <Import Project="$(MSBuildToolsPath)\Microsoft.CSharp.targets" />
</Project>
```

### Example 3: AddinList Static Registry Entry in .vdproj

The AddinList key goes under the EXISTING Resiliency key hierarchy that was created in Plan 01-02.
The key path is: `HKPU\Software\Microsoft\Office\16.0\Outlook\Resiliency\AddinList`

```
"{60EA8692-D2D5-43EB-80DC-7906BF13D6EF}:_NEW_GUID_ADDINLIST"
{
"Name" = "8:AddinList"
"Condition" = "8:"
"AlwaysCreate" = "11:FALSE"
"DeleteAtUninstall" = "11:FALSE"
"Transitive" = "11:FALSE"
    "Keys"
    {
    }
    "Values"
    {
        "{ADCFDA98-8FDD-45E4-90BC-E3D20B029870}:_NEW_GUID_ADDINLIST_VALUE"
        {
        "Name" = "8:ZeroD.PhishReporter"
        "Condition" = "8:"
        "Transitive" = "11:FALSE"
        "ValueTypes" = "3:1"
        "Value" = "8:1"
        }
    }
}
```

**Key differences from DoNotDisableAddinList:**
- `ValueTypes = 3:1` (REG_SZ string) NOT `3:3` (DWORD)
- `Value = 8:1` (string "1") NOT `3:1` (integer 1)
- Key name is `AddinList` NOT `DoNotDisableAddinList`

### Example 4: Wiring Custom Action into .vdproj

The CustomAction section of the .vdproj (currently empty at line 288-290) must be populated:

```
"CustomAction"
{
    "{4AA51A2D-7D85-4A59-BA75-B0809FC8B380}:_NEW_GUID_CUSTOMACTION"
    {
    "Name" = "8:RegistryRemediation"
    "Condition" = "8:"
    "Object" = "8:_GUID_OF_INSTALLERACTIONS_OUTPUT"
    "FileType" = "3:1"
    "InstallAction" = "3:1"
    "Arguments" = "8:"
    "EntryPoint" = "8:"
    "InstallerClass" = "11:TRUE"
    }
}
```

Note: The actual GUID for the Object reference is generated when you add the InstallerActions project output to the setup project via the VS IDE. This is best done through the Visual Studio UI rather than manual .vdproj editing.

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| GPO-only AddinList policy | AddinList + DoNotDisableAddinList + code fixes | Office 2013+ | Belt-and-suspenders: policy keys prevent future disabling; code fixes prevent the cause |
| Manual user re-enable via COM Add-ins dialog | MSI upgrade with automated remediation | Standard enterprise practice | Zero-touch remediation on upgrade; no user action required |
| Separate 32-bit and 64-bit MSI packages | Single MSI with HKCU registry (architecture-neutral) | VSTO best practice for per-user install | Simplifies SCCM/GPO deployment; one package handles both Office bitnesses |
| VBScript custom actions | C# Installer Class | .NET Framework era | Type-safe, debuggable, IDE-integrated |

**Deprecated/outdated:**
- VBScript MSI custom actions: Still functional but no IDE support, no debugging, error-prone
- HKLM-based add-in registration with WOW6432Node: Only needed for per-machine installs; this project uses per-user
- Office 15.0 (2013) registry paths: Target environment is Office 2016+ (16.0 only)

## Open Questions

1. **Should the custom action also target Office 15.0 (Outlook 2013)?**
   - What we know: The project research and ROADMAP explicitly target "Outlook 2016/2019/Microsoft 365" which all use the `16.0` registry path. Outlook 2013 uses `15.0`.
   - What's unclear: Whether any target enterprise machines still run Outlook 2013.
   - Recommendation: Target `16.0` only per existing project scope. Adding `15.0` is a one-line constant change if needed later.

2. **Should the DisabledItems cleanup target CrashingAddinList too?**
   - What we know: Outlook maintains separate `DisabledItems` (soft-disable for slow startup) and `CrashingAddinList` (hard-disable for crashes). DoNotDisableAddinList does NOT protect against crash-based disabling.
   - What's unclear: Whether any target machines have entries in CrashingAddinList.
   - Recommendation: Include CrashingAddinList cleanup in the custom action as defensive measure. If the add-in was previously crashing (pre-Phase 5 fixes), the crash entry would persist even after the code is fixed.

3. **Build environment for InstallerActions project**
   - What we know: The build environment has been flagged as potentially lacking "VSTO Office Tools workload" (from STATE.md blockers).
   - What's unclear: Whether the build environment can compile a plain C# class library that references only BCL assemblies (System.Configuration.Install).
   - Recommendation: InstallerActions is a plain class library, NOT a VSTO project. It only needs the .NET Framework 4.8 SDK, which is available in any VS 2022 installation. It should build regardless of whether Office Tools are installed.

## Sources

### Primary (HIGH confidence)
- [Microsoft Docs -- Registry entries for VSTO Add-ins](https://learn.microsoft.com/en-us/visualstudio/vsto/registry-entries-for-vsto-add-ins?view=vs-2022) -- LoadBehavior values (2=unloaded/load-at-startup, 3=loaded/load-at-startup), HKCU override behavior, WOW6432Node guidance for HKCU (not needed)
- [Microsoft Docs -- Support for keeping add-ins enabled](https://learn.microsoft.com/en-us/office/vba/outlook/concepts/getting-started/support-for-keeping-add-ins-enabled) -- AddinList key spec (REG_SZ, values 0/1/2), DoNotDisableAddinList key spec (DWORD, values 0x01-0x0A), disable reason codes
- [Microsoft Docs -- Deploy a VSTO Solution with Windows Installer](https://learn.microsoft.com/en-us/visualstudio/vsto/deploying-a-vsto-solution-by-using-windows-installer?view=vs-2022) -- Setup Project registry configuration, custom action wiring, 32/64-bit TargetPlatform guidance, Installer Class pattern
- [Microsoft Q&A -- Office addin Resiliency\DisabledItems registry](https://learn.microsoft.com/en-us/answers/questions/184800/office-addin-resiliencydisableditems-registry) -- Confirms binary format is undocumented; REG_BINARY entries per disabled add-in

### Secondary (MEDIUM confidence)
- [Red Gate Simple Talk -- Visual Studio Setup Projects and Custom Actions](https://www.red-gate.com/simple-talk/development/dotnet-development/visual-studio-setup-projects-and-custom-actions/) -- Installer Class wiring, CustomActionData format, deferred execution behavior, repair condition ("Not Installed")
- [Add-in Express Forum -- LoadBehavior set to 2 after MSI update](https://www.add-in-express.com/forum/read.php?FID=5&TID=15071) -- Confirms LoadBehavior=2 after upgrade is a known issue; caused by add-in exception during Outlook close
- Codebase inspection -- .vdproj confirms: TargetPlatform=1 (x64), ProgramFiles64Folder, HKPU registry, RemovePreviousVersions=TRUE, CustomAction section is currently empty, ProgID is ZeroD.PhishReporter

### Tertiary (LOW confidence)
- Community consensus on DisabledItems deletion -- Multiple sources recommend deleting entire DisabledItems key rather than parsing; no official Microsoft documentation on the binary format exists

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH -- all components are .NET Framework BCL; no new external dependencies
- Architecture: HIGH -- Installer Class pattern verified via Microsoft official VSTO deployment docs; registry key formats verified via Microsoft Docs
- Pitfalls: HIGH -- LoadBehavior/DisabledItems interaction verified via multiple sources; AddinList value type (REG_SZ vs DWORD) distinction verified via official Microsoft Docs
- 32/64-bit: MEDIUM -- HKCU architecture neutrality verified via Microsoft Docs; actual validation requires manual testing on both Office bitnesses

**Research date:** 2026-02-26
**Valid until:** 2026-03-28 (stable domain; registry key formats and .vdproj tooling do not change)
