# Phase 1: Foundation - Research

**Researched:** 2026-02-25
**Domain:** .NET Framework retargeting, NLog file logging, MSI registry deployment, NuGet package update
**Confidence:** HIGH

## Summary

Phase 1 establishes the project's technical foundation: retarget to .NET Framework 4.8, add structured file logging via NLog, update HtmlAgilityPack to 1.12.x, and embed the DoNotDisableAddinList registry key into the MSI installer. These are four independent, low-risk changes that unlock all subsequent phases.

The .NET 4.6.1 to 4.8 retarget is an in-place upgrade with no breaking changes relevant to this VSTO add-in. NLog requires careful setup because VSTO add-ins run inside `OUTLOOK.EXE` — the NLog config file must be loaded programmatically from the add-in's assembly directory, not from the host's working directory. The HtmlAgilityPack 1.12.x upgrade introduces no API changes affecting the existing `SelectNodes("//a[@href]")` pattern. The MSI registry key requires building a nested key hierarchy under HKCU in the `.vdproj` Registry Editor.

**Primary recommendation:** Execute all four changes in parallel within a single plan since they are independent. Validate the build compiles and the existing report-phishing workflow still works after each change.

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|-----------------|
| INFR-01 | Project upgraded to .NET Framework 4.8 | Section: .NET Framework 4.8 Retargeting — details exact csproj/packages.config/bootstrapper changes |
| INFR-02 | NLog file logging to %AppData%\PhishingReporter\logs\ with 7-day retention | Section: NLog Logging Setup — covers isolated LogFactory pattern, file target with archive, VSTO-specific config loading |
| INFR-03 | Structured log entries for all report workflow steps | Section: NLog Logging Setup — covers logger placement in Ribbon.cs, GoPhishIntegration.cs, and ThisAddIn.cs |
| INFR-04 | HtmlAgilityPack updated to latest stable version (1.12.x) | Section: HtmlAgilityPack Update — confirms no API breaking changes from 1.11.23 to 1.12.x |
| STRT-04 | Registry DoNotDisableAddinList key deployed via MSI | Section: MSI Registry Key — covers exact registry path, DWORD value, .vdproj structure |
</phase_requirements>

## Standard Stack

### Core

| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| .NET Framework | 4.8 | Runtime target | Final supported version for VSTO; pre-installed on Windows 10/11; in-place upgrade over 4.6.1 |
| NLog | 5.4.0 | File-based structured logging | Supports .NET Framework 3.5-4.8 with zero dependencies; FileTarget built-in; XML config adjustable post-deployment without rebuild; safer than 6.x for .NET Framework due to FileTarget rewrite in 6.0 |
| HtmlAgilityPack | 1.12.4 | HTML parsing for URL extraction | Drop-in upgrade from 1.11.23; no API changes affecting SelectNodes/XPath; targets Net45+ |

### Supporting

| Library | Version | Purpose | When to Use |
|---------|---------|---------|-------------|
| NLog.Schema | 5.4.0 | NLog.xsd for IntelliSense in NLog.config | Install as dev dependency for config editing; not required at runtime |

### Alternatives Considered

| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| NLog 5.4.0 | NLog 6.1.0 | 6.x has rewritten FileTarget (removed ConcurrentWrites, split into separate NuGet packages); higher risk for a first logging integration; 5.4.0 is battle-tested and sufficient |
| NLog 5.4.0 | Serilog 4.x | Serilog requires code-first config; cannot adjust log levels post-deployment via config file drop; NLog's XML config is an operational advantage in enterprise deployments |
| NLog 5.4.0 | System.Diagnostics.Trace | No structured logging, no file rotation, no retention policy; insufficient for diagnosing enterprise add-in failures |

**Installation (via NuGet Package Manager Console):**
```
Install-Package NLog -Version 5.4.0
Install-Package HtmlAgilityPack -Version 1.12.4
```

Note: This project uses `packages.config` (not PackageReference). NuGet commands will update both `packages.config` and the `.csproj` reference.

## Architecture Patterns

### Recommended Project Structure After Phase 1

```
PhishingReporter/
├── ThisAddIn.cs          # NLog LogFactory initialized here; shutdown on Shutdown event
├── Ribbon.cs             # Log entries added at each workflow step (existing code unchanged)
├── GoPhishIntegration.cs # Log entries for HTTP call attempt/result (existing code unchanged)
├── NLog.config           # XML config file — Copy to Output Directory = Copy Always
├── Properties/
│   ├── Settings.settings
│   └── AssemblyInfo.cs
├── packages.config       # Updated: NLog 5.4.0, HtmlAgilityPack 1.12.4
└── PhishingReporter.csproj  # Updated: TargetFrameworkVersion v4.8
```

### Pattern 1: Isolated LogFactory for VSTO Add-in

**What:** Use a dedicated `LogFactory` instance instead of the global `LogManager` to avoid conflicts with other add-ins that may use NLog in the same Outlook process.

**When to use:** Always in VSTO add-ins — the host process (`OUTLOOK.EXE`) may load multiple add-ins, each potentially using NLog with different configurations.

**Example:**
```csharp
// Source: https://github.com/NLog/NLog/wiki/Configure-component-logging
internal static class AppLogger
{
    public static LogFactory Instance { get { return _instance.Value; } }

    private static readonly Lazy<LogFactory> _instance =
        new Lazy<LogFactory>(BuildLogFactory);

    private static LogFactory BuildLogFactory()
    {
        var thisAssembly = System.Reflection.Assembly.GetExecutingAssembly();
        var configFilePath = System.IO.Path.Combine(
            System.IO.Path.GetDirectoryName(thisAssembly.Location),
            "NLog.config");

        var logFactory = new LogFactory();
        logFactory.Configuration = new NLog.Config.XmlLoggingConfiguration(
            configFilePath, logFactory);
        return logFactory;
    }
}

// Usage in any class:
private static readonly NLog.Logger Logger =
    AppLogger.Instance.GetCurrentClassLogger();
```

**Confidence:** HIGH — verified via [NLog Component Logging Wiki](https://github.com/NLog/NLog/wiki/Configure-component-logging)

### Pattern 2: NLog.config for %AppData% File Target with 7-Day Retention

**What:** Configure NLog FileTarget to write to the user's AppData directory with automatic daily archiving and 7-day retention.

**When to use:** For the INFR-02 requirement — persistent logging to `%AppData%\PhishingReporter\logs\`.

**Example:**
```xml
<?xml version="1.0" encoding="utf-8" ?>
<nlog xmlns="http://www.nlog-project.org/schemas/NLog.xsd"
      xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">

  <targets>
    <target name="logfile"
            xsi:type="File"
            fileName="${specialfolder:folder=ApplicationData}/PhishingReporter/logs/phishingreporter.log"
            layout="${longdate} [${level:uppercase=true}] ${logger} - ${message}${onexception:inner= ${exception:format=tostring}}"
            archiveFileName="${specialfolder:folder=ApplicationData}/PhishingReporter/logs/phishingreporter.{#}.log"
            archiveEvery="Day"
            archiveNumbering="Rolling"
            maxArchiveFiles="7"
            concurrentWrites="false"
            keepFileOpen="true" />
  </targets>

  <rules>
    <logger name="PhishingReporter.*" minlevel="Info" writeTo="logfile" />
  </rules>
</nlog>
```

**Confidence:** HIGH — verified via [NLog File Target Wiki](https://github.com/nlog/NLog/wiki/File-target) and [FileTarget Archive Examples](https://github.com/NLog/NLog/wiki/FileTarget-Archive-Examples)

### Pattern 3: Log Placement for Workflow Steps (INFR-03)

**What:** Add log entries at each step of the report-phishing workflow to enable enterprise diagnosis.

**Where to log:**
```
1. ThisAddIn_Startup       → "Add-in startup begin" / "Add-in startup complete"
2. reportPhishing()        → "Report phishing initiated by user"
3. reportPhishingEmailToSecurityTeam() → "Processing selected email: {subject}"
4. GoPhish header check    → "GoPhish header check: {found/not found}"
5. GoPhish HTTP call       → "GoPhish notification sent: {url}" / "GoPhish notification failed: {error}"
6. Report email compose    → "Report email composed for: {subject}"
7. Report email send       → "Report email sent to: {recipient}"
8. Error handler           → "Error during report processing: {exception}"
9. ThisAddIn_Shutdown      → "Add-in shutdown, NLog flushing"
```

**Confidence:** HIGH — derived from reading the actual codebase workflow in `Ribbon.cs` and `GoPhishIntegration.cs`

### Anti-Patterns to Avoid

- **Do NOT use `LogManager.GetCurrentClassLogger()`** — use the isolated `AppLogger.Instance.GetCurrentClassLogger()` to avoid config conflicts with other Outlook add-ins.
- **Do NOT put NLog.config in the project root expecting auto-discovery** — VSTO runs inside `OUTLOOK.EXE`, so NLog's default config discovery (looking next to the entry assembly) will look in the Office install directory, not your add-in directory.
- **Do NOT use `NLog.Extensions.Logging` on .NET Framework 4.8** — it has documented DLL version conflicts with `Microsoft.Extensions.Logging 2.1.x` on .NET Framework.
- **Do NOT change the add-in ProgID** (`ZeroD.PhishReporter`) during retargeting — the ProgID is deployment-permanent and is baked into the MSI registry entries and existing GPO configurations.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| File logging with rotation | Custom StreamWriter with date checks | NLog FileTarget with `archiveEvery` + `maxArchiveFiles` | File locking, concurrent access, archive cleanup, encoding — all handled by NLog |
| Log file path resolution | Hardcoded `%AppData%` string expansion | NLog `${specialfolder:folder=ApplicationData}` layout renderer | Cross-platform, handles missing directories gracefully |
| Config file discovery in add-in | `Path.Combine(Application.StartupPath, ...)` | `Assembly.GetExecutingAssembly().Location` + isolated LogFactory | `Application.StartupPath` resolves to OUTLOOK.EXE directory, not the add-in directory |
| Registry key deployment | Custom installer action in C# | `.vdproj` Registry Editor built-in feature | Custom actions in Setup Projects are fragile and cannot be debugged; registry keys are a first-class vdproj feature |

**Key insight:** The VSTO hosting model means standard .NET assumptions about file paths, config discovery, and assembly loading all break. NLog's isolated LogFactory pattern exists specifically for this scenario.

## Common Pitfalls

### Pitfall 1: NLog.config Not Found at Runtime

**What goes wrong:** NLog silently produces no output. No exceptions are thrown. The add-in appears to work but no log files appear.

**Why it happens:** NLog's auto-discovery algorithm looks for `NLog.config` next to the entry assembly (OUTLOOK.EXE) or in the app's base directory. In a VSTO add-in, neither location contains your config file.

**How to avoid:** Use the isolated LogFactory pattern (Pattern 1 above) that explicitly loads the config from `Assembly.GetExecutingAssembly().Location`. Set `NLog.config` file properties to `Build Action = Content` and `Copy to Output Directory = Copy Always` in the `.csproj`.

**Warning signs:** No log files created in `%AppData%\PhishingReporter\logs\` after clicking the report button.

### Pitfall 2: NLog Assembly Not Deployed by MSI

**What goes wrong:** `FileNotFoundException` for `NLog.dll` at runtime on user machines. Works fine on the developer machine.

**Why it happens:** The Visual Studio Setup Project (.vdproj) only includes project output. If the NLog NuGet package is not properly referenced as a dependency in the project output, the MSI won't include `NLog.dll`.

**How to avoid:** After installing the NLog NuGet package, verify that `NLog.dll` appears in the `bin\Release\` output. Then verify the Setup Project's "Primary Output from PhishingReporter" detected output group includes `NLog.dll`. If it does not, manually add `NLog.dll` to the Setup Project's Application Folder.

**Warning signs:** Build succeeds but MSI is suspiciously small; `NLog.dll` missing from `bin\Release\`.

### Pitfall 3: .NET Framework 4.8 Retarget Breaks Bootstrapper

**What goes wrong:** The MSI setup project fails to build because the bootstrapper package for `.NETFramework,Version=v4.8` is not installed in the Visual Studio bootstrapper cache.

**Why it happens:** Visual Studio Setup Projects have a bootstrapper system that bundles prerequisite installers. Changing from `v4.6.1` to `v4.8` requires updating the `BootstrapperPackage` element in the `.csproj` AND ensuring the VS bootstrapper cache contains the 4.8 prerequisite.

**How to avoid:** Update the `.csproj` `BootstrapperPackage` to reference `.NETFramework,Version=v4.8`. Since .NET 4.8 is pre-installed on Windows 10 1903+ and all Windows 11, you can set the bootstrapper to `Install = false` (don't include the redistributable in the MSI) — the runtime is already present on all target machines.

**Warning signs:** MSI build error mentioning bootstrapper package not found.

### Pitfall 4: HtmlAgilityPack Assembly Version Mismatch

**What goes wrong:** The project compiles but at runtime throws `FileLoadException` because the assembly version in the reference doesn't match the version in the deployed DLL.

**Why it happens:** The `.csproj` has a hard-coded `HintPath` and `Version` attribute in the `<Reference>` for HtmlAgilityPack. Simply dropping in a new DLL without updating the reference causes a version mismatch.

**How to avoid:** Use NuGet Package Manager to perform the upgrade (`Update-Package HtmlAgilityPack`). This automatically updates `packages.config`, the `.csproj` reference (including `Version` and `HintPath`), and downloads the correct DLL. Do NOT manually edit these files.

**Warning signs:** `FileLoadException` mentioning HtmlAgilityPack version; runtime crash on first URL extraction.

### Pitfall 5: DoNotDisableAddinList ProgID Mismatch

**What goes wrong:** The registry key is deployed but Outlook still disables the add-in.

**Why it happens:** The DWORD value name under `DoNotDisableAddinList` must exactly match the add-in's ProgID as registered under `HKCU\Software\Microsoft\Office\Outlook\Addins\`. The current installer registers the add-in as `ZeroD.PhishReporter`. If the registry entry uses a different string (e.g., the assembly name `PhishingReporter` or a typo), Outlook ignores it.

**How to avoid:** The DWORD value name must be exactly `ZeroD.PhishReporter` (matching the existing registration in the installer). Verify by examining the `.vdproj` Registry section which already creates `HKPU\Software\Microsoft\Office\Outlook\Addins\ZeroD.PhishReporter`.

**Warning signs:** Registry key exists but `ZeroD.PhishReporter` not listed when viewed in RegEdit under `DoNotDisableAddinList`.

## Code Examples

### Example 1: .csproj TargetFrameworkVersion Change

```xml
<!-- Before -->
<TargetFrameworkVersion>v4.6.1</TargetFrameworkVersion>

<!-- After -->
<TargetFrameworkVersion>v4.8</TargetFrameworkVersion>
```

Also update the BootstrapperPackage:
```xml
<!-- Before -->
<BootstrapperPackage Include=".NETFramework,Version=v4.6.1">
  <Visible>False</Visible>
  <ProductName>Microsoft .NET Framework 4.6.1 %28x86 and x64%29</ProductName>
  <Install>true</Install>
</BootstrapperPackage>

<!-- After -->
<BootstrapperPackage Include=".NETFramework,Version=v4.8">
  <Visible>False</Visible>
  <ProductName>Microsoft .NET Framework 4.8 %28x86 and x64%29</ProductName>
  <Install>false</Install>
</BootstrapperPackage>
```

Source: Codebase inspection of `PhishingReporter.csproj` (line 29, 49-53)

### Example 2: packages.config After Updates

```xml
<?xml version="1.0" encoding="utf-8"?>
<packages>
  <package id="HtmlAgilityPack" version="1.12.4" targetFramework="net48" />
  <package id="NLog" version="5.4.0" targetFramework="net48" />
</packages>
```

### Example 3: Adding Log Entries to Existing Code (Non-Breaking)

```csharp
// In Ribbon.cs — reportPhishing method
private static readonly NLog.Logger Logger =
    AppLogger.Instance.GetCurrentClassLogger();

public void reportPhishing(Office.IRibbonControl control)
{
    Logger.Info("Report phishing button clicked");

    var areYouSure = MessageBox.Show(
        "Do you want to report this email to the Information Security Team as a potential phishing attempt?",
        "Are you sure?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

    if (areYouSure == DialogResult.Yes)
    {
        Logger.Info("User confirmed report submission");
        reportPhishingEmailToSecurityTeam(control);
    }
    else
    {
        Logger.Info("User cancelled report submission");
    }
}
```

### Example 4: NLog Initialization and Shutdown in ThisAddIn.cs

```csharp
private void ThisAddIn_Startup(object sender, System.EventArgs e)
{
    var logger = AppLogger.Instance.GetCurrentClassLogger();
    logger.Info("PhishingReporter add-in startup begin");
    // ... existing startup code (currently empty) ...
    logger.Info("PhishingReporter add-in startup complete");
}

private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
{
    var logger = AppLogger.Instance.GetCurrentClassLogger();
    logger.Info("PhishingReporter add-in shutdown");
    AppLogger.Instance.Shutdown();
}
```

### Example 5: .vdproj Registry Key Structure for DoNotDisableAddinList

The MSI must create the following nested registry path under HKCU:

```
HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList
    ZeroD.PhishReporter (DWORD) = 1
```

In the `.vdproj` Registry Editor, this requires creating a hierarchy of nested keys:
1. Under `HKCU` > `Software` (already exists)
2. `Microsoft` > `Office` > `16.0` > `Outlook` > `Resiliency` > `DoNotDisableAddinList`
3. Add a DWORD value named `ZeroD.PhishReporter` with value `1`

The value `1` means "do not disable for boot load (LoadBehavior = 3)".

**Important notes:**
- The path uses `16.0` which covers Office 2016, 2019, and Microsoft 365.
- Use HKCU, not HKLM — the `DoNotDisableAddinList` key is per-user by design (under Resiliency).
- The existing installer already writes to `HKPU` (per-user) at `Software\Microsoft\Office\Outlook\Addins\ZeroD.PhishReporter`.
- The `DoNotDisableAddinList` is a separate key under `Resiliency`, NOT under `Addins`.

Source: [Microsoft Docs — Support for keeping add-ins enabled](https://learn.microsoft.com/en-us/office/vba/outlook/concepts/getting-started/support-for-keeping-add-ins-enabled)

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| .NET Framework 4.6.1 | .NET Framework 4.8 | 2019 (4.8 RTM) | Final .NET Framework; pre-installed on Win10/11; no JIT overhead from framework download |
| NLog.Config NuGet package | Manual NLog.config creation | NLog 5.0 (2022) | NLog.Config package deprecated; config files must be manually created and added to project |
| NLog global LogManager | Isolated LogFactory per component | NLog 4.x+ | Prevents config conflicts in shared host processes like OUTLOOK.EXE |
| HtmlAgilityPack 1.11.x (Net45 target) | HtmlAgilityPack 1.12.x | March 2025 | Dropped NetStandard 1.3/1.6; added .NET 8; no API breaking changes for Net45+ consumers |

**Deprecated/outdated:**
- `NLog.Config` NuGet package: No longer published. Create `NLog.config` manually.
- `HttpWebRequest` + `HttpWebResponse`: Still functional but replaced by `HttpClient` in Phase 3. Not changed in this phase.
- .NET Framework 4.6.1 targeting: Still builds but 4.8 is strictly superior (pre-installed on all target machines, extended support until 2029+).

## Open Questions

1. **NLog 5.4.0 vs 6.1.0 for .NET Framework 4.8**
   - What we know: Both support .NET Framework 4.8. NLog 6.x rewrote FileTarget (removed ConcurrentWrites, split into separate packages). NLog 5.4.0 is the last 5.x release and is battle-tested.
   - What's unclear: Whether NLog 6.x's FileTarget changes introduce any edge cases in VSTO hosting scenarios.
   - Recommendation: Use NLog 5.4.0. It is stable, well-documented, and the FileTarget behavior is proven. Upgrading to 6.x can be done in a future milestone if needed. The previous project research recommended 6.1.0, but for a first-time logging integration in a reliability-critical add-in, the conservative choice is safer.

2. **Outlook 2016 (15.0) vs. Office 365/2019 (16.0) registry key path**
   - What we know: The `DoNotDisableAddinList` uses `x.0` version numbering. Office 2016/2019/365 all use `16.0`. Outlook 2013 uses `15.0`.
   - What's unclear: Whether the target enterprise still has any Outlook 2013 (15.0) deployments.
   - Recommendation: Write the registry key for `16.0` only. The PROJECT.md specifies "Outlook 2016/2019/Microsoft 365" as the target environment. If 15.0 support is needed later, add a second key path.

3. **MSI installer rebuild validation**
   - What we know: The `.vdproj` Setup Project must include the new NLog.dll and updated HtmlAgilityPack.dll in its detected dependencies.
   - What's unclear: Whether the existing Setup Project auto-detects NuGet package DLLs or requires manual inclusion.
   - Recommendation: After NuGet package install, verify `bin\Release\` contains `NLog.dll`. Then rebuild the Setup Project and check that the MSI file size increased appropriately. If NLog.dll is missing, manually add it to the Setup Project's Application Folder.

## Sources

### Primary (HIGH confidence)
- [Microsoft Docs — Retargeting changes for .NET Framework 4.8.x](https://learn.microsoft.com/en-us/dotnet/framework/migration-guide/retargeting/4.8.x) — Breaking changes from 4.6.1 to 4.8: only FIPS mode cryptography relaxation, WPF accessibility, and WF checksum changes; none affect VSTO/COM add-ins
- [Microsoft Docs — Support for keeping add-ins enabled](https://learn.microsoft.com/en-us/office/vba/outlook/concepts/getting-started/support-for-keeping-add-ins-enabled) — DoNotDisableAddinList registry key specification, exact path, DWORD values
- [Microsoft Docs — Registry entries for VSTO Add-ins](https://learn.microsoft.com/en-us/visualstudio/vsto/registry-entries-for-vsto-add-ins?view=vs-2022) — ProgID format, LoadBehavior values, HKLM vs HKCU deployment
- [NuGet — NLog 5.4.0](https://www.nuget.org/packages/NLog/5.4.0) — Framework support (.NET Framework 3.5-4.8), zero dependencies
- [NuGet — HtmlAgilityPack 1.12.4](https://www.nuget.org/packages/HtmlAgilityPack/) — Framework support (Net45+), zero dependencies, published 2025-10-03
- [NLog Wiki — Configure component logging](https://github.com/NLog/NLog/wiki/Configure-component-logging) — Isolated LogFactory pattern for add-ins/plugins
- [NLog Wiki — File target](https://github.com/nlog/NLog/wiki/File-target) — Archive configuration, layout renderers, specialfolder usage
- [NLog Wiki — FileTarget Archive Examples](https://github.com/NLog/NLog/wiki/FileTarget-Archive-Examples) — Rolling archive with maxArchiveFiles

### Secondary (MEDIUM confidence)
- [GitHub — HtmlAgilityPack releases](https://github.com/zzzprojects/html-agility-pack/releases) — 1.12.x changelog showing no breaking API changes for SelectNodes/XPath patterns
- [NLog 6.0 release notes](https://nlog-project.org/2025/06/21/nlog-6-0-released.html) — FileTarget ConcurrentWrites removal, package split; informed NLog 5.x recommendation
- Codebase inspection — .vdproj Registry section confirms existing ProgID `ZeroD.PhishReporter` under `HKPU\Software\Microsoft\Office\Outlook\Addins`

### Tertiary (LOW confidence)
- [WebSearch — VSTO assembly loading issues with NLog](https://learn.microsoft.com/en-us/archive/msdn-technet-forums/7970e0fd-7297-440a-8bd9-09ba1a0f4a09) — Assembly resolution challenges in VSTO; confirms need for explicit config loading

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH — NLog 5.4.0 and HtmlAgilityPack 1.12.4 verified via NuGet; .NET 4.8 retargeting verified via Microsoft migration docs
- Architecture: HIGH — NLog isolated LogFactory pattern verified via official NLog wiki; VSTO-specific concerns documented
- Pitfalls: HIGH — NLog config discovery, assembly deployment, ProgID mismatch all verified via official sources and codebase inspection

**Research date:** 2026-02-25
**Valid until:** 2026-03-27 (stable libraries, no fast-moving dependencies)
