# Phase 5: Startup Reliability - Research

**Researched:** 2026-02-26
**Domain:** VSTO add-in startup optimization, Outlook resiliency measurement, deferred initialization, ribbon reflection bypass, CLR cold-start mitigation
**Confidence:** HIGH

## Summary

Phase 5 addresses the root cause of Outlook auto-disabling the add-in: the measured startup time exceeds Outlook's hard-coded 1,000 ms resiliency threshold. Outlook measures the time from when it calls the COM `IDTExtensibility2.OnConnection` method (which for VSTO add-ins triggers the VSTO runtime initialization chain: assembly loading, `Initialize()`, `CreateRibbonExtensibilityObject()`, `FinishInitialization()`, `InternalStartup()`, `ThisAddIn_Startup`) until that entire sequence returns control to Outlook. The "Boot Time (Milliseconds)" value logged in Event ID 45 captures this duration. If the median Boot Time over 5 consecutive cold starts exceeds 1,000 ms, Outlook disables the add-in with reason code 0x00000001 (Boot load).

The current codebase already has a correctly implemented `CreateRibbonExtensibilityObject()` override in `ThisAddIn.cs` that returns `new Ribbon()` directly (Ribbon XML approach using `IRibbonExtensibility`). This means the VSTO Ribbon Designer reflection scan is already bypassed -- the runtime does not scan assemblies for `IRibbonExtension` implementations. However, there are two remaining performance concerns that could push boot time over the threshold:

1. **Static constructor chains triggered during startup**: `GoPhishIntegration` has a static constructor that creates an `HttpClient`, configures `ServicePointManager`, and builds a Polly `ResiliencePipeline`. If any code path during startup triggers CLR type-loading of `GoPhishIntegration`, its static constructor runs inside the measurement window. Similarly, `AppLogger.BuildLogFactory()` loads NLog configuration from disk.

2. **CLR cold-start overhead**: The first time the .NET CLR loads into the Outlook process after system restart, JIT compilation of referenced assemblies (NLog, Polly, HtmlAgilityPack, System.Net.Http) adds 2-5 seconds. This is compounded by Authenticode certificate revocation list (CRL) checking, which can add 15+ seconds if the machine has no internet or slow connectivity.

The strategy is threefold: (a) ensure `ThisAddIn_Startup` does essentially nothing (it already does only two log calls -- verify this is truly sub-5 ms); (b) add `<generatePublisherEvidence enabled="false"/>` to app.config to eliminate CRL check delays; (c) defer any remaining heavy initialization (e.g., ensure `GoPhishIntegration` static constructor is not triggered until first user interaction rather than during startup). If cold-start JIT remains a problem after these changes, NGen via MSI Custom Action is the established solution.

**Primary recommendation:** Add `<generatePublisherEvidence enabled="false"/>` to app.config, verify no static constructors trigger during the VSTO startup path, add high-resolution timing instrumentation to the startup sequence, and optionally add NGen install/uninstall Custom Actions to the MSI for pre-JIT compilation.

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|-----------------|
| STRT-01 | Add-in initialization deferred to Application.Startup event (outside Outlook's resiliency measurement window) | The VSTO startup measurement window covers `OnConnection` through `ThisAddIn_Startup` return. Any non-trivial initialization must be deferred to `Application.Startup` (fires after all add-ins finish loading) or a timer callback. Current `ThisAddIn_Startup` is already near-empty (two log calls). Research identifies `GoPhishIntegration` static constructor and NLog config load as potential startup-path hazards that must be verified and, if triggered, deferred. |
| STRT-02 | CreateRibbonExtensibilityObject overridden in ThisAddIn to eliminate VSTO reflection scan from startup path | Already implemented in the current codebase (`ThisAddIn.cs` line 35-38). The override returns `new Ribbon()` directly, bypassing the VSTO runtime's assembly-scanning reflection. Phase 5 must verify this remains correct and document it as a validated requirement. |
| STRT-03 | Add-in startup time stays under Outlook's 1,000 ms resiliency threshold on typical enterprise hardware | Multiple optimization vectors researched: `generatePublisherEvidence` app.config setting to eliminate CRL delays, NGen pre-JIT for cold-start mitigation, deferred static initialization to keep heavy code out of the measurement window, and Event ID 45 monitoring for verification. Success requires median Boot Time < 1,000 ms over 5 cold starts. |
</phase_requirements>

## Standard Stack

### Core
| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| .NET Framework | 4.8 | Runtime (already targeted) | CRL bypass via `generatePublisherEvidence` is a .NET Framework app.config feature; NGen is a .NET Framework tool |
| VSTO Runtime | 10.0 (VS 2010 Tools for Office) | Add-in lifecycle management | Controls the startup sequence that Outlook measures; `CreateRibbonExtensibilityObject` override is a VSTO runtime feature |
| NLog | 5.4.0 | Startup timing instrumentation | Already installed; high-resolution timestamps in log entries enable sub-ms measurement of startup phases |

### Supporting
| Library | Version | Purpose | When to Use |
|---------|---------|---------|-------------|
| System.Diagnostics.Stopwatch | .NET Fx built-in | High-resolution timing | Measure exact duration of `ThisAddIn_Startup` and individual initialization steps |
| NGen.exe | .NET Fx 4.8 SDK tool | Pre-JIT native image generation | MSI Custom Action for cold-start mitigation on enterprise machines |

### Alternatives Considered
| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| `generatePublisherEvidence` in app.config | Registry-based CRL disable (`HKCU\...\WinTrust\...\State`) | Registry change is machine-wide and affects all .NET apps; app.config is scoped to this add-in only. Use app.config. |
| NGen via MSI Custom Action | No NGen (rely on JIT warmup) | JIT warmup takes 2-5 seconds on cold start, which alone can exceed 1,000 ms threshold. NGen eliminates this. But NGen adds MSI complexity and must be re-run on .NET Framework updates. |
| Lazy<T> for deferred init | Manual first-use initialization | `Lazy<T>` is thread-safe by default and is already used in the codebase (`AppLogger._instance`). Consistent pattern. |
| Timer-based deferred init | `Application.Startup` event | Timer adds non-determinism (what if timer fires before Outlook is ready?). `Application.Startup` fires after all add-ins load -- it is the correct event for deferred init. |

**Installation:** No new NuGet packages needed. All changes use existing .NET Framework BCL and VSTO runtime features.

## Architecture Patterns

### Recommended Changes (Phase 5)
```
PhishingReporter/
  ThisAddIn.cs              # MODIFY: Add Stopwatch instrumentation, Application.Startup handler
  GoPhishIntegration.cs     # VERIFY: Confirm static ctor not triggered during startup path
  AppLogger.cs              # VERIFY: Confirm Lazy<LogFactory> not triggered during startup
  Ribbon.cs                 # VERIFY: No static field initializers that trigger heavy work
  app.config                # MODIFY: Add <runtime> section with generatePublisherEvidence
  NLog.config               # UNCHANGED
  PhishingReporter.csproj   # UNCHANGED (unless NGen MSI work is included)
```

### Pattern 1: Minimal ThisAddIn_Startup with Stopwatch Instrumentation
**What:** Add `System.Diagnostics.Stopwatch` to measure the exact wall-clock time of `ThisAddIn_Startup`, and log it with NLog. Keep the method body to absolute minimum: start stopwatch, log "begin", log "end", stop stopwatch, log elapsed.
**When to use:** During the VSTO startup measurement window. The Stopwatch timestamps allow correlating with Event ID 45 Boot Time.

**Example:**
```csharp
// Source: .NET Framework System.Diagnostics.Stopwatch
// https://learn.microsoft.com/en-us/dotnet/api/system.diagnostics.stopwatch
private void ThisAddIn_Startup(object sender, System.EventArgs e)
{
    var sw = System.Diagnostics.Stopwatch.StartNew();
    Logger.Info("PhishingReporter add-in startup begin");
    // NO initialization here -- everything deferred to Application.Startup
    Logger.Info("PhishingReporter add-in startup complete ({0:F1} ms)", sw.Elapsed.TotalMilliseconds);
}
```

**Confidence:** HIGH -- Stopwatch is the standard .NET high-resolution timer; the current `ThisAddIn_Startup` is already near-empty.

### Pattern 2: Deferred Initialization via Application.Startup Event
**What:** The `Outlook.Application.Startup` event fires after all add-ins have finished loading and Outlook's UI is ready. This is OUTSIDE Outlook's resiliency measurement window. Any initialization that takes non-trivial time should be triggered here.
**When to use:** When initialization work (database connections, network configuration, file I/O) must happen at application start but should not count against the boot time.

**Example:**
```csharp
// Source: Microsoft Docs - Application.Startup Event
// https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.outlook.applicationevents_11_event.startup
private void ThisAddIn_Startup(object sender, System.EventArgs e)
{
    var sw = System.Diagnostics.Stopwatch.StartNew();
    Logger.Info("PhishingReporter add-in startup begin");

    // Subscribe to Application.Startup for deferred init
    this.Application.Startup += Application_Startup;

    Logger.Info("PhishingReporter add-in startup complete ({0:F1} ms)", sw.Elapsed.TotalMilliseconds);
}

private void Application_Startup()
{
    Logger.Info("PhishingReporter deferred initialization begin");
    // Any heavy initialization goes here -- OUTSIDE measurement window
    // Example: Force-load GoPhishIntegration to warm up its static ctor
    // Example: Pre-validate configuration settings
    Logger.Info("PhishingReporter deferred initialization complete");
}
```

**Important nuance:** In the current codebase, `ThisAddIn_Startup` already does essentially nothing (two log calls). The real question is whether any code path during the VSTO startup sequence (before `ThisAddIn_Startup` returns) inadvertently triggers type-loading that pulls in heavy static constructors. This must be verified via profiling or Stopwatch instrumentation at each stage.

**Confidence:** HIGH -- `Application.Startup` being outside the measurement window is confirmed by Microsoft documentation and multiple community sources.

### Pattern 3: generatePublisherEvidence Bypass in app.config
**What:** The .NET CLR performs Authenticode signature verification (including CRL download) for signed assemblies during loading. This can add 15+ seconds if the machine has slow/no internet. The `<generatePublisherEvidence enabled="false"/>` element in app.config disables this check for the current application domain.
**When to use:** Always, for VSTO add-ins deployed via MSI (Windows Installer). MSI deployment already bypasses VSTO manifest validation, so the publisher evidence check is redundant.

**Example:**
```xml
<!-- In app.config, add <runtime> section -->
<configuration>
  <runtime>
    <generatePublisherEvidence enabled="false"/>
  </runtime>
  <!-- existing <configSections> and <userSettings> -->
</configuration>
```

**Confidence:** HIGH -- documented by Microsoft in the VSTO performance troubleshooting guide. The `generatePublisherEvidence` element is a standard .NET Framework configuration option.

### Pattern 4: NGen Pre-JIT via MSI Custom Action
**What:** NGen (Native Image Generator) pre-compiles the add-in assembly and its dependencies into native images, eliminating JIT compilation on first load. This reduces cold-start time by 1-5 seconds.
**When to use:** When cold-start JIT is a significant contributor to boot time. Particularly effective for enterprise deployments where cold starts happen after Windows Updates or reboots.

**Example MSI Custom Action (WiX):**
```xml
<!-- Install: pre-compile to native image -->
<CustomAction Id="NGenInstall"
  Directory="INSTALLDIR"
  ExeCommand="[WindowsFolder]Microsoft.NET\Framework64\v4.0.30319\ngen.exe install PhishingReporter.dll"
  Execute="deferred"
  Impersonate="no"
  Return="ignore" />

<!-- Uninstall: remove native image -->
<CustomAction Id="NGenRemove"
  Directory="INSTALLDIR"
  ExeCommand="[WindowsFolder]Microsoft.NET\Framework64\v4.0.30319\ngen.exe uninstall PhishingReporter.dll"
  Execute="deferred"
  Impersonate="no"
  Return="ignore" />
```

**Note:** NGen must target the correct bitness (Framework vs Framework64) matching the Office installation. This is better addressed in Phase 6 (Enterprise Deployment) since it requires MSI changes and bitness detection. Phase 5 should focus on code-level changes; NGen is an optional optimization if code changes alone do not achieve the < 1,000 ms target.

**Confidence:** MEDIUM -- NGen is well-documented but adds deployment complexity. The primary code-level fixes (generatePublisherEvidence, deferred init) should be attempted first.

### Anti-Patterns to Avoid
- **Putting ANY initialization in ThisAddIn_Startup:** Even seemingly fast operations (reading a config file, creating an HttpClient) can cascade into type-loading that triggers static constructors across multiple assemblies. Keep `ThisAddIn_Startup` to logging only.
- **Using LoadBehavior=0x10 (demand load) to avoid the threshold:** This changes the add-in to load only when the user clicks a ribbon button. The ribbon button would not appear until the add-in loads, creating a confusing UX. Not appropriate for this add-in which needs its ribbon button visible at all times.
- **Disabling CRL check via machine-wide registry:** The `HKCU\...\WinTrust\...\State` registry change affects ALL .NET applications on the machine. Use app.config `generatePublisherEvidence` instead -- it is scoped to this add-in only.
- **Moving heavy code into GetCustomUI/CreateRibbonExtensibilityObject:** These methods run INSIDE the measurement window. They must return immediately. The current implementation correctly returns `new Ribbon()` (cheap object allocation).
- **Triggering GoPhishIntegration static constructor during startup:** If any code in the startup path references `GoPhishIntegration` (even to read a constant or call a method), the CLR loads the type and runs its static constructor, which creates HttpClient, configures ServicePoint, and builds a Polly ResiliencePipeline -- all potentially slow operations.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| High-resolution timing | DateTime.Now subtraction | System.Diagnostics.Stopwatch | DateTime has ~15ms resolution on Windows; Stopwatch uses QueryPerformanceCounter with sub-microsecond resolution |
| CRL bypass | Custom assembly loading | `<generatePublisherEvidence enabled="false"/>` in app.config | Microsoft-documented solution; scoped to the application; no code required |
| Deferred initialization | Custom Timer / BackgroundWorker | `Application.Startup` event | Fires at the correct time (after measurement window closes, before user interaction); no timing guesswork |
| Pre-JIT compilation | Custom ahead-of-time compilation | NGen.exe via MSI Custom Action | NGen is the standard .NET Framework AOT tool; integrates with the GAC and native image cache |

**Key insight:** Phase 5 is primarily about what NOT to do during startup, rather than building new functionality. The optimization is subtractive: remove or defer anything that adds time to the measurement window.

## Common Pitfalls

### Pitfall 1: CLR Cold-Start JIT Tax
**What goes wrong:** The first Outlook launch after a reboot loads the .NET CLR into the Outlook process. JIT compilation of all referenced assemblies (NLog 5.4, Polly 8.4, HtmlAgilityPack 1.12, System.Net.Http, etc.) occurs on the first code path that touches them. This can add 2-5 seconds to the very first startup.
**Why it happens:** JIT is a fundamental aspect of .NET Framework. The CLR compiles IL to native code on first execution of each method.
**How to avoid:** (1) `generatePublisherEvidence=false` eliminates CRL-related cold start. (2) NGen pre-compiles to native images, eliminating JIT entirely. (3) Ensure heavy assemblies (Polly, HtmlAgilityPack) are not loaded during the startup path -- defer type references until first user interaction.
**Warning signs:** Event ID 45 Boot Time is 2,000-5,000 ms on first launch after reboot but drops to 200-500 ms on subsequent launches.

### Pitfall 2: Static Constructor Chains in the Startup Path
**What goes wrong:** A static field initializer or static constructor runs when the CLR first loads a type. If `ThisAddIn_Startup` or `CreateRibbonExtensibilityObject()` references any type with a heavy static constructor, that constructor runs inside the measurement window.
**Why it happens:** Static constructors are invisible in control flow -- they are triggered by the CLR's type loader, not by explicit method calls. Simply declaring a variable of a type (even without assigning it) can trigger loading.
**How to avoid:** Audit every type referenced in `ThisAddIn.cs`, `Ribbon.cs` constructor, and `Ribbon.GetCustomUI()`. Ensure none of them trigger loading of `GoPhishIntegration`, `UrlExtractor`, `AttachmentHasher`, `ReportOrchestrator`, `EmailReport`, or any other type with a non-trivial static initializer. The current `ThisAddIn_Startup` only references `Logger` (which triggers `AppLogger.Instance` -> `Lazy<LogFactory>` -> NLog XML config load). This NLog load may be significant on cold start.
**Warning signs:** Adding a seemingly innocent log call in `ThisAddIn_Startup` increases boot time by 200+ ms.

### Pitfall 3: Authenticode CRL Check Timeout
**What goes wrong:** When a signed .NET assembly loads, the CLR checks the Certificate Revocation List (CRL) online. If the CRL server is unreachable (common on enterprise machines behind proxies or air-gapped networks), the check times out after 15 seconds.
**Why it happens:** The .NET CLR's default Authenticode verification includes CRL checking. Enterprise firewalls or proxy configurations may block the CRL URLs.
**How to avoid:** Add `<generatePublisherEvidence enabled="false"/>` to app.config. This is a per-application setting that disables the CRL check.
**Warning signs:** Boot Time is 15,000+ ms on machines without internet access but < 1,000 ms on developer workstations with fast internet.

### Pitfall 4: NLog Configuration File I/O During Measurement Window
**What goes wrong:** `AppLogger.BuildLogFactory()` reads `NLog.config` from disk. On the first access to `AppLogger.Instance`, this triggers file I/O. If the first access happens during `ThisAddIn_Startup` (via the `Logger` field initializer on `ThisAddIn`), this file read occurs inside the measurement window.
**Why it happens:** `AppLogger.Instance` is a `Lazy<LogFactory>`, so it only runs `BuildLogFactory` on first access. But `ThisAddIn.Logger` is `private static readonly NLog.Logger Logger = AppLogger.Instance.GetCurrentClassLogger()` -- this static field initializer runs when the `ThisAddIn` type is first loaded by the CLR, which is at the very start of the VSTO initialization sequence.
**How to avoid:** Either (a) accept the NLog init cost as small (typically < 50 ms, acceptable budget), or (b) defer the Logger initialization using `Lazy<NLog.Logger>` pattern, or (c) move logging initialization to `Application.Startup`. Option (a) is recommended -- NLog config file parsing is fast and the logging is needed throughout startup for diagnostics.
**Warning signs:** Removing the Logger field from `ThisAddIn` reduces boot time by 30-50 ms.

### Pitfall 5: Measuring on Developer Workstations Instead of Enterprise Hardware
**What goes wrong:** Developer machines have fast SSDs, warm CLR caches, and direct internet. Boot Time shows 200-500 ms. Enterprise machines with spinning disks, cold boots, and proxy-blocked CRL show 2,000-15,000 ms.
**Why it happens:** The resiliency threshold is designed to protect users on representative hardware, not developer machines.
**How to avoid:** Always test on representative enterprise hardware (or a VM simulating enterprise conditions: HDD, cold boot, no direct internet) for final validation. The success criteria explicitly require "a representative enterprise machine (not a developer workstation)."
**Warning signs:** Tests pass on developer machines but fail in production deployment.

## Code Examples

### Example 1: Updated app.config with generatePublisherEvidence
```xml
<?xml version="1.0" encoding="utf-8" ?>
<configuration>
  <runtime>
    <generatePublisherEvidence enabled="false"/>
  </runtime>
  <configSections>
    <sectionGroup name="userSettings" ... >
      <section name="PhishingReporter.Properties.Settings" ... />
    </sectionGroup>
  </configSections>
  <userSettings>
    <!-- existing settings unchanged -->
  </userSettings>
</configuration>
```

**Note:** The `<runtime>` section must be a direct child of `<configuration>`, before or after `<configSections>`. The exact ordering within `<configuration>` does not matter as long as `<configSections>` remains the first element if present (XML schema requirement). Actually, `<configSections>` must be the FIRST child element of `<configuration>`. So `<runtime>` should come AFTER `<configSections>`.

**Corrected order:**
```xml
<configuration>
  <configSections>
    <!-- existing -->
  </configSections>
  <runtime>
    <generatePublisherEvidence enabled="false"/>
  </runtime>
  <userSettings>
    <!-- existing -->
  </userSettings>
</configuration>
```

### Example 2: ThisAddIn.cs with Stopwatch and Deferred Init
```csharp
// Source: Codebase analysis + Microsoft Docs
public partial class ThisAddIn
{
    private static readonly NLog.Logger Logger = AppLogger.Instance.GetCurrentClassLogger();

    private void ThisAddIn_Startup(object sender, System.EventArgs e)
    {
        var sw = System.Diagnostics.Stopwatch.StartNew();
        Logger.Info("PhishingReporter add-in startup begin");

        // Deferred init: Application.Startup fires AFTER Outlook finishes
        // loading all add-ins -- outside the resiliency measurement window
        this.Application.Startup += Application_Startup;

        sw.Stop();
        Logger.Info("PhishingReporter add-in startup complete ({0:F1} ms)", sw.Elapsed.TotalMilliseconds);
    }

    private void Application_Startup()
    {
        Logger.Info("PhishingReporter deferred initialization begin");
        // Force-load any types with heavy static constructors here
        // This ensures JIT and static init happen OUTSIDE the measurement window
        // Currently nothing needs explicit warm-up, but this is the right place
        // if future initialization is added
        Logger.Info("PhishingReporter deferred initialization complete");
    }

    private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
    {
        Logger.Info("PhishingReporter add-in shutdown");
        AppLogger.Instance.Shutdown();
    }

    protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
    {
        // STRT-02: Direct return bypasses VSTO Ribbon Designer reflection scan
        return new Ribbon();
    }

    #region VSTO generated code
    private void InternalStartup()
    {
        this.Startup += new System.EventHandler(ThisAddIn_Startup);
        this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
    }
    #endregion
}
```

### Example 3: Verifying Event ID 45 Boot Time
```powershell
# PowerShell command to check add-in boot time from Event Viewer
Get-WinEvent -LogName Application |
  Where-Object { $_.Id -eq 45 -and $_.ProviderName -eq 'Outlook' } |
  Select-Object -First 5 |
  ForEach-Object { $_.Message }
```

The output shows each loaded add-in with its `Boot Time (Milliseconds)` value. The target is a median under 1,000 ms across 5 consecutive cold starts.

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| No CRL bypass | `generatePublisherEvidence enabled="false"` in app.config | Available since .NET 3.5 SP1 (2008) | Eliminates 15-second timeout on air-gapped networks |
| VSTO Ribbon Designer with reflection scan | Override `CreateRibbonExtensibilityObject()` or use Ribbon XML directly | VSTO 2010 runtime (2010) | Eliminates assembly scanning; saves 100-500 ms on large add-ins |
| Heavy initialization in `ThisAddIn_Startup` | Defer to `Application.Startup` event | Best practice since Outlook 2013 resiliency (2012) | Moves work outside the 1,000 ms measurement window |
| JIT compilation on every cold start | NGen pre-compilation via MSI | Available since .NET Framework 1.0 (2002) | Eliminates 2-5 second JIT penalty on cold starts |
| Managed VSTO add-in (slow CLR bootstrap) | Native COM add-in | N/A -- not applicable | Native COM avoids CLR startup entirely but requires C++ rewrite; not viable for this project |

**Deprecated/outdated:**
- VSTO "Fast Path" loading (VSTO SP1): Superseded by .NET 4.0+ improvements. Already benefiting from this since the project targets .NET 4.8.
- `<loadFromRemoteSources>` config element: Not relevant for locally-installed MSI-deployed add-ins.
- Demand loading (LoadBehavior=0x10): Not appropriate for this add-in because the ribbon button must be visible at all times.

## Open Questions

1. **Is NLog initialization a significant contributor to boot time?**
   - What we know: `ThisAddIn.Logger` is a static readonly field initialized via `AppLogger.Instance.GetCurrentClassLogger()`. The `Lazy<LogFactory>` triggers `BuildLogFactory()` which reads `NLog.config` from disk. This runs when the CLR first loads the `ThisAddIn` type -- very early in the VSTO startup sequence.
   - What's unclear: Whether this file I/O adds 10 ms or 200 ms on enterprise hardware. On an SSD it is likely negligible; on a spinning disk with antivirus scanning, it could be significant.
   - Recommendation: Measure with Stopwatch instrumentation before and after. Accept the cost if < 100 ms (logging is needed throughout startup for diagnostics). If > 100 ms, consider deferring Logger initialization. Confidence: MEDIUM -- need empirical measurement.

2. **Does GoPhishIntegration static constructor trigger during startup?**
   - What we know: `GoPhishIntegration` has a static constructor that creates `HttpClient`, configures `ServicePointManager.FindServicePoint`, and builds a Polly `ResiliencePipeline`. None of the startup code paths (`ThisAddIn_Startup`, `CreateRibbonExtensibilityObject`, `Ribbon()` constructor, `Ribbon.GetCustomUI`) directly reference `GoPhishIntegration`.
   - What's unclear: Whether any indirect reference (e.g., through `using` statements, type reflection, or assembly metadata scanning) triggers CLR type-loading of `GoPhishIntegration` during startup.
   - Recommendation: Verify empirically by adding a log statement to the `GoPhishIntegration` static constructor and checking if it appears in logs before "startup complete". If it does, restructure to use `Lazy<T>` for the heavy resources. Confidence: HIGH that it does NOT trigger (no direct references in startup path), but must verify.

3. **Is NGen necessary or is generatePublisherEvidence sufficient?**
   - What we know: `generatePublisherEvidence` eliminates the CRL timeout (up to 15 seconds). NGen eliminates JIT compilation (2-5 seconds on cold start). These are independent optimizations.
   - What's unclear: Whether `generatePublisherEvidence` alone brings boot time under 1,000 ms on enterprise hardware. If CRL was the primary cause of the 50% failure rate, fixing it alone may suffice.
   - Recommendation: Implement `generatePublisherEvidence` first. Measure on representative hardware. If still over 1,000 ms, add NGen in Phase 6 (Enterprise Deployment) since it requires MSI changes. Confidence: MEDIUM -- depends on empirical measurement.

4. **What is the actual current boot time?**
   - What we know: The STATE.md records a blocker: "Actual current startup time baseline is unknown -- Event ID 45 will reveal this after Phase 1 deployment."
   - What's unclear: Whether the add-in's boot time is 1,200 ms (close to threshold) or 15,000 ms (CRL timeout dominating). The optimization strategy differs significantly.
   - Recommendation: The first task in Phase 5 should be establishing a baseline measurement on representative hardware. All subsequent optimization decisions depend on this number. If the add-in was being disabled for crash reasons (0x3) rather than slow boot (0x1), the optimization strategy is different entirely.

## Sources

### Primary (HIGH confidence)
- [Microsoft Docs -- Improve the performance of a VSTO Add-in](https://learn.microsoft.com/en-us/visualstudio/vsto/improving-the-performance-of-a-vsto-add-in?view=vs-2022) -- CreateRibbonExtensibilityObject override, bypass ribbon reflection, load on demand, Windows Installer deployment benefits
- [Microsoft Docs -- Application event log entries for Outlook add-in load time](https://learn.microsoft.com/en-us/microsoft-365-apps/outlook/performance/log-entries-for-add-ins) -- Event ID 45 format, Boot Time measurement, Application event log location, DisableAddinLogging policy
- [Microsoft Docs -- Support for keeping add-ins enabled](https://learn.microsoft.com/en-us/office/vba/outlook/concepts/getting-started/support-for-keeping-add-ins-enabled) -- DoNotDisableAddinList registry key, AddinList managed add-in policy, disable reason codes (0x1-0xA), 1,000 ms threshold, 5-iteration median
- [Microsoft Docs -- Program VSTO add-ins with the ThisAddIn class](https://learn.microsoft.com/en-us/visualstudio/vsto/programming-vsto-add-ins?view=vs-2022) -- ThisAddIn lifecycle, RequestService timing, CreateRibbonExtensibilityObject in startup sequence, Application.Startup event
- [Microsoft Docs -- Troubleshooting Outlook COM Addins](https://learn.microsoft.com/en-us/archive/blogs/dvespa/troubleshooting-outlook-com-addins-introduction) -- IDTExtensibility2 lifecycle, OnConnection as entry point, VSTO runtime chain, COM activation sequence

### Secondary (MEDIUM confidence)
- [Microsoft Archive -- Resolving performance issues with loading Office add-ins](https://learn.microsoft.com/en-us/archive/blogs/vsod/resolving-performance-issues-with-loading-office-add-ins-vsto-add-ins-or-shared-add-ins) -- 7 causes of slow loading (CLR cold start, CRL check, .NET 3.5 crypto issue, ribbon reflection, Automatic Root Certificate Update), generatePublisherEvidence fix, NGen recommendation
- [EricPhan.net -- VSTO Tip: Stop the built-in Ribbon reflection from slowing your add-in's load time](http://ericphan.net/blog/2010/11/22/vsto-tip-stop-the-built-in-ribbon-reflection-from-slowing-yo.html) -- CreateRibbonExtensibilityObject and CreateRibbonObjects overrides, Ribbon XML vs Designer performance
- [Microsoft Q&A -- VSTO Outlook: Improve and accelerate Add-in startup process](https://learn.microsoft.com/en-us/answers/questions/1056423/vsto-outlook-improve-and-accelerate-add-in-startup) -- Application.Startup event for deferred init, timer-based delay pattern, WPF control initialization as real-world bottleneck example
- Codebase inspection -- ThisAddIn.cs (CreateRibbonExtensibilityObject already overridden), GoPhishIntegration.cs (static constructor with HttpClient + Polly), AppLogger.cs (Lazy LogFactory), app.config (no runtime section)

### Tertiary (LOW confidence)
- [Microsoft Q&A -- Default VSTO project causing Outlook slow startup](https://learn.microsoft.com/en-us/answers/questions/1377543/the-default-project-for-my-visual-studio-outlook-v) -- Even blank VSTO add-ins show 1.62s average boot time; NGen and GAC installation did not resolve; suggests CLR bootstrap overhead is the floor for managed add-ins. Unresolved question -- needs validation.

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH -- No new dependencies; all optimizations use existing .NET Framework configuration and tools
- Architecture: HIGH -- Deferred initialization via Application.Startup and generatePublisherEvidence are Microsoft-documented best practices with extensive community validation
- Pitfalls: HIGH -- All pitfalls derived from Microsoft documentation (CRL timeout, CLR cold start, static constructor chains) and direct codebase analysis of startup code paths
- Open questions: MEDIUM -- Actual boot time baseline and NLog initialization cost require empirical measurement on representative hardware

**Research date:** 2026-02-26
**Valid until:** 2026-03-28 (stable domain, 30-day validity)
