# Stack Research

**Domain:** VSTO Outlook Add-in Reliability — C# / .NET Framework
**Researched:** 2026-02-25
**Confidence:** HIGH (core recommendations verified via Microsoft official docs and NuGet package pages)

---

## Context

This is a subsequent-milestone research file. The existing add-in runs on C# / .NET Framework 4.6.1 / VSTO 4.0. The scope is reliability fixes only: eliminate Outlook's auto-disable behavior and fix UI freezes from synchronous HTTP calls. No migration to Office.js, no new features.

Two root causes drive the entire stack decision:

1. **Outlook resiliency disabling** — add-in median startup exceeds the hard-coded 1,000 ms threshold across 50% of enterprise machines due to .NET Framework JIT overhead and synchronous work in `ThisAddIn_Startup`.
2. **UI thread freeze** — `GoPhishIntegration.sendReportNotificationToServer()` uses `HttpWebRequest.GetResponse()` synchronously on the STA (UI) thread, blocking Outlook for the entire HTTP round-trip.

---

## Recommended Stack

### Core Technologies

| Technology | Version | Purpose | Why Recommended |
|------------|---------|---------|-----------------|
| .NET Framework | **4.8** | Runtime for all VSTO add-in code | 4.8 is the final and highest supported version for VSTO. Microsoft has stated the VSTO runtime will not move beyond 4.8. It is an in-place upgrade over 4.6.1 — no recompilation of Outlook PIA or VSTO assemblies required. Ships with Windows 10/11 and is pre-installed on most enterprise machines, eliminating the cold-start JIT overhead introduced by loading a fresh framework. (HIGH confidence — verified: [VSTO Runtime Lifecycle](https://learn.microsoft.com/en-us/visualstudio/vsto/visual-studio-tools-for-office-runtime?view=visualstudio)) |
| VSTO 4.0 | **4.0** (no change) | Outlook add-in host framework | No alternative for classic desktop Outlook. Microsoft has confirmed VSTO + .NET Framework is the only supported path for COM add-ins on Outlook 2016/2019/365 classic. (HIGH confidence — verified: Microsoft Q&A) |
| C# | **Latest supported by VS 2022** (C# 10 on .NET FW 4.8) | Implementation language | No change. Needed for `async`/`await` syntax that fixes the UI freeze. C# 10 features are available when targeting .NET Framework 4.8 with a modern SDK-style project or VS 2022. |

### HTTP Client Library

| Library | Version | Purpose | Why Recommended |
|---------|---------|---------|-----------------|
| `System.Net.Http.HttpClient` | Built into .NET Framework 4.8 | Async HTTP calls to GoPhish | `HttpClient` is the only .NET-framework-built-in library with native `async`/`await` support (`PostAsync`, `GetAsync`). The existing `HttpWebRequest.GetResponse()` pattern has no async counterpart in .NET Framework without significant wrapping. `HttpClient` must be instantiated as a **static singleton** (one per add-in lifetime) to avoid socket exhaustion — `SocketsHttpHandler.PooledConnectionLifetime` is not available in .NET Framework 4.8, so a singleton is the correct pattern. (HIGH confidence — verified: [HttpClient Guidelines](https://learn.microsoft.com/en-us/dotnet/fundamentals/networking/http/httpclient-guidelines)) |

**Do NOT use** `IHttpClientFactory` from `Microsoft.Extensions.Http` in this add-in. While it technically runs on .NET Framework 4.8 via `netstandard2.0`, Microsoft explicitly states these packages are **not officially supported on .NET Framework** and exist only to help internal teams migrate. Wiring up a `ServiceCollection` in a VSTO add-in adds unnecessary DLL complexity with documented runtime conflicts. A static `HttpClient` singleton achieves the same port-exhaustion safety without those risks.

**Do NOT use** `HttpWebRequest` (the existing code). It has no true async counterpart (`BeginGetResponse`/`EndGetResponse` is APM, not `async`/`await`), leading to deadlock risks on the STA thread. Replace entirely.

### Logging Framework

| Library | Version | Purpose | Why Recommended |
|---------|---------|---------|-----------------|
| **NLog** | **6.1.0** | Persistent file-based logging for diagnosing load failures and GoPhish errors | NLog 6.1.0 (released January 31, 2026) explicitly supports .NET Framework 3.5–4.8 with no caveats. Its `FileTarget` is built-in — no extra sink package required. Its **XML-based configuration** (`NLog.config`) lets support staff change log levels post-deployment without a rebuild or MSI re-deployment, which is critical for diagnosing enterprise add-in load failures. Configuration changes take effect at runtime. (HIGH confidence — verified: [NuGet NLog 6.1.0](https://www.nuget.org/packages/NLog/)) |

**Do NOT use** Serilog for this project. Serilog requires code-first configuration, which cannot be adjusted after deployment without a rebuild. In an enterprise environment where the add-in is pushed via SCCM, the ability to drop an updated `NLog.config` alongside the DLL (or into the user profile) to toggle `DEBUG` logging for a specific machine is a significant operational advantage NLog provides over Serilog.

**Do NOT use** `Microsoft.Extensions.Logging` as a primary logger. Version 10.x targets .NET 8 and .NET Standard 2.0 only; its .NET Framework path is unsupported and has documented DLL version-conflict issues when combined with NLog's extension package on .NET Framework 4.8. Use NLog directly.

### Supporting Libraries

| Library | Version | Purpose | When to Use |
|---------|---------|---------|-------------|
| **HtmlAgilityPack** | **1.12.4** | HTML parsing for URL extraction from email bodies | Already in use; update from 1.11.23. Version 1.12.4 supports .NET Framework 3.5+ with no breaking changes for existing XPath queries. The existing `//a[@href]` extraction is compatible. (HIGH confidence — verified: [NuGet HtmlAgilityPack 1.12.4](https://www.nuget.org/packages/htmlagilitypack/)) |
| **Polly** | **8.6.5** | Retry + timeout policies for GoPhish HTTP calls | Use when the GoPhish call is converted to async `HttpClient`. Adds timeout (prevents infinite hang if GoPhish is unreachable), retry with exponential backoff (handles transient network errors). Polly 8.6.5 supports .NET Framework 4.6.2 and higher. (HIGH confidence — verified: [NuGet Polly 8.6.5](https://www.nuget.org/packages/Polly)) |

### Development Tools

| Tool | Purpose | Notes |
|------|---------|-------|
| **Visual Studio 2022** | Primary IDE; required for VSTO project support | VS 2022 17.x provides VSTO project templates and the "Microsoft Visual Studio Installer Projects" extension for `.vdproj` MSI builds. Required for .NET Framework 4.8 target. |
| **Ngen.exe** (post-install action) | Pre-JIT native image generation to reduce cold-start load time | Run `ngen.exe install <assembly>` as a post-install action in the MSI. Eliminates per-machine JIT compilation from the startup measurement window. This directly addresses the 1-second resiliency threshold. Command: `%windir%\Microsoft.NET\Framework64\v4.0.30319\ngen.exe install PhishingReporter.dll`. (MEDIUM confidence — widely documented but per-machine benefit varies) |
| **Registry deployment via MSI custom action** | Writes `DoNotDisableAddinList` and `AddinList` policy keys | The MSI installer must write these registry keys to prevent Outlook from auto-disabling the add-in. This is a mandatory complement to the startup optimization, not a replacement. (HIGH confidence — verified: [Support for keeping add-ins enabled](https://learn.microsoft.com/en-us/office/vba/outlook/concepts/getting-started/support-for-keeping-add-ins-enabled)) |

---

## Installation

This project uses `packages.config` with NuGet. Install via Visual Studio NuGet Package Manager or Package Manager Console:

```powershell
# Update existing dependency
Update-Package HtmlAgilityPack -Version 1.12.4

# Add new dependencies
Install-Package NLog -Version 6.1.0
Install-Package Polly -Version 8.6.5
```

For the `packages.config` format (existing project style):

```xml
<packages>
  <package id="HtmlAgilityPack" version="1.12.4" targetFramework="net48" />
  <package id="NLog" version="6.1.0" targetFramework="net48" />
  <package id="Polly" version="8.6.5" targetFramework="net48" />
</packages>
```

The `TargetFrameworkVersion` in `PhishingReporter.csproj` must be changed from `v4.6.1` to `v4.8`:

```xml
<TargetFrameworkVersion>v4.8</TargetFrameworkVersion>
```

---

## Alternatives Considered

| Recommended | Alternative | Why Not |
|-------------|-------------|---------|
| .NET Framework 4.8 | Stay on 4.6.1 | 4.6.1 reached end-of-life in 2019 with no security updates. Upgrading to 4.8 is the minimum safe move that keeps VSTO compatibility. The upgrade is in-place and backward-compatible with all existing PIA references. |
| .NET Framework 4.8 | Migrate to .NET 6/8 | VSTO cannot run on modern .NET (Core/5+/6+/8+/10). Outlook loads add-ins in-process and requires the .NET Framework runtime. Attempting this causes assembly load failures. This migration path does not exist until Office.js replaces VSTO. |
| Static singleton `HttpClient` | `HttpWebRequest` | `HttpWebRequest` has no `async`/`await` API. Its APM-based `BeginGetResponse` cannot be safely `await`ed on an STA thread without deadlocking. The add-in already has a resource disposal bug with it (undisposed `HttpWebResponse`). |
| Static singleton `HttpClient` | `IHttpClientFactory` | `IHttpClientFactory` requires a `ServiceCollection` DI bootstrap that is awkward in a VSTO entry point and introduces unsupported `Microsoft.Extensions.*` assembly references with documented DLL conflicts on .NET Framework 4.8. |
| NLog | Serilog | Serilog requires code-first configuration. Changing log verbosity after SCCM deployment requires a rebuild. NLog's XML config file can be updated independently of the DLL. |
| NLog | log4net | log4net has not had a stable release since 2022 and is in maintenance mode. NLog 6.1.0 is actively maintained with .NET 10 support. |
| Polly 8.6.5 | Manual retry loop | A hand-rolled retry loop requires writing timeout logic, cancellation token handling, exponential backoff, and exception filtering — all solved problems in Polly. The GoPhish call has currently no timeout at all; Polly fixes this in three lines. |

---

## What NOT to Use

| Avoid | Why | Use Instead |
|-------|-----|-------------|
| `HttpWebRequest.GetResponse()` | Synchronous; blocks the Outlook STA (UI) thread; no async API; active resource disposal bug (undisposed response stream) confirmed in CONCERNS.md | `HttpClient.PostAsync()` with `await` |
| `Microsoft.Extensions.Http` (`IHttpClientFactory`) | Not officially supported on .NET Framework; documented runtime DLL conflicts when mixed with NLog.Extensions.Logging; requires DI container setup that is foreign to VSTO | Static singleton `HttpClient` |
| `Microsoft.Extensions.Logging` | Not officially supported on .NET Framework 4.8; exists to support Microsoft-internal migration, not external production use; version conflicts between the 10.x package and .NET Framework assembly binder | NLog 6.1.0 directly |
| NLog.Extensions.Logging | Known DLL conflict with `Microsoft.Extensions.Logging 2.1.x` on .NET Framework 4.8 (documented in NLog GitHub issue #740) | Use NLog directly without the Extensions.Logging bridge |
| `SocketsHttpHandler` with `PooledConnectionLifetime` | Not available in .NET Framework 4.8 — this is a .NET Core/5+ API. Using it compiles but silently falls back to old handler behavior | Static singleton `HttpClient` (which naturally pools connections for its lifetime) |
| Modern .NET (5, 6, 7, 8, 9, 10) as runtime target | VSTO/COM add-ins are .NET Framework-only. Attempting to target modern .NET causes Outlook to fail loading the add-in entirely. Microsoft confirmed VSTO will not move beyond .NET Framework 4.8. | .NET Framework 4.8 |
| Serilog | Code-first configuration cannot be adjusted post-deployment without a rebuild; harder to diagnose enterprise add-in issues where support staff need to enable debug logging on one machine | NLog 6.1.0 with XML config |

---

## Stack Patterns by Variant

**For the GoPhish HTTP call (current: synchronous on UI thread):**
- Use `HttpClient` static singleton initialized in `ThisAddIn_Startup`
- Use `async Task` method signature for the GoPhish notification method
- Use `await client.PostAsync(...)` inside an `async` ribbon callback
- Use Polly pipeline wrapping the `HttpClient` call: 3-second timeout, 2 retries with exponential backoff
- Marshal any Outlook OM access (reading email properties) on the UI thread before the `await`; the async continuation after `await` can stay on a thread pool thread as long as it does not touch the Outlook object model

**For startup time (current: exceeds 1,000 ms on cold machines):**
- Move all non-essential initialization out of `ThisAddIn_Startup` into lazy-loaded properties or first-use initialization
- Do not open any network connections, validate certificates, or read remote resources during startup
- Add `ngen.exe` as a post-install action in the MSI to eliminate per-machine JIT cost
- Write registry keys during MSI installation to add the add-in to `DoNotDisableAddinList` (user-level) and `AddinList` (policy-level) as a belt-and-suspenders approach
- The 1,000 ms threshold is hard-coded in Outlook and cannot be changed; both code optimization AND registry keys are required

**For logging (current: none):**
- Initialize NLog in `ThisAddIn_Startup` before any other code
- Write log to `%APPDATA%\PhishingReporter\logs\` with 7-day rolling file retention
- Use `NLog.config` placed alongside the DLL in the install directory so support staff can toggle log level without reinstalling
- Log add-in load entry/exit times in `ThisAddIn_Startup` to measure actual startup duration in production

---

## Version Compatibility

| Package | Compatible With | Notes |
|---------|-----------------|-------|
| NLog 6.1.0 | .NET Framework 3.5–4.8 | No breaking changes from NLog 4.x for basic `FileTarget` and `Logger` usage. No `NLog.Extensions.Logging` package needed. |
| Polly 8.6.5 | .NET Framework 4.6.2+ | Polly 8.x API differs from 7.x — uses `ResiliencePipeline` builder API instead of the old fluent `Policy.Handle()` chain. Use the v8 API; do not mix v7 and v8 patterns. |
| HtmlAgilityPack 1.12.4 | .NET Framework 3.5+ | No breaking changes for existing `SelectNodes("//a[@href]")` XPath usage. Safe drop-in upgrade from 1.11.23. |
| .NET Framework 4.8 | Outlook 2016 (16.x), Outlook 2019 (16.x), Microsoft 365 Classic Desktop | All three ship with .NET Framework 4.8 pre-installed on Windows 10/11 enterprise machines. .NET Framework 4.8 is NOT compatible with New Outlook (Monarch) — but the project scope explicitly excludes that. |
| .NET Framework 4.8 | `Microsoft.Office.Interop.Outlook` v15.0.0.0 (PIA) | The Office 2013 PIA (v15.0) is forward-compatible with Outlook 2016, 2019, and 365. Upgrading the framework version does not require upgrading the PIA reference. |

---

## Resiliency Mechanism Reference

The Outlook add-in resiliency mechanism that causes the auto-disable is documented behavior, not a bug. Understanding it is required to fix it correctly:

- Outlook 2013+ measures the **median startup time over 5 successive Outlook launches**.
- If the median exceeds **1,000 ms**, Outlook prompts the user to disable the add-in. If the user confirms (or if the add-in crashes), it is written to the `DisabledItems` registry key.
- The threshold is **hard-coded** and cannot be increased.
- Two registry key families control the behavior:
  - `HKCU\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList` — per-user whitelist (DWORD, value = 0x01 for boot-load exemption)
  - `HKCU\Software\Policies\Microsoft\Office\16.0\Outlook\Resiliency\AddinList` — GPO-managed list (REG_SZ, value = "1" to always enable)
- The MSI installer must write **both** keys (for Outlook 16.0/Office 2016/2019/365) to cover cases where Group Policy is not deployed.
- The `LoadBehavior` DWORD under `HKCU\Software\Microsoft\Office\Outlook\Addins\<ProgID>` must be set to `3` (load at startup). When Outlook disables an add-in it changes this to `2`.

Source: [Microsoft Docs — Support for keeping add-ins enabled](https://learn.microsoft.com/en-us/office/vba/outlook/concepts/getting-started/support-for-keeping-add-ins-enabled)

---

## Sources

- [Microsoft Docs — Support for keeping add-ins enabled](https://learn.microsoft.com/en-us/office/vba/outlook/concepts/getting-started/support-for-keeping-add-ins-enabled) — HIGH confidence. Official registry key reference for `DoNotDisableAddinList` and `AddinList`.
- [Microsoft Docs — HttpClient Guidelines for .NET](https://learn.microsoft.com/en-us/dotnet/fundamentals/networking/http/httpclient-guidelines) — HIGH confidence. Official guidance on singleton `HttpClient` and `IHttpClientFactory` by framework version. Last updated October 2025.
- [Microsoft Docs — VSTO Runtime Lifecycle Policy](https://learn.microsoft.com/en-us/visualstudio/vsto/visual-studio-tools-for-office-runtime?view=visualstudio) — HIGH confidence. Confirms .NET Framework 4.8 is the final VSTO-supported version.
- [NuGet — NLog 6.1.0](https://www.nuget.org/packages/NLog/) — HIGH confidence. Confirmed version, release date (January 31, 2026), and .NET Framework 3.5–4.8 support.
- [NuGet — HtmlAgilityPack 1.12.4](https://www.nuget.org/packages/htmlagilitypack/) — HIGH confidence. Confirmed latest version and .NET Framework compatibility.
- [NuGet — Polly 8.6.5](https://www.nuget.org/packages/Polly) — HIGH confidence. Confirmed latest stable version (November 23, 2025) and .NET Framework 4.6.2+ support.
- [Microsoft Docs — .NET Framework & Windows OS versions](https://learn.microsoft.com/en-us/dotnet/framework/install/versions-and-dependencies) — HIGH confidence. In-place upgrade behavior of 4.8 over 4.6.1 confirmed.
- [Microsoft Q&A — Default VSTO project causes Outlook to start slowly](https://learn.microsoft.com/en-us/answers/questions/1377543/the-default-project-for-my-visual-studio-outlook-v) — MEDIUM confidence. Confirms even a blank VSTO template exceeds 1,000 ms threshold on some machines.
- [Developer Messaging Blog — Outlook slow add-ins resiliency logic](https://developermessaging.azurewebsites.net/2017/08/02/outlooks-slow-add-ins-resiliency-logic-and-how-to-always-enable-slow-add-ins/) — MEDIUM confidence. Explains the 5-iteration median measurement and the 1,000 ms threshold in detail.
- [NLog GitHub Issue #740 — DLL conflicts with Extensions.Logging on .NET Framework 4.8](https://github.com/NLog/NLog.Extensions.Logging/issues/740) — MEDIUM confidence. Documents why `NLog.Extensions.Logging` should be avoided in this context.
- [Microsoft Docs — Threading support in Office (VSTO)](https://learn.microsoft.com/en-us/visualstudio/vsto/threading-support-in-office?view=vs-2022) — HIGH confidence. STA model and async/await marshaling requirements.

---

*Stack research for: VSTO Outlook Add-in Reliability Fixes*
*Researched: 2026-02-25*
