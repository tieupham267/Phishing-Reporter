# Architecture Research

**Domain:** VSTO Outlook Add-in reliability refactoring
**Researched:** 2026-02-25
**Confidence:** HIGH (Microsoft official docs + verified community patterns)

## Standard Architecture

### System Overview

The recommended architecture separates the current monolithic `Ribbon.cs` into focused components
with clear ownership. The Outlook UI thread constraint is the central concern: Outlook's object
model (OOM) must be called on the UI thread, but network I/O must never block it.

```
┌─────────────────────────────────────────────────────────────────────┐
│                     Outlook Process (UI Thread)                      │
│                                                                       │
│  ┌──────────────┐   ┌──────────────┐   ┌────────────────────────┐  │
│  │  ThisAddIn   │   │    Ribbon    │   │   MailItemExtensions   │  │
│  │  (Lifecycle) │   │  (UI Only)   │   │   (Header Parsing)     │  │
│  └──────┬───────┘   └──────┬───────┘   └────────────────────────┘  │
│         │ startup          │ button click                            │
│         ▼                  ▼                                         │
│  ┌──────────────────────────────────────────────────────────────┐   │
│  │               ReportOrchestrator (async coordinator)          │   │
│  │  - Validates selection on UI thread                           │   │
│  │  - Reads Outlook data on UI thread                            │   │
│  │  - Dispatches network work to background thread               │   │
│  └──┬────────────┬──────────────────┬──────────────────────────┘   │
│     │            │                  │                                │
│     ▼            ▼                  ▼                                │
│  ┌────────┐  ┌──────────┐  ┌───────────────┐                        │
│  │ Email  │  │ GoPhish  │  │  Settings     │                        │
│  │Extractor│  │Detector  │  │  Validator    │                        │
│  └────────┘  └──────────┘  └───────────────┘                        │
└───────────────────────────────────┬─────────────────────────────────┘
                                    │ async (Task / HttpClient)
                                    ▼
┌─────────────────────────────────────────────────────────────────────┐
│                    Background Thread Pool                            │
│                                                                       │
│  ┌──────────────────────────────────────────────────────────────┐   │
│  │               GoPhishHttpClient                               │   │
│  │  - async HTTP GET to GoPhish listener                         │   │
│  │  - Timeout enforced                                           │   │
│  │  - HttpClient singleton (not per-call)                        │   │
│  └──────────────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────────────┘
```

### Component Responsibilities

| Component | Responsibility | Communicates With |
|-----------|----------------|-------------------|
| `ThisAddIn` | VSTO lifecycle only: attach Application.Startup event, validate settings, no heavy work | Ribbon, SettingsValidator |
| `Ribbon` | UI callbacks only: button clicks, image loading, show dialogs | ReportOrchestrator |
| `ReportOrchestrator` | Coordinate the report workflow; read Outlook data on UI thread; hand off network to background | EmailExtractor, GoPhishDetector, GoPhishHttpClient, Settings |
| `EmailExtractor` | Extract URLs, domains, attachments, headers, hashes from a MailItem | HtmlAgilityPack, System.IO |
| `GoPhishDetector` | Parse custom header to detect simulated campaign, construct report URL | (no external deps — pure string logic) |
| `GoPhishHttpClient` | Async HTTP call to GoPhish server; timeout; dispose correctly | System.Net.Http.HttpClient |
| `SettingsValidator` | Validate required settings at startup, surface configuration errors early | Properties.Settings |
| `MailItemExtensions` | Extension methods for MAPI header access (already well-isolated, keep as-is) | Outlook Interop |

## Recommended Project Structure

```
PhishingReporter/
├── ThisAddIn.cs                  # VSTO lifecycle only (startup event hookup, settings validation)
├── ThisAddIn.Designer.cs         # VSTO generated (do not touch)
├── Ribbon.cs                     # UI callbacks only — thin, no logic
├── Ribbon.xml                    # Ribbon XML (no changes needed)
│
├── Orchestration/
│   └── ReportOrchestrator.cs    # Async workflow coordinator
│
├── Email/
│   ├── EmailExtractor.cs        # URL, domain, attachment, hash extraction
│   ├── MailItemExtensions.cs    # MAPI header extension methods (move from Ribbon.cs)
│   └── EmailReport.cs          # Immutable data record for extracted report data
│
├── GoPhish/
│   ├── GoPhishDetector.cs       # Header parsing, URL construction (no HTTP)
│   └── GoPhishHttpClient.cs     # async HTTP; owns HttpClient singleton
│
├── Configuration/
│   └── SettingsValidator.cs    # Validate settings on startup, fail fast with clear messages
│
└── Properties/
    ├── Settings.Designer.cs     # Auto-generated
    └── Settings.settings        # Configuration values
```

### Structure Rationale

- **Orchestration/:** Single class that understands sequence; all other classes are ignorant of workflow order
- **Email/:** Groups Outlook-touching extraction logic; all runs on UI thread; testable with mocked MailItem
- **GoPhish/:** Separates pure string logic (Detector) from network I/O (HttpClient); Detector is unit-testable without HTTP
- **Configuration/:** Startup validation concentrated here; surfaced before any user action is possible

## Architectural Patterns

### Pattern 1: Deferred Initialization via Application.Startup Event

**What:** Hook Outlook's `Application.Startup` event (not `ThisAddIn_Startup`) for any non-trivial
work. `ThisAddIn_Startup` runs while Outlook is measuring startup time. `Application.Startup` fires
after all add-ins are loaded, outside the resiliency measurement window.

**When to use:** Any initialization that takes more than a few milliseconds — settings validation,
logging setup, pre-warming HttpClient.

**Why it works:** Outlook's resiliency mechanism measures the elapsed time inside its own
initialization sequence, not the Application.Startup event handler. Work done in Application.Startup
does not count toward the ~1000 ms threshold.

**Trade-offs:** The add-in's extended objects are not ready during the very first moments of Outlook
startup. For a phishing reporter that only activates on user click, this is irrelevant.

**Example:**
```csharp
// ThisAddIn.cs — keep Startup handler near-empty
private void ThisAddIn_Startup(object sender, EventArgs e)
{
    // Hook Application.Startup instead — runs after resiliency measurement window
    this.Application.Startup += Application_AfterOutlookReady;
}

private void Application_AfterOutlookReady()
{
    // Safe to do non-trivial work here — no resiliency penalty
    SettingsValidator.ValidateOrWarn();
    // Pre-create HttpClient singleton so first HTTP call is not slower
    GoPhishHttpClient.Initialize();
}
```

**Confidence:** HIGH — pattern directly from Microsoft Q&A answer for the same symptom
(VSTO startup disabling).

Source: https://learn.microsoft.com/en-us/answers/questions/1056423/vsto-outlook-improve-and-accelerate-add-in-startup

---

### Pattern 2: Async Network Call with Captured SynchronizationContext

**What:** Capture `SynchronizationContext.Current` on the UI thread before going async. Perform HTTP
work on a thread pool thread using `HttpClient.GetAsync`. Marshal results back to the UI thread via
the captured context for any Outlook OOM calls that must follow.

**When to use:** Any outbound network call (GoPhish notification, future integrations).

**Why it works:** Outlook OOM is not thread-safe — it must be accessed on the main thread. Network
I/O never belongs on the UI thread. The pattern cleanly separates where each kind of work happens.

**Trade-offs:** Slightly more verbose than a synchronous call. The `async void` entry point (on the
button handler) must have its own top-level try-catch to prevent unobserved exceptions from crashing
the add-in.

**Example:**
```csharp
// ReportOrchestrator.cs — async orchestration
public async Task ReportAsync(MailItem mailItem)
{
    // STEP 1: All Outlook OOM access happens here — still on UI thread
    var report = EmailExtractor.Extract(mailItem);
    string headers = mailItem.HeaderString();

    // STEP 2: Detect GoPhish — pure string logic, no thread concern
    string reportUrl = GoPhishDetector.GetReportUrl(headers);

    if (reportUrl != null)
    {
        // STEP 3: HTTP call leaves the UI thread — ConfigureAwait(false) is correct here
        // because we don't need to update Outlook OOM after this await
        bool success = await GoPhishHttpClient.NotifyAsync(reportUrl)
                                              .ConfigureAwait(false);
        // STEP 4: If we need Outlook OOM after await, we must be back on UI thread.
        // Either: don't use ConfigureAwait(false), or capture SynchronizationContext first.
    }
    else
    {
        // Standard report: all Outlook OOM (reportEmail.Send()) before any await
        // No async needed for email-only path — it's fast
        SendReportEmail(report, mailItem);
    }
}

// Ribbon.cs — async void is appropriate here because it IS an event handler
public async void reportPhishing(IRibbonControl control)
{
    try
    {
        var result = MessageBox.Show("...", "Are you sure?", MessageBoxButtons.YesNo);
        if (result == DialogResult.Yes)
        {
            await _orchestrator.ReportAsync(GetSelectedMailItem());
        }
    }
    catch (Exception ex)
    {
        ErrorReporter.Send(ex); // no await — fire-and-forget for error report
    }
}
```

**Confidence:** HIGH — SynchronizationContext pattern for VSTO confirmed by multiple official
Microsoft and expert sources.

Sources:
- https://learn.microsoft.com/en-us/answers/questions/78894/vsto-outlook-addin-how-to-update-ui-from-asynchron
- https://www.add-in-express.com/creating-addins-blog/office-addins-threads/

---

### Pattern 3: HttpClient as Singleton (Not Per-Call)

**What:** Create one static `HttpClient` instance for the lifetime of the add-in. Do not create a
new `HttpClient` per GoPhish notification call.

**When to use:** Always for VSTO add-ins that make HTTP calls. `HttpWebRequest` (the current
implementation) blocks the thread and does not support async naturally. `HttpClient` was designed
for async from the start.

**Why it works:** `HttpClient` is explicitly designed to be instantiated once and reused. Per-call
instantiation causes socket exhaustion (TIME_WAIT) under repeated use. The singleton also allows
setting default headers, timeout, and base address once.

**Trade-offs:** Singleton requires thread-safe initialization. Use `Lazy<HttpClient>` or initialize
in `Application_AfterOutlookReady`.

**Example:**
```csharp
// GoPhishHttpClient.cs
internal static class GoPhishHttpClient
{
    private static readonly HttpClient _client = new HttpClient
    {
        Timeout = TimeSpan.FromSeconds(10) // never block forever
    };

    public static async Task<bool> NotifyAsync(string reportUrl)
    {
        try
        {
            HttpResponseMessage response = await _client
                .GetAsync(reportUrl)
                .ConfigureAwait(false);
            return response.IsSuccessStatusCode;
        }
        catch (TaskCanceledException)
        {
            // Timeout — GoPhish server unreachable, not a crash condition
            return false;
        }
        catch (HttpRequestException)
        {
            // Network error — not a crash condition
            return false;
        }
    }
}
```

**Confidence:** HIGH — HttpClient singleton is official Microsoft guidance for .NET Framework 4.6+.

---

### Pattern 4: Error Isolation at Every Entry Point

**What:** Wrap every public method that Outlook can invoke (every ribbon callback, every event
handler) in its own top-level try-catch. Do not rely on a single outer catch to handle everything.

**When to use:** All ribbon button callbacks, all Application.* event handlers.

**Why it works:** VSTO add-ins run in their own AppDomain but unhandled exceptions in event
handlers still surface to Outlook and can trigger hard disabling. A caught exception that is handled
gracefully (log, show message, continue) does not trigger resiliency disabling.

**Trade-offs:** Slightly repetitive. Use a helper method or extension to centralize the error
reporting side effect without losing the per-entry-point isolation.

**Example:**
```csharp
// Every ribbon callback follows this structure
public async void reportPhishing(IRibbonControl control)
{
    try
    {
        await _orchestrator.ReportAsync(GetSelectedMailItem());
    }
    catch (COMException comEx)
    {
        // Outlook OOM error — log and show user message
        Logger.Error("Outlook COM error during report", comEx);
        MessageBox.Show("Could not access the selected email. Please try again.");
    }
    catch (Exception ex)
    {
        // Unexpected error — log details, send error report, show user message
        Logger.Error("Unexpected error during phishing report", ex);
        ErrorReporter.SendAsync(ex); // fire-and-forget — do not await in catch
        MessageBox.Show("An error occurred. The support team has been notified.");
    }
}
```

**Confidence:** HIGH — pattern confirmed by Microsoft VSTO programming documentation and exception
handling guidance.

Source: https://learn.microsoft.com/en-us/visualstudio/vsto/programming-vsto-add-ins?view=vs-2022

---

### Pattern 5: Registry-Based Resiliency Exemption (Deployment Layer)

**What:** Deploy a registry key via the MSI installer that exempts the add-in from Outlook's
automatic resiliency disabling. This is a defense-in-depth measure — code fixes are primary,
registry is a safety net for the transition period.

**When to use:** MSI installer, deployed via GPO alongside the code fixes.

**Key registry paths:**
```
# Per-user (non-GPO, installed by MSI):
HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Outlook\Resiliency\DoNotDisableAddinList
DWORD: PhishingReporter  = 1

# GPO-managed (set by admin policy — preferred for enterprise):
HKEY_CURRENT_USER\Software\Policies\Microsoft\Office\16.0\Outlook\Resiliency\AddinList
String: PhishingReporter = 1   (1 = always enabled)
```

**Important caveat:** The GPO/DoNotDisableAddinList protects against performance-based disabling
but NOT against crash-based disabling. If the add-in crashes Outlook, it will still be hard-disabled.
Code-level fixes are required to eliminate crash risk. The registry entry is a safety net for the
remaining ~50% startup failure rate during the transition.

**Confidence:** HIGH — directly from Microsoft's official resiliency documentation.

Source: https://learn.microsoft.com/en-us/office/vba/outlook/concepts/getting-started/support-for-keeping-add-ins-enabled

## Data Flow

### Phishing Report Submission Flow (Current — Broken)

```
User Click (UI Thread)
    ↓
Ribbon.reportPhishing()
    ↓
reportPhishingEmailToSecurityTeam() — synchronous, UI thread
    ↓
EmailExtractor methods — UI thread, OK
    ↓
GoPhishIntegration.sendReportNotificationToServer() — BLOCKS UI THREAD
    ↓ (1–5 second HTTP call on UI thread)
MessageBox.Show / reportEmail.Send() — UI thread, OK
```

### Phishing Report Submission Flow (Target — Fixed)

```
User Click (UI Thread)
    ↓
Ribbon.reportPhishing() — async void, top-level try-catch
    ↓
await ReportOrchestrator.ReportAsync() — still on UI thread here
    ↓
EmailExtractor.Extract(mailItem) — UI thread, reads Outlook OOM
    ↓
GoPhishDetector.GetReportUrl(headers) — UI thread, pure string logic
    ↓
if GoPhish campaign:
    await GoPhishHttpClient.NotifyAsync(url) — switches to background thread
    ↓ (HTTP call runs without blocking UI)
    returns bool success — no Outlook OOM access needed after await
    ↓
    Show MessageBox on UI thread (MessageBox.Show is UI-safe)
else:
    SendReportEmail(report, mailItem) — UI thread, Outlook OOM
    ↓
    mailItem.Delete() — UI thread
```

### State Flow (Counter Persistence Fix)

```
Report succeeds
    ↓
Properties.Settings.Default.counter++
    ↓
Properties.Settings.Default.Save()   ← currently missing, fix is one line
    ↓
Counter persists across Outlook restarts
```

### Startup Flow (Current — Slow)

```
Outlook starts → measures time →
ThisAddIn_Startup() → [currently empty, but .NET cold start adds overhead] →
Ribbon constructor → [currently empty] →
Outlook resiliency check: median startup > 1000ms? → disable
```

### Startup Flow (Target — Fast)

```
Outlook starts → measures time →
ThisAddIn_Startup() → hook Application.Startup event → return immediately (<5ms)
                                                       ↑
                                    Outlook resiliency check: ~5ms << 1000ms threshold
                                    Add-in stays enabled

Application.Startup fires (after resiliency window) →
SettingsValidator.ValidateOrWarn() →
GoPhishHttpClient.Initialize() (pre-warm) →
Done — add-in ready
```

### Key Data Flows

1. **Email extraction flow:** MailItem (Outlook OOM) → EmailExtractor → EmailReport record (immutable POCO) → ReportOrchestrator formats and sends. All Outlook OOM access is complete before any async boundary.

2. **GoPhish detection flow:** Raw header string → GoPhishDetector → report URL string or null. Pure function, no I/O, no Outlook dependency. Testable in isolation.

3. **GoPhish notification flow:** Report URL string → GoPhishHttpClient.NotifyAsync() → bool success. Runs entirely on thread pool, never touches Outlook OOM.

4. **Error flow:** Exception at any layer → caught in Ribbon callback → logged (file) → ErrorReporter.SendAsync() (fire-and-forget) → MessageBox to user. The add-in never propagates an exception to Outlook's COM boundary.

## Scaling Considerations

This is a client-side add-in with no server-side scaling concerns. The relevant dimension is
behavior under degraded conditions, not user count.

| Condition | Current Behavior | Target Behavior |
|-----------|-----------------|-----------------|
| GoPhish server down/slow | UI freezes for entire timeout (infinite) | UI responsive; async call times out after 10s; user sees failure message |
| GoPhish server unreachable | Same freeze | Same graceful timeout; error logged |
| Large email (100+ URLs) | String += loops slow; noticeable lag | StringBuilder; still on UI thread but measurably faster |
| Large attachment (100MB+) | Streaming hash already implemented; no change needed | No change; streaming is already correct |
| Outlook restart | Counters reset (bug) | Counters persist after Settings.Save() fix |

### Scaling Priorities (by impact)

1. **First priority — async GoPhish call:** Eliminates 1–5 second UI freeze on every simulated campaign report. Direct user impact.
2. **Second priority — startup deferred initialization:** Eliminates ~50% load failure rate. The root resiliency issue.
3. **Third priority — registry key in MSI:** Defense-in-depth for any residual startup slowness from .NET cold start that code alone cannot fix.

## Anti-Patterns

### Anti-Pattern 1: Synchronous HTTP on the UI Thread

**What people do:** Call `HttpWebRequest.GetResponse()` or `HttpClient.GetAsync().Result` inside a
ribbon button callback.

**Why it's wrong:** The ribbon button callback runs on Outlook's main UI thread. Any blocking call
on this thread freezes the entire Outlook application. Outlook's resiliency mechanism also measures
time spent in event handlers; a 5-second freeze can trigger soft-disabling. Users perceive the
add-in as crashed.

**Do this instead:** Use `await HttpClient.GetAsync()` in an `async Task` method called from an
`async void` event handler. Complete all Outlook OOM access before the first `await`.

---

### Anti-Pattern 2: Heavy Work in ThisAddIn_Startup

**What people do:** Database connections, HTTP calls, WPF control initialization, or settings
loading in `ThisAddIn_Startup`.

**Why it's wrong:** Outlook measures elapsed time during its own startup sequence. `ThisAddIn_Startup`
runs inside this measured window. Microsoft's own default VSTO template causes ~1.62 second startup
— enough to cross the 1000 ms threshold and trigger resiliency disabling.

**Do this instead:** Return from `ThisAddIn_Startup` immediately after hooking `Application.Startup`.
Move all non-trivial work to the `Application.Startup` handler, which runs after the measurement
window.

---

### Anti-Pattern 3: Calling Outlook OOM from Background Threads

**What people do:** Move email processing to `Task.Run()` to avoid blocking the UI, then access
`mailItem.HTMLBody`, `mailItem.Attachments`, or other Outlook OOM from that background task.

**Why it's wrong:** Outlook's object model is not thread-safe. Outlook 2013+ returns
`E_RPC_WRONG_THREAD` when OOM is called from a non-UI thread. This throws a `COMException` and
can hard-crash the add-in.

**Do this instead:** Collect all data from Outlook OOM on the UI thread, store it in a plain C#
object (the `EmailReport` record), then pass that object to background threads for network calls.
Background threads never touch Outlook objects directly.

---

### Anti-Pattern 4: Per-Call HttpClient Instantiation

**What people do:** `new HttpClient()` inside the method that makes the HTTP call, often wrapped in
`using`.

**Why it's wrong:** Each `new HttpClient()` creates a new `HttpClientHandler` which holds a socket.
`Dispose()` does not release the socket immediately — it enters TIME_WAIT for up to 4 minutes. Under
any reasonable report volume this exhausts available sockets.

**Do this instead:** One static `HttpClient` singleton for the add-in lifetime, initialized in
`Application_AfterOutlookReady`.

---

### Anti-Pattern 5: Magic String Return Values for Error States

**What people do:** Return `"NaN"` or `"ERROR"` as sentinel values from methods that can fail.

**Why it's wrong:** Callers must string-compare against undocumented magic values. Typos cause
silent failures. The type system cannot help. The current `GoPhishIntegration` class returns `"NaN"`
and `"ERROR"` — callers do not distinguish network failure from valid responses.

**Do this instead:** Return `null` for "not found" cases (nullable reference types are the idiomatic
.NET pattern). Return `bool` or a `Result<T>` for success/failure. Reserve exceptions for genuine
exceptional conditions.

## Integration Points

### External Services

| Service | Integration Pattern | Notes |
|---------|---------------------|-------|
| GoPhish server | Async HTTP GET via `HttpClient.GetAsync()` | 10-second timeout; failure is not fatal; caller logs and continues |
| Exchange / Active Directory | Outlook OOM `GetExchangeUser()` — UI thread only | Returns null for non-Exchange accounts; null-check each property individually |
| Infosec email (report destination) | `MailItem.Send()` via Outlook OOM — UI thread | Validate address format at startup, not at send time |
| Support email (error destination) | `MailItem.Send()` via Outlook OOM — fire-and-forget | Send error report asynchronously; do not block user on error reporting |

### Internal Boundaries

| Boundary | Communication | Notes |
|----------|---------------|-------|
| Ribbon -> ReportOrchestrator | Direct method call (async Task) | Ribbon holds orchestrator as field; no global state |
| ReportOrchestrator -> EmailExtractor | Direct method call; returns immutable EmailReport | EmailExtractor has no async methods — all Outlook OOM |
| ReportOrchestrator -> GoPhishDetector | Direct method call; returns string? (null if no header) | Pure function; no I/O |
| ReportOrchestrator -> GoPhishHttpClient | await Task<bool> | Network boundary; all async |
| ThisAddIn -> SettingsValidator | Called once in Application.Startup | SettingsValidator reads Properties.Settings; no Outlook OOM |

## Suggested Build Order

Based on component dependencies and risk, implement in this sequence:

**Phase 1 — Foundation (no behavior change, enables everything else)**
1. `MailItemExtensions.cs` — move from Ribbon.cs end; no logic change
2. `EmailReport.cs` — create immutable data record for extracted email data
3. `SettingsValidator.cs` — startup validation; standalone; zero risk

**Phase 2 — Extraction Layer (no async required; pure refactoring)**
4. `EmailExtractor.cs` — extract URL/attachment/header methods from Ribbon.cs
5. `GoPhishDetector.cs` — extract pure string logic from GoPhishIntegration.cs

**Phase 3 — Async Network Layer (highest impact; isolated change)**
6. `GoPhishHttpClient.cs` — replace `HttpWebRequest` with `HttpClient.GetAsync()`; add timeout; singleton

**Phase 4 — Orchestration Layer (wires everything together)**
7. `ReportOrchestrator.cs` — async coordinator; calls Phase 2 and 3 components
8. `Ribbon.cs` — thin to UI callbacks only; delegates to ReportOrchestrator

**Phase 5 — Startup Fix (resiliency threshold fix)**
9. `ThisAddIn.cs` — move work to `Application.Startup`; add deferred initialization pattern

**Phase 6 — Deployment (registry safety net)**
10. MSI installer update — add `DoNotDisableAddinList` registry key for both Office 15.0 and 16.0

**Why this order:**
- Phases 1-2 are pure refactoring with no runtime behavior change — lower risk, establishes structure
- Phase 3 is the async fix; doing it in isolation means GoPhishHttpClient can be reviewed and tested before wiring it up
- Phase 4 wires everything; Ribbon.cs shrinks dramatically
- Phase 5 is the startup fix; doing it last means the full async stack is already in place
- Phase 6 is additive; MSI change deploys alongside Phase 5 code

## Sources

- [Support for keeping add-ins enabled | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/outlook/concepts/getting-started/support-for-keeping-add-ins-enabled) — HIGH confidence, official docs
- [Architecture of VSTO Add-ins | Microsoft Learn](https://learn.microsoft.com/en-us/visualstudio/vsto/architecture-of-vsto-add-ins?view=vs-2022) — HIGH confidence, official docs
- [Improve and accelerate Add-in startup process | Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/1056423/vsto-outlook-improve-and-accelerate-add-in-startup) — HIGH confidence, Application.Startup pattern
- [VSTO Outlook addin — how to update UI from async process | Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/78894/vsto-outlook-addin-how-to-update-ui-from-asynchron) — MEDIUM confidence, async/SynchronizationContext
- [How to work with threads in Office COM add-ins | Add-in Express](https://www.add-in-express.com/creating-addins-blog/office-addins-threads/) — MEDIUM confidence, threading rules
- [Default VSTO project causes Outlook to start slowly | Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/1377543/the-default-project-for-my-visual-studio-outlook-v) — HIGH confidence, 1000ms threshold documented
- [DoNotDisableAddinList | Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/4511521/donotdisableaddinlist) — HIGH confidence, registry key details
- [Always Load an Outlook Addin | Slipstick Systems](https://www.slipstick.com/outlook/always-load-an-outlook-addin/) — MEDIUM confidence, enterprise GPO pattern

---
*Architecture research for: VSTO Outlook add-in reliability refactoring*
*Researched: 2026-02-25*
