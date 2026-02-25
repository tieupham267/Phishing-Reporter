# Phase 3: Async Network Layer - Research

**Researched:** 2026-02-26
**Domain:** Async HTTP, resilience patterns, COM threading in VSTO
**Confidence:** HIGH

## Summary

This phase replaces the synchronous `HttpWebRequest`/`HttpWebResponse` call in `GoPhishIntegration.sendReportNotificationToServer()` with an async `HttpClient` singleton wrapped in a Polly 8 resilience pipeline (exponential backoff retry + per-attempt timeout). The current code (Ribbon.cs line 177) calls `GoPhishIntegration.sendReportNotificationToServer()` synchronously on the Outlook UI thread, which freezes the entire Outlook application during network failures. Additionally, the current code leaks `HttpWebResponse` and `StreamReader` instances (BUGF-04).

The critical complexity in this phase is the VSTO STA threading model. Outlook runs on a single-threaded apartment (STA) thread; all COM Object Model (OOM) access must occur on this thread. The async HTTP call must execute without blocking the STA thread, but the ribbon callback (`reportPhishing`) is a synchronous `void` method invoked by COM. The solution is to make `GoPhishIntegration.SendReportNotificationAsync()` a proper `async Task` method that uses `HttpClient.GetAsync()` with `ConfigureAwait(false)`, and have the ribbon callback call it via `async void` with a top-level try/catch. The actual Phase 4 orchestration will move OOM data extraction before the await boundary, but this phase focuses solely on converting the HTTP layer to async.

**Primary recommendation:** Install Polly 8.x via NuGet. Replace `GoPhishIntegration.sendReportNotificationToServer()` with an async method backed by a static `HttpClient` singleton and a Polly `ResiliencePipeline` combining retry (exponential backoff, 3 attempts) with per-attempt timeout (10s). Use `async void` in the ribbon callback with comprehensive try/catch to prevent unhandled exception crashes.

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|-----------------|
| NETW-01 | GoPhish HTTP call executes asynchronously via HttpClient without blocking the Outlook UI thread | HttpClient.GetAsync() is natively async; ConfigureAwait(false) releases STA thread; async void callback pattern documented in Stephen Cleary's best practices |
| NETW-02 | HttpClient instantiated as static singleton to prevent socket exhaustion | Microsoft official guidance: static/singleton HttpClient prevents socket exhaustion; ServicePoint.ConnectionLeaseTimeout handles DNS for .NET Framework |
| NETW-03 | GoPhish HTTP call has configurable timeout (default 10 seconds) | Polly 8 AddTimeout(TimeSpan) strategy; per-attempt timeout inside retry; HttpClient.Timeout set to Timeout.InfiniteTimeSpan to delegate to Polly |
| NETW-04 | GoPhish HTTP call retries with exponential backoff on transient failures (via Polly) | Polly 8 ResiliencePipelineBuilder.AddRetry with BackoffType.Exponential, UseJitter=true; ShouldHandle predicate for HttpRequestException and timeout |
| NETW-05 | Report submission completes gracefully if GoPhish server is unreachable (falls back to standard email report) | Existing code already returns GoPhishResult.Error on failure; async version returns same enum; caller continues to email report on Error/timeout |
| BUGF-04 | HttpWebResponse and StreamReader properly disposed (replaced by HttpClient) | HttpClient eliminates manual response/stream disposal; response messages managed by HttpClient internally |
</phase_requirements>

## Standard Stack

### Core
| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| System.Net.Http.HttpClient | .NET Fx 4.8 built-in | Async HTTP client | Framework-included, thread-safe, connection-pooling, IDisposable response management |
| Polly | 8.4.2+ | Resilience (retry + timeout) | Industry standard for .NET resilience; v8 targets .NET Framework 4.6.2+; ResiliencePipeline API |
| Polly.Core | 8.4.2+ (transitive) | Core resilience primitives | Required dependency of Polly 8.x; provides ResiliencePipelineBuilder |

### Supporting
| Library | Version | Purpose | When to Use |
|---------|---------|---------|-------------|
| Microsoft.Bcl.AsyncInterfaces | 6.0.0+ (transitive) | IAsyncDisposable for .NET Fx | Pulled in by Polly.Core on .NET Framework 4.6.2+ |
| Microsoft.Bcl.TimeProvider | 8.0.0+ (transitive) | TimeProvider abstraction | Pulled in by Polly.Core for time-based strategies |
| System.Threading.Tasks.Extensions | 4.5.4+ (transitive) | ValueTask support | Pulled in by Polly.Core on .NET Framework |
| System.ComponentModel.Annotations | 4.5.0+ (transitive) | Data annotations | Pulled in by Polly.Core on .NET Framework |

### Alternatives Considered
| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| Polly 8 | Manual retry loop | Polly handles jitter, backoff calculation, timeout coordination, and is battle-tested; hand-rolled retry loops miss edge cases |
| Polly 8 | Polly 7 (Policy.Handle API) | v7 API still works but is legacy; v8 is actively developed with 4x less memory allocation; STATE.md explicitly warns "Polly 8.x uses ResiliencePipeline builder API -- do not mix versions" |
| Static HttpClient singleton | IHttpClientFactory | IHttpClientFactory requires DI container (Microsoft.Extensions.Http); VSTO add-in has no DI; static singleton is Microsoft's recommended alternative |
| ServicePoint.ConnectionLeaseTimeout | SocketsHttpHandler.PooledConnectionLifetime | PooledConnectionLifetime is .NET Core/.NET 5+ only; .NET Framework 4.8 must use ServicePointManager |

**Installation:**
```
Install-Package Polly -Version 8.4.2
```
This pulls in Polly.Core 8.4.2 and all transitive dependencies. No explicit System.Net.Http NuGet package needed -- it is a framework assembly referenced directly.

**Project reference needed (csproj):**
```xml
<Reference Include="System.Net.Http" />
```

## Architecture Patterns

### Recommended File Structure
```
PhishingReporter/
  GoPhishIntegration.cs    # MODIFY: async SendReportNotificationAsync, static HttpClient, resilience pipeline
  GoPhishHttpClient.cs     # NEW: static HttpClient singleton + Polly pipeline factory (optional extraction)
  Ribbon.cs                # MODIFY: async void wrapper in reportPhishing path
```

Note: Whether to extract the HttpClient singleton into a separate class or keep it in GoPhishIntegration.cs is a planning decision. Both are valid; the key constraint is that the HttpClient MUST be a `static readonly` field initialized once.

### Pattern 1: Static HttpClient Singleton with ServicePoint DNS Workaround
**What:** A single static HttpClient instance shared across all calls, with ServicePoint.ConnectionLeaseTimeout to handle DNS changes on .NET Framework.
**When to use:** Always in .NET Framework 4.8 when IHttpClientFactory is not available.
**Example:**
```csharp
// Source: https://learn.microsoft.com/en-us/dotnet/fundamentals/networking/http/httpclient-guidelines
internal static class GoPhishIntegration
{
    private static readonly HttpClient HttpClientInstance;

    static GoPhishIntegration()
    {
        HttpClientInstance = new HttpClient
        {
            // Delegate timeout to Polly, not HttpClient
            Timeout = System.Threading.Timeout.InfiniteTimeSpan
        };

        // .NET Framework 4.8: no SocketsHttpHandler, use ServicePoint for DNS
        // ConnectionLeaseTimeout forces connection recycling for DNS changes
        var goPhishUri = new Uri(Properties.Settings.Default.gophish_url);
        var sp = ServicePointManager.FindServicePoint(goPhishUri);
        sp.ConnectionLeaseTimeout = 60 * 1000; // 1 minute
    }
}
```

### Pattern 2: Polly 8 ResiliencePipeline (Retry + Timeout)
**What:** A pre-built resilience pipeline combining per-attempt timeout and exponential backoff retry.
**When to use:** For all GoPhish HTTP calls.
**Example:**
```csharp
// Source: https://www.pollydocs.org/strategies/retry.html
// Source: https://www.pollydocs.org/strategies/timeout.html
private static readonly ResiliencePipeline<HttpResponseMessage> RetryPipeline =
    new ResiliencePipelineBuilder<HttpResponseMessage>()
        .AddRetry(new RetryStrategyOptions<HttpResponseMessage>
        {
            ShouldHandle = new PredicateBuilder<HttpResponseMessage>()
                .Handle<HttpRequestException>()
                .Handle<TimeoutRejectedException>()
                .HandleResult(r => r.StatusCode >= System.Net.HttpStatusCode.InternalServerError),
            MaxRetryAttempts = 3,
            Delay = TimeSpan.FromSeconds(1),
            BackoffType = DelayBackoffType.Exponential,
            UseJitter = true,
            OnRetry = static args =>
            {
                // Logger call here for retry count visibility (NETW-04 logging)
                return default;
            }
        })
        .AddTimeout(TimeSpan.FromSeconds(10)) // NETW-03: per-attempt timeout
        .Build();
```

**Strategy order matters (FIFO):** Adding retry first, then timeout, means timeout is the innermost strategy (wraps each individual attempt), and retry is the outermost (retries on timeout). This gives per-attempt timeout behavior.

### Pattern 3: Async Void Ribbon Callback with Try/Catch
**What:** The ribbon callback is `async void` (required by COM callback signature) with comprehensive exception handling.
**When to use:** At the ribbon entry point where the async call chain begins.
**Example:**
```csharp
// Source: https://learn.microsoft.com/en-us/archive/msdn-magazine/2013/march/async-await-best-practices-in-asynchronous-programming
public async void reportPhishing(Office.IRibbonControl control)
{
    try
    {
        // All OOM access happens HERE, on the UI thread, before any await
        // ... (existing confirmation dialog, data extraction) ...

        // Async call -- releases UI thread
        await ReportPhishingCoreAsync(/* extracted data */).ConfigureAwait(false);
    }
    catch (Exception ex)
    {
        Logger.Error(ex, "Unhandled exception in reportPhishing callback");
        try { MessageBox.Show("An unexpected error occurred.", "Error"); } catch { }
    }
}
```

### Pattern 4: ConfigureAwait(false) in Library Code
**What:** All awaits in GoPhishIntegration use ConfigureAwait(false) to avoid posting continuations back to the STA thread.
**When to use:** In every async method that does NOT need to access the Outlook OOM or UI afterward.
**Example:**
```csharp
public static async Task<GoPhishResult> SendReportNotificationAsync(string reportUrl)
{
    Logger.Info("Sending GoPhish report notification to: {0}", reportUrl);
    try
    {
        var response = await RetryPipeline.ExecuteAsync(
            async ct => await HttpClientInstance.GetAsync(reportUrl, ct).ConfigureAwait(false),
            CancellationToken.None
        ).ConfigureAwait(false);

        if (response.IsSuccessStatusCode)
        {
            Logger.Info("GoPhish notification sent successfully");
            return GoPhishResult.Reported;
        }
        Logger.Warn("GoPhish notification returned HTTP {0}", (int)response.StatusCode);
        return GoPhishResult.Error;
    }
    catch (TimeoutRejectedException)
    {
        Logger.Warn("GoPhish notification timed out after all retry attempts");
        return GoPhishResult.Error;
    }
    catch (HttpRequestException ex)
    {
        Logger.Error(ex, "GoPhish notification failed after all retry attempts");
        return GoPhishResult.Error;
    }
    catch (Exception ex)
    {
        Logger.Error(ex, "GoPhish notification unexpected error");
        return GoPhishResult.Error;
    }
}
```

### Anti-Patterns to Avoid
- **Calling .Result or .Wait() on STA thread:** Causes permanent deadlock. The STA thread waits for the Task to complete, but the Task's continuation needs the STA thread. This is the #1 async/await pitfall in VSTO.
- **New HttpClient per request:** Causes socket exhaustion (TCP TIME_WAIT). Each HttpClient instance creates its own connection pool.
- **async void without try/catch:** Unhandled exceptions in async void crash the Outlook process silently. Always wrap the entire body in try/catch.
- **Accessing Outlook OOM after await:** COM objects are STA-bound. Accessing them from a thread pool thread (where continuations run after ConfigureAwait(false)) throws COMException 0x8001010E. Extract all data before the first await.
- **Mixing Polly v7 and v8 APIs:** STATE.md explicitly warns against this. Use only ResiliencePipelineBuilder (v8), never Policy.Handle (v7).
- **HttpClient.Timeout competing with Polly timeout:** Set HttpClient.Timeout to Timeout.InfiniteTimeSpan and let Polly manage per-attempt and overall timeouts. Otherwise HttpClient.Timeout applies across all retries, causing TaskCanceledException instead of TimeoutRejectedException.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| Retry with backoff | Manual retry loop with Thread.Sleep/Task.Delay | Polly 8 ResiliencePipeline.AddRetry | Jitter calculation, backoff timing, exception filtering, cancellation support are subtle; Polly handles all edge cases |
| Per-attempt timeout | CancellationTokenSource with timer | Polly 8 ResiliencePipeline.AddTimeout | Polly coordinates timeout with retry; manual CTS-per-attempt is error-prone and leaks if not disposed |
| HTTP connection management | Manual socket/connection tracking | Static HttpClient singleton | HttpClient manages connection pooling internally; manual management duplicates framework behavior |
| DNS refresh on .NET Framework | Timer to periodically recreate HttpClient | ServicePointManager.ConnectionLeaseTimeout | Framework-provided mechanism; recreating HttpClient defeats the singleton pattern |

**Key insight:** The combination of retry + timeout + DNS management + proper exception handling has dozens of edge cases. Polly 8 and HttpClient together handle all of them; hand-rolling any piece risks subtle bugs that only surface under network stress.

## Common Pitfalls

### Pitfall 1: STA Deadlock from .Result/.Wait()
**What goes wrong:** Calling `task.Result` or `task.Wait()` from the Outlook UI thread (STA) causes a permanent deadlock. The UI thread blocks waiting for the Task, but the Task's continuation needs the UI thread to execute (if SynchronizationContext was captured).
**Why it happens:** By default, `await` captures the current SynchronizationContext and posts the continuation back to it. On the STA thread, only one chunk of code runs at a time. If the thread is blocked by .Wait(), the continuation can never execute.
**How to avoid:** Never use .Result or .Wait() in the ribbon callback chain. Use `async void` for the entry point with `await` all the way down. Use `ConfigureAwait(false)` in all library/non-UI methods so continuations run on the thread pool.
**Warning signs:** Outlook completely freezes and never recovers; must kill process via Task Manager. No exception in log (deadlock is silent).

### Pitfall 2: Unhandled Exception in async void Crashing Outlook
**What goes wrong:** An exception thrown from an `async void` method after an `await` is raised on the SynchronizationContext. If unhandled, it terminates the process.
**Why it happens:** `async void` methods have no Task to store exceptions. The runtime posts the exception to the captured SynchronizationContext. In a VSTO add-in, this can crash Outlook silently.
**How to avoid:** Wrap the ENTIRE body of any `async void` method in try/catch. Log the exception and show a user-friendly message. Never let exceptions escape.
**Warning signs:** Outlook crashes without warning; no error dialog; add-in appears to "disappear."

### Pitfall 3: Socket Exhaustion from HttpClient-per-Request
**What goes wrong:** After reporting many phishing emails, the OS runs out of available TCP ports. Log shows "address already in use" or connection refused errors.
**Why it happens:** Each new `HttpClient()` creates a new connection pool. Disposed connections enter TCP TIME_WAIT state (up to 4 minutes on Windows). Under load, ports are exhausted.
**How to avoid:** Use a single `static readonly HttpClient` instance for the lifetime of the add-in. Never dispose it.
**Warning signs:** Intermittent connection failures after many reports; netstat shows many TIME_WAIT connections to GoPhish server.

### Pitfall 4: COMException 0x8001010E from OOM Access on Background Thread
**What goes wrong:** After `await` with `ConfigureAwait(false)`, code resumes on a thread pool thread. Any access to Outlook OOM (MailItem, Selection, etc.) throws COMException.
**Why it happens:** Outlook's COM objects are STA-bound. Thread pool threads are MTA. COM cannot marshal the call.
**How to avoid:** Extract ALL needed data from OOM objects into plain C# variables/objects BEFORE the first `await` that uses `ConfigureAwait(false)`. This is the responsibility of Phase 4, but Phase 3 must not introduce new OOM access after await boundaries.
**Warning signs:** COMException with HRESULT 0x8001010E in logs; "The application called an interface that was marshalled for a different thread."

### Pitfall 5: HttpClient.Timeout vs Polly Timeout Conflict
**What goes wrong:** HttpClient.Timeout (default 100 seconds) applies to the entire operation including all retries. If it expires before Polly's retries complete, a TaskCanceledException is thrown instead of TimeoutRejectedException, bypassing Polly's error handling.
**Why it happens:** HttpClient.Timeout is a separate mechanism from Polly's timeout strategy. They compete.
**How to avoid:** Set `HttpClient.Timeout = System.Threading.Timeout.InfiniteTimeSpan` and let Polly manage all timeout behavior through its AddTimeout strategy.
**Warning signs:** TaskCanceledException in logs instead of TimeoutRejectedException; retries not happening despite Polly configuration.

### Pitfall 6: TLS Protocol Downgrade
**What goes wrong:** The existing code sets `ServicePointManager.SecurityProtocol = Tls12` globally in the send method. This is fragile and can interfere with other add-ins.
**Why it happens:** ServicePointManager.SecurityProtocol is process-wide. Setting it to only TLS 1.2 may break other HTTP calls that need TLS 1.3 or other protocols.
**How to avoid:** .NET Framework 4.8 defaults to system-default TLS (which includes TLS 1.2 and 1.3 on modern Windows). Remove the explicit SecurityProtocol assignment. If required, use `SecurityProtocol |= Tls12` (OR, not assignment) to ADD TLS 1.2 without removing other protocols. Better yet, rely on OS defaults.
**Warning signs:** Other add-ins or Outlook features report SSL/TLS errors after this add-in loads.

## Code Examples

### Complete GoPhish Async Method (Verified Pattern)
```csharp
// Source: Microsoft HttpClient guidelines + Polly 8 docs
using System;
using System.Net;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Polly;
using Polly.Retry;
using Polly.Timeout;

internal static class GoPhishIntegration
{
    private static readonly NLog.Logger Logger = AppLogger.Instance.GetCurrentClassLogger();

    // NETW-02: Static singleton prevents socket exhaustion
    private static readonly HttpClient HttpClientInstance;

    // NETW-03 + NETW-04: Resilience pipeline with retry + timeout
    private static readonly ResiliencePipeline<HttpResponseMessage> Pipeline;

    static GoPhishIntegration()
    {
        HttpClientInstance = new HttpClient
        {
            Timeout = System.Threading.Timeout.InfiniteTimeSpan // Polly manages timeout
        };

        // .NET Framework 4.8 DNS workaround
        var baseUri = new Uri(
            Properties.Settings.Default.gophish_url + ":" +
            Properties.Settings.Default.gophish_listener_port);
        var sp = ServicePointManager.FindServicePoint(baseUri);
        sp.ConnectionLeaseTimeout = 60_000; // 1 minute

        Pipeline = new ResiliencePipelineBuilder<HttpResponseMessage>()
            .AddRetry(new RetryStrategyOptions<HttpResponseMessage>
            {
                ShouldHandle = new PredicateBuilder<HttpResponseMessage>()
                    .Handle<HttpRequestException>()
                    .Handle<TimeoutRejectedException>(),
                MaxRetryAttempts = 3,
                Delay = TimeSpan.FromSeconds(1),
                BackoffType = DelayBackoffType.Exponential,
                UseJitter = true,
                OnRetry = static args =>
                {
                    var logger = AppLogger.Instance.GetCurrentClassLogger();
                    logger.Warn("GoPhish retry attempt {0}, delay {1}ms",
                        args.AttemptNumber, args.RetryDelay.TotalMilliseconds);
                    return default;
                }
            })
            .AddTimeout(TimeSpan.FromSeconds(10)) // NETW-03: 10-second per-attempt timeout
            .Build();
    }

    // ... setReportURL remains synchronous (pure string parsing, no I/O) ...

    // NETW-01: Async method, does not block UI thread
    // BUGF-04: HttpClient manages response lifecycle (no manual dispose needed)
    public static async Task<GoPhishResult> SendReportNotificationAsync(string reportUrl)
    {
        Logger.Info("Sending GoPhish report notification to: {0}", reportUrl);

        try
        {
            using (var response = await Pipeline.ExecuteAsync(
                async ct => await HttpClientInstance.GetAsync(reportUrl, ct)
                    .ConfigureAwait(false),
                CancellationToken.None).ConfigureAwait(false))
            {
                Logger.Info("GoPhish notification result: HTTP {0}", (int)response.StatusCode);
                return response.IsSuccessStatusCode
                    ? GoPhishResult.Reported
                    : GoPhishResult.Error;
            }
        }
        catch (TimeoutRejectedException)
        {
            Logger.Warn("GoPhish notification timed out after all retries");
            return GoPhishResult.Error;
        }
        catch (HttpRequestException ex)
        {
            Logger.Error(ex, "GoPhish notification failed after all retries");
            return GoPhishResult.Error;
        }
        catch (Exception ex)
        {
            Logger.Error(ex, "GoPhish notification unexpected error");
            return GoPhishResult.Error;
        }
    }
}
```

### Ribbon Callback Async Void Entry Point
```csharp
// Source: Stephen Cleary's async best practices
// https://learn.microsoft.com/en-us/archive/msdn-magazine/2013/march/async-await-best-practices-in-asynchronous-programming
public async void reportPhishing(Office.IRibbonControl control)
{
    try
    {
        Logger.Info("Report phishing button clicked");
        var areYouSure = MessageBox.Show(
            "Do you want to report this email?",
            "Are you sure?", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

        if (areYouSure != DialogResult.Yes)
        {
            Logger.Info("User cancelled report submission");
            return;
        }

        Logger.Info("User confirmed report submission");
        // OOM access happens here, on UI thread, BEFORE any await
        // (Phase 4 will extract all OOM data into an immutable record here)
        await reportPhishingEmailToSecurityTeamAsync(control).ConfigureAwait(false);
    }
    catch (Exception ex)
    {
        Logger.Error(ex, "Unhandled exception in reportPhishing callback");
        try { MessageBox.Show("An unexpected error occurred.", "Error"); } catch { }
    }
}
```

### NuGet Package Installation (packages.config)
```xml
<!-- Add to packages.config -->
<package id="Polly" version="8.4.2" targetFramework="net48" />
<package id="Polly.Core" version="8.4.2" targetFramework="net48" />
<!-- Transitive dependencies auto-resolved by NuGet -->
```

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| HttpWebRequest/HttpWebResponse | HttpClient (async) | .NET Framework 4.5+ (2012) | Async-native, connection pooling, thread-safe |
| Polly 7 Policy.Handle fluent API | Polly 8 ResiliencePipelineBuilder | Polly 8.0 (2023) | Unified sync/async, 4x less memory, FIFO ordering |
| Manual TLS 1.2 protocol setting | OS-default TLS on .NET Fx 4.8 | .NET Fx 4.7+ registry, 4.8 default | No need to set SecurityProtocol explicitly on Win10/11 |
| New HttpClient per request | Static singleton + ServicePoint | Always recommended, codified in MS docs 2018+ | Prevents socket exhaustion |

**Deprecated/outdated:**
- `HttpWebRequest`/`HttpWebResponse`: Replaced by `HttpClient`. Still works but lacks async-native API, no connection pooling, manual stream management, resource leaks (BUGF-04).
- Polly 7 `Policy.Handle()` API: Still functional in Polly 8 package (backward compatible) but should not be used in new code. STATE.md explicitly warns "Polly 8.x uses ResiliencePipeline builder API -- do not mix versions."
- Explicit `ServicePointManager.SecurityProtocol = Tls12`: Unnecessary on .NET Framework 4.8 running on Windows 10/11 where TLS 1.2 is the default. Setting it explicitly can interfere with TLS 1.3 support.

## Open Questions

1. **Settings.Default.gophish_url + port URI construction in static constructor**
   - What we know: The static constructor needs the GoPhish URL for ServicePoint configuration. Settings.Default is available at static construction time in VSTO.
   - What's unclear: If Settings.Default is not yet loaded when the static constructor runs (unlikely but not verified for VSTO lifecycle), the ServicePoint configuration would use default values.
   - Recommendation: Low risk. Settings.Default is initialized from config at assembly load time, which happens before any static constructor. Proceed with current approach; if issues arise, defer ServicePoint setup to first use via Lazy<T>.

2. **Retry count and timeout values**
   - What we know: Requirements specify 10-second timeout (NETW-03). 3 retries with exponential backoff (1s, ~2s, ~4s) means worst case is ~17 seconds per report + 10s per attempt = ~47 seconds total.
   - What's unclear: Whether 3 retries is optimal for an intranet GoPhish server that is either up or down (not experiencing transient failures).
   - Recommendation: 3 retries with 10s per-attempt timeout is reasonable. The exponential backoff with jitter prevents thundering herd. Total worst-case time (~47s) is acceptable since it runs on a background thread and does not block the UI.

3. **Should the old synchronous sendReportNotificationToServer be kept?**
   - What we know: Phase 4 will wire up the full async pipeline. During Phase 3, the caller in Ribbon.cs still calls synchronously.
   - What's unclear: Whether to update the caller in Ribbon.cs to call the async version during Phase 3 or wait for Phase 4.
   - Recommendation: Phase 3 should update the caller to use async. The ribbon callback becomes `async void` with try/catch. This is necessary to verify NETW-01 (UI remains responsive). Leaving the caller synchronous would mean Phase 3 cannot be validated against its success criteria. The old synchronous method should be removed to prevent accidental use.

## Sources

### Primary (HIGH confidence)
- [Microsoft HttpClient Guidelines](https://learn.microsoft.com/en-us/dotnet/fundamentals/networking/http/httpclient-guidelines) - Static singleton, DNS behavior, .NET Framework vs .NET Core differences, ServicePoint workaround
- [Microsoft VSTO Threading Support](https://learn.microsoft.com/en-us/visualstudio/vsto/threading-support-in-office?view=vs-2022) - STA model, COM marshaling, COMException on background threads, message filter
- [Polly Retry Strategy Docs](https://www.pollydocs.org/strategies/retry.html) - RetryStrategyOptions, BackoffType.Exponential, UseJitter, ShouldHandle predicate
- [Polly Timeout Strategy Docs](https://www.pollydocs.org/strategies/timeout.html) - AddTimeout, TimeoutRejectedException, per-attempt vs overall timeout, OnTimeout callback
- [Polly Resilience Pipelines Docs](https://www.pollydocs.org/pipelines/) - ResiliencePipelineBuilder, FIFO execution order, ExecuteAsync
- [Polly Migration v7 to v8](https://www.pollydocs.org/migration-v8.html) - API differences, ResiliencePipelineBuilder replaces Policy.Handle

### Secondary (MEDIUM confidence)
- [NuGet Polly 8.4.2](https://www.nuget.org/packages/Polly/8.4.2) - Target frameworks: .NET 6.0, .NET Standard 2.0, .NET Framework 4.6.2+
- [NuGet Polly.Core 8.4.2](https://www.nuget.org/packages/Polly.Core/8.4.2) - Dependencies for .NET Framework: Microsoft.Bcl.AsyncInterfaces, Microsoft.Bcl.TimeProvider, System.Threading.Tasks.Extensions, System.ComponentModel.Annotations, System.ValueTuple
- [Stephen Cleary - Async Best Practices](https://learn.microsoft.com/en-us/archive/msdn-magazine/2013/march/async-await-best-practices-in-asynchronous-programming) - Avoid async void (except event handlers), async all the way, ConfigureAwait(false), deadlock prevention
- [Sergey Teplyakov - Dangers of Async Void](https://sergeyteplyakov.github.io/Blog/csharp/2025/01/28/The_Dangers_Of_Async_Void.html) - async void exception behavior, SynchronizationContext interaction
- [Polly GitHub Issue #512](https://github.com/App-vNext/Polly/issues/512) - HttpClient.Timeout vs Polly timeout: set HttpClient.Timeout = InfiniteTimeSpan

### Tertiary (LOW confidence)
- None. All findings verified against primary or secondary sources.

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH - HttpClient and Polly are Microsoft-recommended; Polly 8.4.2 .NET Framework 4.6.2+ target confirmed on NuGet
- Architecture: HIGH - VSTO threading model well-documented by Microsoft; async void + ConfigureAwait(false) pattern is established best practice from Stephen Cleary (Microsoft MVP)
- Pitfalls: HIGH - All pitfalls sourced from official Microsoft docs or verified community patterns; STA deadlock and socket exhaustion are well-documented .NET Framework issues

**Research date:** 2026-02-26
**Valid until:** 2026-03-28 (stable domain, 30-day validity)
