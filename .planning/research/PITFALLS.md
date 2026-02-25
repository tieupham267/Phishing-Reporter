# Pitfalls Research

**Domain:** VSTO Outlook Add-in Reliability (Enterprise)
**Researched:** 2026-02-25
**Confidence:** HIGH — primary sources are official Microsoft documentation and confirmed community patterns

---

## Critical Pitfalls

### Pitfall 1: Treating the Resiliency Disable as a Symptom, Not the Root Cause

**What goes wrong:**
Teams add GPO registry overrides (`AddinList` = 1 or `DoNotDisableAddinList`) to force the add-in to stay enabled without fixing the actual startup time violation. The add-in continues to exceed Outlook's 1000ms median startup threshold. Users are protected from the disable dialog, but Outlook still records Event ID 59 warnings, the add-in still runs slowly at startup, and antivirus or hardware variation can still trigger edge-case crashes that the policy override cannot suppress.

**Why it happens:**
The registry fix is fast, visible, and deployable via GPO — it feels like a complete solution. The startup time regression requires profiling and code changes, which take longer and carry risk. Teams reach for the administrative lever first and never revisit the root cause.

**How to avoid:**
Fix the startup time first. Registry overrides are a legitimate last resort for the 1-second threshold when genuinely necessary, but they must be layered on top of a fast-loading add-in. The correct sequence is: measure actual startup time using Outlook's Application Event Log (Event ID 45 records load time in milliseconds), identify what is slow in `ThisAddIn_Startup`, move all non-essential work out of the startup path, then apply registry overrides only if the 1-second threshold is genuinely impossible to meet.

**Warning signs:**
- Event ID 59 appearing in Application Event Log after every restart
- Add-in disabling every 5 Outlook launches (Outlook uses median over 5 successive iterations)
- Startup takes 1.62+ seconds even on a blank VSTO template (documented case)
- Antivirus or certificate revocation checking adding latency at startup

**Phase to address:** Phase addressing Outlook resiliency (startup reliability fix)

---

### Pitfall 2: Deferring Heavy Work to a Background Thread and Then Calling Outlook Object Model on It

**What goes wrong:**
Developers correctly identify that startup work must be deferred off the critical path. They fire off a `Task.Run()` or `Thread` to do the heavy work, and then — inside that background thread — call into the Outlook Object Model (OOM) to read mailbox properties or access session information. In Outlook 2013+, OOM calls from a non-main thread return `E_RPC_WRONG_THREAD` (a `COMException`). The add-in crashes silently or throws unhandled exceptions that appear as Outlook instability.

**Why it happens:**
Async programming instinct says "move slow work off the main thread." The `Task.Run()` pattern works correctly for pure .NET work (HTTP calls, file I/O, hashing). Developers apply the same pattern to OOM calls without knowing that the Outlook COM server is STA-only and will reject calls marshaled from MTA thread pool threads.

**How to avoid:**
Split background work into two categories:
1. Work that does NOT touch Outlook OOM (HTTP calls to GoPhish, file hashing, email serialization) — safe to run on `Task.Run()` or `Thread` with MTA.
2. Work that DOES touch Outlook OOM (accessing `Application.Session`, reading email properties, `ExchangeUser` lookup) — must run on the main STA thread.

For the GoPhish HTTP call specifically: the correct pattern is to read all required data from the OOM on the main thread first, then pass immutable values to `Task.Run()` for the actual network call. Never pass live OOM objects across thread boundaries.

**Warning signs:**
- `COMException` with message "The message filter indicated that the application is busy" in logs
- `E_RPC_WRONG_THREAD` HRESULT 0x8001010E appearing in crash reports
- Outlook hangs intermittently rather than deterministically (indicates thread contention)
- Background thread works in development but fails for some users in production (hardware/load variance)

**Phase to address:** Phase converting synchronous HTTP calls to async

---

### Pitfall 3: Using `async void` on VSTO Event Handlers

**What goes wrong:**
Ribbon button click handlers and other COM event callbacks must match the void-returning delegate signature, so converting them to async produces `async void`. When an exception is thrown inside an `async void` method after an `await` point, it is raised on the `SynchronizationContext` that was active at the method start — not caught by any surrounding try/catch. In practice this crashes Outlook. The exception bypasses the add-in's own error handling (the `catch (Exception ex)` block in `reportPhishingEmailToSecurityTeam`) and propagates directly to the host process.

**Why it happens:**
VSTO event handler signatures are fixed by COM. `async Task` cannot be used for event handlers. The C# compiler does not warn about `async void`. Developers converting synchronous code to async naturally convert the event handler itself rather than extracting the async work into a separate `Task`-returning method.

**How to avoid:**
Keep event handlers `async void` but immediately delegate to an `async Task` method and wrap the entire call in try/catch:

```csharp
// Event handler — must be void
private async void ReportButton_Click(object sender, RibbonControlEventArgs e)
{
    try
    {
        await ReportPhishingEmailAsync();
    }
    catch (Exception ex)
    {
        // Handle here — this IS reached because the try/catch wraps the await
        HandleError(ex);
    }
}

// Actual async work — Task-returning, testable, exceptions propagate correctly
private async Task ReportPhishingEmailAsync()
{
    // ... async work here
}
```

Never put the `await` call outside a try/catch in an `async void` handler. Also register `AppDomain.CurrentDomain.UnhandledException` as a last-resort logger at add-in startup.

**Warning signs:**
- Outlook crashes without error dialog after clicking the report button
- Event ID 1000 in Windows Application Event Log pointing at Outlook with `mscoree.dll` as faulting module
- Error email never arrives when a failure should have been caught
- Stack trace missing or shows `AggregateException` with no useful inner exception

**Phase to address:** Phase converting synchronous HTTP calls to async

---

### Pitfall 4: The `LoadBehavior` HKCU Override Silently Defeats MSI Deployments

**What goes wrong:**
The MSI installer writes `LoadBehavior=3` to `HKEY_LOCAL_MACHINE`. An enterprise user's machine where the add-in was previously auto-disabled has `LoadBehavior=2` (load on startup, failed) or `LoadBehavior=0` (do not load) written to `HKEY_CURRENT_USER`. The `HKCU` key silently wins over `HKLM`. After deploying the updated MSI to fix reliability issues, the add-in still does not load for the 50% of users whose per-user registry is in a disabled state. The support team cannot reproduce the issue because test machines have clean `HKCU` keys.

**Why it happens:**
VSTO's registry architecture has `HKCU` taking precedence over `HKLM` by design. MSI packages typically write only to `HKLM` for machine-wide installs. When Outlook auto-disables the add-in, it writes the disabled `LoadBehavior` to `HKCU` — and MSI upgrades do not touch `HKCU` user-specific keys, leaving the stale disabled state intact.

**How to avoid:**
The updated MSI installer must include a Custom Action that resets the `HKCU\SOFTWARE\Microsoft\Office\<version>\Outlook\Addins\<ProgID>\LoadBehavior` to `3` during installation. This must run for each Office version (16.0 for Outlook 2016/2019/365) and must clear `HKCU\SOFTWARE\Microsoft\Office\<version>\Outlook\Resiliency\DisabledItems` entries for the add-in. Additionally, deploy a GPO entry under `HKCU\Software\Policies\Microsoft\Office\16.0\Outlook\Resiliency\AddinList` with the ProgID set to `1` to prevent future auto-disabling by policy.

Do not rely solely on HKLM writes — test the upgrade path from a machine that has a previously-disabled state, not just a clean install.

**Warning signs:**
- MSI deployment reports success but 50% of users still report missing add-in
- Add-in visible in COM Add-ins dialog as unchecked (disabled)
- Re-enabling by hand works but reverts on next Outlook restart
- Test machines work fine; production machines do not (test machines have clean registry)

**Phase to address:** Phase addressing enterprise deployment

---

### Pitfall 5: Blocking `.Result` or `.Wait()` on an Async Task from the UI Thread

**What goes wrong:**
During async conversion, developers may wrap an async call synchronously to avoid changing callers. Calling `.Result` or `.Wait()` on a `Task` from the STA UI thread causes a deadlock: the task's continuation needs to marshal back to the STA context, but the STA context is blocked waiting for the task to complete. Outlook freezes permanently rather than temporarily. This is worse than the original synchronous HTTP call because the original at least eventually returned; the deadlock never resolves without killing the process.

**Why it happens:**
Partial async conversions are common when refactoring incrementally. A developer makes `GoPhishIntegration.sendReportNotificationToServer()` return a `Task` but the caller still expects a synchronous string result. The quick fix is `.Result` — which compiles and appears to work in unit tests (which run on MTA threadpool threads) but deadlocks in production (which runs on the STA main thread).

**How to avoid:**
The async conversion must be complete end-to-end ("async all the way"). There is no safe way to call `.Wait()` or `.Result` on a Task from a VSTO STA event handler. If a method returns `Task`, every caller must `await` it, and every caller must itself be `async`. The refactoring is a vertical slice through the call stack.

Additionally, use `ConfigureAwait(false)` on all `await` calls inside library and service methods (the `GoPhishIntegration` class) to avoid needlessly marshaling continuations back to the STA context.

**Warning signs:**
- Outlook freezes permanently (requires Task Manager kill) as opposed to freezing briefly
- Deadlock is 100% reproducible on the main thread but never in unit tests
- Thread dump shows main thread blocked on `Task.Result` or `Monitor.Wait`
- Conversion compiled and passed CI but fails in Outlook immediately

**Phase to address:** Phase converting synchronous HTTP calls to async

---

### Pitfall 6: 32-bit vs. 64-bit Office Mismatch in MSI Deployment

**What goes wrong:**
The add-in MSI targets one bitness. Enterprise environments with mixed 32-bit and 64-bit Office installs have the add-in fail to load on the mismatched machines. 64-bit Office reads the 64-bit registry hive for add-in registration; 32-bit Office reads the 32-bit (Wow6432Node) hive. A single-bitness MSI registers in only one hive. The add-in is silently absent for half the user base with no error visible to the user.

**Why it happens:**
Development machines typically have a single Office bitness. The mixed-bitness problem only surfaces in large enterprise deployments. The omission is invisible because the COM registration appears to succeed and the add-in works on the developer's machine.

**How to avoid:**
For an enterprise-wide MSI deployment targeting mixed Office versions, either create two separate MSI packages (one for 32-bit Office, one for 64-bit Office) with WMI filtering in SCCM/GPO to deploy the correct one, or create a single MSI that registers in both hives. The build target bitness of the VSTO project must match the Office bitness — a 32-bit add-in will not load in 64-bit Office even if the registry entry exists.

Check the actual distribution of Office bitness in the environment before assuming. Modern enterprise M365 installs are predominantly 64-bit; legacy Office 2016 licenses have more 32-bit deployments.

**Warning signs:**
- Add-in absent from COM Add-ins list on some machines but not others
- Pattern correlates with machine age (older hardware more likely to have 32-bit Office)
- SCCM deployment shows success but user complaints persist
- Registry key exists under `HKLM\SOFTWARE\Microsoft\Office\Outlook\Addins` but not under `HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\Outlook\Addins` (or vice versa)

**Phase to address:** Phase addressing enterprise deployment

---

### Pitfall 7: Changing the Assembly Name or ProgID After Initial Deployment

**What goes wrong:**
A refactoring renames the assembly or changes the ProgID (e.g., from `PhishingReporter` to `PhishingReporter.Connect`). The new MSI registers a new ProgID while the old ProgID entry remains in the registry from the previous install. Outlook sees two add-in registrations and attempts to load both. The old one fails to load (DLL is gone), triggering a crash or load failure that disables the new one via guilt-by-association in Outlook's resiliency tracking. Alternatively, the new ProgID is not in the `DoNotDisableAddinList` GPO, so it gets auto-disabled immediately despite the old ProgID having exemption.

**Why it happens:**
Refactoring naturally produces better names. Developers do not think of assembly name and ProgID as deployment-permanent identifiers. The problem only manifests in production on machines that have an existing installation, not on test machines with clean installs.

**How to avoid:**
The ProgID and assembly name are deployment-permanent once the first version ships to users. Do not change them. If a rename is unavoidable, the MSI upgrade must explicitly remove all old registry entries for the old ProgID as a Custom Action before registering the new ProgID. Update all GPO entries for the new ProgID before deploying.

**Warning signs:**
- Both old and new add-in names visible in COM Add-ins dialog
- Load failure for the old add-in causes Outlook to start the resiliency tracking clock for both
- Machines with prior install behave differently from machines with fresh install

**Phase to address:** Phase addressing enterprise deployment

---

## Technical Debt Patterns

| Shortcut | Immediate Benefit | Long-term Cost | When Acceptable |
|----------|-------------------|----------------|-----------------|
| GPO `AddinList=1` override without fixing startup speed | Add-in stays enabled without code changes | Hides the performance problem; add-in still slow; crashes still cause disable | Never as a standalone fix; acceptable as defense-in-depth after startup is fixed |
| `async void` on event handler with all logic inside it | Minimal refactoring effort | Unhandled exceptions crash Outlook silently | Never — always extract to `async Task` method |
| `.Result` / `.Wait()` on async call from UI thread | Avoids propagating async up the call stack | Guaranteed deadlock on STA thread | Never in VSTO event handlers |
| Writing only to HKLM in MSI | Simpler installer | HKCU disabled state from prior install survives upgrade | Never if users may have had previous disabled state |
| Single-bitness MSI | Simpler build | Fails silently on mismatched Office bitness | Never in mixed-bitness enterprise environments |
| Synchronous HTTP call wrapped in try/catch | Simple, readable | Freezes Outlook UI; Outlook counts this time against resiliency threshold | Never — even if GoPhish is internal/fast |

---

## Integration Gotchas

| Integration | Common Mistake | Correct Approach |
|-------------|----------------|------------------|
| GoPhish HTTP call | Calling `GetResponse()` synchronously on UI thread | Use `HttpClient.PostAsync()` with `await`; read all OOM data before entering async context |
| GoPhish HTTP call | Passing Outlook `MailItem` object into `Task.Run()` lambda | Extract all needed values (string, byte[]) from the MailItem on the main thread before calling `Task.Run()` |
| Exchange/AD user lookup via `GetExchangeUser()` | Calling from background thread after async handoff | Perform `GetExchangeUser()` synchronously on main thread before `await`; pass the result struct into async method |
| `Properties.Settings.Default.Save()` | Never calling `Save()` after incrementing counters | Call `Save()` immediately after every counter increment; settings changes are in-memory only until `Save()` |
| `HttpWebRequest` / `HttpWebResponse` | Not disposing response stream | Replace with `HttpClient` which handles disposal; or wrap in `using` blocks if staying on `HttpWebRequest` |

---

## Performance Traps

| Trap | Symptoms | Prevention | When It Breaks |
|------|----------|------------|----------------|
| Heavy `ThisAddIn_Startup` work | Add-in disabled after 5 Outlook launches; Event ID 59 | Defer all non-essential startup work to `Application.Startup` event or a timer | Any time median over 5 starts exceeds 1000ms |
| WPF UserControl in TaskPane initialized at startup | 2+ second startup time even on fast hardware | Initialize TaskPane lazily in a timer callback after Outlook startup completes | Always — WPF initialization alone can exceed 1000ms on cold start |
| Synchronous regex on full email headers in `setReportURL()` | Noticeable delay on forwarded emails with large headers | Compile `Regex` as a static readonly field; limit input length before matching | Emails with 3+ forward chains (~10KB+ headers) |
| String `+=` concatenation in report-building loops | Lag building report for emails with 50+ URLs | Replace with `StringBuilder`; use `string.Join()` for URL lists | Emails with 50+ URLs or attachments |
| Loading entire attachment into memory for hashing | OOM on systems with limited RAM when users forward large attachments | Already uses `FileStream` streaming — adequate; add a max-size guard with user warning | Attachments over ~500MB on machines with 4GB RAM |

---

## Security Mistakes

| Mistake | Risk | Prevention |
|---------|------|------------|
| Hardcoded TLS 1.2 via unsafe enum cast | When TLS 1.2 is deprecated or restricted by policy, add-in breaks network calls with no upgrade path | Switch to `HttpClient` which negotiates protocol version automatically; or make the security protocol version a configurable setting |
| No certificate validation configured for GoPhish HTTPS | Default behavior depends on Windows certificate store; corporate proxy MitM goes undetected | Explicitly document certificate requirements; add logging of certificate subject on connection |
| MSI write to HKCU with credentials or API keys | Extractable from registry by any user on the machine | Never store secrets in registry or Settings.settings; use Windows DPAPI (`ProtectedData.Protect()`) if storage is unavoidable |
| GPO `AddinList=1` applied globally without policy scope | Forces add-in enabled even when crashing, removing Outlook's crash protection | Scope the GPO to the add-in's OU; test crash recovery behavior before deploying force-enable policy |

---

## UX Pitfalls

| Pitfall | User Impact | Better Approach |
|---------|-------------|-----------------|
| Outlook freezes for 1-5 seconds during GoPhish HTTP call | User thinks Outlook or the add-in has crashed; may force-close Outlook | Show a brief "Reporting..." status in the ribbon label during async operation; restore when done |
| No feedback after reporting succeeds | User re-clicks the report button thinking the first click did nothing | Show a success confirmation (notification or dialog) that clears after 3 seconds |
| Error email goes to support silently with no user notice | User does not know the report failed; phishing email goes unreported | Show an error dialog telling the user the report failed and asking them to retry or contact support |
| Add-in disabled dialog from Outlook with no context | User clicks "Disable" to dismiss — permanently disabling the add-in | The GPO override prevents this dialog; but users who disabled it manually are already lost without admin remediation |

---

## "Looks Done But Isn't" Checklist

- [ ] **Async conversion:** The GoPhish HTTP call is `await`-able, BUT all Outlook Object Model access (MailItem properties, ExchangeUser) still occurs on the main thread BEFORE the `await` — verify no OOM objects are captured in lambda closures passed to `Task.Run()`
- [ ] **Startup optimization:** Work moved out of `ThisAddIn_Startup`, BUT verify it was moved to `Application.Startup` event (fires after all add-ins loaded, not timed) and not to a background thread that still calls OOM
- [ ] **MSI deployment:** Registry entries written, BUT verify both HKLM and HKCU `LoadBehavior` are set to 3 AND that `DisabledItems` entries for the ProgID are cleared as a Custom Action during upgrade
- [ ] **GPO override deployed:** `AddinList` entry set to 1 in policy, BUT verify it is present for EVERY Outlook version in the environment (15.0 for 2013, 16.0 for 2016/2019/365 — they all use 16.0 registry key)
- [ ] **Mixed bitness:** MSI deploys to 32-bit and 64-bit Office correctly — verify by checking both `HKLM\SOFTWARE\Microsoft\Office\Outlook\Addins` and `HKLM\SOFTWARE\Wow6432Node\Microsoft\Office\Outlook\Addins` on a test machine
- [ ] **Exception handling in async:** `try/catch` exists inside the `async void` event handler wrapping the `await` of the `Task`-returning method — unhandled exceptions in `async void` crash Outlook without calling the existing error-report mechanism

---

## Recovery Strategies

| Pitfall | Recovery Cost | Recovery Steps |
|---------|---------------|----------------|
| Add-in disabled on user machines due to HKCU LoadBehavior | MEDIUM | Deploy a remediation script via SCCM that sets `HKCU\...\LoadBehavior=3` and clears `DisabledItems`; or re-run updated MSI with Custom Action |
| Deadlock from `.Result` on UI thread | LOW (code fix) | Remove `.Result`/`.Wait()` entirely; propagate `async` up the call stack; test on STA thread explicitly |
| `async void` crash taking down Outlook | LOW (code fix) | Extract async logic to `Task`-returning method; wrap `await` in try/catch in the `async void` handler |
| Wrong bitness MSI deployed | MEDIUM | Identify affected machines via SCCM inventory; deploy correct-bitness MSI as a remediation package |
| Assembly name / ProgID changed mid-deployment | HIGH | Create a cleanup MSI that removes all old ProgID registry entries; deploy before new version MSI; update GPO entries |
| Startup time exceeds 1000ms despite fixes | LOW | Profile with Stopwatch in Event Log; most likely culprit is certificate revocation check on HTTPS connection — disable by setting `ServicePointManager.CheckCertificateRevocationList = false` for internal GoPhish servers, or move to `HttpClient` with explicit handler |

---

## Pitfall-to-Phase Mapping

| Pitfall | Prevention Phase | Verification |
|---------|------------------|--------------|
| Registry override without fixing root cause startup speed | Phase: Startup reliability fix | Measure Event ID 45 load time < 1000ms on 5 consecutive Outlook starts after fix |
| Background thread calling Outlook OOM | Phase: Async HTTP conversion | Code review: grep for any Outlook Interop type accessed inside `Task.Run()` or after `await` on non-main thread |
| `async void` with unhandled exceptions | Phase: Async HTTP conversion | Unit test that simulates GoPhish server failure during async call; verify error is caught and routed to error handler |
| HKCU `LoadBehavior` not reset on upgrade | Phase: Enterprise deployment | Test upgrade path from machine with previously-disabled add-in, not just clean install |
| `.Result`/`.Wait()` deadlock | Phase: Async HTTP conversion | Integration test running the full report flow on an STA thread (use `[STAThread]` test harness) |
| 32-bit vs. 64-bit MSI mismatch | Phase: Enterprise deployment | Verify add-in loads on both 32-bit and 64-bit Office machines in test environment |
| Assembly name / ProgID change | Phase: Enterprise deployment | Verify no dual-registration occurs on machines upgraded from previous install |
| `async void` crash bypassing error handler | Phase: Async HTTP conversion | Confirm `AppDomain.CurrentDomain.UnhandledException` is registered at add-in startup for last-resort logging |

---

## Sources

- [Microsoft Learn: Support for keeping add-ins enabled](https://learn.microsoft.com/en-us/office/vba/outlook/concepts/getting-started/support-for-keeping-add-ins-enabled) — HIGH confidence, official docs, defines resiliency disable codes and registry keys
- [Microsoft Learn: Add-ins are user re-enabled after being disabled](https://learn.microsoft.com/en-us/troubleshoot/outlook/performance/add-ins-are-user-re-enabled-after-being-disabled) — HIGH confidence, official support article, defines Event ID 45/59, threshold values (500ms shutdown, 1000ms startup)
- [Microsoft Learn: Threading support in Office (VSTO)](https://learn.microsoft.com/en-us/visualstudio/vsto/threading-support-in-office?view=vs-2022) — HIGH confidence, official docs, defines STA model, `COMException` on background thread calls, `E_RPC_WRONG_THREAD`
- [Microsoft Learn: Registry entries for VSTO Add-ins](https://learn.microsoft.com/en-us/visualstudio/vsto/registry-entries-for-vsto-add-ins?view=vs-2022) — HIGH confidence, official docs, defines `LoadBehavior` values and HKLM vs HKCU precedence
- [Microsoft Learn: Deploy a VSTO Solution with Windows Installer](https://learn.microsoft.com/en-us/visualstudio/vsto/deploying-a-vsto-solution-by-using-windows-installer?view=vs-2022) — HIGH confidence, official docs, assembly name change warning
- [Microsoft Learn: Office primary interop assemblies](https://learn.microsoft.com/en-us/visualstudio/vsto/office-primary-interop-assemblies?view=vs-2022) — HIGH confidence, official docs, PIA version binding and GAC pitfalls
- [Microsoft Learn: VSTO Outlook addin — how to update UI from async process](https://learn.microsoft.com/en-us/answers/questions/78894/vsto-outlook-addin-how-to-update-ui-from-asynchron) — MEDIUM confidence, official Q&A forum
- [Microsoft Learn: VSTO Outlook: Improve and accelerate Add-in startup](https://learn.microsoft.com/en-us/answers/questions/1056423/vsto-outlook-improve-and-accelerate-add-in-startup) — MEDIUM confidence, official Q&A forum, startup optimization strategies
- [Delay Loading Outlook Add-ins (theofficecontext.com)](https://theofficecontext.com/2017/06/23/delay-loading-outlook-add-ins/) — MEDIUM confidence, community post from Office MVP, timer-based deferred initialization pattern
- [Outlook's slow add-ins resiliency logic (osict.com PDF mirror)](https://www.osict.com/ufc/file2/osict_sites/michel/39952e7597d8d18608299fa0fd5e55f0/pu/Outlook_resiliency_v1_.pdf) — MEDIUM confidence, mirrors Developer Messaging blog content on 1000ms threshold and median over 5 iterations
- [Add-in Express: Async/await with COM plugins](https://www.add-in-express.com/forum/read.php?FID=5&TID=15848) — MEDIUM confidence, community forum with practical VSTO async experience
- [Add-in Express: Threading in managed Office extensions](https://www.add-in-express.com/creating-addins-blog/2010/11/04/threads-managed-office-extensions/) — MEDIUM confidence, widely-cited community resource on OOM threading rules
- [Enterprise Craftsmanship: Async/await pitfalls](https://enterprisecraftsmanship.com/posts/pitfalls-of-async-await/) — MEDIUM confidence, general C# async pitfalls applicable to VSTO

---
*Pitfalls research for: VSTO Outlook Add-in Reliability (Enterprise)*
*Researched: 2026-02-25*
