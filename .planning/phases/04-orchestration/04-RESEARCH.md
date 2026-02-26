# Phase 4: Orchestration - Research

**Researched:** 2026-02-26
**Domain:** VSTO async orchestration, COM threading isolation, immutable data extraction, single-responsibility refactoring
**Confidence:** HIGH

## Summary

Phase 4 completes the architectural transformation that Phases 2 and 3 started. Phase 2 extracted single-responsibility components (UrlExtractor, AttachmentHasher, GoPhishIntegration) from Ribbon.cs. Phase 3 made the GoPhish HTTP call async and wired the ribbon callback as `async void`. But Ribbon.cs still contains 380+ lines of email processing logic -- OOM data extraction, report body composition, error email construction, item type detection, and user information gathering -- all interleaved with COM object lifecycle management. The goal of Phase 4 is to extract ALL of this into two new classes: an immutable `EmailReport` data object that captures every piece of Outlook OOM data on the UI thread before any await boundary, and a `ReportOrchestrator` that composes the async workflow (GoPhish detection, report dispatch, error handling) using only the plain C# data from `EmailReport` -- never touching OOM.

The critical technical constraint is the VSTO STA threading model. After `await` with `ConfigureAwait(false)` in the GoPhish branch, execution resumes on a thread pool thread. Any Outlook OOM access from that thread throws COMException 0x8001010E. The current code already handles this for the GoPhish branch (mailItem.Delete() was moved before the await in Phase 3), but the architecture is fragile -- the safety depends on control-flow analysis of which branch contains an await. Phase 4 makes the safety structural: ALL OOM access happens in a single extraction step that produces an immutable DTO, and everything after that step operates on plain C# objects.

The second constraint is the C# language version. .NET Framework 4.8 uses C# 7.3 by default. C# records (C# 9.0) and init-only properties (C# 9.0) are NOT available. The immutable `EmailReport` must use the traditional pattern: a `sealed class` with `readonly` backing fields and get-only properties, initialized via constructor.

**Primary recommendation:** Create two new files: `EmailReport.cs` (immutable sealed class capturing all OOM-extracted data) and `ReportOrchestrator.cs` (async Task method that accepts EmailReport and Outlook.Application, performs GoPhish detection + notification, composes and sends the report email, and handles errors). Ribbon.cs shrinks to: confirm dialog, OOM data extraction into EmailReport, delegate to ReportOrchestrator, COM cleanup. All OOM access in Ribbon.cs occurs before the single `await` call to the orchestrator.

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|-----------------|
| QUAL-01 | Email processing logic extracted from Ribbon.cs into dedicated EmailProcessor class | EmailReport sealed class captures all OOM data (headers, URLs, hashes, sender info, basic info); ReportOrchestrator handles all workflow logic (GoPhish detection, report composition, email sending). Ribbon.cs reduced to UI callbacks + OOM extraction + delegation. The "EmailProcessor" name from requirements maps to the combination of EmailReport (data) + ReportOrchestrator (behavior). |
| QUAL-02 | GoPhish integration refactored with async HttpClient and proper result types | GoPhishIntegration already has async HttpClient + GoPhishResult enum (from Phases 2-3). Phase 4 wires the orchestrator to call `GoPhishIntegration.SendReportNotificationAsync` and `GoPhishIntegration.setReportURL` using data from the immutable EmailReport, completing the async pipeline end-to-end with no OOM access after the await boundary. |
</phase_requirements>

## Standard Stack

### Core
| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| .NET Framework | 4.8 | Runtime (C# 7.3 language version) | Already targeted; immutable DTO uses sealed class + get-only properties (no records) |
| Microsoft.Office.Interop.Outlook | 15.0 (PIA) | OOM access for data extraction | Already referenced; EmailReport extraction reads OOM properties on UI thread |
| System.Threading.Tasks | .NET Fx 4.8 built-in | Async Task for ReportOrchestrator | Already used in Phase 3; orchestrator returns Task |

### Supporting
| Library | Version | Purpose | When to Use |
|---------|---------|---------|-------------|
| System.Collections.ObjectModel | .NET Fx built-in | ReadOnlyCollection for immutable lists in EmailReport | Wrap extracted URL/hash lists |
| NLog | 5.4.0 | Logging in new classes | Already referenced; ReportOrchestrator and EmailReport follow AppLogger pattern |

### Alternatives Considered
| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| Sealed class with get-only properties | `readonly struct` (C# 7.2+) | Struct would cause boxing when passed as method argument; EmailReport has many fields (large struct penalty); sealed class is clearer and matches existing UrlExtractionResult/AttachmentHashResult pattern in codebase |
| Sealed class with get-only properties | C# 9 record (via LangVersion override) | Forcing LangVersion=9.0 on .NET Fx 4.8 is unsupported by Microsoft; risks compiler errors with missing runtime types; not worth the risk for a simple DTO |
| Single ReportOrchestrator class | Separate classes per workflow step | Over-engineering for current scope; ReportOrchestrator is ~150 lines; can split later if needed |
| Outlook.Application passed to orchestrator | Passing only primitive data | Some OOM operations (CreateItem, MailItem.Send) MUST occur on UI thread; orchestrator needs Application reference for report email creation, but only accesses it from the UI-thread caller context |

**Installation:** No new packages needed. Phase 4 uses only what previous phases already installed.

## Architecture Patterns

### Recommended Project Structure After Phase 4
```
PhishingReporter/
  ThisAddIn.cs              # Unchanged
  Ribbon.cs                 # THIN: confirm dialog + OOM extraction + delegate to orchestrator
  EmailReport.cs            # NEW: immutable sealed class holding all extracted email data
  ReportOrchestrator.cs     # NEW: async workflow (GoPhish check, report compose, send, errors)
  GoPhishIntegration.cs     # Unchanged from Phase 3
  UrlExtractor.cs           # Unchanged from Phase 2
  AttachmentHasher.cs       # Unchanged from Phase 2
  AppLogger.cs              # Unchanged from Phase 1
  NLog.config               # Unchanged
  Properties/               # Unchanged
  PhishingReporter.csproj   # Updated: new .cs file entries
```

### Pattern 1: Immutable EmailReport DTO (C# 7.3 Compatible)
**What:** A sealed class with constructor-initialized get-only properties that captures ALL data extracted from Outlook OOM objects. Once constructed, the object is completely immutable and safe to pass across thread boundaries.
**When to use:** As the boundary between UI-thread OOM access and background-thread async processing.
**Why sealed class, not readonly struct:** EmailReport has 10+ fields including collections (URLs, domains, attachment hashes). A struct that large would be copied on every method call, causing unnecessary allocation. A sealed class is heap-allocated once and passed by reference. This also matches the existing immutable DTO pattern in the codebase (UrlExtractionResult, AttachmentHashResult).

**Example:**
```csharp
// Source: .NET Framework 4.8 / C# 7.3 immutable class pattern
// https://learn.microsoft.com/en-us/dotnet/csharp/language-reference/keywords/readonly
namespace PhishingReporter
{
    /// <summary>
    /// Immutable snapshot of all data extracted from an Outlook email item.
    /// Constructed on the UI thread from OOM objects; safe to use on any thread.
    /// </summary>
    internal sealed class EmailReport
    {
        // Item identity
        public string ItemType { get; }        // "MailItem", "MeetingItem", etc.
        public bool IsMailItem { get; }

        // Mail-specific data (null if not a MailItem)
        public string Subject { get; }
        public string Headers { get; }
        public string HtmlBody { get; }
        public string FolderPath { get; }

        // GoPhish detection (pre-computed on UI thread)
        public string GoPhishReportUrl { get; }  // null if not a GoPhish campaign

        // Extracted analysis (pre-computed on UI thread)
        public UrlExtractionResult UrlAnalysis { get; }
        public IReadOnlyList<AttachmentHashResult> AttachmentHashes { get; }

        // User/environment info (extracted on UI thread)
        public string UserInfoSection { get; }
        public string BasicInfoSection { get; }
        public string PluginDetailsSection { get; }

        // Counters (read on UI thread for thread safety)
        public int SuspiciousReportsCount { get; }
        public int GoPhishReportsCount { get; }

        public EmailReport(
            string itemType,
            bool isMailItem,
            string subject,
            string headers,
            string htmlBody,
            string folderPath,
            string goPhishReportUrl,
            UrlExtractionResult urlAnalysis,
            IReadOnlyList<AttachmentHashResult> attachmentHashes,
            string userInfoSection,
            string basicInfoSection,
            string pluginDetailsSection,
            int suspiciousReportsCount,
            int goPhishReportsCount)
        {
            ItemType = itemType;
            IsMailItem = isMailItem;
            Subject = subject;
            Headers = headers;
            HtmlBody = htmlBody;
            FolderPath = folderPath;
            GoPhishReportUrl = goPhishReportUrl;
            UrlAnalysis = urlAnalysis;
            AttachmentHashes = attachmentHashes;
            UserInfoSection = userInfoSection;
            BasicInfoSection = basicInfoSection;
            PluginDetailsSection = pluginDetailsSection;
            SuspiciousReportsCount = suspiciousReportsCount;
            GoPhishReportsCount = goPhishReportsCount;
        }
    }
}
```

**Design decision -- pre-compute vs raw data:** The EmailReport stores pre-computed string sections (UserInfoSection, BasicInfoSection) rather than raw OOM property values (ExchangeUser.Name, ExchangeUser.PrimarySmtpAddress, etc.). This is because:
1. The raw OOM properties require COM objects (ExchangeUser, MAPIFolder) that cannot be stored safely (they are RCW-wrapped COM references).
2. The formatting logic (string concatenation for report body) is simple and deterministic -- no benefit to deferring it.
3. The existing methods `GetCurrentUserInfos()`, `GetBasicInfo()`, `GetPluginDetails()` already produce the formatted strings. Reusing their output minimizes code changes.

However, the URL analysis and attachment hashes are stored as structured objects (UrlExtractionResult, IReadOnlyList<AttachmentHashResult>) because the orchestrator needs to format them into the report body, and these objects are already immutable plain-C# types with no COM references.

**Confidence:** HIGH -- immutable sealed class is a standard C# 7.3 pattern; aligns with existing codebase DTOs.

### Pattern 2: ReportOrchestrator Async Workflow
**What:** A static class with a single public async Task method that accepts an EmailReport and an Outlook.Application reference, then executes the full reporting workflow: GoPhish notification (async), report email composition (sync, using Application.CreateItem), email sending, counter persistence, and user feedback.
**When to use:** Called from Ribbon.cs after OOM data extraction.

**Critical threading insight:** The orchestrator needs `Outlook.Application.CreateItem()` to create the report MailItem and `.Send()` to dispatch it. These are OOM calls that must happen on the STA thread. In the current code flow:
- The GoPhish branch: `await` occurs, so post-await code runs on thread pool. But post-await, the only OOM needed is creating/sending report email -- which in the GoPhish branch is NOT done (no email report, just GoPhish notification).
- The non-GoPhish branch: NO `await` occurs before report email creation/send. All OOM access is on the UI thread.

This means the orchestrator's architecture must ensure that `Application.CreateItem()` and `MailItem.Send()` are only called when no await has preceded them in the current execution path. The simplest correct design: the orchestrator returns to the caller (Ribbon.cs) with instructions about what happened, and Ribbon.cs performs the final OOM operations (create report email, send it) on the UI thread. BUT this would keep email composition logic in Ribbon.cs, violating QUAL-01.

**Alternative design (recommended):** Split the orchestrator into two phases:
1. **Phase A (before await):** GoPhish detection (pure string parsing, already computed in EmailReport), report email creation + composition + send (all OOM, all on UI thread), OR
2. **Phase B (after await):** GoPhish async notification only (no OOM access needed).

In practice, the orchestrator method is called ON the UI thread. The non-GoPhish path executes entirely synchronously (no await hit), keeping all OOM access on the UI thread. The GoPhish path: create/send is NOT needed (just the async notification), and mailItem.Delete() was already done in Ribbon.cs before calling the orchestrator.

Wait -- re-analyzing the actual code flow. Let me be precise about what each branch does:

**GoPhish branch (simulatedPhishingURL != null):**
1. Delete the reported email (OOM) -- must be on UI thread
2. Send GoPhish notification (async HTTP) -- must be async, releases UI thread
3. Increment gophish_reports_counter + Save() -- safe from any thread
4. Show success MessageBox -- safe from any thread

**Non-GoPhish branch (simulatedPhishingURL == null):**
1. Increment suspecious_reports_counter + Save() -- safe from any thread
2. Compose report email body (uses UserInfo, BasicInfo, URLs, Hashes, Headers, PluginDetails) -- pure string building, no OOM IF data is pre-extracted
3. Create report MailItem, set To/Subject/Body, attach original, Save, Send -- OOM access
4. Delete the reported email -- OOM access

The key realization: if ALL the data is pre-extracted into EmailReport, then in the non-GoPhish branch, the report body composition is pure string concatenation using EmailReport fields -- no OOM. But `Application.CreateItem()`, `reportEmail.Send()`, and `mailItem.Delete()` are still OOM calls.

**The non-GoPhish branch never hits an await**, so it executes entirely on the UI thread. The OOM calls in this branch are safe.

**The GoPhish branch hits an await** (SendReportNotificationAsync), so post-await code runs on a thread pool. But post-await, only Settings and MessageBox are accessed -- both safe from any thread.

**Therefore:** The orchestrator CAN contain all the logic as long as:
1. It is called from the UI thread
2. All OOM access occurs before the first `await` in any execution path
3. The GoPhish branch's await is the ONLY await, and all code after it avoids OOM

This is exactly the current pattern, just moved to a new class. The orchestrator receives an `Outlook.Application` reference to create the report email, and it receives the original item (or a reference to it) for deletion. The original MailItem must be deleted on the UI thread -- in the GoPhish branch this is done before the await, in the non-GoPhish branch there is no await.

**Example:**
```csharp
// ReportOrchestrator.cs
namespace PhishingReporter
{
    internal static class ReportOrchestrator
    {
        private static readonly NLog.Logger Logger =
            AppLogger.Instance.GetCurrentClassLogger();

        /// <summary>
        /// Executes the full report workflow using pre-extracted data.
        /// MUST be called from the UI thread. All OOM access occurs
        /// before any await boundary.
        /// </summary>
        public static async Task ExecuteAsync(
            EmailReport report,
            Outlook.Application application,
            object originalItem,
            Outlook.MailItem mailItem)
        {
            if (report.GoPhishReportUrl != null)
            {
                // GoPhish branch -- delete BEFORE await (UI thread)
                mailItem.Delete();
                Logger.Info("Reported email deleted from mailbox");

                // AWAIT BOUNDARY -- after this, code runs on thread pool
                GoPhishResult result = await GoPhishIntegration
                    .SendReportNotificationAsync(report.GoPhishReportUrl)
                    .ConfigureAwait(false);
                Logger.Info("GoPhish notification result: {0}", result);

                // Safe from any thread
                Properties.Settings.Default.gophish_reports_counter++;
                Properties.Settings.Default.Save();
                MessageBox.Show("Good job!...", "We have a winner!");
            }
            else
            {
                // Non-GoPhish branch -- NO await, stays on UI thread
                Properties.Settings.Default.suspecious_reports_counter++;
                Properties.Settings.Default.Save();

                // Compose report body from pre-extracted data (pure strings)
                string body = ComposeReportBody(report);

                // OOM access -- safe, still on UI thread (no await in this path)
                Outlook.MailItem reportEmail = null;
                try
                {
                    reportEmail = (Outlook.MailItem)application
                        .CreateItem(Outlook.OlItemType.olMailItem);
                    reportEmail.To = Properties.Settings.Default.infosec_email;
                    reportEmail.Subject = "[POTENTIAL PHISH] " + report.Subject;
                    reportEmail.Attachments.Add(originalItem);
                    reportEmail.Body = body;
                    reportEmail.Save();
                    reportEmail.Send();
                    Logger.Info("Report email sent to: {0}",
                        Properties.Settings.Default.infosec_email);

                    mailItem.Delete();
                    Logger.Info("Reported email deleted from mailbox");
                }
                finally
                {
                    if (reportEmail != null)
                    {
                        try { Marshal.ReleaseComObject(reportEmail); }
                        catch { }
                        reportEmail = null;
                    }
                }
            }
        }

        private static string ComposeReportBody(EmailReport report)
        {
            // Pure string building -- no OOM access
            // Uses report.UserInfoSection, report.BasicInfoSection, etc.
            // ... (format the report body from pre-extracted data) ...
        }
    }
}
```

**Confidence:** HIGH -- follows established OOM-before-await pattern from Phase 3; no new threading concepts introduced.

### Pattern 3: Thin Ribbon.cs Callback Pattern
**What:** After Phase 4, the Ribbon.cs `reportPhishingEmailToSecurityTeamAsync` method shrinks to: validate selection, extract OOM data into EmailReport, delegate to ReportOrchestrator, cleanup COM objects.
**When to use:** The ribbon callback is the only place where OOM data is extracted.

**Example (conceptual):**
```csharp
private async Task reportPhishingEmailToSecurityTeamAsync(IRibbonControl control)
{
    Selection selection = null;
    MailItem mailItem = null;

    try
    {
        selection = Globals.ThisAddIn.Application.ActiveExplorer().Selection;

        // ... validate selection count and item type ...

        mailItem = selection[1] as MailItem;

        // Extract ALL OOM data into immutable object (UI thread)
        EmailReport report = ExtractEmailReport(selection[1], mailItem);

        // Delegate to orchestrator
        await ReportOrchestrator.ExecuteAsync(
            report,
            Globals.ThisAddIn.Application,
            selection[1],
            mailItem).ConfigureAwait(false);
    }
    catch (Exception ex)
    {
        // Error handling (error email, user notification)
    }
    finally
    {
        // COM cleanup
        if (mailItem != null) { try { Marshal.ReleaseComObject(mailItem); } catch { } }
        if (selection != null) { try { Marshal.ReleaseComObject(selection); } catch { } }
    }
}

private EmailReport ExtractEmailReport(object selectedItem, MailItem mailItem)
{
    // All OOM access happens here, on the UI thread
    // Calls GetCurrentUserInfos(), GetBasicInfo(), GetURLsAndAttachmentsInfo()
    // Returns immutable EmailReport
}
```

**Confidence:** HIGH -- direct simplification of existing code structure.

### Pattern 4: OOM Data Extraction Before Await Boundary
**What:** The pattern of extracting ALL Outlook Object Model data into plain C# objects on the UI thread, before any code path that contains an `await`, ensuring COM objects are never accessed from thread pool threads.
**When to use:** At every entry point where async operations follow OOM access. In this codebase, only in Ribbon.cs.

**The boundary rule:** Draw a line in the code at the first possible `await`. Everything above the line can access OOM. Everything below the line (including code inside the awaited method) must work only with plain C# objects.

```
UI THREAD (before await)          |  THREAD POOL (after await)
----------------------------------|----------------------------------
selection = ActiveExplorer()      |  GoPhish HTTP notification
mailItem = selection[1]           |  Settings counter increment
headers = mailItem.HeaderString() |  MessageBox.Show (auto-marshals)
htmlBody = mailItem.HTMLBody      |
urls = UrlExtractor.ExtractUrls() |
hashes = AttachmentHasher.Compute |
userInfo = GetCurrentUserInfos()  |
basicInfo = GetBasicInfo()        |
report = new EmailReport(...)     |
mailItem.Delete() [GoPhish only]  |
                                  |
--- await boundary ----           |
```

**Confidence:** HIGH -- this is the core architectural pattern established in Phase 3 research (Pitfall 4).

### Anti-Patterns to Avoid
- **Storing COM object references in EmailReport:** Never store MailItem, Attachment, MAPIFolder, ExchangeUser, or any other OOM object in the DTO. Extract the string/int/bool value and release the COM object immediately. COM objects are RCW wrappers -- storing them in a DTO that crosses thread boundaries causes COMException.
- **Deferring formatting to the orchestrator with raw COM data:** Do not pass `mailItem.Parent` or `session.CurrentUser` to the orchestrator. Extract the formatted string on the UI thread.
- **Making EmailReport mutable:** Do not add setters. The entire point is that once constructed, the report cannot change. This prevents subtle bugs where async code modifies shared state.
- **Moving COM cleanup to the orchestrator:** COM objects must be released on the thread that owns them (the UI thread). Keep all `Marshal.ReleaseComObject` calls in Ribbon.cs `finally` blocks, not in the orchestrator.
- **Creating the report MailItem in the orchestrator's GoPhish branch post-await:** The GoPhish branch does not create a report email (it only sends the GoPhish notification). But if future changes add report email creation in the GoPhish branch, it must be done before the await, not after.
- **Passing mailItem to the orchestrator for post-await access:** The orchestrator receives mailItem for `.Delete()` -- but this call must happen BEFORE the await (GoPhish branch) or in a path with NO await (non-GoPhish branch). This is architecturally safe but fragile. Document the invariant clearly.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| Immutable data container | Mutable class with setters | Sealed class with get-only properties + constructor | Immutability is the safety guarantee against cross-thread data races; mutable DTOs defeat the purpose |
| Report body formatting | Inline string concatenation in orchestrator with OOM access | Pre-computed string sections in EmailReport | Keeps OOM access in one place (Ribbon.cs); orchestrator never touches COM |
| Thread marshaling back to UI | SynchronizationContext.Post / Control.Invoke | Keep OOM calls before the await boundary | Marshaling adds complexity and latency; the simpler pattern is to extract data first and avoid needing to marshal back |
| COM object lifecycle | IDisposable wrapper around OOM objects | Manual Marshal.ReleaseComObject in try/finally | Matches existing codebase pattern (Phase 2); IDisposable wrapper adds infrastructure code with no benefit for the small number of COM objects here |

**Key insight:** Phase 4 does NOT introduce any new async patterns or threading mechanisms. It restructures EXISTING code so that the OOM-before-await invariant is structural (enforced by the EmailReport boundary) rather than incidental (enforced by control-flow analysis of which branch has an await).

## Common Pitfalls

### Pitfall 1: Storing COM References in the Immutable DTO
**What goes wrong:** EmailReport stores a `MailItem` or `Attachment` reference instead of extracting the string/int value. When the orchestrator accesses the property on a background thread, COMException 0x8001010E is thrown.
**Why it happens:** It feels natural to pass the "rich" object rather than extracting each property individually. But COM objects are STA-bound -- they cannot be safely used from another thread.
**How to avoid:** EmailReport MUST contain only: primitive types (string, int, bool), immutable C# objects (UrlExtractionResult, AttachmentHashResult, IReadOnlyList), and null. No COM types.
**Warning signs:** Any `Microsoft.Office.Interop.Outlook.*` type in EmailReport's property types.

### Pitfall 2: Constructor Parameter Explosion
**What goes wrong:** EmailReport has 14+ constructor parameters, making it easy to swap arguments (e.g., passing headers as subject or vice versa).
**Why it happens:** C# 7.3 has no named/required initialization syntax (no `init`, no records). All properties must be set via constructor.
**How to avoid:** Use a builder pattern or a static factory method with clearly-named local variables at the call site. Alternatively, group related properties into sub-objects (e.g., a `UserContext` class for environment info). For this codebase, the simplest approach: keep the constructor but use named arguments at the call site (`new EmailReport(itemType: ..., isMailItem: ..., subject: ...)`). Named arguments are available in C# 7.3.
**Warning signs:** Bugs where report subject appears in the headers field or vice versa.

### Pitfall 3: Forgetting to Extract New OOM Properties
**What goes wrong:** A future developer adds a new report field (e.g., `mailItem.ReceivedTime`) but accesses it directly in the orchestrator instead of extracting it into EmailReport. On the GoPhish path, this throws COMException.
**Why it happens:** The pattern is non-obvious -- you have to know that the orchestrator runs on a thread pool after an await.
**How to avoid:** Document the invariant clearly in EmailReport and ReportOrchestrator. Add a comment at the await boundary in the orchestrator: "WARNING: After this line, no Outlook OOM access is permitted."
**Warning signs:** COMException 0x8001010E appearing in logs after adding a new feature.

### Pitfall 4: Error Email Creation After Await
**What goes wrong:** The error handling path creates an error email (`Application.CreateItem`) after an await has occurred, causing COMException because the code is on a thread pool thread.
**Why it happens:** The current error handling in the catch block creates an error email to notify support. If this catch block runs after an await, it is on the wrong thread.
**How to avoid:** Move error email creation to Ribbon.cs (which is always on the UI thread for the catch block of the outer try/catch). The orchestrator logs the error and rethrows or returns a result; Ribbon.cs handles error email creation.
**Warning signs:** Error reporting itself fails with COMException in logs.

### Pitfall 5: COM Cleanup in Wrong Scope
**What goes wrong:** COM objects are released in the orchestrator's finally block, but the orchestrator runs on a thread pool after an await. `Marshal.ReleaseComObject` from a non-owner thread causes unpredictable behavior.
**Why it happens:** Moving code from Ribbon.cs to ReportOrchestrator moves the finally block too.
**How to avoid:** Keep ALL `Marshal.ReleaseComObject` calls in Ribbon.cs. The orchestrator never releases COM objects -- it only reads data from EmailReport and performs non-COM operations.
**Warning signs:** Intermittent `InvalidComObjectException` or `AccessViolationException` after reporting.

### Pitfall 6: Breaking the Attachment Add with Extracted Data
**What goes wrong:** The report email `reportEmail.Attachments.Add(selection[1])` requires the ORIGINAL Outlook item (COM object) to be attached. This cannot be replaced with data from EmailReport -- you cannot reconstruct an Outlook item from extracted strings.
**Why it happens:** The developer extracts everything into EmailReport and assumes the orchestrator can work without any COM reference.
**How to avoid:** The orchestrator receives the original item reference (`object originalItem`) specifically for `Attachments.Add()`. This OOM call happens in the non-GoPhish branch, which has no await, so it is safe on the UI thread. Document this exception clearly.
**Warning signs:** Report email arrives without the original email attachment, or ArgumentException when trying to attach a string instead of a COM object.

## Code Examples

### Example 1: EmailReport Construction in Ribbon.cs
```csharp
// Source: Codebase analysis + C# 7.3 immutable class pattern
// Called on UI thread in Ribbon.cs, before any await
private EmailReport ExtractEmailReport(object selectedItem, MailItem mailItem)
{
    string itemType = DetectItemType(selectedItem);
    bool isMailItem = (itemType == "MailItem");

    string subject = isMailItem ? mailItem.Subject : itemType;
    string headers = isMailItem ? mailItem.HeaderString() : null;
    string htmlBody = isMailItem ? mailItem.HTMLBody : null;

    // GoPhish detection (pure string parsing, no I/O)
    string goPhishReportUrl = (headers != null)
        ? GoPhishIntegration.setReportURL(headers)
        : null;

    // Pre-compute analysis (these methods access OOM internally)
    UrlExtractionResult urlAnalysis = (htmlBody != null)
        ? UrlExtractor.ExtractUrls(htmlBody)
        : new UrlExtractionResult(Array.Empty<string>(), Array.Empty<string>());

    IReadOnlyList<AttachmentHashResult> attachmentHashes =
        isMailItem ? ExtractAttachmentHashes(mailItem) : Array.Empty<AttachmentHashResult>();

    // Pre-compute formatted sections (these methods access OOM)
    string userInfoSection = GetCurrentUserInfos();
    string basicInfoSection = isMailItem ? GetBasicInfo(mailItem) : null;
    string pluginDetailsSection = GetPluginDetails();

    return new EmailReport(
        itemType: itemType,
        isMailItem: isMailItem,
        subject: subject,
        headers: headers,
        htmlBody: htmlBody,
        folderPath: isMailItem ? GetFolderPath(mailItem) : null,
        goPhishReportUrl: goPhishReportUrl,
        urlAnalysis: urlAnalysis,
        attachmentHashes: attachmentHashes,
        userInfoSection: userInfoSection,
        basicInfoSection: basicInfoSection,
        pluginDetailsSection: pluginDetailsSection,
        suspiciousReportsCount: Properties.Settings.Default.suspecious_reports_counter,
        goPhishReportsCount: Properties.Settings.Default.gophish_reports_counter);
}
```

### Example 2: ReportOrchestrator.ComposeReportBody
```csharp
// Source: Extracted from current Ribbon.cs report composition
// Pure string building -- no OOM access
private static string ComposeReportBody(EmailReport report)
{
    var body = report.UserInfoSection;
    body += "\n";

    if (report.BasicInfoSection != null)
    {
        body += report.BasicInfoSection;
        body += "\n";
    }

    // URLs and Attachments section from pre-extracted data
    body += "---------- URLs and Attachments ----------";
    body += "\n # of unique Domains: " + report.UrlAnalysis.UniqueDomains.Count;
    foreach (string domain in report.UrlAnalysis.UniqueDomains)
    {
        body += "\n --> Domain: " + domain.Replace(":", "[:]");
    }

    body += "\n\n # of URLs: " + report.UrlAnalysis.Urls.Count;
    foreach (string url in report.UrlAnalysis.Urls)
    {
        body += "\n --> URL: " + url.Replace(":", "[:]");
    }

    body += "\n\n # of Attachments: " + report.AttachmentHashes.Count;
    foreach (var hash in report.AttachmentHashes)
    {
        body += "\n --> Attachment: " + hash.FileName
            + " (" + hash.SizeBytes + " bytes)"
            + "\n\t\tMD5: " + hash.Md5
            + "\n\t\tSha256: " + hash.Sha256 + "\n";
    }

    body += "\n---------- Headers ----------";
    body += "\n" + report.Headers;
    body += "\n";
    body += report.PluginDetailsSection + "\n\n";

    return body;
}
```

### Example 3: Thin Ribbon.cs After Phase 4
```csharp
// Source: Codebase refactoring following QUAL-01
private async Task reportPhishingEmailToSecurityTeamAsync(IRibbonControl control)
{
    Logger.Info("Processing selected email for phishing report");

    Selection selection = null;
    MailItem mailItem = null;

    try
    {
        selection = Globals.ThisAddIn.Application.ActiveExplorer().Selection;

        if (selection.Count < 1)
        {
            MessageBox.Show("Select an email before reporting.", "Error");
            return;
        }
        if (selection.Count > 1)
        {
            MessageBox.Show("You can report 1 email at a time.", "Error");
            return;
        }

        if (!(selection[1] is Outlook.MeetingItem
            || selection[1] is Outlook.ContactItem
            || selection[1] is Outlook.AppointmentItem
            || selection[1] is Outlook.TaskItem
            || selection[1] is Outlook.MailItem))
        {
            MessageBox.Show("You cannot report this item", "Error");
            return;
        }

        mailItem = selection[1] as MailItem;

        // ALL OOM extraction happens here, on the UI thread
        EmailReport report = ExtractEmailReport(selection[1], mailItem);
        Logger.Info("Email data extracted, item type: {0}", report.ItemType);

        // Delegate to orchestrator
        await ReportOrchestrator.ExecuteAsync(
            report,
            Globals.ThisAddIn.Application,
            selection[1],
            mailItem).ConfigureAwait(false);
    }
    catch (Exception ex)
    {
        Logger.Error(ex, "Error during report processing");
        // Error email creation stays here (UI thread for catch block)
        SendErrorEmail(ex);
    }
    finally
    {
        if (mailItem != null) { try { Marshal.ReleaseComObject(mailItem); } catch { } mailItem = null; }
        if (selection != null) { try { Marshal.ReleaseComObject(selection); } catch { } selection = null; }
    }
}
```

### Example 4: Error Email Helper in Ribbon.cs
```csharp
// Stays in Ribbon.cs because it needs OOM (Application.CreateItem)
// and must run on UI thread
private void SendErrorEmail(Exception ex)
{
    MessageBox.Show(
        "There was an error! An automatic email was sent to the support to resolve the issue.",
        "Do not worry");

    MailItem errorEmail = null;
    try
    {
        errorEmail = (MailItem)Globals.ThisAddIn.Application
            .CreateItem(OlItemType.olMailItem);
        errorEmail.To = Properties.Settings.Default.support_email;
        errorEmail.Subject = "[Outlook Addin Error]";
        errorEmail.Body = "Addin error message: " + ex;
        errorEmail.Save();
        errorEmail.Send();
    }
    finally
    {
        if (errorEmail != null)
        {
            try { Marshal.ReleaseComObject(errorEmail); } catch { }
            errorEmail = null;
        }
    }
}
```

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| Monolithic ribbon class with interleaved OOM + logic | Orchestrator pattern with DTO boundary | Industry best practice since async/await in .NET 4.5 (2012) | Clean separation of UI-thread OOM extraction from background async work |
| Mutable MailItem passed across async boundary | Immutable DTO extracted on UI thread | Became critical with ConfigureAwait(false) adoption | Prevents COMException 0x8001010E; data is thread-safe by construction |
| C# 9 records for immutable DTOs | Sealed class with get-only properties (C# 7.3) | C# 7.3 is the ceiling for .NET Framework 4.8 | Same immutability guarantees; more boilerplate but no functional difference |

**Deprecated/outdated:**
- C# records: Would be ideal for EmailReport but require C# 9.0 (.NET 5+). Not available on .NET Framework 4.8.
- `init` accessors: Would simplify EmailReport construction but require C# 9.0. Not available.
- Forcing `<LangVersion>9.0</LangVersion>` in csproj: Technically compiles for some features but is unsupported by Microsoft for .NET Framework targets; risks runtime failures when using features that depend on missing BCL types.

## Open Questions

1. **Should the orchestrator receive `Outlook.Application` or should Ribbon.cs create the report MailItem?**
   - What we know: The orchestrator needs `Application.CreateItem()` for the non-GoPhish branch. This is an OOM call that must happen on the UI thread. Since the non-GoPhish branch has no await, the orchestrator runs entirely on the UI thread for that branch.
   - What's unclear: Whether passing `Outlook.Application` to a class named "orchestrator" is a design smell. It creates a dependency on OOM in a class that is supposed to be the "clean" async layer.
   - Recommendation: Accept the dependency. The alternative (Ribbon.cs creates the MailItem and passes it to the orchestrator) would mean Ribbon.cs knows about the report email creation -- which is exactly the logic we are trying to extract. The orchestrator can document that it requires UI-thread invocation and that `Application` is used only in the no-await branch. Confidence: MEDIUM -- acceptable tradeoff; could revisit if the orchestrator grows.

2. **Should GetCurrentUserInfos, GetBasicInfo, GetPluginDetails move to a new class?**
   - What we know: These methods access OOM and produce formatted strings. They are called from Ribbon.cs during EmailReport extraction. Moving them to a separate "EmailDataExtractor" class would further thin Ribbon.cs.
   - What's unclear: Whether they should be static methods on EmailReport (factory pattern), on a dedicated extractor class, or remain in Ribbon.cs as private helpers.
   - Recommendation: Move them to a dedicated static class (e.g., `EmailDataExtractor` or keep them as private methods in Ribbon.cs called during extraction). Since the success criteria says "Ribbon.cs contains no email parsing, URL extraction, hash calculation, or HTTP logic," these formatting methods are arguably "email processing" and should move. Extract them to a static `EmailDataExtractor` class that returns the formatted strings. Confidence: MEDIUM -- the success criteria wording suggests they should move out of Ribbon.cs.

3. **reportEmail.Attachments.Add(selection[1]) -- should this be in the orchestrator or Ribbon.cs?**
   - What we know: This requires the original COM object (selection[1]) to attach the reported email to the report. The original item cannot be serialized into EmailReport. The non-GoPhish branch creates the report email in what will be the orchestrator.
   - What's unclear: Whether passing the raw COM object to the orchestrator violates the "no OOM after await" principle.
   - Recommendation: Pass `originalItem` (type `object`) to the orchestrator. It is only used in the non-GoPhish branch, which has no await. Document the invariant. This is the same pattern as passing `mailItem` for `.Delete()`. Confidence: HIGH -- the non-GoPhish branch is synchronous.

4. **Should the GoPhish branch's reportEmail creation be removed?**
   - What we know: Currently, the GoPhish branch creates a reportEmail (lines 154-159 in current Ribbon.cs: `Application.CreateItem`, `.Attachments.Add`, `.To`, `.Subject`) but NEVER sends it. The report email is simply released in the finally block.
   - What's unclear: Whether this was intentional (the email was prepared "just in case") or a leftover from an earlier code structure.
   - Recommendation: In the orchestrator, do NOT create a report email in the GoPhish branch. This eliminates unnecessary OOM access and simplifies the GoPhish path. If the GoPhish notification fails, the user still sees the success message (existing behavior -- GoPhish result is logged but does not gate the user message). Confidence: HIGH -- removing dead code.

## Sources

### Primary (HIGH confidence)
- [Microsoft Docs -- Threading support in Office](https://learn.microsoft.com/en-us/visualstudio/vsto/threading-support-in-office?view=vs-2022) -- STA model, COM marshaling, COMException on background threads, IMessageFilter
- [Microsoft Docs -- Configure language version](https://learn.microsoft.com/en-us/dotnet/csharp/language-reference/configure-language-version) -- .NET Framework 4.8 defaults to C# 7.3; records require C# 9.0
- [Microsoft Docs -- readonly keyword](https://learn.microsoft.com/en-us/dotnet/csharp/language-reference/keywords/readonly) -- Readonly fields and get-only properties for immutable classes
- [Microsoft Docs -- Structure types (readonly struct)](https://learn.microsoft.com/en-us/dotnet/csharp/language-reference/builtin-types/struct) -- readonly struct available in C# 7.2+, but not appropriate for large DTOs
- Phase 3 Research (.planning/phases/03-async-network-layer/03-RESEARCH.md) -- OOM-before-await pattern, COMException 0x8001010E, ConfigureAwait(false) semantics
- Phase 2 Research (.planning/phases/02-code-extraction/02-RESEARCH.md) -- UrlExtractionResult and AttachmentHashResult immutable DTO patterns, COM cleanup patterns
- Codebase inspection -- Current Ribbon.cs (481 lines), all OOM access points cataloged

### Secondary (MEDIUM confidence)
- [Add-in Express -- Threading in managed Office extensions](https://www.add-in-express.com/creating-addins-blog/2010/11/04/threads-managed-office-extensions/) -- "Calling the Outlook object model in a thread is a wrong solution because the object model is almost not thread-safe"; confirms extract-first-then-async pattern
- [GitHub bpatra/Hookmainthread](https://github.com/bpatra/Hookmainthread) -- Sample VSTO Outlook add-in for testing async processing; confirms the pattern of separating OOM access from async work
- [Microsoft Q&A -- VSTO Outlook addin update UI from async process](https://learn.microsoft.com/en-us/answers/questions/78894/vsto-outlook-addin-how-to-update-ui-from-asynchron) -- Community confirmation of STA threading constraints with async/await in VSTO

### Tertiary (LOW confidence)
- None. All findings verified against primary or secondary sources.

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH -- No new dependencies; all patterns use existing .NET Framework 4.8 BCL and C# 7.3 features
- Architecture: HIGH -- EmailReport + ReportOrchestrator pattern is a direct evolution of the OOM-before-await pattern established in Phase 3; the immutable DTO pattern matches existing codebase conventions (UrlExtractionResult, AttachmentHashResult)
- Pitfalls: HIGH -- All pitfalls derived from Phase 3 research (COMException 0x8001010E) and direct codebase analysis of OOM access points; threading constraints verified via Microsoft Docs

**Research date:** 2026-02-26
**Valid until:** 2026-03-28 (stable domain, 30-day validity)
