# Phase 2: Code Extraction - Research

**Researched:** 2026-02-26
**Domain:** C# class extraction / refactoring, VSTO exception-safety, COM object lifecycle, HtmlAgilityPack URL parsing, Settings persistence, GoPhish result typing
**Confidence:** HIGH

## Summary

Phase 2 decomposes the 458-line `Ribbon.cs` monolith into three single-responsibility classes (UrlExtractor, AttachmentHasher, GoPhishDetector) while simultaneously fixing five bugs and wrapping all ribbon callbacks in exception-safe try/catch. The extraction is a pure refactoring of synchronous code -- no async conversion happens here (that is Phase 3). Because these classes remain synchronous and operate entirely on the UI thread within existing Outlook COM calls, the refactoring carries low risk.

The five bugs addressed are: (1) the `Contains("a")` URL filter that silently drops URLs not containing the letter "a", (2) missing `Settings.Save()` call causing counter values to be lost on restart, (3) GoPhish returning magic strings "OK"/"ERROR"/"NaN" instead of typed results, (4) temp attachment files not cleaned up in a `finally` block, and (5) COM objects from Outlook OOM not released via `Marshal.ReleaseComObject`. Additionally, all ribbon callback entry points (`reportPhishing`, `getGroup1Image`, `Ribbon_Load`, `GetCustomUI`) must be wrapped in try/catch to prevent unhandled exceptions from triggering Outlook's soft-disable mechanism.

**Primary recommendation:** Extract one class at a time, fix its associated bug during extraction, and keep each extraction as a separate plan. Wrap ribbon callbacks in try/catch as a dedicated first plan since it is the highest-impact safety change and enables safer iteration on subsequent extractions.

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|-----------------|
| QUAL-03 | URL extraction logic extracted into UrlExtractor class | Section: Architecture Patterns -- UrlExtractor class design, removes `Contains("a")` bug during extraction |
| QUAL-04 | Hash calculation logic extracted into AttachmentHasher class | Section: Architecture Patterns -- AttachmentHasher class design, adds `finally` cleanup during extraction |
| QUAL-05 | COM objects properly released via Marshal.ReleaseComObject in all processing loops | Section: COM Object Cleanup -- Marshal.ReleaseComObject pattern for Outlook OOM objects |
| BUGF-01 | URL detection correctly captures all links (remove broken Contains("a") filter) | Section: Common Pitfalls, Pitfall 1 -- root cause analysis and fix pattern |
| BUGF-02 | Report counters persist across Outlook sessions (call Settings.Save() after increment) | Section: Common Pitfalls, Pitfall 2 -- ApplicationSettingsBase.Save() must be called explicitly |
| BUGF-03 | GoPhish integration returns enum/bool instead of magic strings | Section: Architecture Patterns -- GoPhishResult enum design |
| BUGF-05 | Temporary attachment files cleaned up in finally block | Section: Common Pitfalls, Pitfall 3 -- try/finally pattern for temp file lifecycle |
| STRT-05 | All ribbon event handler entry points wrapped in try/catch | Section: Architecture Patterns -- Exception-safe ribbon callback pattern |
</phase_requirements>

## Standard Stack

### Core

| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| .NET Framework | 4.8 | Runtime (already targeted in Phase 1) | Final VSTO-compatible runtime; all Phase 2 code is pure C# with no new dependencies |
| HtmlAgilityPack | 1.12.4 | HTML parsing for URL extraction (already referenced) | XPath-based `SelectNodes("//a[@href]")` is the correct, battle-tested pattern for link extraction |
| NLog | 5.4.0 | Logging in new extracted classes (already referenced) | Isolated LogFactory pattern established in Phase 1; new classes follow same logger pattern |

### Supporting

No new libraries are required for Phase 2. All work uses existing .NET Framework BCL types:
- `System.Runtime.InteropServices.Marshal` -- COM object cleanup
- `System.Security.Cryptography.MD5` / `SHA256` -- hash computation
- `System.IO.File` / `Path` -- temp file management
- `System.Configuration.ApplicationSettingsBase` -- settings persistence

### Alternatives Considered

| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| Manual `Marshal.ReleaseComObject` in finally blocks | `WithComCleanup()` IDisposable wrapper from Jake Ginnivan's pattern | Wrapper is cleaner but adds infrastructure code; manual release is sufficient for the small number of COM objects in this codebase. Wrapper could be added in Phase 4 if needed. |
| Enum for GoPhish result | Boolean return + out parameter for error message | Enum is more expressive (can distinguish NotFound/Success/Error) and matches the three existing states; bool loses the NotFound vs Error distinction |
| `SHA256.Create()` factory method | Keep `SHA256Managed` | `SHA256Managed` works on .NET 4.8 but is obsolete on modern .NET; `SHA256.Create()` is the recommended factory pattern and works identically on .NET Framework 4.8 |

**Installation:** No new packages needed. Phase 2 uses only what Phase 1 already installed.

## Architecture Patterns

### Recommended Project Structure After Phase 2

```
PhishingReporter/
├── ThisAddIn.cs              # Unchanged from Phase 1
├── Ribbon.cs                 # Thin coordination layer: try/catch wrappers + delegation
├── GoPhishIntegration.cs     # Refactored: returns GoPhishResult enum, no magic strings
├── UrlExtractor.cs           # NEW: extracted from GetURLsAndAttachmentsInfo()
├── AttachmentHasher.cs       # NEW: extracted from GetURLsAndAttachmentsInfo()
├── AppLogger.cs              # Unchanged from Phase 1
├── NLog.config               # Unchanged from Phase 1
├── Properties/
│   ├── Settings.settings     # Unchanged
│   └── AssemblyInfo.cs       # Unchanged
├── packages.config           # Unchanged
└── PhishingReporter.csproj   # Updated: new .cs file entries
```

### Pattern 1: Exception-Safe Ribbon Callback (STRT-05)

**What:** Wrap every public method in `Ribbon.cs` that Outlook calls (ribbon callbacks) in a top-level try/catch to prevent unhandled exceptions from reaching the COM boundary and triggering Outlook's soft-disable mechanism.

**When to use:** Every ribbon callback registered in Ribbon.xml -- `reportPhishing`, `getGroup1Image`, `Ribbon_Load`, `GetCustomUI`.

**Why critical:** Microsoft documentation states soft disabling occurs when "a VSTO Add-in produces an error that does not cause the application to unexpectedly close." An unhandled exception in a ribbon callback is exactly this scenario. The exception crosses the COM/managed boundary and Outlook records it as an add-in failure.

**Example:**
```csharp
// Source: Microsoft Docs — Re-enable a VSTO Add-in that has been disabled
// https://learn.microsoft.com/en-us/visualstudio/vsto/how-to-re-enable-a-vsto-add-in-that-has-been-disabled

public void reportPhishing(Office.IRibbonControl control)
{
    try
    {
        Logger.Info("Report phishing button clicked");
        // ... existing logic ...
    }
    catch (Exception ex)
    {
        Logger.Error(ex, "Unhandled exception in reportPhishing callback");
        try
        {
            MessageBox.Show(
                "An unexpected error occurred. Please try again or contact IT support.",
                "Phishing Reporter Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
        catch (Exception)
        {
            // Even MessageBox can fail in degraded COM states; swallow to protect add-in
        }
    }
}

public Bitmap getGroup1Image(IRibbonControl control)
{
    try
    {
        return Resources.phishing;
    }
    catch (Exception ex)
    {
        Logger.Error(ex, "Unhandled exception in getGroup1Image callback");
        return null; // Outlook handles null gracefully (no image shown)
    }
}
```

**Key insight:** The inner try/catch around MessageBox is deliberate. If the COM state is sufficiently degraded, even showing a dialog can fail. The outer catch must never re-throw.

**Confidence:** HIGH -- verified via Microsoft Docs on VSTO soft-disable behavior.

### Pattern 2: UrlExtractor Class (QUAL-03, BUGF-01)

**What:** A pure static class that accepts an HTML string and returns a structured result containing all extracted URLs and their domains. Fixes the `Contains("a")` bug by removing the filter entirely.

**Why static:** The class has no state; it transforms input HTML to output data. Making it static makes the single-responsibility obvious and avoids unnecessary instantiation.

**Example:**
```csharp
// UrlExtractor.cs
using System;
using System.Collections.Generic;
using System.Linq;
using HtmlAgilityPack;

namespace PhishingReporter
{
    /// <summary>
    /// Result of URL extraction from an email HTML body.
    /// Immutable data transfer object.
    /// </summary>
    internal sealed class UrlExtractionResult
    {
        public IReadOnlyList<string> Urls { get; }
        public IReadOnlyList<string> UniqueDomains { get; }

        public UrlExtractionResult(
            IReadOnlyList<string> urls,
            IReadOnlyList<string> uniqueDomains)
        {
            Urls = urls;
            UniqueDomains = uniqueDomains;
        }
    }

    /// <summary>
    /// Extracts URLs and domains from email HTML body.
    /// Replaces the broken Contains("a") filter from the original Ribbon.cs implementation.
    /// </summary>
    internal static class UrlExtractor
    {
        private static readonly NLog.Logger Logger =
            AppLogger.Instance.GetCurrentClassLogger();

        /// <summary>
        /// Extracts all href values from anchor tags in the provided HTML.
        /// </summary>
        /// <param name="emailHtmlBody">Raw HTML body of the email.</param>
        /// <returns>Extraction result with URLs and unique domains.</returns>
        public static UrlExtractionResult ExtractUrls(string emailHtmlBody)
        {
            if (string.IsNullOrEmpty(emailHtmlBody))
            {
                return new UrlExtractionResult(
                    Array.Empty<string>(),
                    Array.Empty<string>());
            }

            var doc = new HtmlDocument();
            doc.LoadHtml(emailHtmlBody);

            var urlNodes = doc.DocumentNode.SelectNodes("//a[@href]");
            if (urlNodes == null)
            {
                return new UrlExtractionResult(
                    Array.Empty<string>(),
                    Array.Empty<string>());
            }

            var urls = new List<string>();
            var domains = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var link in urlNodes)
            {
                string href = link.GetAttributeValue("href", "");
                if (string.IsNullOrWhiteSpace(href))
                    continue;

                // BUG FIX (BUGF-01): No Contains("a") filter.
                // All href values are captured regardless of content.
                urls.Add(href);

                // Domain extraction
                try
                {
                    domains.Add(new Uri(href).Host);
                }
                catch (UriFormatException)
                {
                    // Handle mailto: links — extract domain from email address
                    int atIndex = href.IndexOf('@');
                    if (atIndex >= 0)
                    {
                        string emailDomain = href.Substring(atIndex + 1);
                        // Trim any trailing path/query from the domain
                        int slashIndex = emailDomain.IndexOf('/');
                        if (slashIndex >= 0)
                            emailDomain = emailDomain.Substring(0, slashIndex);

                        if (!string.IsNullOrWhiteSpace(emailDomain))
                            domains.Add(emailDomain);
                    }
                    else
                    {
                        Logger.Warn("Unparseable URL skipped for domain extraction: {0}", href);
                    }
                }
            }

            Logger.Info("Extracted {0} URLs and {1} unique domains from email body",
                urls.Count, domains.Count);

            return new UrlExtractionResult(
                urls.AsReadOnly(),
                domains.ToList().AsReadOnly());
        }
    }
}
```

**Confidence:** HIGH -- derived from direct codebase inspection of the bug at Ribbon.cs line 280.

### Pattern 3: AttachmentHasher Class (QUAL-04, BUGF-05)

**What:** A static class that computes MD5 and SHA256 hashes of Outlook attachment content, using temp files with guaranteed cleanup in a `finally` block.

**Example:**
```csharp
// AttachmentHasher.cs
using System;
using System.IO;
using System.Security.Cryptography;
using Microsoft.Office.Interop.Outlook;

namespace PhishingReporter
{
    /// <summary>
    /// Hash results for a single attachment. Immutable.
    /// </summary>
    internal sealed class AttachmentHashResult
    {
        public string FileName { get; }
        public int SizeBytes { get; }
        public string Md5 { get; }
        public string Sha256 { get; }

        public AttachmentHashResult(string fileName, int sizeBytes, string md5, string sha256)
        {
            FileName = fileName;
            SizeBytes = sizeBytes;
            Md5 = md5;
            Sha256 = sha256;
        }
    }

    internal static class AttachmentHasher
    {
        private static readonly NLog.Logger Logger =
            AppLogger.Instance.GetCurrentClassLogger();

        /// <summary>
        /// Saves attachment to a temp file, computes hashes, and cleans up.
        /// The temp file is ALWAYS deleted, even if hashing throws.
        /// </summary>
        public static AttachmentHashResult ComputeHashes(Attachment attachment)
        {
            // Use Path.GetTempFileName() for uniqueness instead of predictable names
            string tempPath = Path.Combine(
                Path.GetTempPath(),
                "Outlook-Phishaddin-" + Guid.NewGuid().ToString("N") + ".tmp");

            try
            {
                attachment.SaveAsFile(tempPath);

                string md5Hash = ComputeMd5(tempPath);
                string sha256Hash = ComputeSha256(tempPath);

                Logger.Debug("Computed hashes for attachment: {0}", attachment.FileName);

                return new AttachmentHashResult(
                    attachment.FileName,
                    attachment.Size,
                    md5Hash,
                    sha256Hash);
            }
            finally
            {
                // BUGF-05: Guaranteed cleanup regardless of exceptions
                try
                {
                    if (File.Exists(tempPath))
                    {
                        File.Delete(tempPath);
                    }
                }
                catch (IOException ex)
                {
                    Logger.Warn(ex, "Failed to delete temp file: {0}", tempPath);
                }
            }
        }

        private static string ComputeMd5(string filePath)
        {
            using (var md5 = MD5.Create())
            using (var stream = File.OpenRead(filePath))
            {
                byte[] hash = md5.ComputeHash(stream);
                return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
            }
        }

        private static string ComputeSha256(string filePath)
        {
            using (var sha256 = SHA256.Create())
            using (var stream = File.OpenRead(filePath))
            {
                byte[] hash = sha256.ComputeHash(stream);
                return BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
            }
        }
    }
}
```

**Key changes from original:**
1. `finally` block guarantees temp file deletion (BUGF-05)
2. `SHA256.Create()` replaces `new SHA256Managed()` (obsolete API)
3. `Guid.NewGuid()` in filename prevents collisions if two reports run concurrently
4. `BitConverter.ToString()` replaces manual byte-to-hex loop for SHA256 (matches existing MD5 pattern)

**Confidence:** HIGH -- patterns verified via Microsoft Docs on SHA256.Create() and File.Delete best practices.

### Pattern 4: GoPhishResult Enum (BUGF-03)

**What:** Replace the magic strings "OK", "ERROR", and "NaN" with a typed enum. The three existing return values map to: `NotFound` (was "NaN"), `Reported` (was "OK"), `Error` (was "ERROR").

**Example:**
```csharp
// In GoPhishIntegration.cs
namespace PhishingReporter
{
    /// <summary>
    /// Result of GoPhish campaign detection and reporting.
    /// Replaces magic strings "OK", "ERROR", and "NaN".
    /// </summary>
    internal enum GoPhishResult
    {
        /// <summary>No GoPhish header found in email — not a simulated campaign.</summary>
        NotFound,

        /// <summary>GoPhish campaign detected and successfully reported to server.</summary>
        Reported,

        /// <summary>GoPhish campaign detected but reporting failed (network error, timeout).</summary>
        Error
    }
}
```

Updated method signatures:
```csharp
// setReportURL returns null instead of "NaN" when no header found
public static string setReportURL(string headers)
{
    // ... same regex logic ...
    return null; // was: return "NaN"
}

// sendReportNotificationToServer returns GoPhishResult instead of string
public static GoPhishResult sendReportNotificationToServer(string reportURL)
{
    // ... same HTTP logic ...
    return GoPhishResult.Reported;  // was: return "OK"
    // in catch:
    return GoPhishResult.Error;     // was: return "ERROR"
}
```

Caller update in Ribbon.cs:
```csharp
string simulatedPhishingURL = GoPhishIntegration.setReportURL(reportedItemHeaders);
if (simulatedPhishingURL != null)  // was: != "NaN"
{
    GoPhishResult result = GoPhishIntegration.sendReportNotificationToServer(simulatedPhishingURL);
    Logger.Info("GoPhish notification result: {0}", result);
    // ...
}
```

**Confidence:** HIGH -- straightforward enum mapping of existing string literals.

### Pattern 5: COM Object Cleanup (QUAL-05)

**What:** Release COM objects obtained from Outlook OOM properties using `Marshal.ReleaseComObject` in `finally` blocks. The critical objects in the current code are: `Selection`, `MailItem`, `MAPIFolder` (from `mailItem.Parent`), `AddressEntry`, `ExchangeUser`, and `Attachments` collection items.

**Rules:**
1. Never chain property accesses (`foo.Bar.Baz`) -- store each intermediate COM object in a local variable
2. Release in reverse order of acquisition
3. Set variable to `null` after release
4. Wrap in try/finally, never let exceptions from ReleaseComObject propagate

**Example for the main processing method:**
```csharp
Selection selection = null;
MailItem mailItem = null;
MailItem reportEmail = null;

try
{
    selection = Globals.ThisAddIn.Application.ActiveExplorer().Selection;
    // ... processing ...
    mailItem = selection[1] as MailItem;
    // ... use mailItem ...
}
finally
{
    if (reportEmail != null) { Marshal.ReleaseComObject(reportEmail); reportEmail = null; }
    if (mailItem != null) { Marshal.ReleaseComObject(mailItem); mailItem = null; }
    if (selection != null) { Marshal.ReleaseComObject(selection); selection = null; }
}
```

**For the GetCurrentUserInfos method:**
```csharp
AddressEntry addrEntry = null;
ExchangeUser currentUser = null;

try
{
    addrEntry = Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry;
    if (addrEntry.Type == "EX")
    {
        currentUser = addrEntry.GetExchangeUser();
        // ... use currentUser ...
    }
}
finally
{
    if (currentUser != null) { Marshal.ReleaseComObject(currentUser); currentUser = null; }
    if (addrEntry != null) { Marshal.ReleaseComObject(addrEntry); addrEntry = null; }
}
```

**Important:** The `Explorer` object from `ActiveExplorer()` is a long-lived Outlook object owned by the application -- do NOT release it. Only release objects obtained within the scope of the report operation.

**Confidence:** HIGH -- pattern verified via Jake Ginnivan's VSTO COM Interop guide and Microsoft VSTO documentation.

### Pattern 6: Settings Persistence (BUGF-02)

**What:** Call `Properties.Settings.Default.Save()` after modifying counter values so they persist to the user.config file across Outlook sessions.

**Example:**
```csharp
// After incrementing any counter:
Properties.Settings.Default.suspecious_reports_counter++;
Properties.Settings.Default.Save();  // BUGF-02: Persist to disk

// Same for GoPhish counter:
Properties.Settings.Default.gophish_reports_counter++;
Properties.Settings.Default.Save();  // BUGF-02: Persist to disk
```

**Where Settings.Save() writes:** User-scoped settings are persisted to `user.config` in the user's `AppData\Local` directory, under a path derived from the application's assembly identity. For VSTO add-ins, the path includes the Outlook version, so an Outlook major version upgrade may cause the settings to appear "lost" (the user.config is in a folder named after the old version). This is a known VSTO limitation, not a bug to fix in this phase.

**Confidence:** HIGH -- verified via Microsoft Docs ApplicationSettingsBase.Save().

### Anti-Patterns to Avoid

- **Do NOT make extracted classes depend on Outlook OOM types in their public API.** `UrlExtractor.ExtractUrls()` takes a `string` (HTML body), not a `MailItem`. `AttachmentHasher.ComputeHashes()` takes an `Attachment` because it needs `SaveAsFile()`, but this is the minimum COM dependency. Phase 4 will further decouple by extracting all OOM data into an immutable `EmailReport` before processing.
- **Do NOT catch and re-throw exceptions in ribbon callbacks.** The try/catch at the callback boundary must be the terminal handler. Re-throwing defeats the purpose of preventing COM-boundary exception propagation.
- **Do NOT use `Marshal.FinalReleaseComObject`.** It aggressively releases ALL references, which can crash if any other code path still holds a reference to the RCW. Use `Marshal.ReleaseComObject` (single decrement) instead.
- **Do NOT release COM objects obtained from `Globals.ThisAddIn.Application`.** The Application object is managed by the VSTO runtime and lives for the add-in's lifetime. Releasing it would crash subsequent operations.
- **Do NOT put `Settings.Save()` in a loop or call it after every individual property change.** Call it once after all counters for a single report operation are updated.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| URL extraction from HTML | Regex-based href parser | HtmlAgilityPack `SelectNodes("//a[@href]")` | Regex cannot correctly parse HTML; HAP handles malformed markup, nested tags, attribute quoting variations |
| File hash computation | Manual byte-to-hex conversion loop | `BitConverter.ToString(hash).Replace("-","").ToLowerInvariant()` | Already used for MD5 in codebase; consistent, one-liner, no off-by-one risk |
| Temp file unique naming | DisplayName-based file paths (current: `Outlook-Phishaddin-{DisplayName}.txt`) | `Guid.NewGuid().ToString("N") + ".tmp"` | DisplayName can contain path-illegal characters (`<`, `>`, `:`); GUID guarantees uniqueness across concurrent operations |
| Settings persistence | Custom file I/O for counters | `ApplicationSettingsBase.Save()` | Already using the Settings infrastructure; just need to call `.Save()` |

**Key insight:** Phase 2 adds zero new dependencies. Every fix and extraction uses libraries and patterns already in the project. The value is in correctness (bug fixes) and structure (single-responsibility classes), not in new technology.

## Common Pitfalls

### Pitfall 1: The Contains("a") URL Filter Bug (BUGF-01)

**What goes wrong:** The current code at Ribbon.cs line 280 filters URLs with `att.Value.Contains("a")`. This means any URL that does not contain the lowercase letter "a" is silently dropped from the report. For example, `https://evil.com/login` is captured (contains "a" in... wait, no -- "login" has no "a"... it is in the path? No: "https://evil.com/login" does NOT contain "a"). Actually, re-reading: the href value `https://evil.com/login` does NOT contain "a", so it would be DROPPED. But `https://example.com` WOULD be captured because "example" contains "a".

**Why it happens:** The filter appears to be a failed attempt to check if the node is an `<a>` tag, but `att.Value` is the href attribute VALUE, not the tag name. The HtmlAgilityPack `SelectNodes("//a[@href]")` already guarantees only `<a>` tags are selected, making ANY additional filter redundant.

**How to avoid:** Remove the `Contains("a")` check entirely. The XPath selector `//a[@href]` already constrains to anchor elements with href attributes. No further filtering is needed.

**Warning signs:** Security team reports that phishing email analysis is "missing URLs" compared to what they see in the raw email source.

### Pitfall 2: Missing Settings.Save() (BUGF-02)

**What goes wrong:** The `suspecious_reports_counter` and `gophish_reports_counter` values increment correctly during a single Outlook session but reset to their default values (0) when Outlook is restarted.

**Why it happens:** `ApplicationSettingsBase` loads user-scoped settings automatically on first access, but does NOT save automatically. The current code increments `Properties.Settings.Default.suspecious_reports_counter++` (Ribbon.cs line 168) and `Properties.Settings.Default.gophish_reports_counter++` (Ribbon.cs line 159) but never calls `Properties.Settings.Default.Save()`. The in-memory values are correct but never persisted to the `user.config` file.

**How to avoid:** Call `Properties.Settings.Default.Save()` after incrementing counters. Place the call after both counter paths (GoPhish and suspicious) so it executes regardless of which branch is taken.

**Warning signs:** Counter displays correct values during a session but shows 0 after Outlook restart.

### Pitfall 3: Temp File Leak on Exception (BUGF-05)

**What goes wrong:** If `CalculateMD5()` or `GetHashSha256()` throws an exception (e.g., file locked by antivirus, disk full), the temp file at `%TEMP%\Outlook-Phishaddin-{name}.txt` is never deleted because `File.Delete(filePath)` at line 338 is inside the success path, not in a `finally` block.

**Why it happens:** The original code structure is:
```
SaveAsFile(filePath)
if (File.Exists(filePath))
    hash = ComputeHash(filePath)
    File.Delete(filePath)          // <-- only reached on success
```

If `ComputeHash` throws, execution jumps to the outer catch block, and the temp file remains.

**How to avoid:** Move file deletion to a `finally` block. The `AttachmentHasher.ComputeHashes()` pattern (see Architecture Patterns) demonstrates this. Wrap the `try` around `SaveAsFile` + hash computation, and put `File.Delete` in `finally`.

**Warning signs:** Accumulating `Outlook-Phishaddin-*.txt` files in the user's temp directory over time.

### Pitfall 4: COM Object Leak in GetCurrentUserInfos()

**What goes wrong:** The `GetCurrentUserInfos()` method at Ribbon.cs line 233 accesses `Globals.ThisAddIn.Application.Session.CurrentUser.AddressEntry` and potentially `.GetExchangeUser()`. These intermediate COM objects (`Session`, `CurrentUser`, `AddressEntry`, `ExchangeUser`) are never released, creating RCW leaks.

**Why it happens:** The COM object lifecycle requirement is non-obvious -- in pure C#, the GC handles cleanup, but COM objects use reference counting. Each property access on an Outlook OOM object returns a new RCW wrapping a new COM reference. Without explicit `Marshal.ReleaseComObject`, these references leak until the next GC cycle, and in Outlook this can cause "ghost" inspector windows and other symptoms.

**How to avoid:** Store each intermediate COM object in a named local variable, use it, then release in a `finally` block in reverse order of acquisition. See Pattern 5 in Architecture Patterns.

**Warning signs:** Outlook becomes sluggish after many report operations; inspector windows fail to close; Outlook hangs on shutdown.

### Pitfall 5: SHA256Managed Not Disposed

**What goes wrong:** The `GetHashSha256()` method at Ribbon.cs line 414 creates `new SHA256Managed()` without a `using` statement. The `SHA256Managed` class implements `IDisposable` and holds unmanaged cryptographic resources.

**Why it happens:** The original code uses `SHA256Managed sha = new SHA256Managed()` as a local variable without `using` or `try/finally`. The GC will eventually collect it, but the unmanaged resources may linger.

**How to avoid:** Use `using (var sha256 = SHA256.Create())` pattern (see AttachmentHasher code example). This replaces the obsolete `SHA256Managed` constructor with the recommended factory method AND ensures proper disposal.

**Warning signs:** No immediate symptoms, but represents a resource leak in a potentially hot path (many attachments).

### Pitfall 6: Predictable Temp File Names

**What goes wrong:** The current temp file naming `Outlook-Phishaddin-{DisplayName}.txt` uses the attachment's display name, which can: (a) contain path-illegal characters like `<`, `>`, `:`, `"`, (b) collide if two attachments have the same display name, and (c) be predictable for local privilege escalation attacks (an attacker could pre-create a symlink at the known path).

**Why it happens:** The original code uses `a.DisplayName` directly in the path without sanitization.

**How to avoid:** Use `Guid.NewGuid().ToString("N")` for the temp filename. This guarantees uniqueness and eliminates illegal character issues. The original filename is recorded in the hash result metadata, not in the temp path.

**Warning signs:** `IOException` or `ArgumentException` when processing attachments with special characters in their names.

## Code Examples

### Example 1: Complete Ribbon.cs reportPhishing After Extraction

```csharp
// Source: Codebase inspection + Microsoft VSTO disable docs
public void reportPhishing(Office.IRibbonControl control)
{
    try
    {
        Logger.Info("Report phishing button clicked");
        var areYouSure = MessageBox.Show(
            "Do you want to report this email to the Information Security Team as a potential phishing attempt?",
            "Are you sure?",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question);

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
    catch (Exception ex)
    {
        Logger.Error(ex, "Unhandled exception in reportPhishing callback");
        try
        {
            MessageBox.Show(
                "An unexpected error occurred. Please try again or contact IT support.",
                "Phishing Reporter Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }
        catch (Exception)
        {
            // Protect add-in stability above all else
        }
    }
}
```

### Example 2: Using UrlExtractor in GetURLsAndAttachmentsInfo

```csharp
// After extraction, Ribbon.cs delegates URL work to UrlExtractor
public string GetURLsAndAttachmentsInfo(MailItem mailItem)
{
    string result = "---------- URLs and Attachments ----------";

    // URL extraction delegated to UrlExtractor (QUAL-03)
    var urlResult = UrlExtractor.ExtractUrls(mailItem.HTMLBody);

    result += "\n # of unique Domains: " + urlResult.UniqueDomains.Count;
    foreach (string domain in urlResult.UniqueDomains)
    {
        result += "\n --> Domain: " + domain.Replace(":", "[:]");
    }

    result += "\n\n # of URLs: " + urlResult.Urls.Count;
    foreach (string url in urlResult.Urls)
    {
        result += "\n --> URL: " + url.Replace(":", "[:]");
    }

    // Attachment hashing delegated to AttachmentHasher (QUAL-04)
    result += "\n\n # of Attachments: " + mailItem.Attachments.Count;
    foreach (Attachment a in mailItem.Attachments)
    {
        var hashResult = AttachmentHasher.ComputeHashes(a);
        result += "\n --> Attachment: " + hashResult.FileName
            + " (" + hashResult.SizeBytes + " bytes)"
            + "\n\t\tMD5: " + hashResult.Md5
            + "\n\t\tSha256: " + hashResult.Sha256 + "\n";
    }

    return result;
}
```

### Example 3: Settings.Save() After Counter Increment

```csharp
// GoPhish branch:
Properties.Settings.Default.gophish_reports_counter++;
Properties.Settings.Default.Save(); // BUGF-02: persist immediately

// Suspicious email branch:
Properties.Settings.Default.suspecious_reports_counter++;
Properties.Settings.Default.Save(); // BUGF-02: persist immediately
```

### Example 4: COM Cleanup in GetCurrentUserInfos

```csharp
public string GetCurrentUserInfos()
{
    string str = "---------- User Information ----------";
    str += "\n - Domain:" + Environment.UserDomainName;
    str += "\n - Username:" + Environment.UserName;
    str += "\n - Machine name:" + Environment.MachineName;

    Outlook.NameSpace session = null;
    Outlook.Recipient currentUserRecipient = null;
    Outlook.AddressEntry addrEntry = null;
    Outlook.ExchangeUser currentUser = null;

    try
    {
        session = Globals.ThisAddIn.Application.Session;
        currentUserRecipient = session.CurrentUser;
        addrEntry = currentUserRecipient.AddressEntry;

        if (addrEntry.Type == "EX")
        {
            currentUser = addrEntry.GetExchangeUser();
            if (currentUser != null)
            {
                str += "\n - Name: " + currentUser.Name;
                str += "\n - STMP address: " + currentUser.PrimarySmtpAddress;
                str += "\n - Title: " + currentUser.JobTitle;
                str += "\n - Department: " + currentUser.Department;
                str += "\n - Location: " + currentUser.OfficeLocation;
                str += "\n - Business phone: " + currentUser.BusinessTelephoneNumber;
                str += "\n - Mobile phone: " + currentUser.MobileTelephoneNumber;
            }
        }
    }
    finally
    {
        if (currentUser != null) { Marshal.ReleaseComObject(currentUser); currentUser = null; }
        if (addrEntry != null) { Marshal.ReleaseComObject(addrEntry); addrEntry = null; }
        if (currentUserRecipient != null) { Marshal.ReleaseComObject(currentUserRecipient); currentUserRecipient = null; }
        if (session != null) { Marshal.ReleaseComObject(session); session = null; }
    }

    return str + "\n";
}
```

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| `new SHA256Managed()` | `SHA256.Create()` | .NET 6+ (SYSLIB0021) | Functionally identical on .NET Framework 4.8; forward-compatible if ever migrated |
| Magic string returns ("OK", "ERROR", "NaN") | Typed enums | C# best practice since C# 1.0 | Eliminates typo bugs, enables switch exhaustiveness checking |
| Inline COM object access (property chaining) | Named locals with try/finally release | VSTO best practice since Outlook 2007 | Prevents RCW leaks, ghost inspectors, shutdown hangs |
| No Settings.Save() call | Explicit Save() after mutations | ApplicationSettingsBase design since .NET 2.0 | User-scoped settings were never auto-saved; the API requires explicit persistence |

**Deprecated/outdated:**
- `SHA256Managed`: Marked `[Obsolete]` in .NET 6+ with SYSLIB0021. On .NET Framework 4.8 it compiles without warnings, but `SHA256.Create()` is the recommended replacement and is functionally identical.
- `HttpWebRequest`/`HttpWebResponse`: Still used in `GoPhishIntegration.sendReportNotificationToServer()`. NOT changed in Phase 2 -- that is Phase 3 (async HttpClient conversion). Phase 2 only changes the return type from string to enum.

## Open Questions

1. **Explorer COM object release**
   - What we know: `Globals.ThisAddIn.Application.ActiveExplorer()` returns a COM object. The VSTO runtime manages `Application`, but `ActiveExplorer()` returns a new RCW on each call.
   - What's unclear: Whether releasing the Explorer object causes issues with Outlook's ribbon state (since the ribbon is hosted by the Explorer).
   - Recommendation: Store `ActiveExplorer()` in a local variable but do NOT release it -- it is a long-lived UI object. Release only the `Selection` obtained from it. If testing reveals issues, revisit.

2. **Attachment COM objects in foreach loop**
   - What we know: `mailItem.Attachments` returns a COM collection. Each `Attachment` in the loop is a separate COM object.
   - What's unclear: Whether releasing individual `Attachment` objects mid-loop invalidates the `Attachments` collection enumerator.
   - Recommendation: Collect attachment data (via `AttachmentHasher.ComputeHashes()`) in the loop, but defer release of individual `Attachment` objects to after the loop. Release the `Attachments` collection reference after the loop completes. If the foreach enumerator owns the reference, explicit release may cause `InvalidComObjectException`. Test iteratively.

3. **MAPIFolder from mailItem.Parent**
   - What we know: `GetBasicInfo()` accesses `mailItem.Parent as MAPIFolder` to get FolderPath. This returns a COM object that is never released.
   - What's unclear: Whether the cast `as MAPIFolder` creates a new RCW or reuses the existing one.
   - Recommendation: Store in a local, release in finally. This is a short-lived reference with clear scope.

## Sources

### Primary (HIGH confidence)
- [Microsoft Docs -- Re-enable a VSTO Add-in that has been disabled](https://learn.microsoft.com/en-us/visualstudio/vsto/how-to-re-enable-a-vsto-add-in-that-has-been-disabled?view=vs-2022) -- Soft-disable vs hard-disable criteria; unhandled exception causes soft-disable
- [Microsoft Docs -- ApplicationSettingsBase.Save()](https://learn.microsoft.com/en-us/dotnet/api/system.configuration.applicationsettingsbase.save?view=windowsdesktop-9.0) -- Save() must be called explicitly; values not auto-persisted
- [Microsoft Docs -- SHA256.Create()](https://learn.microsoft.com/en-us/dotnet/api/system.security.cryptography.sha256.create?view=net-8.0) -- Recommended factory method; SHA256Managed is obsolete
- [Microsoft Docs -- SHA256Managed obsolete (SYSLIB0021)](https://learn.microsoft.com/en-us/dotnet/api/system.security.cryptography.sha256managed?view=net-7.0) -- Obsolete annotation details
- Codebase inspection -- Ribbon.cs line 280 `Contains("a")` bug, line 159/168 missing Save(), GoPhishIntegration.cs "OK"/"ERROR"/"NaN" strings, line 328-338 temp file lifecycle

### Secondary (MEDIUM confidence)
- [Jake Ginnivan -- VSTO and COM Interop](http://jake.ginnivan.net/vsto-com-interop/) -- COM object cleanup patterns, single-dot rule, scope constraints for ReleaseComObject
- [Add-in Express -- Releasing COM objects](https://www.add-in-express.com/creating-addins-blog/releasing-com-objects-garbage-collector-marshal-relseasecomobject/) -- GC.Collect vs Marshal.ReleaseComObject tradeoffs
- [HtmlAgilityPack -- SelectNodes documentation](https://html-agility-pack.net/select-nodes) -- XPath pattern `//a[@href]` selects all anchor elements with href attribute
- [MSDN Forums -- User-scoped settings in VSTO Add-in](https://social.msdn.microsoft.com/Forums/vstudio/en-US/69121741-df1e-4141-bf45-d664db72a77c/) -- VSTO-specific user.config path and version sensitivity

### Tertiary (LOW confidence)
- [Stack Overflow via Kiwix -- VSTO Outlook addin settings best way](https://kiwix.ounapuu.ee/content/stackoverflow.com_en_all_2023-11/questions/8332568/vsto-outlook-addin-need-to-save-settings-best-way) -- Community confirmation that Settings.Default.Save() works in VSTO but path changes on Outlook upgrade

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH -- No new dependencies; all patterns use existing .NET Framework 4.8 BCL and already-referenced NuGet packages
- Architecture: HIGH -- Extraction patterns are straightforward class decomposition with immutable result types; verified against codebase inspection
- Pitfalls: HIGH -- All six pitfalls verified against actual Ribbon.cs source code with exact line numbers; COM cleanup patterns verified via VSTO COM Interop documentation
- Bug fixes: HIGH -- All five bugs have root causes identified from direct code inspection; fixes are deterministic (remove filter, add Save() call, change return type, add finally block, add ReleaseComObject)

**Research date:** 2026-02-26
**Valid until:** 2026-03-28 (stable codebase, no external dependency changes)
