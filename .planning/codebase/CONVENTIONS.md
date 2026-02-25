# Coding Conventions

**Analysis Date:** 2026-02-25

## Language & Framework

**Primary Language:** C# (.NET Framework 4.6.1)

**Framework:** VSTO (Visual Studio Tools for Office) - Microsoft Office Add-in development framework

**Project Type:** Outlook Add-in (COM-visible class library)

## Naming Patterns

**Files:**
- PascalCase for all file names: `Ribbon.cs`, `GoPhishIntegration.cs`, `ThisAddIn.cs`
- Generated designer files follow pattern: `[Name].Designer.cs` (e.g., `ThisAddIn.Designer.cs`)
- Generated resource files: `[Name].Designer.cs` in Properties folder

**Classes:**
- PascalCase: `Ribbon`, `GoPhishIntegration` (see `Ribbon.cs` line 46, `GoPhishIntegration.cs` line 18)
- Static classes use PascalCase with `static` keyword: `static class GoPhishIntegration` (line 18)
- Extension classes follow naming convention with "Extensions" suffix: `MailItemExtensions` (`Ribbon.cs` line 409)

**Methods:**
- PascalCase for public methods: `reportPhishing()`, `setReportURL()`, `sendReportNotificationToServer()` (`Ribbon.cs` lines 60, 73; `GoPhishIntegration.cs` lines 28, 53)
- Prefix methods by action: `set`, `get`, `send`, `report`, `Get`, `Calculate`
- Example pattern: `GetBasicInfo()`, `GetCurrentUserInfos()`, `GetURLsAndAttachmentsInfo()`, `GetPluginDetails()` (`Ribbon.cs` lines 202-340)

**Variables:**
- camelCase for local variables and parameters: `headers`, `user_id`, `report_url`, `reportedItemType` (`GoPhishIntegration.cs` lines 28-42; `Ribbon.cs` lines 76-77)
- Underscore_separated for static configuration variables: `GoPhishURL`, `URLrequest`, `GoPhishHeader`, `WebExpID` (`GoPhishIntegration.cs` lines 21-25)
- String variables often descriptive: `reportedItemHeaders`, `simulatedPhishingURL`, `simulatedPhishingResponse` (`Ribbon.cs` lines 78, 138, 142)

**Constants:**
- PascalCase with `public const`: `_strTls12`, `Tls12` (`GoPhishIntegration.cs` lines 50-51)
- Regex patterns as constants: `HeaderRegex`, `TransportMessageHeadersSchema` (`Ribbon.cs` lines 411-415)
- Configuration string constants like `emailAtChar = "@"` (line 275)

**Properties/Settings:**
- Access via `Properties.Settings.Default`: `Properties.Settings.Default.gophish_url`, `Properties.Settings.Default.infosec_email` (`GoPhishIntegration.cs` line 21; `Ribbon.cs` line 124)
- Settings are snake_case: `gophish_url`, `infosec_email`, `gophish_listener_port`, `gophish_custom_header` (see `app.config`)

## Code Style

**Formatting:**
- No explicit formatter detected (no .editorconfig, prettier, or StyleCop configuration)
- Indentation: 4 spaces (standard .NET)
- Brace style: Allman (opening brace on new line) - typical for C#
  ```csharp
  if(condition)
  {
      // code
  }
  ```

**Namespaces:**
- Single namespace per project: `namespace PhishingReporter` (line 16, `Ribbon.cs`; line 15, `GoPhishIntegration.cs`)
- No nested namespaces used

**Using Statements:**
- Organized alphabetically in most files
- Mixed ordering observed: `using System` statements often come first, then Office-specific imports
- Example from `Ribbon.cs` (lines 7-21):
  ```csharp
  using Microsoft.Office.Core;
  using PhishingReporter.Properties;
  using System;
  using System.Drawing;
  using System.IO;
  using System.Linq;
  // ... etc
  ```

## Import Organization

**Order (observed pattern):**
1. Microsoft.Office namespace imports
2. PhishingReporter application namespaces
3. System namespace imports (System, System.Collections, System.IO, etc.)
4. Specialized imports (HtmlAgilityPack, Security)

**Aliases:**
- Office interop libraries use aliases for clarity:
  ```csharp
  using Outlook = Microsoft.Office.Interop.Outlook;
  using Office = Microsoft.Office.Core;
  ```

**No explicit path aliases:** Direct namespace imports used throughout.

## Error Handling

**Patterns:**
- Broad `catch (System.Exception)` blocks used (`Ribbon.cs` lines 121-193, `GoPhishIntegration.cs` lines 57-68)
- Generic exception catching without specific exception type logging
- Silent failures in some cases: `catch (System.Exception exc) { return "ERROR"; }` (`GoPhishIntegration.cs` lines 64-67)
- User-facing error messages via `MessageBox.Show()` (`Ribbon.cs` lines 82, 86, 150, 184, 197)
- Nested try-catch blocks for multiple operation safety (`Ribbon.cs` lines 268-290)

**Error Recovery:**
- Try-catch-finally pattern with `using` statements for resource cleanup (`Ribbon.cs` lines 385-392)
- File operations wrapped in try-catch: URL parsing and domain extraction (`Ribbon.cs` lines 268-290)

**Validation:**
- Null checks used: `if(urlNodes != null)` (`Ribbon.cs` line 258)
- Selection count validation: `if(selection.Count < 1)`, `if(selection.Count > 1)` (`Ribbon.cs` lines 80, 84)
- Type checking with `is` operator: `if (selection[1] is Outlook.MeetingItem || ...)` (`Ribbon.cs` line 90)
- String comparison: `if(group.ToString().Trim()!=string.Empty)` (`GoPhishIntegration.cs` line 35)

## Logging

**Framework:** `Console.WriteLine()` only (see `Ribbon.cs` line 287)

**Patterns:**
- Minimal logging in production code
- Debug comments for disabled functionality: `// MessageBox.Show(...)` for debug output (`Ribbon.cs` lines 144, 171, 191)
- Console.WriteLine used for exception logging: `Console.WriteLine("Bad url: {0}", emailDomain);` (`Ribbon.cs` line 287)

**Message Formats:**
- String formatting with placeholders: `Console.WriteLine("Bad url: {0}", emailDomain);`
- Direct string concatenation for multi-line messages

## Comments

**When to Comment:**
- File headers with developer info and license:
  ```csharp
  /*
   * Developer: Abdulla Albreiki
   * Github: https://github.com/0dteam
   * licensed under the GNU General Public License v3.0
   */
  ```
- Inline comments for non-obvious logic: `// Extract GoPhish Custom Header (X-GOPHISH-AJSMN: USERID0123)` (`GoPhishIntegration.cs` line 30)
- Region markers for code organization: `#region Helpers`, `#region IRibbonExtensibility Members` (`Ribbon.cs` lines 342, 361)
- TODO comments for incomplete features: `// TODO: Follow these steps to enable the Ribbon (XML) item:` (`Ribbon.cs` line 24)

**JSDoc/TSDoc:**
- Not used. C# uses XML documentation comments sparingly if at all.
- No `///` summary tags observed in main code

**Debug Comments:**
- Debug statements left in production code:
  ```csharp
  // DEBUG: to check if reporting email reaches GoPhish Portal
  // MessageBox.Show(simulatedPhishingURL + " --- " + simulatedPhishingResponse);
  ```
  (`Ribbon.cs` lines 143-144)

## Function Design

**Size:**
- Functions range from 10-80+ lines
- `reportPhishingEmailToSecurityTeam()` is 127 lines (lines 73-199 in `Ribbon.cs`) - large, handling multiple concerns
- Helper functions like `GetBasicInfo()` are 13 lines (lines 202-214)
- `setReportURL()` is 20 lines (lines 28-48 in `GoPhishIntegration.cs`)

**Parameters:**
- Single parameter pattern common: `string headers`, `string reportURL` (`GoPhishIntegration.cs` lines 28, 53)
- No default parameters observed
- Object casting used for type flexibility: `Object mailItemObj = (selection[1] as object) as Object;` (`Ribbon.cs` line 115)

**Return Values:**
- String returns for status: "OK", "ERROR", "NaN" (`GoPhishIntegration.cs` line 62; `GoPhishIntegration.cs` line 47)
- String returns for UI text: formatted report content
- Void returns for UI events: `reportPhishing()`, `Ribbon_Load()` (`Ribbon.cs` lines 60, 354)
- Boolean returns in LINQ operations: `.Any()`, `.Contains()` (`Ribbon.cs` lines 420, 258)

## Module Design

**Exports:**
- Public static methods for utility functions: `setReportURL()`, `sendReportNotificationToServer()` (`GoPhishIntegration.cs`)
- Public instance methods for Ribbon callbacks: `reportPhishing()`, `getGroup1Image()` (`Ribbon.cs` lines 60, 54)
- Private helper methods for internal operations: `reportPhishingEmailToSecurityTeam()` (`Ribbon.cs` line 73)

**Access Modifiers:**
- Public for interface-facing methods
- Private for internal helpers
- Static for utility classes: `static class GoPhishIntegration` (`GoPhishIntegration.cs` line 18)

**Barrel Files:**
- Not used. Each class in separate file.
- Extension methods added to namespace directly: `MailItemExtensions` in `Ribbon.cs` (lines 409-442)

## Patterns Observed

**Static Utility Classes:**
- `GoPhishIntegration` contains static methods for GoPhish integration logic
- Configuration accessed via `Properties.Settings.Default` static property

**Extension Methods:**
- `MailItemExtensions` class extends `MailItem` with header parsing: `Headers()`, `HeaderLookup()`, `HeaderString()`
- Enables: `mailItem.HeaderString()` pattern (`Ribbon.cs` line 130)

**Regex Patterns:**
- Compiled regex constants: `WebExpID`, `WebExpPrefix`, `HeaderRegex` with raw string patterns
- Example: `@"^(?<header_key>[-A-Za-z0-9]+)..."` (`Ribbon.cs` line 412)

**String Obfuscation:**
- Hyperlinks escaped to prevent accidental clicks: `att.Value.Replace(":", "[:]")` (`Ribbon.cs` line 266)
- Used for reporting safety: domains and URLs displayed as `https[:]//example.com`

**Configuration Pattern:**
- Settings stored in `app.config` and accessed via `Properties.Settings.Default`
- Mutable counters: `Properties.Settings.Default.gophish_reports_counter++` (`Ribbon.cs` line 147)

---

*Convention analysis: 2026-02-25*
