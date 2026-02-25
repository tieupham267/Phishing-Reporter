# Architecture

**Analysis Date:** 2026-02-25

## Pattern Overview

**Overall:** Plugin-based event-driven architecture with single-responsibility integration points

**Key Characteristics:**
- VSTO (Visual Studio Tools for Office) add-in architecture
- Event-driven UI with ribbon extensibility
- Layered separation between Outlook interop, email processing, and external integrations
- Stateful configuration through application settings
- Synchronous request-response model with GoPhish integration

## Layers

**Presentation Layer (Ribbon UI):**
- Purpose: User interface for reporting phishing emails; provides menu buttons and right-click context menu
- Location: `PhishingReporter/Ribbon.cs`, `PhishingReporter/Ribbon.xml`
- Contains: Ribbon UI callbacks, button event handlers, image resources
- Depends on: Outlook interop, email processing logic, configuration settings
- Used by: Outlook application event system

**Email Processing Layer:**
- Purpose: Extract, analyze, and prepare phishing email details for reporting
- Location: `PhishingReporter/Ribbon.cs` (helper methods within reportPhishingEmailToSecurityTeam)
- Contains: Email parsing, URL extraction, attachment analysis, hash calculation
- Depends on: HtmlAgilityPack, System.Security.Cryptography, Outlook interop
- Used by: Ribbon event handlers

**Integration Layer:**
- Purpose: Handle external service communication (GoPhish detection and notification)
- Location: `PhishingReporter/GoPhishIntegration.cs`
- Contains: GoPhish header parsing, HTTP reporting, TLS configuration
- Depends on: System.Net, configuration settings
- Used by: Email processing layer (called during report workflow)

**Add-In Lifecycle Layer:**
- Purpose: Manage VSTO add-in initialization and cleanup
- Location: `PhishingReporter/ThisAddIn.cs`, `PhishingReporter/ThisAddIn.Designer.cs`
- Contains: Startup/shutdown event handlers, ribbon extensibility object creation
- Depends on: Office interop, Ribbon class
- Used by: Outlook application runtime

**Configuration Layer:**
- Purpose: Centralized settings management and persistence
- Location: `PhishingReporter/Properties/Settings.Designer.cs`, `PhishingReporter/app.config`
- Contains: User-scoped settings (email addresses, URLs, counters), default values
- Depends on: System.Configuration
- Used by: All layers for reading configuration values

## Data Flow

**Standard Phishing Report Workflow:**

1. User selects email in Outlook and clicks "Report Phishing" button (from ribbon or context menu)
2. `Ribbon.reportPhishing()` displays confirmation dialog
3. On confirmation, `reportPhishingEmailToSecurityTeam()` executes:
   - Validates single email is selected
   - Creates new report email to infosec_email
   - Attaches original email
   - Prepares report body with extracted email details
4. Email details extracted:
   - User information via Exchange API (`GetCurrentUserInfos()`)
   - Basic metadata (`GetBasicInfo()`)
   - URLs and attachments with hashes (`GetURLsAndAttachmentsInfo()`)
   - Email headers via MAPI (`HeaderString()`)
5. GoPhish header check via `GoPhishIntegration.setReportURL()`:
   - If custom header found → simulated campaign detected
   - Calls `GoPhishIntegration.sendReportNotificationToServer()`
   - Updates `gophish_reports_counter`
   - Shows success message
6. If not GoPhish campaign:
   - Prepares full report email body with all extracted details
   - Saves and auto-sends report email to infosec address
   - Updates `suspecious_reports_counter`
   - Optionally shows thank you message
7. Original email deleted from Inbox
8. On error: Auto-sends error report to support_email with exception details

**GoPhish Integration Workflow:**

1. `GoPhishIntegration.setReportURL()` extracts custom header via regex:
   - Pattern: `X-GOPHISH-AJSMN: [0-9a-zA-Z]+`
   - Extracts user ID from header value
   - Constructs report URL: `http://gophish_url:gophish_listener_port/report?rid=USERID`
2. `GoPhishIntegration.sendReportNotificationToServer()` sends HTTP GET request:
   - Enforces TLS 1.2
   - Returns "OK" on success, "ERROR" on failure

**State Management:**
- Settings stored in Windows user profile via VSTO application settings
- Counters (`suspecious_reports_counter`, `gophish_reports_counter`) track reporting activity
- No in-memory state; all state persisted to configuration

## Key Abstractions

**Email Item Extension:**
- Purpose: Parse and extract email headers from MAPI storage
- Examples: `PhishingReporter/Ribbon.cs` (lines 409-441)
- Pattern: Extension methods on MailItem (HeaderString, HeaderLookup, Headers)
- Encapsulates regex-based header parsing and MAPI property access

**Report Email Generator:**
- Purpose: Compose formatted report email body with standardized sections
- Examples: `GetCurrentUserInfos()`, `GetBasicInfo()`, `GetURLsAndAttachmentsInfo()`, `GetPluginDetails()`
- Pattern: Multiple focused helper methods that build sections independently
- Each method returns formatted string section; combined in main workflow

**GoPhish Header Parser:**
- Purpose: Detect simulated phishing campaigns by extracting custom tracking headers
- Examples: `PhishingReporter/GoPhishIntegration.cs` (lines 28-48)
- Pattern: Static utility class with regex matching and string manipulation
- Returns "NaN" when header not found (no type checking; string sentinel value)

**Hash Calculators:**
- Purpose: Generate MD5 and SHA256 hashes for attachment forensics
- Examples: `CalculateMD5()`, `GetHashSha256()`
- Pattern: Helper methods that read file, compute hash, return hex string
- Used for attachment integrity verification in reports

## Entry Points

**Ribbon Button Click:**
- Location: `PhishingReporter/Ribbon.cs` - `reportPhishing()` method (lines 60-67)
- Triggers: User clicks "Report Phishing" button in Home ribbon, Read Message ribbon, or email context menu
- Responsibilities: Display confirmation dialog, route to main report workflow

**Context Menu Action:**
- Location: `PhishingReporter/Ribbon.xml` - ContextMenuMailItem binding
- Triggers: User right-clicks email and selects "Report Phishing"
- Responsibilities: Same as ribbon button (routed to same reportPhishing handler)

**Add-In Startup:**
- Location: `PhishingReporter/ThisAddIn.cs` - `ThisAddIn_Startup()` method (lines 19-22)
- Triggers: Outlook application starts and loads add-in
- Responsibilities: Currently empty (no initialization logic)

## Error Handling

**Strategy:** Try-catch with user-facing error notification and automatic error email to support team

**Patterns:**

1. **URI Parsing Errors** (lines 268-290 in Ribbon.cs):
   - Catches `UriFormatException` during domain extraction
   - Falls back to alternative parsing for mailto: links
   - Silently ignores malformed URLs with console output

2. **Workflow Errors** (lines 182-193 in Ribbon.cs):
   - Catches generic `System.Exception` in main report workflow
   - Shows user: "There was an error! An automatic email was sent to the support to resolve the issue."
   - Creates and sends error report to support_email with exception message
   - Error email subject: `[Outlook Addin Error]`

3. **GoPhish Network Errors** (lines 64-67 in GoPhishIntegration.cs):
   - Catches generic exception from HTTP request
   - Returns "ERROR" string (no exception thrown to caller)
   - Caller doesn't distinguish between network failure and successful response

## Cross-Cutting Concerns

**Logging:** Console.WriteLine used in one location (malformed URL logging in Ribbon.cs line 287); no centralized logging framework

**Validation:**
- Selection count validation (must be exactly 1 email selected)
- Item type validation (only MailItem, MeetingItem, ContactItem, etc. supported)
- Configuration validation: relies on default values; no startup validation of required settings

**Authentication:**
- Uses Outlook's built-in authentication for Exchange user information
- GoPhish integration uses unauthenticated HTTP (custom header provides tracking, not security)
- No API keys or bearer tokens used

**Configuration Access:**
- All config reads via `Properties.Settings.Default.setting_name`
- Settings are user-scoped and persisted in Windows user profile
- No validation that required settings are configured before use

---

*Architecture analysis: 2026-02-25*
