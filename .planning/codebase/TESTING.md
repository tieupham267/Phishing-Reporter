# Testing Patterns

**Analysis Date:** 2026-02-25

## Test Framework Status

**Current State:** No testing framework detected

**Test Infrastructure Absent:**
- No `.csproj` references to NUnit, xUnit, or MSTest
- No test project files found in repository
- No test configuration files (`.runsettings`, `nunit.config`, `xunit.config`)
- No `*Test.cs` or `*Spec.cs` files in codebase
- No test assembly references in main project dependencies

**Packages Currently Deployed:**
- HtmlAgilityPack 1.11.23 (HTML parsing library)
- Microsoft Office Interop libraries (VSTO)
- .NET Framework 4.6.1 built-ins

## Test Framework Recommendation

**Recommended Approach for C# VSTO Add-in:**

Given the nature of this Outlook Add-in project, unit testing is challenging due to:
1. Heavy dependency on Microsoft Office Interop (COM objects)
2. UI-driven event handlers (Ribbon callbacks, button clicks)
3. File I/O operations (attachment processing, hash calculation)
4. External HTTP calls (GoPhish integration)

**Recommended Test Framework Stack:**
- **Unit Testing:** MSTest (built into Visual Studio, minimal setup)
- **Mocking:** Moq (can mock HTTP calls, file operations)
- **Alternative:** NUnit + NSubstitute for more flexible mocking

**Example Structure:** Create `PhishingReporter.Tests` project alongside main project

## Architecture Barriers to Testing

**Hard-to-Test Components:**

1. **Ribbon.cs (442 lines)** - `Ribbon` class (`Ribbon.cs` lines 46-407)
   - Inherits from `Office.IRibbonExtensibility`
   - Depends on `Globals.ThisAddIn` (static reference to add-in instance)
   - Creates `MailItem` objects via Office Interop
   - Calls `MessageBox.Show()` for UI feedback
   - Barrier: Cannot instantiate without active Outlook

2. **GoPhishIntegration.cs (71 lines)** - `GoPhishIntegration` static class (`GoPhishIntegration.cs` lines 18-70)
   - Static methods use static configuration variables
   - Hardcoded static field access: `Properties.Settings.Default`
   - Makes HTTP requests via `HttpWebRequest`
   - Barrier: Static state makes mocking difficult

3. **File Operations** (`Ribbon.cs` lines 312-325)
   - Attachments saved to temp folder: `Environment.ExpandEnvironmentVariables(@"%TEMP%\...")`
   - Hash calculation on files: `CalculateMD5()`, `GetHashSha256()`
   - Barrier: Requires actual file system access

4. **Outlook Interop Dependencies** (throughout `Ribbon.cs`)
   - Direct COM object interaction: `Globals.ThisAddIn.Application`
   - Selection management: `Selection selection = Globals.ThisAddIn.Application.ActiveExplorer().Selection`
   - Barrier: Requires Outlook running, cannot mock COM objects easily

## Testable Logic Identified

**Functions with testable logic (with refactoring):**

1. **Regex Parsing** - `GoPhishIntegration.setReportURL()` (`GoPhishIntegration.cs` lines 28-48)
   - Input: email headers string
   - Output: report URL or "NaN"
   - Testable with: String fixtures
   - Issues: Hard-coded Regex patterns as class members makes testing verbose

2. **Header Parsing** - `MailItemExtensions` class (`Ribbon.cs` lines 409-442)
   - `HeaderLookup()` method parses email header string
   - Input: email header raw string
   - Output: `ILookup<string, string>`
   - Testable with: String fixtures of email headers

3. **URL Extraction** - `GetURLsAndAttachmentsInfo()` (`Ribbon.cs` lines 245-326)
   - Extracts URLs, domains, and attachment info from email
   - Testable component: HTML parsing and domain extraction logic
   - Issues: Tightly coupled to MailItem object and file system

4. **Hash Calculation** - `CalculateMD5()`, `GetHashSha256()` (`Ribbon.cs` lines 383-403)
   - Input: file path
   - Output: hash string
   - Testable with: Temporary test files
   - Issues: Depends on actual file system

## Suggested Test Structure (if implemented)

### Project Layout:
```
PhishingReporter/
├── PhishingReporter.csproj
├── GoPhishIntegration.cs
├── Ribbon.cs
├── ThisAddIn.cs
└── Properties/

PhishingReporter.Tests/
├── PhishingReporter.Tests.csproj
├── GoPhishIntegrationTests.cs
├── MailItemExtensionsTests.cs
├── HashCalculationTests.cs
└── Fixtures/
    ├── EmailHeaderFixtures.cs
    └── TestData/
        ├── sample_headers.txt
        └── test_file.txt
```

### Example Test Cases (Pseudo-code):

**GoPhishIntegrationTests.cs:**
```csharp
[TestClass]
public class GoPhishIntegrationTests
{
    [TestMethod]
    public void SetReportURL_ValidGoPhishHeader_ReturnsURL()
    {
        // Arrange
        string headers = "X-GOPHISH-AJSMN: USER123";

        // Act
        string result = GoPhishIntegration.setReportURL(headers);

        // Assert
        Assert.IsTrue(result.Contains("rid=USER123"));
        Assert.AreNotEqual("NaN", result);
    }

    [TestMethod]
    public void SetReportURL_NoGoPhishHeader_ReturnsNaN()
    {
        // Arrange
        string headers = "From: test@example.com\r\nTo: user@example.com";

        // Act
        string result = GoPhishIntegration.setReportURL(headers);

        // Assert
        Assert.AreEqual("NaN", result);
    }
}
```

**MailItemExtensionsTests.cs:**
```csharp
[TestClass]
public class MailItemExtensionsTests
{
    [TestMethod]
    public void HeaderLookup_ValidHeaders_ParsesCorrectly()
    {
        // Arrange
        string headerString = "From: test@example.com\r\nSubject: Test Email\r\n";

        // Act
        var lookup = /* parse headers */;

        // Assert
        Assert.IsTrue(lookup.Contains("From"));
        Assert.IsTrue(lookup.Contains("Subject"));
    }
}
```

## Current Code Testing Gaps

**Untested Areas (High Risk):**

1. **GoPhish Integration Flow** (`GoPhishIntegration.sendReportNotificationToServer()`)
   - HTTP request sending
   - TLS 1.2 security protocol setup
   - Exception handling for network failures
   - Gap: No verification of successful reporting to GoPhish

2. **Email Processing Pipeline** (`reportPhishingEmailToSecurityTeam()`)
   - Email type detection (MailItem, ContactItem, etc.)
   - Email body construction with user/security info
   - Report email sending
   - Original email deletion
   - Gap: No verification of email construction or sending

3. **Data Extraction** (`GetURLsAndAttachmentsInfo()`)
   - HTML parsing and link extraction
   - Domain extraction from URLs
   - Attachment hash calculation (MD5, SHA256)
   - Gap: No validation of extracted data accuracy

4. **Configuration Validation**
   - Settings loaded from `app.config`
   - Default values if settings missing
   - Gap: No fallback or validation for required settings

5. **User Input Validation**
   - Selection validation (empty, multiple, invalid types)
   - Header validation before processing
   - URL validation before extraction
   - Gap: Minimal validation, user-facing only

## Code Quality Observations

**Testability Issues:**

1. **Static Dependencies:**
   - `Properties.Settings.Default` accessed directly throughout code
   - Cannot inject test configurations
   - Recommendation: Extract to injectable `IConfiguration` interface

2. **Global State:**
   - `Globals.ThisAddIn` referenced directly
   - Cannot test without Outlook running
   - Recommendation: Inject `IOutlookApplication` dependency

3. **Hard-coded String Constants:**
   - Report URL template: `URLrequest = ... + "/report?rid=USERID"`
   - Custom header name: `GoPhishHeader = Properties.Settings.Default.gophish_custom_header`
   - Recommendation: Move to configuration object

4. **Mixed Concerns:**
   - `Ribbon.cs` handles UI, business logic, and data extraction
   - Single responsibility violated
   - Recommendation: Extract business logic to separate service classes

5. **Exception Swallowing:**
   - `catch (System.Exception exc) { return "ERROR"; }` in `sendReportNotificationToServer()`
   - Exception details lost
   - Recommendation: Log exception details before returning error

## Manual Testing Evidence

**Tested Configuration (from README.md):**
- GoPhish v0.12.1 Windows version (December 2023)
- Outlook 2019 (x64)
- Windows OS (.NET Framework 4.6.1 compatible)

**Manual Test Scenarios Identified:**
1. GoPhish integration report flow verification
2. Email selection and reporting
3. Attachment hash calculation and reporting
4. URL and domain extraction from email HTML
5. Configuration updates and persistence

## Refactoring Path for Testability

**Phase 1: Extract Business Logic**
1. Create `GoPhishReportService` class to handle GoPhish HTTP requests
2. Create `EmailProcessingService` class to extract email data
3. Create `ConfigurationService` to replace direct `Properties.Settings` access

**Phase 2: Introduce Dependency Injection**
1. Add interfaces: `IGoPhishReportService`, `IEmailProcessingService`, `IConfigurationService`
2. Inject into `Ribbon` class constructor
3. Remove static `Globals.ThisAddIn` dependency

**Phase 3: Add Unit Tests**
1. Test service classes with mocked dependencies
2. Test Regex patterns with string fixtures
3. Test hash calculations with temporary files

**Phase 4: Integration Tests**
1. Test complete report flow with mocked Outlook
2. Test file operations with temp directories

---

*Testing analysis: 2026-02-25*
