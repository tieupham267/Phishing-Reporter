# Codebase Structure

**Analysis Date:** 2026-02-25

## Directory Layout

```
Phishing-Reporter/
├── PhishingReporter/              # Main VSTO add-in project
│   ├── GoPhishIntegration.cs      # GoPhish integration logic
│   ├── Ribbon.cs                  # UI event handlers and email processing
│   ├── Ribbon.xml                 # Ribbon UI definition
│   ├── ThisAddIn.cs               # VSTO add-in lifecycle
│   ├── ThisAddIn.Designer.cs      # Auto-generated VSTO code
│   ├── ThisAddIn.Designer.xml     # VSTO manifest
│   ├── app.config                 # Application configuration (settings)
│   ├── packages.config            # NuGet dependencies
│   ├── PhishingReporter.csproj    # C# project file
│   ├── phishing.ico               # Application icon
│   ├── Properties/                # Project properties and settings
│   │   ├── AssemblyInfo.cs        # Assembly metadata
│   │   ├── Settings.Designer.cs   # Auto-generated settings class
│   │   ├── Settings.settings      # Settings definition
│   │   ├── Resources.Designer.cs  # Auto-generated resources class
│   │   └── Resources.resx         # Resource file (contains phishing icon)
│   ├── Resources/                 # Embedded resources
│   │   └── phishing.png           # Icon for UI buttons
│   └── CREDITS.txt                # Attribution file
├── Installer/                      # WiX installer project
│   └── Installer.vdproj           # Visual Studio installer definition
├── PhishingReporter.sln           # Solution file
├── README.md                       # Project documentation
├── LICENSE                         # GNU GPL v3.0
├── phishing.ico                   # Icon (also in Properties)
├── phishing.png                   # Logo
└── splash.psd                     # Installer splash screen design
```

## Directory Purposes

**PhishingReporter/:**
- Purpose: Core VSTO add-in implementation
- Contains: C# source files, configuration, embedded resources
- Key files: `Ribbon.cs` (900+ lines), `ThisAddIn.cs` (50 lines), `GoPhishIntegration.cs` (70 lines)

**PhishingReporter/Properties/:**
- Purpose: Project metadata, settings, and embedded resources
- Contains: Assembly info, auto-generated settings classes, resource files
- Key files: `Settings.Designer.cs` (auto-generated from Settings.settings)

**PhishingReporter/Resources/:**
- Purpose: Embedded image resources for UI
- Contains: PNG icon file for ribbon buttons
- Key files: `phishing.png` (icon displayed in Home and Read Message ribbon)

**Installer/:**
- Purpose: Windows installer package definition
- Contains: Visual Studio Installer Projects (.vdproj) configuration
- Key files: `Installer.vdproj` (defines MSI generation)

## Key File Locations

**Entry Points:**
- `PhishingReporter/ThisAddIn.cs`: VSTO add-in initialization and ribbon creation
- `PhishingReporter/Ribbon.cs` (lines 60-67): `reportPhishing()` method - main user-triggered entry point

**Configuration:**
- `PhishingReporter/app.config`: Default and user-scoped settings (email addresses, GoPhish URL, port, counters)
- `PhishingReporter/Properties/Settings.settings`: Settings definition file
- `PhishingReporter/Properties/Settings.Designer.cs`: Auto-generated settings accessor class

**Core Logic:**
- `PhishingReporter/Ribbon.cs` (lines 73-200): `reportPhishingEmailToSecurityTeam()` - main workflow
- `PhishingReporter/Ribbon.cs` (lines 202-327): Helper methods for extracting email details
- `PhishingReporter/GoPhishIntegration.cs` (lines 28-68): GoPhish campaign detection and reporting

**UI Definition:**
- `PhishingReporter/Ribbon.xml`: Ribbon UI layout and button definitions
- `PhishingReporter/ThisAddIn.Designer.xml`: VSTO manifest

**Extensions:**
- `PhishingReporter/Ribbon.cs` (lines 409-441): `MailItemExtensions` class - email header parsing

## Naming Conventions

**Files:**
- PascalCase for C# source files: `ThisAddIn.cs`, `Ribbon.cs`, `GoPhishIntegration.cs`
- Designer-generated files suffixed with `.Designer.cs`: `ThisAddIn.Designer.cs`, `Settings.Designer.cs`
- XML files for VSTO: `Ribbon.xml`, `ThisAddIn.Designer.xml`
- Config files: lowercase `app.config`, `packages.config`

**Classes:**
- PascalCase: `Ribbon`, `GoPhishIntegration`, `MailItemExtensions`
- Static utility classes: `GoPhishIntegration` (static methods only)
- Extension methods: `MailItemExtensions` (static class with extension methods)

**Methods:**
- camelCase for public methods: `reportPhishing()`, `getGroup1Image()`, `sendReportNotificationToServer()`
- camelCase for private helpers: `reportPhishingEmailToSecurityTeam()`, `CalculateMD5()`, `GetURLsAndAttachmentsInfo()`
- Mix of naming styles (some PascalCase in helpers like `GetBasicInfo()`)

**Variables:**
- camelCase for local variables: `domainsInEmail`, `emailHTML`, `urlsText`
- camelCase for parameters: `mailItem`, `control`, `reportURL`
- Sentinel values as strings: `"NaN"` (GoPhish integration returns "NaN" when header not found)

**Constants/Settings:**
- snake_case for configuration keys: `infosec_email`, `gophish_url`, `gophish_listener_port`, `gophish_custom_header`
- Accessed via `Properties.Settings.Default.setting_name`

## Where to Add New Code

**New Feature (Email Analysis):**
- Primary code: `PhishingReporter/Ribbon.cs` (add method in helper methods section, lines 201+)
- Call from: Main workflow in `reportPhishingEmailToSecurityTeam()` (around line 160+)
- Email body building: String concatenation to `reportEmail.Body`

**New External Integration:**
- Implementation: Create new file `PhishingReporter/[Service]Integration.cs` (follow GoPhishIntegration pattern)
- Pattern: Static class with public static methods
- Call from: `reportPhishingEmailToSecurityTeam()` after GoPhish check (lines 138-151)
- Configuration: Add settings to `app.config` and `Properties/Settings.settings`

**New UI Control:**
- Ribbon definition: `PhishingReporter/Ribbon.xml` (add button or menu item in tabs or contextMenus)
- Callback method: Add new method to `Ribbon` class (e.g., `onAction="newButtonHandler"`)
- Image resource: Add to `PhishingReporter/Resources/` and reference in `Properties/Resources.resx`

**Utilities/Helpers:**
- Focused helpers: Add to appropriate existing class (e.g., hash functions in Ribbon.cs)
- Reusable extensions: Add to `MailItemExtensions` class in `Ribbon.cs` (lines 409-441)
- Email parsing: Add static methods to `MailItemExtensions`

**Tests:**
- Tests: Not currently present in codebase; should create `PhishingReporter.Tests/` project if adding
- Unit tests would target: Parsing logic (GoPhish header, URLs, domains), hash calculation
- Integration tests would target: Outlook interop, email creation, settings access

## Special Directories

**PhishingReporter/bin/ and obj/:**
- Purpose: Build output and intermediate files
- Generated: Yes (created during build)
- Committed: No (in .gitignore)

**PhishingReporter/Properties/:**
- Purpose: Project configuration, settings, and embedded resources
- Generated: Partially (Settings.Designer.cs, Resources.Designer.cs auto-generated from .settings and .resx)
- Committed: Yes (source .settings and .resx files committed; generated files optional)

**.vdproj Installer Projects:**
- Purpose: Windows installer (MSI) packaging configuration
- Generated: No (manually edited)
- Committed: Yes

## Architecture Dependencies

**External Packages:**
- `HtmlAgilityPack v1.11.23`: HTML parsing for URL extraction from email body
- `Microsoft.Office.Interop.Outlook`: Outlook VSTO interop
- `Microsoft.Office.Tools.v4.0.*`: VSTO framework

**Framework:**
- `.NET Framework 4.6.1`: Target framework
- `System.*`: Built-in assemblies for crypto, networking, XML, Windows Forms

**Build System:**
- MSBuild 15.0
- Visual Studio 2017+ (VSTO support)
- VSTO Runtime 4.0

---

*Structure analysis: 2026-02-25*
