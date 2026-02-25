# Technology Stack

**Analysis Date:** 2026-02-25

## Languages

**Primary:**
- C# (.NET Framework 4.6.1) - Core plugin implementation and business logic
- XML - Ribbon UI customization and configuration

**Secondary:**
- VB/MSBuild - Project configuration and build management

## Runtime

**Environment:**
- .NET Framework 4.6.1
- Microsoft Visual Studio Tools for Office (VSTO) 4.0
- Requires Windows OS with Microsoft Outlook 2010 or later

**Package Manager:**
- NuGet - Manages .NET package dependencies
- Lockfile: `packages.config` - Lists all NuGet package references

## Frameworks

**Core:**
- VSTO (Visual Studio Tools for Office) 4.0 - Outlook add-in framework
  - Enables COM interop with Office Object Models
  - Provides managed code execution within Outlook

**UI:**
- Office Ribbon UI (RibbonX) - Custom ribbon customization in Outlook
  - Defined in `PhishingReporter/Ribbon.xml`
  - Integrates into Mail and Read Message tabs

**HTML Processing:**
- HtmlAgilityPack 1.11.23 - HTML parsing and XPath-based element extraction
  - Used in `PhishingReporter/Ribbon.cs` for extracting URLs from email HTML

## Key Dependencies

**Critical:**
- HtmlAgilityPack 1.11.23 - HTML DOM parsing library for extracting URLs and domains from email bodies
  - Located: `packages\HtmlAgilityPack.1.11.23\lib\Net45\HtmlAgilityPack.dll`

**Framework/Interop:**
- Microsoft.Office.Tools.v4.0.Framework (v10.0.0.0) - VSTO core framework
- Microsoft.Office.Tools.Outlook.v4.0.Utilities (v10.0.0.0) - Outlook-specific VSTO utilities
- Microsoft.Office.Interop.Outlook (v15.0.0.0) - Outlook Object Model interop
- Office Primary Interop Assemblies (PIA) v15.0 - Core Office functionality

**System Libraries:**
- System.Windows.Forms - UI message boxes and dialogs
- System.Security.Cryptography - MD5 and SHA256 hash computation for attachments
- System.Text.RegularExpressions - Regex parsing for email headers and GoPhish integration
- System.Net - HTTP communication with GoPhish server
- System.Xml.Linq - XML processing

## Configuration

**Environment:**
- Configured via `PhishingReporter/app.config` (user settings)
- Settings stored in `PhishingReporter/Properties/Settings.settings`
- Settings are user-configurable at installation time

**Key Configuration:**
- `infosec_email` - Email address of Information Security Team
- `gophish_url` - URL of GoPhish server (e.g., http://gophish.example.com)
- `gophish_listener_port` - GoPhish listener port (typically 80)
- `gophish_custom_header` - Custom header name for tracking (e.g., X-GOPHISH-AJSMN)
- `support_email` - Email for bug reports
- `plugin_version` - Current version identifier

**Build:**
- Project file: `PhishingReporter/PhishingReporter.csproj`
- Solution file: `PhishingReporter.sln`
- Installer project: `Installer/Installer.vdproj`
- Debug output: `PhishingReporter/bin/Debug/`
- Release output: `PhishingReporter/bin/Release/`

## Platform Requirements

**Development:**
- Windows OS (Visual Studio runs on Windows)
- Visual Studio 2010 or later with VSTO support
- Microsoft Visual Studio Installer Projects component (for building installer)
- .NET Framework 4.6.1 or later
- Office 2010+ (32-bit or 64-bit)

**Production:**
- Windows OS
- Microsoft Outlook 2010 or later
- .NET Framework 4.6.1
- Matching Outlook architecture (32-bit plugin for 32-bit Office, 64-bit for 64-bit Office)
- VSTO 4.0 Runtime (installed via bootstrapper)

**Deployment:**
- Installer-based deployment using Visual Studio Installer Projects (`.vdproj`)
- Creates `.msi` installer with bootstrapper for VSTO runtime
- Installation target: `publish\` directory after build

## Build Configuration

**Debug:**
- Configuration: `Debug|AnyCPU`
- Debug symbols: Enabled
- Output: `bin\Debug\`
- Optimization: Off
- Warning level: 4

**Release:**
- Configuration: `Release|AnyCPU`
- Debug symbols: PDB only
- Output: `bin\Release\`
- Optimization: On
- Warning level: 4

**Code Signing:**
- Manifest signing enabled
- Certificate: Temporary key file `PhishingReporter_TemporaryKey.pfx`
- Manifest certificate thumbprint: `49F4622D1A7A5798A0337CEE03366CECC8D37E0C`

---

*Stack analysis: 2026-02-25*
