# External Integrations

**Analysis Date:** 2026-02-25

## APIs & External Services

**GoPhish Phishing Framework:**
- Service: GoPhish (Open Source Phishing Framework)
  - What it's used for: Integration with simulated phishing campaigns for tracking reported emails
  - SDK/Client: Custom HTTP client via `System.Net.HttpWebRequest`
  - Implementation: `PhishingReporter/GoPhishIntegration.cs`
  - Auth: Custom header-based tracking via `gophish_custom_header` setting

**SMTP Email Service:**
- Service: Microsoft Outlook/Exchange SMTP
  - What it's used for: Sending phishing reports to security team and support notifications
  - Protocol: Native Outlook Object Model via COM interop
  - Implementation: `PhishingReporter/Ribbon.cs` - Methods `reportPhishingEmailToSecurityTeam()` and error handling

## Data Storage

**Databases:**
- None - Plugin does not use database storage

**File Storage:**
- Local filesystem only - Temporary attachment storage
  - Location: User's `%TEMP%` directory
  - Usage: Temporary storage of attachments for hash calculation
  - Implementation: `PhishingReporter/Ribbon.cs` line 312

**Configuration Storage:**
- Windows Registry (via user settings)
  - Persists plugin settings like email addresses, GoPhish URL, counters
  - Schema: `HKEY_CURRENT_USER\Software\[Company]\[Product]`

**Caching:**
- None - No caching layer used

## Authentication & Identity

**Auth Provider:**
- Custom (Windows Domain/Outlook Session)
  - Implementation: Uses Outlook session context for current user
  - User info extraction: `PhishingReporter/Ribbon.cs` - `GetCurrentUserInfos()` method
  - Extracts: Domain, username, machine name, email, title, department via Outlook API

**GoPhish Integration Auth:**
- Custom header-based tracking
  - Header name: Configured via `gophish_custom_header` setting
  - Header value: GoPhish tracking ID (`{{.RId}}` template variable)
  - Parsing: Regex-based extraction in `GoPhishIntegration.cs` line 31

## Monitoring & Observability

**Error Tracking:**
- Custom error email notification
  - Implementation: `PhishingReporter/Ribbon.cs` lines 184-192
  - Errors sent to: `support_email` setting value
  - Contents: Exception message and full stack trace

**Logs:**
- No centralized logging
- Error information captured via email to support team
- Debug messages: Commented code shows message box popups for development (line 144)

## CI/CD & Deployment

**Hosting:**
- Desktop/Client-side deployment
- Distributed via Windows Installer (.msi) package

**CI Pipeline:**
- None detected - Manual build via Visual Studio

**Deployment Approach:**
- Visual Studio Installer Projects
  - Installer project: `Installer/Installer.vdproj`
  - Builds to: `PhishingReporter\Installer\Release\` folder
  - Includes bootstrapper for VSTO 4.0 runtime
  - Supports both 32-bit and 64-bit Office installations
  - Update interval: Configurable (default 7 days)

## Environment Configuration

**Required env vars:**
- None - All configuration stored in `app.config` and Windows Registry

**Configuration vars required:**
- `infosec_email` - Information Security Team email (e.g., `Information Security Team <infosec@example.com>`)
- `gophish_url` - GoPhish server URL (e.g., `http://gophish.example.com`)
- `gophish_listener_port` - GoPhish listener port (e.g., `80`)
- `gophish_custom_header` - Custom header name for campaign tracking (e.g., `X-GOPHISH-AJSMN`)
- `support_email` - Plugin support team email for bug reports
- `plugin_version` - Version identifier (e.g., `V1.1`)

**Secrets location:**
- Stored in Windows Registry via user settings
- No secrets in source code
- Email addresses and URLs configured at installation time

## Webhooks & Callbacks

**Incoming:**
- None - Plugin does not expose webhooks

**Outgoing:**
- GoPhish Report Callback
  - Endpoint: `{gophish_url}:{gophish_listener_port}/report?rid={tracking_id}`
  - Triggered by: User clicking "Report Phishing" button in Outlook
  - Implementation: `GoPhishIntegration.cs` - `sendReportNotificationToServer()` method
  - Protocol: HTTP GET request
  - Response: Expected 200 OK response
  - Error handling: Returns "ERROR" on failure, "OK" on success

- Support Email Notification
  - Triggered by: Plugin errors during phishing report process
  - Recipient: `support_email` configuration
  - Implementation: `PhishingReporter/Ribbon.cs` lines 186-192
  - Contents: Exception details with timestamp and user context

## Network Requirements

**TLS/SSL:**
- TLS 1.2 forced for GoPhish communication
  - Configured in: `GoPhishIntegration.cs` line 55
  - `ServicePointManager.SecurityProtocol = Tls12;`

**Firewall/Network:**
- Outbound HTTP/HTTPS to GoPhish server required
- Outbound SMTP to mail server required (via Outlook)
- No inbound ports required

## Integration Flow

**Phishing Report Flow:**

1. User selects email in Outlook
2. User clicks "Report Phishing" button (Ribbon or context menu)
3. Plugin validates single email selected
4. Plugin extracts email headers using Outlook Object Model
5. Plugin searches headers for GoPhish tracking header (e.g., `X-GOPHISH-AJSMN`)
6. If GoPhish header found:
   - Extract tracking ID from header value
   - Build report URL: `{gophish_url}:{port}/report?rid={tracking_id}`
   - Send HTTP GET to GoPhish server via `sendReportNotificationToServer()`
   - Increment `gophish_reports_counter`
   - Show success message to user
   - Delete reported email
7. If GoPhish header NOT found (real phishing):
   - Extract email content:
     - User information (domain, username, title, department via Outlook)
     - Email metadata (sender, subject, folder, OS, Outlook version)
     - URLs and domains from HTML body via HtmlAgilityPack
     - Attachments with MD5/SHA256 hashes
     - Full email headers
   - Compose report email to `infosec_email`
   - Attach original email message
   - Add `[POTENTIAL PHISH]` prefix to subject
   - Send report via Outlook SMTP
   - Increment `suspecious_reports_counter`
   - Delete reported email

**Error Handling Flow:**

1. If exception occurs during report process
2. Catch exception and extract message
3. Create new email to `support_email`
4. Subject: `[Outlook Addin Error]`
5. Body: Full exception message with stack trace
6. Send via Outlook
7. Show user-friendly error message

---

*Integration audit: 2026-02-25*
