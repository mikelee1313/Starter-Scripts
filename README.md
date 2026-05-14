# Starter Scripts

A collection of PowerShell starter templates for Microsoft 365 automation using the Microsoft Graph API, PnP.PowerShell, and various OAuth 2.0 authentication flows. Each script is a ready-to-use baseline — fill in your configuration and add your own logic.

## Scripts

| Script | Auth Flow | Use Case |
|---|---|---|
| `Application-Auth-Graph.ps1` | App-only (Certificate or Client Secret) | Unattended Graph API automation |
| `Application-Auth-PNP.ps1` | App-only (Certificate or Client Secret) | SharePoint Online / PnP.PowerShell automation |
| `DeviceCodeAuth.ps1` | Delegated — Device Code | Graph API from headless/SSH environments |
| `Get-TokenWithPKCE.ps1` | Delegated — Authorization Code + PKCE | Interactive Graph API automation (browser sign-in) |

---

## Script Details

### Application-Auth-Graph.ps1

App-only authentication against Microsoft Graph using the **client credentials** flow. Supports both certificate-based JWT assertions and client secret authentication.

**When to use:** Fully unattended/service scenarios where no user context is needed (e.g., scheduled tasks, Azure Automation).

**Features:**
- Certificate (preferred) or client secret authentication
- Automatic token refresh on expiry
- Throttle-aware retry with exponential back-off (handles 429, 502, 503, 504)
- Automatic pagination via `Invoke-GraphPagedRequest`
- Built-in example calls (paginated GET, single-resource GET, POST, PATCH)

**Prerequisites:**
- Entra ID app registration with **application** permissions granted and admin-consented
- Certificate with private key in the Windows certificate store, **or** client secret stored in the `GRAPH_CLIENT_SECRET` environment variable

---

### Application-Auth-PNP.ps1

App-only authentication using **PnP.PowerShell** for SharePoint Online and Microsoft 365. Supports certificate and client secret authentication.

**When to use:** SharePoint Online automation using rich PnP cmdlets (`Get-PnPListItem`, `Set-PnPTenantSite`, etc.). For pure Graph API work without SharePoint context, use `Application-Auth-Graph.ps1` instead.

**Features:**
- Certificate (preferred) or client secret authentication
- `Connect-ToPnPSite` helper for switching between site collections
- `Invoke-PnPWithRetry` wrapper for throttle-resistant PnP and REST calls
- Supports raw SharePoint REST via `Invoke-PnPSPRestMethod`
- Built-in example calls (list items, enumerate sites, REST call, update item, upload file)

**Prerequisites:**
- PnP.PowerShell module: `Install-Module PnP.PowerShell -Scope CurrentUser`
- Entra ID app registration with SharePoint **application** permissions (e.g., `Sites.FullControl.All`) granted and admin-consented
- Certificate in the Windows certificate store, **or** client secret in the `PNP_CLIENT_SECRET` environment variable

---

### DeviceCodeAuth.ps1

Delegated authentication for Microsoft Graph using the **OAuth 2.0 Device Code** flow. Displays a user code and URL; the user signs in on any device.

**When to use:** Scripts running in SSH sessions, headless servers, or environments that cannot open a browser on the machine running the script.

**Features:**
- No browser required on the script host
- Automatic silent token refresh using the refresh token
- Falls back to a new device code flow if the refresh token expires or is revoked
- Throttle-aware retry with exponential back-off
- Automatic pagination via `Invoke-GraphPagedRequest`

**Prerequisites:**
- Entra ID app registration configured as a **Public Client** with device code flow enabled (`Authentication > Allow public client flows`)
- Delegated Graph API permissions granted and admin-consented

---

### Get-TokenWithPKCE.ps1

Delegated authentication for Microsoft Graph using the **OAuth 2.0 Authorization Code + PKCE** flow. Opens the default browser for interactive sign-in.

**When to use:** Scripts that need to act as the signed-in user (e.g., access their mailbox, OneDrive, presence). No client secret is required — safe for public clients.

**Features:**
- No client secret needed (PKCE — RFC 7636)
- Cryptographically random code verifier and SHA-256 challenge
- Temporary local HTTP listener captures the authorization code from the redirect
- Automatic silent token refresh via refresh token
- Falls back to a new interactive sign-in if the refresh token expires
- Throttle-aware retry with exponential back-off
- Automatic pagination via `Invoke-GraphPagedRequest`

**Prerequisites:**
- Entra ID app registration configured as a **Public Client**
- Redirect URI (default: `http://localhost:8080`) registered under **Mobile and desktop applications** (not as a web redirect)
- Delegated Graph API permissions granted and consented
- Port 8080 (or your chosen port) must be free when the script runs

---

## Quick Start

1. **Clone or download** this repository.
2. **Choose** the script that matches your authentication scenario (see table above).
3. **Fill in** the `Configuration` section at the top of the script:
   ```powershell
   $tenantId  = 'your-tenant-id'   # Tenant ID or domain
   $clientId  = 'your-client-id'   # Entra ID app registration client ID
   ```
4. **Add your logic** inside the `#region Main` block at the bottom of the script.
5. **Run** the script in PowerShell 5.1 or PowerShell 7+.

---

## Choosing the Right Script

```
Do you need to act as a signed-in user?
├── Yes
│   ├── Browser available on the script host?  →  Get-TokenWithPKCE.ps1
│   └── No browser / headless environment?     →  DeviceCodeAuth.ps1
└── No (service / unattended)
    ├── Working with SharePoint / PnP cmdlets?  →  Application-Auth-PNP.ps1
    └── Pure Graph API calls?                   →  Application-Auth-Graph.ps1
```

---

## Security Notes

- **Never hard-code** client secrets in script files or commit them to source control.
- Store secrets in environment variables and reference them in the configuration section:
  ```powershell
  # Set once per session
  Set-Item Env:\GRAPH_CLIENT_SECRET 'your-secret'
  
  # Or persist for the current user
  [System.Environment]::SetEnvironmentVariable('GRAPH_CLIENT_SECRET', 'your-secret', 'User')
  ```
- **Prefer certificate authentication** over client secrets where possible.
- Apply the **principle of least privilege** — request only the Graph scopes your script actually needs.

---

## Requirements

- Windows PowerShell 5.1 or PowerShell 7+
- An [Entra ID app registration](https://learn.microsoft.com/en-us/entra/identity-platform/quickstart-register-app) configured for the appropriate auth flow
- **PnP.PowerShell only:** `Install-Module PnP.PowerShell -Scope CurrentUser`

---

## Author

[mikelee1313](https://github.com/mikelee1313)
