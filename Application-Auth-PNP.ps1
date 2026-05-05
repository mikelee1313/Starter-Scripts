<#
.SYNOPSIS
    Baseline template for SharePoint / Microsoft 365 scripts using PnP.PowerShell
    with app-only authentication.

.DESCRIPTION
    Provides a reusable foundation for PnP.PowerShell automation. Supports both
    certificate-based and client secret authentication via an Entra ID app registration.
    Includes throttle-aware retry with exponential back-off for PnP cmdlets and raw
    SharePoint REST calls.

    PnP.PowerShell manages its own connection state and token refresh internally —
    there is no need to manually acquire or cache tokens. Simply call Connect-ToPnPSite
    to establish a connection; the module handles re-authentication transparently.

    USAGE
    -----
    1. Fill in the Configuration section (Tenant URL, Client ID, auth type).
    2. For scripts that iterate multiple sites, call Connect-ToPnPSite for each site URL
       before making PnP calls against it.
    3. Wrap any PnP cmdlet calls that may throttle inside Invoke-PnPWithRetry.
    4. Use Invoke-PnPSPRestMethod for raw SharePoint REST calls the PnP cmdlets don't cover.
    5. Add your logic inside the #region Main block at the bottom of the file.

    PREREQUISITES
    -------------
    - PnP.PowerShell module installed:
          Install-Module PnP.PowerShell -Scope CurrentUser
    - Entra ID app registration with the required SharePoint / Graph application permissions
      granted and admin-consented.
    - Certificate auth : certificate with private key in the Windows certificate store.
    - Secret auth      : client secret stored in an environment variable — never
                         hard-coded in source. Set PNP_CLIENT_SECRET before running.
    - The app registration must have the Sites.FullControl.All (or appropriate) SharePoint
      application permission, or be granted site collection admin via PnP/admin APIs.

    NOTE ON PNP vs RAW GRAPH
    ------------------------
    Use PnP.PowerShell when working primarily with SharePoint Online, Teams, or M365 Groups
    and you want rich cmdlets (Get-PnPListItem, Set-PnPTenantSite, etc.).
    For pure Graph API automation without SharePoint context, prefer Starting_Script.ps1.
#>

#region Configuration
##############################################################
#                  CONFIGURATION SECTION                     #
##############################################################

# ---- Debug output (set to $true for verbose PnP call tracing) ----
$debug = $false

# ---- Tenant & App Registration ----
$tenantId = ''   # Tenant ID or verified domain, e.g. 'contoso.onmicrosoft.com'
$clientId = ''   # Application (client) ID of the Entra ID app registration

# ---- Tenant root SharePoint URL ----
$tenantUrl = ''   # e.g. 'https://contoso.sharepoint.com'

# ---- Default site to connect to at startup ----
# Use the tenant admin URL for admin-level operations, or a specific site collection URL.
# e.g. 'https://contoso-admin.sharepoint.com'  or  'https://contoso.sharepoint.com/sites/MySite'
$defaultSiteUrl = ''

# ---- Authentication type: 'Certificate' or 'ClientSecret' ----
$AuthType = 'Certificate'

# Certificate thumbprint (used when $AuthType = 'Certificate')
$Thumbprint = 'B696FDCFE1453F3FBC6031F54DE988DA0ED905A9'

# Certificate store location: 'LocalMachine' or 'CurrentUser'
$CertStore = 'LocalMachine'

# Client Secret (used when $AuthType = 'ClientSecret')
# SECURITY: Never hard-code secrets in source files or commit them to source control.
#   Store the secret in an environment variable and reference it here:
#       Set-Item Env:\PNP_CLIENT_SECRET 'your-secret'   (current session)
#       [System.Environment]::SetEnvironmentVariable('PNP_CLIENT_SECRET','...','User')
$clientSecret = $env:PNP_CLIENT_SECRET

# ---- Output folder for any exported files ----
$OutputFolder = $env:TEMP

# ---- Throttle / retry settings ----
$MaxRetries = 15   # Maximum retry attempts per throttled request
$InitialBackoffSec = 3    # Starting back-off in seconds (doubles each retry, caps at 300)

##############################################################
#                END CONFIGURATION SECTION                   #
##############################################################
#endregion Configuration

#region Initialization
$date = Get-Date -Format 'yyyyMMddHHmmss'
$today = (Get-Date).Date

# Verify PnP.PowerShell is available
if (-not (Get-Module -ListAvailable -Name 'PnP.PowerShell')) {
    throw "PnP.PowerShell module not found. Install it with: Install-Module PnP.PowerShell -Scope CurrentUser"
}
Import-Module PnP.PowerShell -ErrorAction Stop
#endregion Initialization

#region Helper Functions

function Connect-ToPnPSite {
    <#
    .SYNOPSIS
        Establishes a PnP.PowerShell connection to the specified SharePoint site URL
        using the configured authentication type (Certificate or ClientSecret).
        Call this function whenever you switch to a different site collection.
    .PARAMETER Url
        The full URL of the SharePoint site or admin centre to connect to.
        Defaults to $defaultSiteUrl if omitted.
    .EXAMPLE
        Connect-ToPnPSite
        Connect-ToPnPSite -Url 'https://contoso.sharepoint.com/sites/HR'
    #>
    [CmdletBinding()]
    param (
        [Parameter()] [string] $Url = $script:defaultSiteUrl
    )

    if ([string]::IsNullOrWhiteSpace($Url)) {
        throw 'No site URL provided and $defaultSiteUrl is empty. Set $defaultSiteUrl in the Configuration section.'
    }

    if ($debug) { Write-Host "  PnP -> Connecting to $Url ($AuthType)" -ForegroundColor DarkGray }

    if ($AuthType -eq 'Certificate') {
        if ([string]::IsNullOrWhiteSpace($Thumbprint)) {
            throw 'Certificate Thumbprint is empty. Fill in $Thumbprint in the Configuration section.'
        }
        try {
            # Verify the certificate exists in the specified store before connecting.
            # Connect-PnPOnline -Thumbprint searches both CurrentUser\My and LocalMachine\My
            # automatically — no store parameter is needed or available.
            if (-not (Test-Path "Cert:\$CertStore\My\$Thumbprint")) {
                throw "Certificate $Thumbprint not found in Cert:\$CertStore\My"
            }
            Connect-PnPOnline -Url $Url -ClientId $clientId -Tenant $tenantId `
                -Thumbprint $Thumbprint -ErrorAction Stop
            Write-Host "  Connected (Certificate) -> $Url" -ForegroundColor Green
        }
        catch {
            Write-Host "  PnP connection failed (Certificate): $($_.Exception.Message)" -ForegroundColor Red
            throw
        }
    }
    elseif ($AuthType -eq 'ClientSecret') {
        if ([string]::IsNullOrWhiteSpace($clientSecret)) {
            throw 'Client secret is empty. Set the PNP_CLIENT_SECRET environment variable before running.'
        }
        try {
            Connect-PnPOnline -Url $Url -ClientId $clientId -ClientSecret $clientSecret -ErrorAction Stop
            Write-Host "  Connected (ClientSecret) -> $Url" -ForegroundColor Green
        }
        catch {
            Write-Host "  PnP connection failed (ClientSecret): $($_.Exception.Message)" -ForegroundColor Red
            throw
        }
    }
    else {
        throw "Invalid AuthType '$AuthType'. Accepted values: 'Certificate' or 'ClientSecret'."
    }
}

function Invoke-PnPWithRetry {
    <#
    .SYNOPSIS
        Executes a script block containing PnP cmdlet or SharePoint REST calls with
        Retry-After / exponential-backoff handling for throttling (429, 503) and
        transient server errors (502, 504).
    .DESCRIPTION
        PnP.PowerShell handles throttling for most of its own cmdlets internally, but
        Invoke-PnPSPRestMethod and bulk operations can still surface throttle exceptions.
        Wrap any call that may be throttled inside this function.
    .PARAMETER ScriptBlock
        The code to execute and potentially retry.
    .EXAMPLE
        $items = Invoke-PnPWithRetry { Get-PnPListItem -List 'Documents' -PageSize 500 }

        $result = Invoke-PnPWithRetry {
            Invoke-PnPSPRestMethod -Method Get -Url '/_api/web/lists'
        }
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [scriptblock] $ScriptBlock,
        [Parameter()]          [int]         $MaxRetries = $script:MaxRetries,
        [Parameter()]          [int]         $InitialBackoffSeconds = $script:InitialBackoffSec
    )

    $retryCount = 0
    $backoffSec = $InitialBackoffSeconds

    while ($retryCount -le $MaxRetries) {
        try {
            return & $ScriptBlock
        }
        catch {
            # Extract HTTP status code from the exception if available
            $statusCode = $null
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }
            # PnP sometimes wraps the status in the message string
            if (-not $statusCode -and $_.Exception.Message -match '(429|502|503|504)') {
                $statusCode = [int]$Matches[1]
            }

            $isRetryable = $statusCode -in @(429, 502, 503, 504) -or
            ($_.Exception -is [System.Net.WebException] -and
            $_.Exception.Status -in @(
                [System.Net.WebExceptionStatus]::Timeout,
                [System.Net.WebExceptionStatus]::ConnectionClosed
            ))

            if (-not $isRetryable) { throw $_ }

            if ($retryCount -ge $MaxRetries) {
                Write-Warning "Max retries ($MaxRetries) reached."
                throw $_
            }

            # Honour Retry-After header when present
            $waitSec = $backoffSec
            if ($statusCode -in @(429, 503)) {
                try {
                    $ra = $_.Exception.Response.Headers['Retry-After']
                    if ($ra) { $waitSec = [int]$ra }
                }
                catch {}
            }

            $retryCount++
            Write-Host "    Throttled ($statusCode). Waiting ${waitSec}s (attempt $retryCount/$MaxRetries)..." `
                -ForegroundColor Yellow
            Start-Sleep -Seconds $waitSec
            $backoffSec = [Math]::Min($backoffSec * 2, 300)
        }
    }
}

#endregion Helper Functions

#region Main
##############################################################
#                    MAIN EXECUTION                          #
##############################################################

try {
    # Connect to the default site at startup.
    # For scripts that iterate multiple sites, call Connect-ToPnPSite inside the loop.
    Connect-ToPnPSite

    #----------------------------------------------------------
    # Add your PnP cmdlet calls below.
    #
    # Tip: Wrap calls that may throttle in Invoke-PnPWithRetry { ... }
    # When switching sites, call Connect-ToPnPSite -Url '<new site url>'
    #----------------------------------------------------------

    # --- Example 1: Get all items from a list ---
    # $items = Invoke-PnPWithRetry {
    #     Get-PnPListItem -List 'Documents' -PageSize 500
    # }
    # Write-Host "Total items: $($items.Count)"

    # --- Example 2: Iterate all site collections in the tenant ---
    # $sites = Invoke-PnPWithRetry {
    #     Get-PnPTenantSite -IncludeOneDriveSites:$false
    # }
    # foreach ($site in $sites) {
    #     Connect-ToPnPSite -Url $site.Url
    #     $web = Invoke-PnPWithRetry { Get-PnPWeb }
    #     Write-Host "$($web.Title) — $($site.Url)"
    # }

    # --- Example 3: Raw SharePoint REST call via Invoke-PnPSPRestMethod ---
    # $result = Invoke-PnPWithRetry {
    #     Invoke-PnPSPRestMethod -Method Get -Url '/_api/web/lists?$select=Title,Id'
    # }
    # $result.value | ForEach-Object { Write-Host $_.Title }

    # --- Example 4: Update a list item ---
    # Invoke-PnPWithRetry {
    #     Set-PnPListItem -List 'Tasks' -Identity 42 -Values @{ Status = 'Completed' }
    # }

    # --- Example 5: Upload a file ---
    # Invoke-PnPWithRetry {
    #     Add-PnPFile -Path 'C:\Reports\output.csv' -Folder 'Shared Documents/Reports'
    # }

    Write-Host "`nScript completed successfully." -ForegroundColor Green
}
catch {
    Write-Host "`nScript failed: $($_.Exception.Message)" -ForegroundColor Red
    throw
}
finally {
    # Always disconnect to release the connection cleanly
    try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch {}
}

#endregion Main
