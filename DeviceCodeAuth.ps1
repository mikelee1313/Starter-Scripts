<#
.SYNOPSIS
    Baseline template for Microsoft Graph API scripts using delegated (user) authentication
    with the OAuth 2.0 Device Code flow.

.DESCRIPTION
    Provides a reusable foundation for Microsoft Graph API automation that requires a signed-in
    user context (delegated permissions) but cannot open a browser on the machine running the
    script — e.g. SSH sessions, headless servers, or environments behind a proxy.

    FLOW OVERVIEW
    -------------
    1. A device code and a short user code are requested from Entra ID.
    2. The user code and verification URL are displayed; the user signs in on any device.
    3. The script polls the token endpoint until the user completes sign-in or the code expires.
    4. Subsequent calls silently refresh the access token using the refresh token; the full
       device code flow is only re-triggered if the refresh token is also expired or revoked.

    USAGE
    -----
    1. Fill in the Configuration section (Tenant ID, Client ID, Scopes).
    2. Add your Graph API calls inside the #region Main block at the bottom of the file.
    3. Use Invoke-GraphPagedRequest for GET calls that may return multiple pages.
    4. Use Invoke-GraphRequestWithThrottleHandling directly for single-resource calls
       (POST / PATCH / DELETE) or when you need the raw response object.

    PREREQUISITES
    -------------
    - Entra ID app registration configured as a Public Client
      (Authentication > Allow public client flows > Enable device code flow).
    - Delegated Graph API permissions granted and admin-consented for the required scopes.
    - No client secret or certificate needed — this is a public client flow.

    NOTE ON DEVICE CODE vs PKCE vs APP-ONLY
    ----------------------------------------
    - Device Code  : delegated / user context; no browser on the script host required.
    - PKCE          : delegated / user context; requires a browser on the script host.
    - App-only      : no user context; requires a certificate or client secret
                      (see Starting_Script.ps1).
#>

#region Configuration
##############################################################
#                  CONFIGURATION SECTION                     #
##############################################################

# ---- Debug output (set to $true for verbose Graph call tracing) ----
$debug = $false

# ---- Tenant & App Registration ----
$tenantId = ''   # Tenant ID or verified domain, e.g. 'contoso.onmicrosoft.com'
$clientId = ''   # Application (client) ID of the Entra ID app registration

# ---- OAuth scopes (space-separated) ----
# Use specific scopes instead of /.default to request only what is needed (least privilege).
# Examples: 'User.Read', 'Mail.Read', 'Files.ReadWrite'
$scope = 'https://graph.microsoft.com/.default'

# ---- Output folder for any exported files ----
$OutputFolder = $env:TEMP

# ---- Throttle / retry settings ----
$MaxRetries = 15    # Maximum retry attempts per request
$InitialBackoffSec = 3     # Starting back-off in seconds (doubles each retry, caps at 300)
$RequestTimeoutSec = 300   # Per-request timeout in seconds

##############################################################
#                END CONFIGURATION SECTION                   #
##############################################################
#endregion Configuration

#region Initialization
$date = Get-Date -Format 'yyyyMMddHHmmss'
$today = (Get-Date).Date

$global:token = $null
$global:tokenExpiry = $null
$global:refreshToken = $null
#endregion Initialization

#region Helper Functions

function Invoke-GraphRequestWithThrottleHandling {
    <#
    .SYNOPSIS
        Wraps Invoke-RestMethod with Retry-After / exponential-backoff throttle handling
        for Microsoft Graph API calls (429, 502, 503, 504, and network timeouts).
    .EXAMPLE
        $headers = Get-GraphAuthHeaders
        $result  = Invoke-GraphRequestWithThrottleHandling `
                       -Uri    'https://graph.microsoft.com/v1.0/me' `
                       -Method 'GET' `
                       -Headers $headers
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [string]    $Uri,
        [Parameter(Mandatory)] [string]    $Method,
        [Parameter()]          [hashtable] $Headers = @{},
        [Parameter()]          [string]    $Body = $null,
        [Parameter()]          [string]    $ContentType = 'application/json',
        [Parameter()]          [int]       $MaxRetries = $script:MaxRetries,
        [Parameter()]          [int]       $InitialBackoffSeconds = $script:InitialBackoffSec,
        [Parameter()]          [int]       $TimeoutSeconds = $script:RequestTimeoutSec
    )

    $retryCount = 0
    $backoffSec = $InitialBackoffSeconds

    if ($debug) { Write-Host "  Graph -> $Method $Uri" -ForegroundColor DarkGray }

    while ($retryCount -le $MaxRetries) {
        try {
            $invokeParams = @{
                Uri         = $Uri
                Method      = $Method
                Headers     = $Headers
                ContentType = $ContentType
                TimeoutSec  = $TimeoutSeconds
                ErrorAction = 'Stop'
                Verbose     = $false
            }
            if ($Body) { $invokeParams['Body'] = $Body }

            return Invoke-RestMethod @invokeParams
        }
        catch {
            $statusCode = $null
            if ($_.Exception.Response) {
                $statusCode = [int]$_.Exception.Response.StatusCode
            }

            $isRetryable = ($statusCode -in @(429, 502, 503, 504)) -or
            ($_.Exception -is [System.Net.WebException] -and
            $_.Exception.Status -in @(
                [System.Net.WebExceptionStatus]::Timeout,
                [System.Net.WebExceptionStatus]::ConnectionClosed
            ))

            if (-not $isRetryable) { throw $_ }

            if ($retryCount -ge $MaxRetries) {
                Write-Warning "Max retries ($MaxRetries) reached for: $Uri"
                throw $_
            }

            # Honour the Retry-After header when present (common on 429 and 503)
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

function Invoke-GraphPagedRequest {
    <#
    .SYNOPSIS
        Executes a Graph GET request and automatically follows @odata.nextLink pages,
        returning all results as a single flat list.
    .EXAMPLE
        $messages = Invoke-GraphPagedRequest `
                        -Uri 'https://graph.microsoft.com/v1.0/me/messages?$select=id,subject&$top=100'
        Write-Host "Total messages: $($messages.Count)"
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)] [string] $Uri
    )

    $results = [System.Collections.Generic.List[object]]::new()
    $nextLink = $Uri

    do {
        Test-ValidToken
        $headers = Get-GraphAuthHeaders
        $response = Invoke-GraphRequestWithThrottleHandling -Uri $nextLink -Method 'GET' -Headers $headers

        if ($null -ne $response.value) {
            $results.AddRange([object[]]$response.value)
        }
        else {
            # Single-object response (no value array) — return directly
            return $response
        }

        $nextLink = $response.'@odata.nextLink'
        if ($debug -and $nextLink) { Write-Host '  Fetching next page...' -ForegroundColor DarkGray }
    } while ($nextLink)

    return $results
}

#endregion Helper Functions

#region Authentication Functions

function Get-TokenWithDeviceCode {
    <#
    .SYNOPSIS
        Performs the OAuth 2.0 Device Code flow to acquire a Microsoft Graph access token
        and refresh token for the signed-in user.
        Displays the user code and verification URL, then polls until the user completes
        sign-in or the device code expires.
    #>
    Write-Host 'Starting Device Code authentication...' -ForegroundColor Cyan

    $deviceCodeUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/devicecode"
    $tokenUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    # Request a device code
    $deviceBody = @{
        client_id = $clientId
        scope     = $scope
    }

    try {
        $deviceResponse = Invoke-RestMethod -Method Post -Uri $deviceCodeUri -Body $deviceBody `
            -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop -Verbose:$false
    }
    catch {
        Write-Host "  Failed to request device code: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }

    # Display sign-in instructions to the user
    Write-Host "`nUser sign-in required:" -ForegroundColor Green
    Write-Host "  1. Open a browser on any device and go to: " -NoNewline -ForegroundColor Cyan
    Write-Host $deviceResponse.verification_uri -ForegroundColor White
    Write-Host "  2. Enter the code: " -NoNewline -ForegroundColor Cyan
    Write-Host $deviceResponse.user_code -ForegroundColor Yellow
    Write-Host "  3. Sign in with your account credentials." -ForegroundColor Cyan
    Write-Host "`nWaiting for authentication" -NoNewline -ForegroundColor Yellow

    # Poll the token endpoint until success, expiry, or a terminal error
    $pollInterval = [int]$deviceResponse.interval
    $expiresIn = [int]$deviceResponse.expires_in
    $maxAttempts = [Math]::Ceiling($expiresIn / $pollInterval)
    $attempts = 0

    $tokenBody = @{
        grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
        client_id   = $clientId
        device_code = $deviceResponse.device_code
    }

    do {
        Start-Sleep -Seconds $pollInterval
        $attempts++

        try {
            $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $tokenBody `
                -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop -Verbose:$false

            # Success — store tokens and return
            $global:token = $resp.access_token
            $global:refreshToken = $resp.refresh_token
            $expiresInSec = if ($resp.expires_in) { [int]$resp.expires_in } else { 3600 }
            $global:tokenExpiry = (Get-Date).AddSeconds($expiresInSec - 300)
            Write-Host "`n  Signed in. Token valid until: $($global:tokenExpiry)" -ForegroundColor Green
            return
        }
        catch {
            $errBody = $null
            try { $errBody = $_.ErrorDetails.Message | ConvertFrom-Json } catch {}

            switch ($errBody.error) {
                'authorization_pending' {
                    # Normal — user hasn't completed sign-in yet
                    Write-Host '.' -NoNewline -ForegroundColor Yellow
                }
                'slow_down' {
                    # Server asked us to back off; increase interval for remainder of session
                    $pollInterval += 5
                    Write-Host '.' -NoNewline -ForegroundColor Yellow
                }
                'authorization_declined' {
                    Write-Host "`n  User declined the sign-in request." -ForegroundColor Red
                    throw 'Device code authorization was declined by the user.'
                }
                'expired_token' {
                    Write-Host "`n  Device code expired. Please re-run the script." -ForegroundColor Red
                    throw 'Device code expired before the user completed sign-in.'
                }
                default {
                    $msg = if ($errBody.error_description) { $errBody.error_description } else { $_.Exception.Message }
                    Write-Host "`n  Authentication error: $msg" -ForegroundColor Red
                    throw "Device code flow failed: $msg"
                }
            }
        }
    } while ($attempts -lt $maxAttempts)

    throw 'Device code authentication timed out. The device code expired before sign-in was completed.'
}

function Update-TokenFromRefreshToken {
    <#
    .SYNOPSIS
        Silently refreshes the access token using the stored refresh token.
        Falls back to a full interactive device code flow if no refresh token is available.
    #>
    if ([string]::IsNullOrWhiteSpace($global:refreshToken)) {
        Write-Host 'No refresh token available — starting device code sign-in...' -ForegroundColor Yellow
        Get-TokenWithDeviceCode
        return
    }

    Write-Host 'Refreshing access token silently...' -ForegroundColor Yellow

    $tokenUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"
    $tokenBody = @{
        grant_type    = 'refresh_token'
        client_id     = $clientId
        refresh_token = $global:refreshToken
        scope         = $scope
    }

    try {
        $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $tokenBody `
            -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop -Verbose:$false
        $global:token = $resp.access_token
        # Entra ID may rotate the refresh token; always store the latest one
        if ($resp.refresh_token) { $global:refreshToken = $resp.refresh_token }
        $expiresIn = if ($resp.expires_in) { [int]$resp.expires_in } else { 3600 }
        $global:tokenExpiry = (Get-Date).AddSeconds($expiresIn - 300)
        Write-Host "  Token refreshed. Valid until: $($global:tokenExpiry)" -ForegroundColor Green
    }
    catch {
        # Refresh token may be expired or revoked — fall back to device code flow
        Write-Host "  Silent refresh failed: $($_.Exception.Message). Falling back to device code sign-in..." `
            -ForegroundColor Yellow
        $global:refreshToken = $null
        Get-TokenWithDeviceCode
    }
}

function Test-ValidToken {
    <#
    .SYNOPSIS
        Checks whether the cached access token is still valid; refreshes silently (or via
        device code as a last resort) if it is expired or missing.
        The token is considered stale 5 minutes before its actual expiry.
    #>
    if ($null -eq $global:tokenExpiry -or (Get-Date) -gt $global:tokenExpiry) {
        Update-TokenFromRefreshToken
    }
}

function Get-GraphAuthHeaders {
    <#
    .SYNOPSIS
        Returns a hashtable containing the Authorization bearer header for Graph API calls.
        Automatically refreshes the token when it is expired or about to expire.
    .EXAMPLE
        $headers = Get-GraphAuthHeaders
        Invoke-GraphRequestWithThrottleHandling -Uri '...' -Method 'GET' -Headers $headers
    #>
    Test-ValidToken
    return @{ Authorization = "Bearer $global:token" }
}

#endregion Authentication Functions

#region Main
##############################################################
#                    MAIN EXECUTION                          #
##############################################################

try {
    # Authenticate once at startup — token refreshes silently (or via device code) when needed
    Get-TokenWithDeviceCode

    #----------------------------------------------------------
    # Add your Graph API calls below.
    #
    # Tip: Invoke-GraphPagedRequest calls Test-ValidToken internally.
    # For manual calls, always start with: $headers = Get-GraphAuthHeaders
    #----------------------------------------------------------

    # --- Example 1: Get the signed-in user's profile ---
    # $headers = Get-GraphAuthHeaders
    # $me = Invoke-GraphRequestWithThrottleHandling `
    #           -Uri     'https://graph.microsoft.com/v1.0/me' `
    #           -Method  'GET' `
    #           -Headers $headers
    # Write-Host "Signed in as: $($me.displayName) ($($me.userPrincipalName))"

    # --- Example 2: Paginated GET (returns all pages as a flat list) ---
    # $messages = Invoke-GraphPagedRequest `
    #                 -Uri 'https://graph.microsoft.com/v1.0/me/messages?$select=id,subject,receivedDateTime&$top=100'
    # Write-Host "Total messages: $($messages.Count)"

    # --- Example 3: POST (send a mail as the signed-in user) ---
    # $headers = Get-GraphAuthHeaders
    # $mail = @{
    #     message = @{
    #         subject      = 'Test from PowerShell'
    #         body         = @{ contentType = 'Text'; content = 'Hello from Graph!' }
    #         toRecipients = @(@{ emailAddress = @{ address = 'someone@contoso.com' } })
    #     }
    # } | ConvertTo-Json -Depth 5
    # Invoke-GraphRequestWithThrottleHandling `
    #     -Uri     'https://graph.microsoft.com/v1.0/me/sendMail' `
    #     -Method  'POST' `
    #     -Headers $headers `
    #     -Body    $mail

    # --- Example 4: List all users (requires User.Read.All) ---
    # $users = Invoke-GraphPagedRequest `
    #              -Uri 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName&$top=999'
    # Write-Host "Total users: $($users.Count)"

    Write-Host "`nScript completed successfully." -ForegroundColor Green
}
catch {
    Write-Host "`nScript failed: $($_.Exception.Message)" -ForegroundColor Red
    throw
}

#endregion Main