<#
.SYNOPSIS
    Baseline template for Microsoft Graph API scripts using delegated (user) authentication
    with the OAuth 2.0 Authorization Code + PKCE flow.

.DESCRIPTION
    Provides a reusable foundation for Microsoft Graph API automation that requires a signed-in
    user context (delegated permissions). Uses PKCE (Proof Key for Code Exchange) — no client
    secret is required, making it safe for public-client / interactive scripts.

    FLOW OVERVIEW
    -------------
    1. A cryptographically random PKCE code verifier and SHA-256 challenge are generated.
    2. The default browser opens the Entra ID authorization URL.
    3. A temporary local HTTP listener captures the authorization code from the redirect.
    4. The code is exchanged for an access token + refresh token.
    5. Subsequent calls silently refresh the access token using the refresh token; a new
       interactive browser sign-in is only triggered if the refresh token is also expired.

    USAGE
    -----
    1. Fill in the Configuration section (Tenant ID, Client ID, Scopes).
    2. Ensure $redirectUri is registered as a redirect URI on the Entra ID app registration
       (under "Mobile and desktop applications" platform — do NOT add it as a web redirect).
    3. Add your Graph API calls inside the #region Main block at the bottom of the file.
    4. Use Invoke-GraphPagedRequest for GET calls that may return multiple pages.
    5. Use Invoke-GraphRequestWithThrottleHandling directly for single-resource calls
       (POST / PATCH / DELETE) or when you need the raw response object.

    PREREQUISITES
    -------------
    - Entra ID app registration configured as a Public Client.
    - Delegated Graph API permissions granted (and consented) for the required scopes.
    - The redirect URI (default: http://localhost:8080) added to the app registration.
    - Port 8080 (or your chosen port) must be free when the script runs.

    NOTE ON PKCE vs APP-ONLY
    ------------------------
    PKCE / delegated auth is appropriate when you need to act as the signed-in user
    (e.g., access the user's mailbox, OneDrive, presence). For fully unattended / service
    scenarios use the certificate or client-secret app-only template (Starting_Script.ps1).
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

# ---- Redirect URI — must match a registered redirect URI on the app registration ----
$redirectUri = 'http://localhost:8080'

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

# ---- Browser sign-in listener timeout (seconds) ----
$AuthListenerTimeoutSec = 120

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

function New-PKCEParameters {
    <#
    .SYNOPSIS
        Generates a cryptographically random PKCE code verifier and its SHA-256 challenge.
        The verifier is a 32-byte random value encoded as base64url (43 characters, no padding).
        The challenge is SHA-256(verifier) encoded as base64url, per RFC 7636.
    #>
    $randomBytes = [byte[]]::new(32)
    [System.Security.Cryptography.RandomNumberGenerator]::Create().GetBytes($randomBytes)
    $codeVerifier = [Convert]::ToBase64String($randomBytes).TrimEnd('=').Replace('+', '-').Replace('/', '_')

    $hasher = [System.Security.Cryptography.SHA256]::Create()
    $codeChallenge = [Convert]::ToBase64String(
        $hasher.ComputeHash([System.Text.Encoding]::UTF8.GetBytes($codeVerifier))
    ).TrimEnd('=').Replace('+', '-').Replace('/', '_')

    return @{
        CodeVerifier  = $codeVerifier
        CodeChallenge = $codeChallenge
    }
}

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

function Get-TokenWithPKCE {
    <#
    .SYNOPSIS
        Performs an interactive OAuth 2.0 Authorization Code + PKCE flow to acquire a
        Microsoft Graph access token and refresh token for the signed-in user.
        Opens the default browser, listens on $redirectUri for the callback, then
        exchanges the authorization code for tokens.
    #>
    Write-Host 'Starting interactive PKCE authentication...' -ForegroundColor Cyan

    $pkce = New-PKCEParameters
    $tokenUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    # Build authorization URL
    $authUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/authorize" +
    "?client_id=$clientId" +
    "&response_type=code" +
    "&redirect_uri=$([uri]::EscapeDataString($redirectUri))" +
    "&scope=$([uri]::EscapeDataString($scope))" +
    "&code_challenge=$($pkce.CodeChallenge)" +
    "&code_challenge_method=S256" +
    "&response_mode=query"

    Write-Host 'Opening browser for sign-in...' -ForegroundColor Yellow
    Write-Host "If the browser does not open automatically, navigate to:`n$authUri" -ForegroundColor Cyan
    Start-Process $authUri

    # Start local HTTP listener to capture the redirect
    $listener = [System.Net.HttpListener]::new()
    $listener.Prefixes.Add("$redirectUri/")
    $listener.Start()

    Write-Host "`nWaiting for browser sign-in (timeout: ${AuthListenerTimeoutSec}s)..." -ForegroundColor Yellow

    # Use async GetContext so we can enforce a timeout
    $asyncResult = $listener.BeginGetContext($null, $null)
    $completed = $asyncResult.AsyncWaitHandle.WaitOne(($AuthListenerTimeoutSec * 1000))

    if (-not $completed) {
        $listener.Stop()
        throw "Authentication timed out after ${AuthListenerTimeoutSec} seconds. No response received from browser."
    }

    $context = $listener.EndGetContext($asyncResult)
    $request = $context.Request
    $response = $context.Response

    # Parse the authorization code (or error) from the redirect query string
    $authCode = $null
    if ($request.QueryString['code']) {
        $authCode = $request.QueryString['code']
        $responseHtml = '<html><body><h1>Authentication Successful</h1><p>You may close this window and return to PowerShell.</p></body></html>'
        Write-Host 'Authorization code received.' -ForegroundColor Green
    }
    else {
        $authError = $request.QueryString['error']
        $authErrorDesc = $request.QueryString['error_description']
        $responseHtml = "<html><body><h1>Authentication Failed</h1><p>$authError</p><p>$authErrorDesc</p></body></html>"
        Write-Host "Authentication failed: $authError — $authErrorDesc" -ForegroundColor Red
    }

    # Send confirmation page to the browser then shut down the listener
    $buffer = [System.Text.Encoding]::UTF8.GetBytes($responseHtml)
    $response.ContentLength64 = $buffer.Length
    $response.OutputStream.Write($buffer, 0, $buffer.Length)
    $response.OutputStream.Close()
    $listener.Stop()

    if (-not $authCode) {
        throw 'Failed to obtain an authorization code. Check the browser for error details.'
    }

    # Exchange authorization code for tokens
    $tokenBody = @{
        grant_type    = 'authorization_code'
        client_id     = $clientId
        code          = $authCode
        redirect_uri  = $redirectUri
        code_verifier = $pkce.CodeVerifier
        scope         = $scope
    }

    try {
        $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $tokenBody `
            -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop -Verbose:$false
        $global:token = $resp.access_token
        $global:refreshToken = $resp.refresh_token
        $expiresIn = if ($resp.expires_in) { [int]$resp.expires_in } else { 3600 }
        $global:tokenExpiry = (Get-Date).AddSeconds($expiresIn - 300)
        Write-Host "  Signed in. Token valid until: $($global:tokenExpiry)" -ForegroundColor Green
    }
    catch {
        Write-Host "  Token exchange failed: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

function Update-TokenFromRefreshToken {
    <#
    .SYNOPSIS
        Silently refreshes the access token using the stored refresh token.
        Falls back to a full interactive PKCE sign-in if no refresh token is available.
    #>
    if ([string]::IsNullOrWhiteSpace($global:refreshToken)) {
        Write-Host 'No refresh token available — starting interactive sign-in...' -ForegroundColor Yellow
        Get-TokenWithPKCE
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
        # Refresh token may be expired or revoked — fall back to interactive sign-in
        Write-Host "  Silent refresh failed: $($_.Exception.Message). Falling back to interactive sign-in..." -ForegroundColor Yellow
        $global:refreshToken = $null
        Get-TokenWithPKCE
    }
}

function Test-ValidToken {
    <#
    .SYNOPSIS
        Checks whether the cached access token is still valid; refreshes silently (or
        interactively as a last resort) if it is expired or missing.
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
    # Authenticate once at startup — token refreshes silently (or interactively) when needed
    Get-TokenWithPKCE

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

    # --- Example 4: Upload a small file to OneDrive ---
    # $headers       = Get-GraphAuthHeaders
    # $headers['Content-Type'] = 'text/plain'
    # $fileBytes     = [System.Text.Encoding]::UTF8.GetBytes('Hello, OneDrive!')
    # Invoke-GraphRequestWithThrottleHandling `
    #     -Uri     'https://graph.microsoft.com/v1.0/me/drive/root:/hello.txt:/content' `
    #     -Method  'PUT' `
    #     -Headers $headers `
    #     -Body    ([System.Text.Encoding]::UTF8.GetString($fileBytes))

    Write-Host "`nScript completed successfully." -ForegroundColor Green
}
catch {
    Write-Host "`nScript failed: $($_.Exception.Message)" -ForegroundColor Red
    throw
}

#endregion Main
