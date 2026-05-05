<#
.SYNOPSIS
    Baseline template for Microsoft Graph API scripts using app-only authentication.

.DESCRIPTION
    Provides a reusable foundation for Microsoft Graph API automation. Supports both
    certificate-based and client secret authentication via an Entra ID app registration.
    Includes automatic token refresh, throttle-aware retry with exponential back-off,
    and automatic pagination handling.

    USAGE
    -----
    1. Fill in the Configuration section (Tenant ID, Client ID, auth type).
    2. Add your Graph API calls inside the #region Main block at the bottom of the file.
    3. Use Invoke-GraphPagedRequest  for GET calls that may return multiple pages.
    4. Use Invoke-GraphRequestWithThrottleHandling directly for single-resource calls
       (POST / PATCH / DELETE) or when you need the raw response object.

    PREREQUISITES
    -------------
    - Entra ID app registration with the required Graph application permissions granted
      and admin-consented. A single registration in the home tenant covers all geos.
    - Certificate auth : certificate with private key in the Windows certificate store.
    - Secret auth      : client secret stored in an environment variable — never
                         hard-coded in source. Set GRAPH_CLIENT_SECRET before running.
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

# ---- Authentication type: 'Certificate' or 'ClientSecret' ----
$AuthType = 'Certificate'

# Certificate thumbprint (used when $AuthType = 'Certificate')
$Thumbprint = ''

# Certificate store location: 'LocalMachine' or 'CurrentUser'
$CertStore = 'LocalMachine'

# Client Secret (used when $AuthType = 'ClientSecret')
# SECURITY: Never hard-code secrets in source files or commit them to source control.
#   Store the secret in an environment variable and reference it here:
#       Set-Item Env:\GRAPH_CLIENT_SECRET 'your-secret'   (current session)
#       [System.Environment]::SetEnvironmentVariable('GRAPH_CLIENT_SECRET','...','User')
$clientSecret = $env:GRAPH_CLIENT_SECRET

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
                       -Uri    'https://graph.microsoft.com/v1.0/users/user@contoso.com' `
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
        $users = Invoke-GraphPagedRequest `
                     -Uri 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName&$top=999'
        Write-Host "Total users: $($users.Count)"
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

function AcquireToken {
    <#
    .SYNOPSIS
        Acquires a Microsoft Graph access token using client credentials flow
        (certificate JWT assertion or client secret). One token covers all Graph
        endpoints and geo datacenters.
    #>
    Write-Host "Authenticating to Microsoft Graph ($AuthType)..." -ForegroundColor Cyan

    $scope = 'https://graph.microsoft.com/.default'
    $tokenUri = "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token"

    if ($AuthType -eq 'ClientSecret') {
        if ([string]::IsNullOrWhiteSpace($clientSecret)) {
            throw 'Client secret is empty. Set the GRAPH_CLIENT_SECRET environment variable before running.'
        }

        $body = @{
            grant_type    = 'client_credentials'
            client_id     = $clientId
            client_secret = $clientSecret
            scope         = $scope
        }
        try {
            $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body `
                -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop -Verbose:$false
            $global:token = $resp.access_token
            $expiresIn = if ($resp.expires_in) { [int]$resp.expires_in } else { 3600 }
            $global:tokenExpiry = (Get-Date).AddSeconds($expiresIn - 300)
            Write-Host "  Connected via Client Secret. Token valid until: $($global:tokenExpiry)" -ForegroundColor Green
        }
        catch {
            Write-Host "  Authentication failed (ClientSecret): $($_.Exception.Message)" -ForegroundColor Red
            throw
        }
    }
    elseif ($AuthType -eq 'Certificate') {
        try {
            $cert = Get-Item -Path "Cert:\$CertStore\My\$Thumbprint" -ErrorAction Stop
        }
        catch {
            Write-Host "  Certificate $Thumbprint not found in Cert:\$CertStore\My" -ForegroundColor Red
            throw
        }

        $rsa = [System.Security.Cryptography.X509Certificates.RSACertificateExtensions]::GetRSAPrivateKey($cert)
        if (-not $rsa) {
            throw "Unable to access RSA private key for certificate $Thumbprint. Ensure the private key is accessible to the current account."
        }

        $now = [System.DateTimeOffset]::UtcNow
        $exp = $now.AddMinutes(10).ToUnixTimeSeconds()
        $nbf = $now.ToUnixTimeSeconds()
        $x5t = [Convert]::ToBase64String($cert.GetCertHash()).TrimEnd('=').Replace('+', '-').Replace('/', '_')

        $header = @{ alg = 'RS256'; typ = 'JWT'; x5t = $x5t } | ConvertTo-Json -Compress
        $payload = @{
            aud = $tokenUri
            exp = $exp
            iss = $clientId
            jti = [System.Guid]::NewGuid().ToString()
            nbf = $nbf
            sub = $clientId
        } | ConvertTo-Json -Compress

        $hB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        $pB64 = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($payload)).TrimEnd('=').Replace('+', '-').Replace('/', '_')
        $toSign = "$hB64.$pB64"

        $sig = $rsa.SignData(
            [System.Text.Encoding]::UTF8.GetBytes($toSign),
            [System.Security.Cryptography.HashAlgorithmName]::SHA256,
            [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)
        $jwt = "$toSign.$([Convert]::ToBase64String($sig).TrimEnd('=').Replace('+', '-').Replace('/', '_'))"

        $body = @{
            client_id             = $clientId
            client_assertion_type = 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer'
            client_assertion      = $jwt
            scope                 = $scope
            grant_type            = 'client_credentials'
        }

        try {
            $resp = Invoke-RestMethod -Method Post -Uri $tokenUri -Body $body `
                -ContentType 'application/x-www-form-urlencoded' -ErrorAction Stop -Verbose:$false
            $global:token = $resp.access_token
            $expiresIn = if ($resp.expires_in) { [int]$resp.expires_in } else { 3600 }
            $global:tokenExpiry = (Get-Date).AddSeconds($expiresIn - 300)
            Write-Host "  Connected via Certificate ($Thumbprint). Token valid until: $($global:tokenExpiry)" -ForegroundColor Green
        }
        catch {
            Write-Host "  Authentication failed (Certificate): $($_.Exception.Message)" -ForegroundColor Red
            throw
        }
    }
    else {
        throw "Invalid AuthType '$AuthType'. Accepted values: 'Certificate' or 'ClientSecret'."
    }
}

function Test-ValidToken {
    <#
    .SYNOPSIS
        Checks whether the cached token is still valid; re-authenticates if expired or missing.
        The token is considered stale 5 minutes before its actual expiry to avoid mid-call failures.
    #>
    if ($null -eq $global:tokenExpiry -or (Get-Date) -gt $global:tokenExpiry) {
        Write-Host 'Token expired or missing — refreshing...' -ForegroundColor Yellow
        AcquireToken
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
    # Authenticate once at startup — token is refreshed automatically when it expires
    AcquireToken

    #----------------------------------------------------------
    # Add your Graph API calls below.
    #
    # Tip: Call Test-ValidToken (or use Invoke-GraphPagedRequest which calls it
    # internally) before any long-running loop to avoid mid-run token expiry.
    #----------------------------------------------------------

    # --- Example 1: Paginated GET (returns all pages as a flat list) ---
    # $users = Invoke-GraphPagedRequest `
    #              -Uri 'https://graph.microsoft.com/v1.0/users?$select=id,displayName,userPrincipalName&$top=999'
    # Write-Host "Total users: $($users.Count)"

    # --- Example 2: Single-resource GET ---
    # $headers = Get-GraphAuthHeaders
    # $org = Invoke-GraphRequestWithThrottleHandling `
    #            -Uri     'https://graph.microsoft.com/v1.0/organization' `
    #            -Method  'GET' `
    #            -Headers $headers
    # Write-Host "Tenant display name: $($org.value[0].displayName)"

    # --- Example 3: POST (create a resource) ---
    # $headers  = Get-GraphAuthHeaders
    # $newGroup = @{
    #     displayName     = 'Test Group'
    #     mailEnabled     = $false
    #     mailNickname    = 'testgroup'
    #     securityEnabled = $true
    # } | ConvertTo-Json
    # $created = Invoke-GraphRequestWithThrottleHandling `
    #                -Uri     'https://graph.microsoft.com/v1.0/groups' `
    #                -Method  'POST' `
    #                -Headers $headers `
    #                -Body    $newGroup
    # Write-Host "Created group: $($created.id)"

    # --- Example 4: PATCH (update a resource) ---
    # $headers = Get-GraphAuthHeaders
    # $patch   = @{ displayName = 'Updated Name' } | ConvertTo-Json
    # Invoke-GraphRequestWithThrottleHandling `
    #     -Uri     "https://graph.microsoft.com/v1.0/groups/<groupId>" `
    #     -Method  'PATCH' `
    #     -Headers $headers `
    #     -Body    $patch

    Write-Host "`nScript completed successfully." -ForegroundColor Green
}
catch {
    Write-Host "`nScript failed: $($_.Exception.Message)" -ForegroundColor Red
    throw
}

#endregion Main

