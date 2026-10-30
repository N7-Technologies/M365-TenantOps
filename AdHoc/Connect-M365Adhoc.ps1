#requires -Version 7.0
<#
.SYNOPSIS
    Tenant-aware ad-hoc Microsoft 365 admin connection, PowerShell 7.
.DESCRIPTION
    Connects Microsoft Graph, Exchange Online, and Teams via device-code flows.
    Graph uses Microsoft's OAuth device-code endpoint and an in-memory refresh
    token to work around the Graph SDK device credential failure.
    Does not connect Azure, SharePoint, or Purview.

    Intended for a new, plain PowerShell 7 terminal per client. Do not run in
    an existing dedicated client terminal. Never log the token variables.
.EXAMPLE
    & ".\AdHoc\Connect-M365Adhoc.ps1"
.EXAMPLE
    Assert-AdhocClientContext
    Repair-AdhocClientSession -Service Graph
    Add-AdhocGraphScopes -Scopes 'Application.Read.All'
#>

$ErrorActionPreference = 'Stop'

foreach ($name in @('Microsoft.Graph.Authentication','Microsoft.Graph.Identity.DirectoryManagement','ExchangeOnlineManagement','MicrosoftTeams')) {
    if (-not (Get-Module -ListAvailable -Name $name)) { throw "Required module is missing: $name" }
}

# Pin the tested Graph versions; do not mix Graph SDK assemblies in one process.
Import-Module Microsoft.Graph.Authentication -RequiredVersion 2.41.2 -ErrorAction Stop
Import-Module Microsoft.Graph.Identity.DirectoryManagement -RequiredVersion 2.41.2 -ErrorAction Stop

$global:AdhocM365 = @{
    ClientId      = '14d82eec-204b-4c2f-b7e8-296a70dab67e'
    Authority     = 'https://login.microsoftonline.com/organizations'
    Scopes        = @('Organization.Read.All','Reports.Read.All','User.Read.All','Group.Read.All','offline_access')
    TenantId      = $null
    TenantName    = $null
    Account       = $null
    RefreshToken  = $null
    ExpiresAt     = $null
    GraphHealthy  = $false
    TeamsVerified = $false
    RefreshBusy   = $false
    RefreshError  = $false
}

function global:Invoke-AdhocDeviceGrant {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$Scopes)

    $s = $global:AdhocM365
    $device = Invoke-RestMethod -Method Post -Uri "$($s.Authority)/oauth2/v2.0/devicecode" -ContentType 'application/x-www-form-urlencoded' -Body @{
        client_id = $s.ClientId
        scope     = ($Scopes -join ' ')
    } -ErrorAction Stop

    Write-Host "`n$($device.message)`n" -ForegroundColor Yellow
    Write-Host 'Waiting for device authentication...' -ForegroundColor Cyan
    $interval = [Math]::Max(5, [int]$device.interval)
    $deadline = (Get-Date).AddSeconds([int]$device.expires_in)

    while ((Get-Date) -lt $deadline) {
        Start-Sleep -Seconds $interval
        $reply = Invoke-WebRequest -Method Post -Uri "$($s.Authority)/oauth2/v2.0/token" -SkipHttpErrorCheck -ContentType 'application/x-www-form-urlencoded' -Body @{
            client_id   = $s.ClientId
            grant_type  = 'urn:ietf:params:oauth:grant-type:device_code'
            device_code = $device.device_code
        } -ErrorAction Stop
        $data = $reply.Content | ConvertFrom-Json
        if ([int]$reply.StatusCode -eq 200) {
            if (-not $data.access_token -or -not $data.refresh_token) { throw 'Token endpoint returned incomplete credentials.' }
            return $data
        }
        switch ($data.error) {
            'authorization_pending' { continue }
            'slow_down' { $interval += 5; continue }
            default { throw "Device authentication failed: $($data.error)" }
        }
    }
    throw 'Device authentication expired. Run the connection command again.'
}

function global:Set-AdhocGraphToken {
    [CmdletBinding()]
    param([Parameter(Mandatory)]$TokenResult, [switch]$Initial)

    $s = $global:AdhocM365
    $secure = ConvertTo-SecureString $TokenResult.access_token -AsPlainText -Force
    try {
        Connect-MgGraph -AccessToken $secure -NoWelcome -ErrorAction Stop
        $org = Get-MgOrganization -Property Id,DisplayName,VerifiedDomains -ErrorAction Stop | Select-Object -First 1
        $me = Invoke-MgGraphRequest -Method Get -Uri 'https://graph.microsoft.com/v1.0/me' -ErrorAction Stop
        if (-not $org.Id -or -not $me.userPrincipalName) { throw 'Graph did not return an organization and signed-in account.' }
        if (-not $Initial -and ([string]$org.Id -ne $s.TenantId -or [string]$me.userPrincipalName -ine $s.Account)) {
            $s.GraphHealthy = $false
            throw 'Graph token identity or tenant differs from the originally verified client.'
        }

        if ($Initial) {
            $s.TenantId = [string]$org.Id
            $s.TenantName = if ($org.DisplayName) { [string]$org.DisplayName } else { [string]$org.Id }
            $s.Account = [string]$me.userPrincipalName
            $initialDomain = ($org.VerifiedDomains | Where-Object IsInitial | Select-Object -First 1).Name
            $global:M365TenantPrefix = if ($initialDomain) { ($initialDomain -split '\.')[0] -replace '[^a-zA-Z0-9]', '' } else { $null }
        }
        if ($TokenResult.refresh_token) { $s.RefreshToken = $TokenResult.refresh_token }
        $s.ExpiresAt = (Get-Date).ToUniversalTime().AddSeconds([int]$TokenResult.expires_in)
        $s.GraphHealthy = $true
        $s.RefreshError = $false
    }
    finally { $secure = $null }
}

function global:Update-AdhocGraphToken {
    [CmdletBinding()]
    param([switch]$Force)
    $s = $global:AdhocM365
    if (-not $s.RefreshToken -or -not $s.TenantId) { throw 'No verified Graph session is available.' }
    if ($s.RefreshBusy) { return }
    if (-not $Force -and $s.ExpiresAt -and (Get-Date).ToUniversalTime() -lt $s.ExpiresAt.AddMinutes(-5)) { return }

    $s.RefreshBusy = $true
    $newToken = $null
    try {
        $newToken = Invoke-RestMethod -Method Post -Uri "$($s.Authority)/oauth2/v2.0/token" -ContentType 'application/x-www-form-urlencoded' -Body @{
            client_id     = $s.ClientId
            grant_type    = 'refresh_token'
            refresh_token = $s.RefreshToken
            scope         = ($s.Scopes -join ' ')
        } -ErrorAction Stop
        if (-not $newToken.access_token) { throw 'The token endpoint returned no access token.' }
        Set-AdhocGraphToken -TokenResult $newToken -ErrorAction Stop
        Write-Host "`n[OK] Graph token renewed and tenant verified" -ForegroundColor Green
    }
    catch {
        $s.RefreshError = $true
        $s.GraphHealthy = $false
        throw "Graph renewal failed: $($_.Exception.Message)"
    }
    finally {
        $newToken = $null
        $s.RefreshBusy = $false
    }
}

function global:Add-AdhocGraphScopes {
    [CmdletBinding()]
    param([Parameter(Mandatory)][string[]]$Scopes)
    $s = $global:AdhocM365
    if (-not $s.TenantId) { throw 'Connect an ad-hoc tenant first.' }
    $combined = @($s.Scopes + $Scopes | Where-Object { $_ } | Sort-Object -Unique)
    $new = @($combined | Where-Object { $_ -notin $s.Scopes })
    if (-not $new.Count) { Write-Host '[OK] Requested Graph scopes already configured.' -ForegroundColor Green; return }

    Write-Host "Requesting additional Graph scopes: $($new -join ', ')" -ForegroundColor Cyan
    $result = $null
    try {
        $result = Invoke-AdhocDeviceGrant -Scopes $combined
        # Identity and tenant must remain the same.
        Set-AdhocGraphToken -TokenResult $result -ErrorAction Stop
        $s.Scopes = $combined
        Write-Host '[OK] Graph scopes updated and tenant verified.' -ForegroundColor Green
    }
    finally { $result = $null }
}

function global:Get-AdhocExoStatus {
    $s = $global:AdhocM365
    if (-not $s.TenantId -or -not (Get-Command Get-ConnectionInformation -ErrorAction SilentlyContinue)) { return $false }
    try {
        $connections = @(Get-ConnectionInformation -ErrorAction Stop | Where-Object { $_.State -eq 'Connected' })
        return [bool]($connections.Count -eq 1 -and [string]$connections[0].TenantID -eq $s.TenantId)
    }
    catch { return $false }
}

function global:Connect-AdhocExchange {
    [CmdletBinding()]
    param()
    if (-not $global:AdhocM365.TenantId) { throw 'Graph tenant verification is required before Exchange.' }
    Import-Module ExchangeOnlineManagement -ErrorAction Stop
    Connect-ExchangeOnline -Device -ShowBanner:$false -ErrorAction Stop
    if (-not (Get-AdhocExoStatus)) { throw 'Exchange Online tenant mismatch or multiple active connections. Do not administer this session.' }
    Write-Host '[OK] Exchange Online verified' -ForegroundColor Green
}

function global:Connect-AdhocTeams {
    [CmdletBinding()]
    param()
    $s = $global:AdhocM365
    $s.TeamsVerified = $false
    if (-not $s.TenantId) { throw 'Graph tenant verification is required before Teams.' }
    Import-Module MicrosoftTeams -ErrorAction Stop
    Connect-MicrosoftTeams -TenantId $s.TenantId -UseDeviceAuthentication -ErrorAction Stop | Out-Null
    $tenant = Get-CsTenant -ErrorAction Stop
    if (-not $tenant -or [string]$tenant.TenantId -ne $s.TenantId) { throw 'Teams tenant verification failed.' }
    $s.TeamsVerified = $true
    Write-Host '[OK] Microsoft Teams verified' -ForegroundColor Green
}

function global:Assert-AdhocClientContext {
    [CmdletBinding()]
    param([ValidateSet('All','Graph','Exchange','Teams')][string]$Service = 'All')
    $s = $global:AdhocM365
    if (-not $s.TenantId) { throw 'No ad-hoc tenant has been verified.' }
    $selection = if ($Service -eq 'All') { @('Graph','Exchange','Teams') } else { @($Service) }
    $results = foreach ($item in $selection) {
        $status = 'FAILED'
        try {
            switch ($item) {
                Graph {
                    Update-AdhocGraphToken -ErrorAction Stop
                    $org = Get-MgOrganization -Property Id -ErrorAction Stop | Select-Object -First 1
                    if ($s.GraphHealthy -and $org -and [string]$org.Id -eq $s.TenantId) { $status = 'VERIFIED' }
                }
                Exchange { if (Get-AdhocExoStatus) { $status = 'VERIFIED' } }
                Teams {
                    # Explicit assertion performs a live remote tenant check.
                    $tenant = Get-CsTenant -ErrorAction Stop
                    if ($tenant -and [string]$tenant.TenantId -eq $s.TenantId) {
                        $s.TeamsVerified = $true
                        $status = 'VERIFIED'
                    } else { $s.TeamsVerified = $false }
                }
            }
        }
        catch {
            if ($item -eq 'Teams') { $s.TeamsVerified = $false }
            Write-Warning "$item validation failed: $($_.Exception.Message)"
        }
        [PSCustomObject]@{ Service = $item; Status = $status; ExpectedTenant = $s.TenantId }
    }
    Write-Host "`nCLIENT CONTEXT: $($s.TenantName)" -ForegroundColor Cyan
    $results | Format-Table -AutoSize
    if (@($results | Where-Object Status -ne 'VERIFIED').Count) { throw 'One or more client contexts failed verification. Do not run administrative changes.' }
    Write-Host '[PASS] Requested services match the expected tenant.' -ForegroundColor Green
}

function global:Repair-AdhocClientSession {
    [CmdletBinding()]
    param([ValidateSet('All','Graph','Exchange','Teams')][string]$Service = 'All', [switch]$Force)
    $s = $global:AdhocM365
    if (-not $s.TenantId) { throw 'No ad-hoc tenant has been verified.' }
    $selection = if ($Service -eq 'All') { @('Graph','Exchange','Teams') } else { @($Service) }
    foreach ($item in $selection) {
        Write-Host "`nChecking $item..." -ForegroundColor Cyan
        try {
            switch ($item) {
                Graph {
                    $graphContext = Get-MgContext -ErrorAction SilentlyContinue
                    if ($Force -or -not $graphContext -or $graphContext.AuthType -ne 'UserProvidedAccessToken' -or
                        $s.RefreshError -or -not $s.GraphHealthy -or
                        (Get-Date).ToUniversalTime() -ge $s.ExpiresAt.AddMinutes(-5)) {
                        Update-AdhocGraphToken -Force -ErrorAction Stop
                    }
                }
                Exchange {
                    if ($Force -or -not (Get-AdhocExoStatus)) { Connect-AdhocExchange -ErrorAction Stop }
                }
                Teams {
                    if ($Force -or -not $s.TeamsVerified) { Connect-AdhocTeams -ErrorAction Stop }
                    else {
                        $tenant = Get-CsTenant -ErrorAction Stop
                        if (-not $tenant -or [string]$tenant.TenantId -ne $s.TenantId) { Connect-AdhocTeams -ErrorAction Stop }
                    }
                }
            }
            Assert-AdhocClientContext -Service $item -ErrorAction Stop
        }
        catch { Write-Warning "Repair failed for ${item}: $($_.Exception.Message)" }
    }
}

function global:prompt {
    $s = $global:AdhocM365
    if (-not $s) { return "PS $((Get-Location).Path)> " }
    if ($s.RefreshToken -and -not $s.RefreshError -and -not $s.RefreshBusy -and $s.ExpiresAt -and
        (Get-Date).ToUniversalTime() -ge $s.ExpiresAt.AddMinutes(-5)) {
        try { Update-AdhocGraphToken -ErrorAction Stop }
        catch { Write-Warning $_.Exception.Message }
    }
    $ctx = Get-MgContext -ErrorAction SilentlyContinue
    $graphOk = [bool]($ctx -and $ctx.AuthType -eq 'UserProvidedAccessToken' -and
        $s.GraphHealthy -and -not $s.RefreshError -and $s.ExpiresAt -and
        (Get-Date).ToUniversalTime() -lt $s.ExpiresAt)
    $exoOk = Get-AdhocExoStatus
    $client = if ($s.TenantName) { $s.TenantName.ToUpperInvariant() } else { 'UNKNOWN' }

    Write-Host "[$client]" -ForegroundColor Cyan -NoNewline
    Write-Host $(if ($graphOk) { ' [GRAPH:OK]' } else { ' [GRAPH:CHECK]' }) -ForegroundColor $(if ($graphOk) { 'Green' } else { 'Red' }) -NoNewline
    Write-Host $(if ($exoOk) { ' [EXO:OK]' } else { ' [EXO:CHECK]' }) -ForegroundColor $(if ($exoOk) { 'Green' } else { 'Red' }) -NoNewline
    Write-Host $(if ($s.TeamsVerified) { ' [TEAMS:VERIFIED]' } else { ' [TEAMS:CHECK]' }) -ForegroundColor $(if ($s.TeamsVerified) { 'Green' } else { 'Red' })
    return "PS $((Get-Location).Path)> "
}

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host '  AD-HOC MICROSOFT 365 CONNECTION' -ForegroundColor Cyan
Write-Host '========================================' -ForegroundColor Cyan

$initialToken = $null
try {
    Write-Host "`n[1/3] Connecting to Microsoft Graph..." -ForegroundColor Yellow
    $initialToken = Invoke-AdhocDeviceGrant -Scopes $global:AdhocM365.Scopes
    Set-AdhocGraphToken -TokenResult $initialToken -Initial
    Write-Host "[OK] Graph verified: $($global:AdhocM365.TenantName)" -ForegroundColor Green
    Write-Host "Tenant: $($global:AdhocM365.TenantId)" -ForegroundColor Cyan
    Write-Host "Account: $($global:AdhocM365.Account)" -ForegroundColor Cyan
}
catch {
    $global:AdhocM365.GraphHealthy = $false
    Write-Warning "Graph setup failed: $($_.Exception.Message)"
    return
}
finally { $initialToken = $null }

Write-Host "`n[2/3] Connecting to Exchange Online..." -ForegroundColor Yellow
try { Connect-AdhocExchange -ErrorAction Stop }
catch { Write-Warning "Exchange failed: $($_.Exception.Message)" }

Write-Host "`n[3/3] Connecting to Microsoft Teams..." -ForegroundColor Yellow
try { Connect-AdhocTeams -ErrorAction Stop }
catch { Write-Warning "Teams failed: $($_.Exception.Message)" }

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host '  CONNECTION SUMMARY' -ForegroundColor Cyan
Write-Host '========================================' -ForegroundColor Cyan
$s = $global:AdhocM365
[PSCustomObject]@{
    Client   = $s.TenantName
    TenantId = $s.TenantId
    Account  = $s.Account
    Graph    = if ($s.GraphHealthy) { 'VERIFIED' } else { 'CHECK' }
    Exchange = if (Get-AdhocExoStatus) { 'VERIFIED' } else { 'CHECK' }
    Teams    = if ($s.TeamsVerified) { 'VERIFIED' } else { 'CHECK' }
} | Format-List
Write-Host 'Use Assert-AdhocClientContext before sensitive changes.' -ForegroundColor Yellow
