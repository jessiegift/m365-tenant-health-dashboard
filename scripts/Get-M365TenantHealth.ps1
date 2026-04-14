param(
    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientSecret
)

# ─────────────────────────────────────────────
# 1. Get Graph token
# ─────────────────────────────────────────────
$body = @{
    grant_type    = "client_credentials"
    scope         = "https://graph.microsoft.com/.default"
    client_id     = $ClientId
    client_secret = $ClientSecret
}

try {
    $tokenResponse = Invoke-RestMethod -Method Post `
        -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body $body
    $accessToken = $tokenResponse.access_token
    Write-Output "✅ Token acquired"
} catch {
    Write-Error "❌ Failed to get token: $($_.Exception.Message)"
    exit 1
}

$headers = @{ Authorization = "Bearer $accessToken" }

# ─────────────────────────────────────────────
# 2. Service Health
# ─────────────────────────────────────────────
try {
    $serviceHealth = Invoke-RestMethod -Headers $headers `
        -Uri "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/healthOverviews" -Method Get
    Write-Output "✅ Service health: retrieved $($serviceHealth.value.Count) services"
} catch {
    Write-Error "❌ Failed to get service health: $($_.Exception.Message)"
    $serviceHealth = @{ value = @() }
}

# ─────────────────────────────────────────────
# 3. License Usage
# ─────────────────────────────────────────────
try {
    $licenses = Invoke-RestMethod -Headers $headers `
        -Uri "https://graph.microsoft.com/v1.0/subscribedSkus" -Method Get
    Write-Output "✅ Licenses: retrieved $($licenses.value.Count) SKUs"
} catch {
    Write-Error "❌ Failed to get licenses: $($_.Exception.Message)"
    $licenses = @{ value = @() }
}

# ─────────────────────────────────────────────
# 4. Users with Sign-in Activity (paginated)
# ─────────────────────────────────────────────
$allUsers = @()
$usersUri = "https://graph.microsoft.com/v1.0/users?`$select=displayName,userPrincipalName,signInActivity&`$top=999"

try {
    do {
        $response = Invoke-RestMethod -Headers $headers -Uri $usersUri -Method Get
        $allUsers += $response.value
        $usersUri = $response.'@odata.nextLink'
    } while ($usersUri)
    Write-Output "✅ Users: retrieved $($allUsers.Count) total users"
} catch {
    Write-Error "❌ Failed to get users: $($_.Exception.Message)"
    $allUsers = @()
}

# Inactive users (30+ days)
$threshold = (Get-Date).AddDays(-30)
$inactiveUsers = @()

foreach ($u in $allUsers) {
    if ($u.signInActivity.lastSignInDateTime) {
        $last = [datetime]$u.signInActivity.lastSignInDateTime
        if ($last -lt $threshold) {
            $inactiveUsers += [pscustomobject]@{
                DisplayName       = $u.displayName
                UserPrincipalName = $u.userPrincipalName
                LastSignIn        = $last.ToString("yyyy-MM-dd")
            }
        }
    }
}

# ─────────────────────────────────────────────
# 5. MFA Adoption Rate
# ─────────────────────────────────────────────
$mfaRate = "N/A"
$mfaRegistered = 0
$mfaTotal = 0

try {
    $authMethods = Invoke-RestMethod -Headers $headers `
        -Uri "https://graph.microsoft.com/v1.0/reports/authenticationMethods/userRegistrationDetails" -Method Get

    $mfaTotal      = $authMethods.value.Count
    $mfaRegistered = ($authMethods.value | Where-Object { $_.isMfaRegistered -eq $true }).Count
    $mfaRate       = if ($mfaTotal -gt 0) { [math]::Round(($mfaRegistered / $mfaTotal) * 100, 1) } else { 0 }

    Write-Output "✅ MFA adoption: $mfaRate% ($mfaRegistered / $mfaTotal users)"
} catch {
    Write-Error "❌ Failed to get MFA data: $($_.Exception.Message)"
}

# ─────────────────────────────────────────────
# 6. Build HTML Dashboard
# ─────────────────────────────────────────────

# Determine MFA colour
$mfaColour = if ($mfaRate -eq "N/A") { "grey" }
             elseif ([double]$mfaRate -ge 90) { "#2e7d32" }
             elseif ([double]$mfaRate -ge 70) { "#f57c00" }
             else { "#c62828" }

$htmlHeader = @"
<html>
<head>
    <title>M365 Tenant Health Dashboard</title>
    <style>
        body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; background: #fafafa; color: #333; }
        h1 { color: #0f4c81; }
        h2 { color: #0f4c81; margin-top: 30px; }
        .summary-cards { display: flex; gap: 20px; margin-bottom: 30px; flex-wrap: wrap; }
        .card { background: white; border-radius: 8px; padding: 20px; box-shadow: 0 2px 4px rgba(0,0,0,0.1); min-width: 200px; }
        .card .number { font-size: 42px; font-weight: bold; }
        .card .label { font-size: 14px; color: #666; margin-top: 5px; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 30px; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
        th, td { border: 1px solid #eee; padding: 10px 14px; font-size: 13px; }
        th { background-color: #0f4c81; color: white; text-align: left; }
        tr:nth-child(even) { background-color: #f9f9f9; }
        .updated { color: #888; font-size: 13px; }
    </style>
</head>
<body>
<h1>Microsoft 365 Tenant Health Dashboard</h1>
<p class="updated">Last updated: $(Get-Date -Format "yyyy-MM-dd HH:mm:ss") UTC</p>

<div class="summary-cards">
    <div class="card">
        <div class="number" style="color: $mfaColour;">$mfaRate%</div>
        <div class="label">MFA Adoption ($mfaRegistered / $mfaTotal users)</div>
    </div>
    <div class="card">
        <div class="number" style="color: #c62828;">$($inactiveUsers.Count)</div>
        <div class="label">Inactive Users (30+ days)</div>
    </div>
    <div class="card">
        <div class="number" style="color: #0f4c81;">$($licenses.value.Count)</div>
        <div class="label">License SKUs</div>
    </div>
</div>
"@

# Service health with colour-coded status
$serviceRows = foreach ($svc in $serviceHealth.value) {
    $icon = switch ($svc.status) {
        "serviceOperational"  { "🟢" }
        "serviceDegradation"  { "🟡" }
        "serviceInterruption" { "🔴" }
        "investigating"       { "🟠" }
        default               { "⚪" }
    }
    "<tr><td>$($svc.service)</td><td>$icon $($svc.status)</td></tr>"
}

$serviceTable = @"
<h2>Service Health</h2>
<table>
<tr><th>Service</th><th>Status</th></tr>
$($serviceRows -join "`n")
</table>
"@

# License table
$licenseRows = foreach ($lic in $licenses.value) {
    $consumed = $lic.consumedUnits
    $total    = $lic.prepaidUnits.enabled
    "<tr><td>$($lic.skuPartNumber)</td><td>$consumed</td><td>$total</td></tr>"
}

$licenseTable = @"
<h2>License Usage</h2>
<table>
<tr><th>SKU</th><th>Consumed</th><th>Total</th></tr>
$($licenseRows -join "`n")
</table>
"@

# Inactive users table
$inactiveTable = $inactiveUsers |
    Sort-Object LastSignIn |
    ConvertTo-Html -Fragment -PreContent "<h2>Inactive Users (30+ days)</h2>"

$htmlFooter = @"
</body>
</html>
"@

$fullHtml = $htmlHeader + $serviceTable + $licenseTable + $inactiveTable + $htmlFooter

# ─────────────────────────────────────────────
# 7. Save to /docs/index.html
# ─────────────────────────────────────────────
$docsPath = Join-Path $PSScriptRoot "..\docs\index.html"
$fullHtml | Out-File -FilePath $docsPath -Encoding UTF8
Write-Output "✅ Dashboard saved to $docsPath"
