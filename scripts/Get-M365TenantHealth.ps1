param(
    [string]$TenantId,
    [string]$ClientId,
    [string]$ClientSecret
)

# 1. Get Graph token using client credentials
$body = @{
    grant_type    = "client_credentials"
    scope         = "https://graph.microsoft.com/.default"
    client_id     = $ClientId
    client_secret = $ClientSecret
}

$tokenResponse = Invoke-RestMethod -Method Post -Uri "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token" -Body $body
$accessToken   = $tokenResponse.access_token

$headers = @{
    Authorization = "Bearer $accessToken"
}

# 2. Call Graph endpoints

# Service health
$serviceHealth = Invoke-RestMethod -Headers $headers -Uri "https://graph.microsoft.com/v1.0/admin/serviceAnnouncement/health" -Method Get

# License usage
$licenses = Invoke-RestMethod -Headers $headers -Uri "https://graph.microsoft.com/v1.0/subscribedSkus" -Method Get

# Users with sign-in activity
$users = Invoke-RestMethod -Headers $headers -Uri "https://graph.microsoft.com/v1.0/users?`$select=displayName,userPrincipalName,signInActivity&`$top=999" -Method Get

# Inactive users (30+ days)
$threshold = (Get-Date).AddDays(-30)
$inactiveUsers = @()

foreach ($u in $users.value) {
    if ($u.signInActivity.lastSignInDateTime) {
        $last = [datetime]$u.signInActivity.lastSignInDateTime
        if ($last -lt $threshold) {
            $inactiveUsers += [pscustomobject]@{
                DisplayName       = $u.displayName
                UserPrincipalName = $u.userPrincipalName
                LastSignIn        = $last
            }
        }
    }
}

# 3. Build HTML dashboard

$htmlHeader = @"
<html>
<head>
    <title>M365 Tenant Health Dashboard</title>
    <style>
        body { font-family: Arial; margin: 20px; }
        h1, h2 { color: #0f4c81; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 30px; }
        th, td { border: 1px solid #ddd; padding: 8px; font-size: 13px; }
        th { background-color: #f2f2f2; text-align: left; }
    </style>
</head>
<body>
<h1>Microsoft 365 Tenant Health Dashboard</h1>
<p>Last updated: $(Get-Date)</p>
"@

$serviceTable = $serviceHealth.value |
    Select-Object service, status, classification |
    ConvertTo-Html -Fragment -PreContent "<h2>Service Health</h2>"

$licenseTable = $licenses.value |
    Select-Object skuPartNumber, consumedUnits, @{n="TotalUnits";e={$_.prepaidUnits.enabled}} |
    ConvertTo-Html -Fragment -PreContent "<h2>License Usage</h2>"

$inactiveTable = $inactiveUsers |
    Sort-Object LastSignIn |
    ConvertTo-Hhtml -Fragment -PreContent "<h2>Inactive Users (30+ days)</h2>"

$htmlFooter = @"
</body>
</html>
"@

$fullHtml = $htmlHeader + $serviceTable + $licenseTable + $inactiveTable + $htmlFooter

# 4. Save to /docs/index.html
$docsPath = Join-Path $PSScriptRoot "..\docs\index.html"
$fullHtml | Out-File -FilePath $docsPath -Encoding UTF8
