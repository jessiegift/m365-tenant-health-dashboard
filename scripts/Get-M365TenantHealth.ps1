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

$updatedTime = Get-Date -Format "dd MMM yyyy, HH:mm UTC"

# Counts for summary cards
$healthyCount   = ($serviceHealth.value | Where-Object { $_.status -eq "serviceOperational" }).Count
$degradedCount  = ($serviceHealth.value | Where-Object { $_.status -ne "serviceOperational" }).Count
$totalServices  = $serviceHealth.value.Count
$totalUsers     = $allUsers.Count
$inactiveCount  = $inactiveUsers.Count
$licenseCount   = $licenses.value.Count

# MFA colour logic
$mfaColourHex = if ($mfaRate -eq "N/A") { "#94a3b8" }
                elseif ([double]$mfaRate -ge 90) { "#10b981" }
                elseif ([double]$mfaRate -ge 70) { "#f59e0b" }
                else { "#ef4444" }

$mfaLabel = if ($mfaRate -eq "N/A") { "N/A" } else { "$mfaRate%" }

# Service health rows
$serviceRows = foreach ($svc in ($serviceHealth.value | Sort-Object status)) {
    $statusText = $svc.status
    $badgeClass = switch ($statusText) {
        "serviceOperational"  { "badge-ok" }
        "serviceDegradation"  { "badge-warn" }
        "serviceInterruption" { "badge-crit" }
        "investigating"       { "badge-warn" }
        default               { "badge-unknown" }
    }
    $dot = switch ($statusText) {
        "serviceOperational"  { "dot-ok" }
        "serviceDegradation"  { "dot-warn" }
        "serviceInterruption" { "dot-crit" }
        default               { "dot-unknown" }
    }
    $friendly = switch ($statusText) {
        "serviceOperational"  { "Operational" }
        "serviceDegradation"  { "Degraded" }
        "serviceInterruption" { "Outage" }
        "investigating"       { "Investigating" }
        default               { $statusText }
    }
    "<tr><td><span class='svc-name'>$($svc.service)</span></td><td><span class='badge $badgeClass'><span class='dot $dot'></span>$friendly</span></td></tr>"
}

# License rows
$licenseRows = foreach ($lic in $licenses.value) {
    $consumed = $lic.consumedUnits
    $total    = $lic.prepaidUnits.enabled
    $pct      = if ($total -gt 0) { [math]::Round(($consumed / $total) * 100) } else { 0 }
    $barColour = if ($pct -ge 90) { "#ef4444" } elseif ($pct -ge 70) { "#f59e0b" } else { "#10b981" }
    $skuName  = $lic.skuPartNumber -replace "_", " "
    @"
<tr>
  <td><span class='svc-name'>$skuName</span></td>
  <td class='num-cell'>$consumed</td>
  <td class='num-cell'>$total</td>
  <td>
    <div class='bar-wrap'><div class='bar-fill' style='width:$pct%;background:$barColour'></div></div>
    <span class='bar-pct'>$pct%</span>
  </td>
</tr>
"@
}

# Inactive users rows
$inactiveRows = foreach ($u in ($inactiveUsers | Sort-Object LastSignIn)) {
    $initials = ($u.DisplayName -split " " | ForEach-Object { $_[0] }) -join ""
    if ($initials.Length -gt 2) { $initials = $initials.Substring(0,2) }
    "<tr><td><div class='user-cell'><div class='avatar'>$initials</div><div><div class='user-name'>$($u.DisplayName)</div><div class='user-upn'>$($u.UserPrincipalName)</div></div></div></td><td class='num-cell last-seen'>$($u.LastSignIn)</td></tr>"
}

$fullHtml = @"
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width,initial-scale=1"/>
<title>M365 Tenant Health</title>
<link rel="preconnect" href="https://fonts.googleapis.com"/>
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet"/>
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --bg:       #0d1117;
    --surface:  #161b22;
    --surface2: #1c2230;
    --border:   #21262d;
    --border2:  #30363d;
    --text:     #e6edf3;
    --muted:    #7d8590;
    --accent:   #2f81f7;
    --ok:       #10b981;
    --warn:     #f59e0b;
    --crit:     #ef4444;
    --radius:   12px;
  }

  body {
    font-family: 'DM Sans', sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
    font-size: 14px;
    line-height: 1.6;
  }

  /* ── Layout ── */
  .shell { display: flex; min-height: 100vh; }

  .sidebar {
    width: 220px;
    flex-shrink: 0;
    background: var(--surface);
    border-right: 1px solid var(--border);
    padding: 28px 0;
    position: sticky;
    top: 0;
    height: 100vh;
    display: flex;
    flex-direction: column;
  }

  .sidebar-logo {
    padding: 0 20px 28px;
    border-bottom: 1px solid var(--border);
    margin-bottom: 16px;
  }
  .sidebar-logo .logo-mark {
    width: 32px; height: 32px;
    background: var(--accent);
    border-radius: 8px;
    display: flex; align-items: center; justify-content: center;
    font-size: 16px; font-weight: 600; color: white;
    margin-bottom: 10px;
  }
  .sidebar-logo .tenant-name { font-size: 13px; font-weight: 600; color: var(--text); }
  .sidebar-logo .tenant-sub  { font-size: 11px; color: var(--muted); margin-top: 2px; }

  .nav-label { font-size: 10px; font-weight: 600; color: var(--muted); letter-spacing: 0.08em; text-transform: uppercase; padding: 0 20px 6px; }
  .nav-item {
    display: flex; align-items: center; gap: 10px;
    padding: 8px 20px; font-size: 13px; color: var(--muted);
    cursor: pointer; transition: color 0.15s, background 0.15s;
    border-left: 2px solid transparent;
  }
  .nav-item:hover { color: var(--text); background: var(--surface2); }
  .nav-item.active { color: var(--accent); border-left-color: var(--accent); background: rgba(47,129,247,0.08); }
  .nav-icon { font-size: 15px; }

  .sidebar-footer { margin-top: auto; padding: 16px 20px 0; border-top: 1px solid var(--border); }
  .updated-label { font-size: 10px; color: var(--muted); line-height: 1.5; }
  .updated-time  { font-family: 'DM Mono', monospace; font-size: 10px; color: var(--muted); }

  /* ── Main content ── */
  .main { flex: 1; padding: 32px 36px; overflow-x: hidden; }

  .page-header { margin-bottom: 32px; }
  .page-title { font-size: 22px; font-weight: 600; color: var(--text); letter-spacing: -0.3px; }
  .page-sub { font-size: 13px; color: var(--muted); margin-top: 4px; }

  /* ── Stat cards ── */
  .cards { display: grid; grid-template-columns: repeat(auto-fit, minmax(160px, 1fr)); gap: 16px; margin-bottom: 36px; }
  .stat-card {
    background: var(--surface);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    padding: 20px;
    position: relative;
    overflow: hidden;
    transition: border-color 0.2s;
  }
  .stat-card:hover { border-color: var(--border2); }
  .stat-card::before {
    content: '';
    position: absolute; top: 0; left: 0; right: 0; height: 2px;
    background: var(--card-accent, var(--accent));
  }
  .stat-label { font-size: 11px; font-weight: 500; color: var(--muted); text-transform: uppercase; letter-spacing: 0.06em; margin-bottom: 10px; }
  .stat-value { font-size: 34px; font-weight: 600; color: var(--card-color, var(--text)); font-variant-numeric: tabular-nums; letter-spacing: -1px; }
  .stat-sub { font-size: 11px; color: var(--muted); margin-top: 6px; }

  /* ── Section ── */
  .section { margin-bottom: 36px; }
  .section-header { display: flex; align-items: center; justify-content: space-between; margin-bottom: 14px; }
  .section-title { font-size: 14px; font-weight: 600; color: var(--text); }
  .section-count { font-size: 12px; color: var(--muted); background: var(--surface2); border: 1px solid var(--border); border-radius: 20px; padding: 2px 10px; }

  /* ── Table ── */
  .table-wrap { background: var(--surface); border: 1px solid var(--border); border-radius: var(--radius); overflow: hidden; }
  table { width: 100%; border-collapse: collapse; }
  thead th {
    background: var(--surface2);
    font-size: 11px; font-weight: 600; color: var(--muted);
    text-transform: uppercase; letter-spacing: 0.06em;
    padding: 10px 16px; text-align: left;
    border-bottom: 1px solid var(--border);
  }
  tbody tr { border-bottom: 1px solid var(--border); transition: background 0.1s; }
  tbody tr:last-child { border-bottom: none; }
  tbody tr:hover { background: var(--surface2); }
  td { padding: 11px 16px; }

  .svc-name { font-size: 13px; color: var(--text); font-weight: 400; }
  .num-cell { font-family: 'DM Mono', monospace; font-size: 13px; color: var(--muted); text-align: right; }

  /* ── Badges ── */
  .badge {
    display: inline-flex; align-items: center; gap: 6px;
    font-size: 11px; font-weight: 500; padding: 3px 10px;
    border-radius: 20px; letter-spacing: 0.02em;
  }
  .badge-ok      { background: rgba(16,185,129,0.12); color: #34d399; border: 1px solid rgba(16,185,129,0.25); }
  .badge-warn    { background: rgba(245,158,11,0.12); color: #fbbf24; border: 1px solid rgba(245,158,11,0.25); }
  .badge-crit    { background: rgba(239,68,68,0.12);  color: #f87171; border: 1px solid rgba(239,68,68,0.25); }
  .badge-unknown { background: rgba(125,133,144,0.12);color: #94a3b8; border: 1px solid rgba(125,133,144,0.25); }

  .dot { width: 6px; height: 6px; border-radius: 50%; display: inline-block; }
  .dot-ok      { background: #10b981; box-shadow: 0 0 6px #10b981; }
  .dot-warn    { background: #f59e0b; box-shadow: 0 0 6px #f59e0b; }
  .dot-crit    { background: #ef4444; box-shadow: 0 0 6px #ef4444; animation: pulse 1.5s infinite; }
  .dot-unknown { background: #94a3b8; }

  @keyframes pulse {
    0%, 100% { box-shadow: 0 0 4px #ef4444; }
    50%       { box-shadow: 0 0 10px #ef4444; }
  }

  /* ── Progress bars ── */
  .bar-wrap { display: inline-block; width: 80px; height: 4px; background: var(--border2); border-radius: 2px; vertical-align: middle; margin-right: 8px; overflow: hidden; }
  .bar-fill { height: 100%; border-radius: 2px; transition: width 0.3s; }
  .bar-pct  { font-family: 'DM Mono', monospace; font-size: 11px; color: var(--muted); }

  /* ── User cells ── */
  .user-cell { display: flex; align-items: center; gap: 10px; }
  .avatar {
    width: 30px; height: 30px; border-radius: 50%;
    background: rgba(47,129,247,0.15); border: 1px solid rgba(47,129,247,0.3);
    display: flex; align-items: center; justify-content: center;
    font-size: 11px; font-weight: 600; color: var(--accent);
    flex-shrink: 0; text-transform: uppercase;
  }
  .user-name { font-size: 13px; color: var(--text); }
  .user-upn  { font-size: 11px; color: var(--muted); font-family: 'DM Mono', monospace; }
  .last-seen { color: var(--warn) !important; }
</style>
</head>
<body>
<div class="shell">

  <!-- Sidebar -->
  <nav class="sidebar">
    <div class="sidebar-logo">
      <div class="logo-mark">M</div>
      <div class="tenant-name">Tenant Health</div>
      <div class="tenant-sub">Microsoft 365</div>
    </div>
    <div class="nav-label">Overview</div>
    <div class="nav-item active"><span class="nav-icon">&#9681;</span> Dashboard</div>
    <div class="nav-item"><span class="nav-icon">&#9432;</span> Service Health</div>
    <div class="nav-item"><span class="nav-icon">&#128100;</span> Users</div>
    <div class="nav-item"><span class="nav-icon">&#128273;</span> Licences</div>
    <div class="nav-item"><span class="nav-icon">&#128737;</span> Security</div>
    <div class="sidebar-footer">
      <div class="updated-label">Last refreshed</div>
      <div class="updated-time">$updatedTime</div>
    </div>
  </nav>

  <!-- Main -->
  <main class="main">
    <div class="page-header">
      <div class="page-title">Tenant Overview</div>
      <div class="page-sub">Real-time health data from Microsoft Graph API</div>
    </div>

    <!-- Stat cards -->
    <div class="cards">
      <div class="stat-card" style="--card-accent:#10b981; --card-color:#10b981;">
        <div class="stat-label">Services healthy</div>
        <div class="stat-value">$healthyCount</div>
        <div class="stat-sub">of $totalServices monitored</div>
      </div>
      <div class="stat-card" style="--card-accent:#ef4444; --card-color:#ef4444;">
        <div class="stat-label">Degraded / issues</div>
        <div class="stat-value">$degradedCount</div>
        <div class="stat-sub">requires attention</div>
      </div>
      <div class="stat-card" style="--card-accent:$mfaColourHex; --card-color:$mfaColourHex;">
        <div class="stat-label">MFA adoption</div>
        <div class="stat-value">$mfaLabel</div>
        <div class="stat-sub">$mfaRegistered of $mfaTotal users</div>
      </div>
      <div class="stat-card" style="--card-accent:#f59e0b; --card-color:#f59e0b;">
        <div class="stat-label">Inactive users</div>
        <div class="stat-value">$inactiveCount</div>
        <div class="stat-sub">no sign-in 30+ days</div>
      </div>
      <div class="stat-card" style="--card-accent:#2f81f7; --card-color:#2f81f7;">
        <div class="stat-label">Total users</div>
        <div class="stat-value">$totalUsers</div>
        <div class="stat-sub">in directory</div>
      </div>
      <div class="stat-card" style="--card-accent:#a78bfa; --card-color:#a78bfa;">
        <div class="stat-label">Licence SKUs</div>
        <div class="stat-value">$licenseCount</div>
        <div class="stat-sub">active subscriptions</div>
      </div>
    </div>

    <!-- Service health -->
    <div class="section">
      <div class="section-header">
        <span class="section-title">Service health</span>
        <span class="section-count">$totalServices services</span>
      </div>
      <div class="table-wrap">
        <table>
          <thead><tr><th>Service</th><th>Status</th></tr></thead>
          <tbody>
            $($serviceRows -join "`n")
          </tbody>
        </table>
      </div>
    </div>

    <!-- Licence usage -->
    <div class="section">
      <div class="section-header">
        <span class="section-title">Licence usage</span>
        <span class="section-count">$licenseCount SKUs</span>
      </div>
      <div class="table-wrap">
        <table>
          <thead><tr><th>SKU</th><th style="text-align:right">Consumed</th><th style="text-align:right">Total</th><th>Utilisation</th></tr></thead>
          <tbody>
            $($licenseRows -join "`n")
          </tbody>
        </table>
      </div>
    </div>

    <!-- Inactive users -->
    <div class="section">
      <div class="section-header">
        <span class="section-title">Inactive users <span style="color:#f59e0b">&#9888;</span></span>
        <span class="section-count">$inactiveCount users</span>
      </div>
      <div class="table-wrap">
        <table>
          <thead><tr><th>User</th><th style="text-align:right">Last sign-in</th></tr></thead>
          <tbody>
            $($inactiveRows -join "`n")
          </tbody>
        </table>
      </div>
    </div>

  </main>
</div>
</body>
</html>
"@

# ─────────────────────────────────────────────
# 7. Save to /docs/index.html
# ─────────────────────────────────────────────
$docsPath = Join-Path $PSScriptRoot "..\docs\index.html"
$fullHtml | Out-File -FilePath $docsPath -Encoding UTF8
Write-Output "✅ Dashboard saved to $docsPath"
