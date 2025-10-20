<#
.SYNOPSIS
  Azure Service Health Report (Management Group Level)
.DESCRIPTION
  Retrieves Azure Service Health issues, advisories, and maintenance events
  across all subscriptions under a management group and emails an HTML report.
#>

# ----------------------------
# CONFIGURATION
# ----------------------------
$ManagementGroupId = "YourMgmtGroupIdHere"
$OutputHtmlFile = "C:\Temp\ServiceHealthReport.html"

# Email details (optional - if you want to send mail)
$To = "you@example.com"
$From = "azurehealth@example.com"
$SmtpServer = "smtp.office365.com"
$Subject = "Azure Service Health Report - $(Get-Date -Format 'dd-MMM-yyyy')"

# ----------------------------
# CONNECT TO AZURE
# ----------------------------
Write-Host "🔹 Connecting to Azure..."
Connect-AzAccount -WarningAction SilentlyContinue | Out-Null
Write-Host "✅ Connected to Azure."

# ----------------------------
# GET SUBSCRIPTIONS UNDER MGMT GROUP
# ----------------------------
Write-Host "🔹 Getting subscriptions under management group: $ManagementGroupId"
$subs = Get-AzManagementGroupSubscription -GroupId $ManagementGroupId -ErrorAction Stop
if (-not $subs) {
    Write-Host "⚠️ No subscriptions found under management group $ManagementGroupId."
    exit
}

# ----------------------------
# COLLECT SERVICE HEALTH EVENTS
# ----------------------------
$allEvents = @()

foreach ($sub in $subs) {
    $subId = $sub.Id -replace ".*/"
    Set-AzContext -SubscriptionId $subId -ErrorAction SilentlyContinue | Out-Null
    Write-Host "🔸 Checking subscription: $subId"

    try {
        $events = Get-AzServiceHealthEvent -Status Active, Resolved -ErrorAction Stop
        foreach ($event in $events) {
            $regions = if ($event.AffectedRegion.Name) { ($event.AffectedRegion.Name -join ', ') } else { 'Global' }
            $impact = if ($event.ImpactType) { $event.ImpactType } else { 'N/A' }

            $allEvents += [PSCustomObject]@{
                SubscriptionId   = $subId
                SubscriptionName = (Get-AzSubscription -SubscriptionId $subId).Name
                Title            = $event.Title
                Impact           = $impact
                Status           = $event.Status
                IncidentType     = $event.IncidentType
                StartTime        = $event.StartTime
                LastUpdateTime   = $event.LastUpdateTime
                Regions          = $regions
            }
        }
    } catch {
        Write-Host "⚠️ Skipping subscription $subId due to permission or access issues."
    }
}

# ----------------------------
# BUILD ENHANCED HTML REPORT
# ----------------------------
if (-not $allEvents) {
    Write-Host "✅ No events found."
    $htmlBody = @"
    <div style='text-align:center; font-family:Segoe UI;'>
        <h2 style='color:#107C10;'>✅ All Systems Healthy</h2>
        <p style='font-size:14px;'>No active or recent Azure Service Health events found under management group 
        <b>$ManagementGroupId</b>.</p>
    </div>
"@
} else {
    Write-Host "📊 Building enhanced HTML report for $($allEvents.Count) event(s)..."

    # Create summary counts
    $summary = $allEvents | Group-Object -Property IncidentType | ForEach-Object {
        [PSCustomObject]@{
            Type  = $_.Name
            Count = $_.Count
        }
    }

    # Type label mapping
    $typeMap = @{
        "ServiceIssue"        = "Service Issues"
        "PlannedMaintenance"  = "Planned Maintenance"
        "HealthAdvisory"      = "Health Advisories"
        "SecurityAdvisory"    = "Security Advisories"
    }

    # Header
    $htmlBody = @"
    <h2 style='color:#0078D4; font-family:Segoe UI;'>Azure Service Health Summary</h2>
    <p style='font-family:Segoe UI; font-size:13px;'>
        <b>Management Group:</b> $ManagementGroupId <br/>
        <b>Generated:</b> $(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss') <br/>
        <b>Total Events:</b> $($allEvents.Count)
    </p>
"@

    # Summary tiles
    $htmlBody += "<div style='display:flex; gap:15px; flex-wrap:wrap; margin-bottom:20px;'>"
    foreach ($s in $summary) {
        $label = $typeMap[$s.Type]
        $color = switch ($s.Type) {
            "ServiceIssue"        { "#E81123" }
            "PlannedMaintenance"  { "#0078D4" }
            "HealthAdvisory"      { "#107C10" }
            "SecurityAdvisory"    { "#FFB900" }
            default               { "#666" }
        }

        $htmlBody += "<div style='flex:1; min-width:150px; background:$color; color:white; 
                        border-radius:10px; padding:12px; text-align:center;'>
                        <div style='font-size:22px; font-weight:bold;'>$($s.Count)</div>
                        <div style='font-size:13px;'>$label</div>
                      </div>"
    }
    $htmlBody += "</div>"

    # Detailed tables by type
    $grouped = $allEvents | Group-Object -Property IncidentType
    foreach ($group in $grouped) {
        $incidentType = $typeMap[$group.Name]
        $color = switch ($group.Name) {
            "ServiceIssue"        { "#FDE7E9" }
            "PlannedMaintenance"  { "#E7F3FD" }
            "HealthAdvisory"      { "#E6F4EA" }
            "SecurityAdvisory"    { "#FFF4CE" }
            default               { "#f9f9f9" }
        }

        $htmlBody += "<h3 style='color:#333;font-family:Segoe UI;margin-top:25px;'>$incidentType</h3>"
        $htmlBody += "<table style='width:100%; border-collapse:collapse; font-family:Segoe UI; font-size:13px;'>
                        <thead>
                            <tr style='background-color:$color; border-bottom:2px solid #ccc;'>
                                <th style='padding:8px; text-align:left;'>Subscription</th>
                                <th style='padding:8px; text-align:left;'>Title</th>
                                <th style='padding:8px; text-align:left;'>Impact</th>
                                <th style='padding:8px; text-align:left;'>Status</th>
                                <th style='padding:8px; text-align:left;'>Start Time</th>
                                <th style='padding:8px; text-align:left;'>Last Update</th>
                                <th style='padding:8px; text-align:left;'>Region(s)</th>
                            </tr>
                        </thead>
                        <tbody>"
        foreach ($e in $group.Group) {
            $statusColor = if ($e.Status -eq "Active") { "color:#E81123;font-weight:bold;" } else { "color:#107C10;" }

            $htmlBody += "<tr style='border-bottom:1px solid #eee;'>
                            <td style='padding:6px;'>$($e.SubscriptionName)</td>
                            <td style='padding:6px;'>$($e.Title)</td>
                            <td style='padding:6px;'>$($e.Impact)</td>
                            <td style='padding:6px;$statusColor'>$($e.Status)</td>
                            <td style='padding:6px;'>$($e.StartTime.ToString('dd-MMM-yyyy HH:mm'))</td>
                            <td style='padding:6px;'>$($e.LastUpdateTime.ToString('dd-MMM-yyyy HH:mm'))</td>
                            <td style='padding:6px;'>$($e.Regions)</td>
                          </tr>"
        }
        $htmlBody += "</tbody></table><br/>"
    }
}

# Final HTML layout
$emailBody = @"
<html>
<head>
<style>
body { font-family:'Segoe UI', Arial, sans-serif; margin: 20px; color:#222; }
a { color:#0078D4; text-decoration:none; }
a:hover { text-decoration:underline; }
table tr:hover { background-color:#f5f5f5; }
</style>
</head>
<body>
$htmlBody
<p style='font-size:11px; color:#999; margin-top:20px;'>
Generated automatically by Azure Automation Runbook • $(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss')
</p>
</body>
</html>
"@

# Save report locally
$emailBody | Out-File -FilePath $OutputHtmlFile -Encoding UTF8
Write-Host "✅ HTML report saved to: $OutputHtmlFile"

# ----------------------------
# OPTIONAL: SEND EMAIL
# ----------------------------
try {
    Send-MailMessage -To $To -From $From -Subject $Subject -BodyAsHtml -Body $emailBody -SmtpServer $SmtpServer -UseSsl
    Write-Host "📧 Email sent successfully to $To"
} catch {
    Write-Host "⚠️ Failed to send email: $($_.Exception.Message)"
}
