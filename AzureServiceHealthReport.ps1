<#
.SYNOPSIS
  Comprehensive Azure Service Health & Security Report across all subscriptions
.DESCRIPTION
  Fetches Active Alerts, Planned Maintenance, Health Advisories, and Security Advisories
  from all accessible subscriptions and sends an HTML report via Microsoft Graph email.
#>

# ----------------------------
# CONFIGURATION
# ----------------------------
$OutputHtmlFile = "C:\Temp\AzureComprehensiveReport.html"
$To = "you@example.com"
$Subject = "Azure Monitor & Service Health Report - $(Get-Date -Format 'dd-MMM-yyyy')"
$ManagedIdentityClientId = "9464e54c-6ec0-4b15-8380-6172a2e3114b"

# ----------------------------
# CONNECT TO AZURE
# ----------------------------
Write-Host "🔹 Connecting to Azure using Managed Identity..."
Connect-AzAccount -Identity
Write-Host "✅ Connected to Azure."

# ----------------------------
# GET ALL SUBSCRIPTIONS
# ----------------------------
$subscriptions = Get-AzSubscription
Write-Host "🔹 Found $($subscriptions.Count) subscriptions."

# Initialize arrays to hold data
$allAlerts = @()
$allMaintenance = @()
$allHealthAdvisories = @()
$allSecurityAdvisories = @()

foreach ($sub in $subscriptions) {
    Write-Host "🔹 Processing subscription: $($sub.Name) ($($sub.Id))..."
    Set-AzContext -SubscriptionId $sub.Id

    # ----------------------------
    # ACTIVE ALERTS
    # ----------------------------
    try {
        $alerts = Get-AzMetricAlertRuleV2 -DetailedOutput | Where-Object { $_.Enabled -eq $true }
        foreach ($alert in $alerts) {
            $allAlerts += [PSCustomObject]@{
                SubscriptionName = $sub.Name
                AlertName        = $alert.Name
                Severity         = $alert.Severity
                State            = $alert.State
                Resource         = $alert.Scopes -join ", "
                LastUpdated      = $alert.LastUpdatedTime
            }
        }
    } catch { Write-Host "⚠️ Alerts fetch failed for $($sub.Name): $($_.Exception.Message)" }

    # ----------------------------
    # PLANNED MAINTENANCE
    # ----------------------------
    try {
        $maintenance = Get-AzActivityLog -StartTime (Get-Date).AddDays(-7) `
                                         -EndTime (Get-Date) `
                                         -MaxRecord 500 `
                                         -Status Active | Where-Object { $_.EventName.Value -like "*Planned Maintenance*" }
        foreach ($m in $maintenance) {
            $allMaintenance += [PSCustomObject]@{
                SubscriptionName = $sub.Name
                ResourceGroup    = $m.ResourceGroupName
                ResourceName     = $m.ResourceId.Split('/')[-1]
                Event            = $m.EventName.Value
                ScheduledTime    = $m.EventTimestamp.ToLocalTime()
            }
        }
    } catch { Write-Host "⚠️ Planned Maintenance fetch failed for $($sub.Name): $($_.Exception.Message)" }

    # ----------------------------
    # HEALTH ADVISORIES
    # ----------------------------
    try {
        $health = Get-AzResourceHealthAvailabilityStatus -DetailedStatus
        foreach ($h in $health | Where-Object { $_.AvailabilityState -ne "Available" }) {
            $allHealthAdvisories += [PSCustomObject]@{
                SubscriptionName = $sub.Name
                ResourceGroup    = $h.ResourceGroupName
                ResourceName     = $h.ResourceName
                Type             = $h.ResourceType
                Status           = $h.AvailabilityState
                Details          = $h.ReasonType
                LastUpdated      = $h.Timestamp.ToLocalTime()
            }
        }
    } catch { Write-Host "⚠️ Health Advisories fetch failed for $($sub.Name): $($_.Exception.Message)" }

    # ----------------------------
    # SECURITY ADVISORIES
    # ----------------------------
    try {
        $securityAlerts = Get-AzSecurityAlert
        foreach ($sa in $securityAlerts) {
            $allSecurityAdvisories += [PSCustomObject]@{
                SubscriptionName = $sub.Name
                ResourceGroup    = $sa.ResourceGroupName
                ResourceName     = $sa.ResourceName
                AlertType        = $sa.AlertType
                Severity         = $sa.Severity
                Status           = $sa.State
                TimeGenerated    = $sa.TimeGenerated.ToLocalTime()
            }
        }
    } catch { Write-Host "⚠️ Security Advisories fetch failed for $($sub.Name): $($_.Exception.Message)" }
}

# ----------------------------
# BUILD HTML REPORT
# ----------------------------
$htmlBody = "<h2 style='color:#0078D4; font-family:Segoe UI;'>Azure Monitor & Service Health Report</h2>"
$htmlBody += "<p style='font-family:Segoe UI; font-size:13px;'>Generated: $(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss')</p>"

# ----- Active Alerts -----
if ($allAlerts.Count -eq 0) { $htmlBody += "<h3 style='color:#107C10;'>✅ No Active Alerts</h3>" }
else {
    $htmlBody += "<h3 style='color:#E81123;'>📢 Active Alerts ($($allAlerts.Count))</h3>"
    $htmlBody += "<table style='width:100%; border-collapse:collapse; font-family:Segoe UI; font-size:13px;'>"
    $htmlBody += "<thead><tr style='background-color:#E7F3FD; border-bottom:2px solid #ccc;'>
                    <th>Subscription</th><th>Alert Name</th><th>Severity</th><th>State</th><th>Resource</th><th>Last Updated</th></tr></thead><tbody>"
    foreach ($a in $allAlerts) {
        $htmlBody += "<tr style='border-bottom:1px solid #eee;'>
                        <td>$($a.SubscriptionName)</td>
                        <td>$($a.AlertName)</td>
                        <td>$($a.Severity)</td>
                        <td>$($a.State)</td>
                        <td>$($a.Resource)</td>
                        <td>$($a.LastUpdated.ToString('dd-MMM-yyyy HH:mm'))</td>
                      </tr>"
    }
    $htmlBody += "</tbody></table><br/>"
}

# ----- Planned Maintenance -----
if ($allMaintenance.Count -eq 0) { $htmlBody += "<h3 style='color:#107C10;'>✅ No Planned Maintenance</h3>" }
else {
    $htmlBody += "<h3 style='color:#FF8C00;'>🛠 Planned Maintenance ($($allMaintenance.Count))</h3>"
    $htmlBody += "<table style='width:100%; border-collapse:collapse; font-family:Segoe UI; font-size:13px;'>"
    $htmlBody += "<thead><tr style='background-color:#FFF4E5; border-bottom:2px solid #ccc;'>
                    <th>Subscription</th><th>Resource Group</th><th>Resource</th><th>Event</th><th>Scheduled Time</th></tr></thead><tbody>"
    foreach ($m in $allMaintenance) {
        $htmlBody += "<tr style='border-bottom:1px solid #eee;'>
                        <td>$($m.SubscriptionName)</td>
                        <td>$($m.ResourceGroup)</td>
                        <td>$($m.ResourceName)</td>
                        <td>$($m.Event)</td>
                        <td>$($m.ScheduledTime.ToString('dd-MMM-yyyy HH:mm'))</td>
                      </tr>"
    }
    $htmlBody += "</tbody></table><br/>"
}

# ----- Health Advisories -----
if ($allHealthAdvisories.Count -eq 0) { $htmlBody += "<h3 style='color:#107C10;'>✅ No Health Advisories</h3>" }
else {
    $htmlBody += "<h3 style='color:#FFD700;'>⚠ Health Advisories ($($allHealthAdvisories.Count))</h3>"
    $htmlBody += "<table style='width:100%; border-collapse:collapse; font-family:Segoe UI; font-size:13px;'>"
    $htmlBody += "<thead><tr style='background-color:#FFFBE6; border-bottom:2px solid #ccc;'>
                    <th>Subscription</th><th>Resource Group</th><th>Resource</th><th>Type</th><th>Status</th><th>Details</th><th>Last Updated</th></tr></thead><tbody>"
    foreach ($h in $allHealthAdvisories) {
        $htmlBody += "<tr style='border-bottom:1px solid #eee;'>
                        <td>$($h.SubscriptionName)</td>
                        <td>$($h.ResourceGroup)</td>
                        <td>$($h.ResourceName)</td>
                        <td>$($h.Type)</td>
                        <td>$($h.Status)</td>
                        <td>$($h.Details)</td>
                        <td>$($h.LastUpdated.ToString('dd-MMM-yyyy HH:mm'))</td>
                      </tr>"
    }
    $htmlBody += "</tbody></table><br/>"
}

# ----- Security Advisories -----
if ($allSecurityAdvisories.Count -eq 0) { $htmlBody += "<h3 style='color:#107C10;'>✅ No Security Advisories</h3>" }
else {
    $htmlBody += "<h3 style='color:#800080;'>🔒 Security Advisories ($($allSecurityAdvisories.Count))</h3>"
    $htmlBody += "<table style='width:100%; border-collapse:collapse; font-family:Segoe UI; font-size:13px;'>"
    $htmlBody += "<thead><tr style='background-color:#F3E5FF; border-bottom:2px solid #ccc;'>
                    <th>Subscription</th><th>Resource Group</th><th>Resource</th><th>Alert Type</th><th>Severity</th><th>Status</th><th>Time</th></tr></thead><tbody>"
    foreach ($s in $allSecurityAdvisories) {
        $htmlBody += "<tr style='border-bottom:1px solid #eee;'>
                        <td>$($s.SubscriptionName)</td>
                        <td>$($s.ResourceGroup)</td>
                        <td>$($s.ResourceName)</td>
                        <td>$($s.AlertType)</td>
                        <td>$($s.Severity)</td>
                        <td>$($s.Status)</td>
                        <td>$($s.TimeGenerated.ToString('dd-MMM-yyyy HH:mm'))</td>
                      </tr>"
    }
    $htmlBody += "</tbody></table><br/>"
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
th, td { padding:6px; text-align:left; }
</style>
</head>
<body>
$htmlBody
<p style='font-size:11px; color:#999; margin-top:20px;'>Generated automatically by Azure Automation Runbook • $(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss')</p>
</body>
</html>
"@

# Save HTML report locally
$emailBody | Out-File -FilePath $OutputHtmlFile -Encoding UTF8
Write-Host "✅ HTML report saved to: $OutputHtmlFile"

# ----------------------------
# SEND EMAIL USING MICROSOFT GRAPH
# ----------------------------
Write-Host "🔹 Connecting to Microsoft Graph..."
Connect-MgGraph -Identity -ClientId $ManagedIdentityClientId
try {
    $emailParams = @{
        Message = @{
            Subject = $Subject
            Body    = @{
                ContentType = 'HTML'
                Content     = $emailBody
            }
            ToRecipients = @(@{EmailAddress=@{Address=$To}})
        }
        SaveToSentItems = $true
    }
    Send-MgUserMail @emailParams
    Write-Host "📧 Email sent successfully to $To"
} catch {
    Write-Host "⚠️ Failed to send email via Microsoft Graph: $($_.Exception.Message)"
}
