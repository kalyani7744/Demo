<#
.SYNOPSIS
  Azure Tenant-Level Service Health Report via Microsoft Graph
.DESCRIPTION
  Fetches tenant-wide Azure Service Health (Service Issues, Planned Maintenance, Health & Security Advisories) using Microsoft Graph
  and sends an HTML report via Microsoft Graph email using a system-assigned managed identity.
#>

# ----------------------------
# CONFIGURATION
# ----------------------------
$OutputHtmlFile = "C:\Temp\TenantServiceHealthReport.html"
$To = "you@example.com"
$Subject = "Azure Tenant Service Health Report - $(Get-Date -Format 'dd-MMM-yyyy')"
$ManagedIdentityClientId = "9464e54c-6ec0-4b15-8380-6172a2e3114b"

# ----------------------------
# CONNECT TO MICROSOFT GRAPH
# ----------------------------
Write-Host "🔹 Connecting to Microsoft Graph with Managed Identity..."
Connect-MgGraph -Identity -ClientId $ManagedIdentityClientId -ErrorAction Stop
Write-Host "✅ Connected to Microsoft Graph."

# ----------------------------
# GET SERVICE HEALTH MESSAGES
# ----------------------------
Write-Host "🔹 Fetching tenant-level service health messages..."
$messages = Get-MgAdminServiceAnnouncementMessage -Status Active,Resolved -ErrorAction Stop

if (-not $messages) {
    Write-Host "✅ No active or recent service health messages found."
    $htmlBody = @"
    <div style='text-align:center; font-family:Segoe UI;'>
        <h2 style='color:#107C10;'>✅ All Systems Healthy</h2>
        <p style='font-size:14px;'>No active or recent Azure Service Health messages found for this tenant.</p>
    </div>
"@
} else {
    Write-Host "📊 Building HTML report for $($messages.Count) messages..."

    # Categorize messages
    $grouped = $messages | Group-Object -Property Service

    $htmlBody = @"
    <h2 style='color:#0078D4; font-family:Segoe UI;'>Azure Tenant Service Health Summary</h2>
    <p style='font-family:Segoe UI; font-size:13px;'>Generated: $(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss')<br/>Total Messages: $($messages.Count)</p>
"@

    foreach ($group in $grouped) {
        $serviceName = $group.Name
        $htmlBody += "<h3 style='color:#333;font-family:Segoe UI;margin-top:20px;'>$serviceName</h3>"
        $htmlBody += "<table style='width:100%; border-collapse:collapse; font-family:Segoe UI; font-size:13px;'>"
        $htmlBody += "<thead><tr style='background-color:#E7F3FD; border-bottom:2px solid #ccc;'>"
        $htmlBody += "<th style='padding:8px; text-align:left;'>Title</th><th style='padding:8px; text-align:left;'>Status</th><th style='padding:8px; text-align:left;'>Start Date</th><th style='padding:8px; text-align:left;'>Last Updated</th><th style='padding:8px; text-align:left;'>Details</th></tr></thead><tbody>"

        foreach ($msg in $group.Group) {
            $statusColor = if ($msg.Status -eq 'Active') { 'color:#E81123;font-weight:bold;' } else { 'color:#107C10;' }
            $htmlBody += "<tr style='border-bottom:1px solid #eee;'>"
            $htmlBody += "<td style='padding:6px;'>$($msg.Title)</td>"
            $htmlBody += "<td style='padding:6px;$statusColor'>$($msg.Status)</td>"
            $htmlBody += "<td style='padding:6px;'>$($msg.StartDateTime.ToLocalTime().ToString('dd-MMM-yyyy HH:mm'))</td>"
            $htmlBody += "<td style='padding:6px;'>$($msg.LastModifiedDateTime.ToLocalTime().ToString('dd-MMM-yyyy HH:mm'))</td>"
            $htmlBody += "<td style='padding:6px;'><a href='$($msg.MicrosoftGraphUrl)' target='_blank'>Details</a></td>"
            $htmlBody += "</tr>"
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
    Write-Host "📧 Email sent successfully to $To via Microsoft Graph"
} catch {
    Write-Host "⚠️ Failed to send email via Microsoft Graph: $($_.Exception.Message)"
}
