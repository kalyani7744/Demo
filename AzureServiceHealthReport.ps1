<#
.SYNOPSIS
  Azure Tenant-Wide Health, Maintenance & Security Report
.DESCRIPTION
  Collects Azure Monitor Alerts, Planned Maintenance, Health Advisories, and Security Advisories
  across all subscriptions, builds an HTML report, and emails it via Microsoft Graph.
#>

# ----------------------------
# CONFIGURATION
# ----------------------------
$OutputHtmlFile = "C:\Temp\AzureTenantHealthSecurityReport.html"
$To = "you@example.com"
$Subject = "Azure Tenant Health & Security Report - $(Get-Date -Format 'dd-MMM-yyyy')"
$ManagedIdentityClientId = "9464e54c-6ec0-4b15-8380-6172a2e3114b"

# ----------------------------
# CONNECT TO AZURE & MICROSOFT GRAPH
# ----------------------------
Write-Host "🔹 Connecting to Azure and Microsoft Graph..."
Connect-AzAccount -Identity -ErrorAction Stop
Connect-MgGraph -Identity -ClientId $ManagedIdentityClientId -ErrorAction Stop
Write-Host "✅ Connected to Azure and Graph."

# ----------------------------
# PREPARE VARIABLES
# ----------------------------
$allAlerts = @()
$allMaintenance = @()
$allHealthAdvisories = @()
$allSecurityAdvisories = @()

# ----------------------------
# FETCH DATA ACROSS SUBSCRIPTIONS
# ----------------------------
$subscriptions = Get-AzSubscription
foreach ($sub in $subscriptions) {
    Write-Host "🔸 Processing subscription: $($sub.Name)"
    Set-AzContext -SubscriptionId $sub.Id | Out-Null

    # ----------------------------
    # 1️⃣ Azure Monitor Alerts
    # ----------------------------
    try {
        $alerts = Get-AzMetricAlertRuleV2 -ErrorAction SilentlyContinue
        foreach ($a in $alerts) {
            $allAlerts += [PSCustomObject]@{
                SubscriptionName = $sub.Name
                AlertName        = $a.Name
                Severity         = $a.Severity
                State            = $a.Enabled
                Resource         = $a.TargetResourceId.Split('/')[-1]
                LastUpdated      = (Get-Date).ToLocalTime()
            }
        }
    } catch {
        Write-Host "⚠️ Alerts fetch failed for $($sub.Name): $($_.Exception.Message)"
    }

    # ----------------------------
    # 2️⃣ Planned Maintenance
    # ----------------------------
    try {
        $maintenanceUpdates = Get-AzMaintenanceUpdate -ErrorAction SilentlyContinue
        foreach ($m in $maintenanceUpdates) {
            $allMaintenance += [PSCustomObject]@{
                SubscriptionName = $sub.Name
                ResourceGroup    = $m.ResourceGroupName
                ResourceName     = $m.ResourceName
                Event            = $m.Status
                ScheduledTime    = $m.StatusDateTime.ToLocalTime()
            }
        }
    } catch {
        Write-Host "⚠️ Planned Maintenance fetch failed for $($sub.Name): $($_.Exception.Message)"
    }

    # ----------------------------
    # 3️⃣ Health Advisories
    # ----------------------------
    try {
        $resources = Get-AzResource
        foreach ($r in $resources) {
            try {
                $health = Get-AzResourceHealth -ResourceId $r.ResourceId -ErrorAction Stop
                if ($health.AvailabilityState -ne "Available") {
                    $allHealthAdvisories += [PSCustomObject]@{
                        SubscriptionName = $sub.Name
                        ResourceGroup    = $r.ResourceGroupName
                        ResourceName     = $r.Name
                        Type             = $r.ResourceType
                        Status           = $health.AvailabilityState
                        Details          = $health.ReasonType
                        LastUpdated      = $health.Timestamp.ToLocalTime()
                    }
                }
            } catch {}
        }
    } catch {
        Write-Host "⚠️ Health Advisories fetch failed for $($sub.Name): $($_.Exception.Message)"
    }

    # ----------------------------
    # 4️⃣ Security Advisories
    # ----------------------------
    try {
        $securityAlerts = Get-AzSecurityAlert -ErrorAction SilentlyContinue
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
    } catch {
        Write-Host "⚠️ Security Advisories fetch failed for $($sub.Name): $($_.Exception.Message)"
    }
}

# ----------------------------
# BUILD HTML REPORT
# ----------------------------
function New-SectionHtml ($title, $icon, $color, $data) {
    if (-not $data -or $data.Count -eq 0) {
        return "<h3 style='color:$color;'>$icon $title</h3><p style='font-size:13px;color:gray;'>✅ None detected</p>"
    }

    $table = "<h3 style='color:$color;'>$icon $title</h3><table style='width:100%;border-collapse:collapse;font-family:Segoe UI;font-size:13px;'>"
    $table += "<thead><tr style='background-color:#f3f3f3;border-bottom:2px solid #ccc;'>"
    $data[0].psobject.Properties.Name | ForEach-Object {
        $table += "<th style='padding:6px;text-align:left;'>$_</th>"
    }
    $table += "</tr></thead><tbody>"
    foreach ($row in $data) {
        $table += "<tr style='border-bottom:1px solid #eee;'>"
        foreach ($val in $row.psobject.Properties.Value) {
            $table += "<td style='padding:6px;'>$val</td>"
        }
        $table += "</tr>"
    }
    $table += "</tbody></table><br/>"
    return $table
}

$htmlBody = @()
$htmlBody += New-SectionHtml "Active Alerts" "📢" "#E81123" $allAlerts
$htmlBody += New-SectionHtml "Planned Maintenance" "🛠" "#F7630C" $allMaintenance
$htmlBody += New-SectionHtml "Health Advisories" "⚠" "#FFB900" $allHealthAdvisories
$htmlBody += New-SectionHtml "Security Advisories" "🔒" "#6B4EFF" $allSecurityAdvisories

$emailBody = @"
<html>
<head>
<style>
body { font-family:'Segoe UI', Arial, sans-serif; margin: 20px; color:#222; }
a { color:#0078D4; text-decoration:none; }
a:hover { text-decoration:underline; }
table tr:hover { background-color:#f9f9f9; }
</style>
</head>
<body>
<h2 style='color:#0078D4;'>Azure Tenant Health & Security Report</h2>
<p style='font-size:13px;'>Generated: $(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss')</p>
$htmlBody
<p style='font-size:11px;color:#888;'>Generated automatically by Azure Automation • $(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss')</p>
</body>
</html>
"@

$emailBody | Out-File -FilePath $OutputHtmlFile -Encoding UTF8
Write-Host "✅ HTML report saved to: $OutputHtmlFile"

# ----------------------------
# SEND EMAIL VIA GRAPH
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
