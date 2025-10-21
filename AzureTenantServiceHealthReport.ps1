<#
.SYNOPSIS
  Azure Tenant-Wide Health, Maintenance & Security Report via Microsoft Graph
.DESCRIPTION
  Collects Service Health (Service Issues, Planned Maintenance, Health & Security Advisories)
  across all subscriptions using Microsoft Graph, builds an HTML report, and emails it.
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
Select-MgProfile -Name "beta"
Write-Host "✅ Connected."

# ----------------------------
# FETCH SERVICE HEALTH DATA
# ----------------------------
Write-Host "🔹 Fetching Service Health messages..."
$serviceIssues = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/admin/serviceAnnouncement/issues" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty value
$maintenance = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/admin/serviceAnnouncement/messages" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty value

# Separate by category
$plannedMaintenance = $maintenance | Where-Object { $_.Category -eq 'plannedMaintenance' }
$healthAdvisories   = $maintenance | Where-Object { $_.Category -eq 'healthAdvisory' }
$securityAdvisories = $maintenance | Where-Object { $_.Category -eq 'securityAdvisory' }

# ----------------------------
# BUILD HTML REPORT
# ----------------------------
function Build-SectionHtml {
    param($Title, $Items, $Color)

    if (-not $Items -or $Items.Count -eq 0) {
        return "<h3 style='color:$Color;'>$Title</h3><p style='font-size:13px;color:gray;'>✅ None detected</p>"
    }

    $html = "<h3 style='color:$Color;'>$Title</h3><table style='width:100%;border-collapse:collapse;font-family:Segoe UI;font-size:13px;'>"
    $html += "<thead><tr style='background-color:#f3f3f3;border-bottom:2px solid #ccc;'>"
    $Items[0].psobject.Properties.Name | ForEach-Object { $html += "<th style='padding:6px;text-align:left;'>$_</th>" }
    $html += "</tr></thead><tbody>"

    foreach ($row in $Items) {
        $html += "<tr style='border-bottom:1px solid #eee;'>"
        foreach ($val in $row.psobject.Properties.Value) {
            $html += "<td style='padding:6px;'>$val</td>"
        }
        $html += "</tr>"
    }

    $html += "</tbody></table><br/>"
    return $html
}

# Convert Graph objects to simple PSCustomObjects for HTML table
$alertsHtml = @()
foreach ($i in $serviceIssues) {
    $alertsHtml += [PSCustomObject]@{
        Title       = $i.Title
        Service     = ($i.ImpactedService -join ", ")
        Region      = ($i.ImpactedRegions -join ", ")
        Impact      = $i.Classification
        Status      = $i.Status
        StartTime   = if ($i.StartDateTime) { [datetime]$i.StartDateTime } else { "" }
        EndTime     = if ($i.EndDateTime) { [datetime]$i.EndDateTime } else { "" }
        Scope       = if ($i.ImpactedResources) { ($i.ImpactedResources -join ", ") } else { "Tenant" }
        LastUpdated = if ($i.LastModifiedDateTime) { [datetime]$i.LastModifiedDateTime } else { "" }
    }
}

$plannedHtml = @()
foreach ($p in $plannedMaintenance) {
    $plannedHtml += [PSCustomObject]@{
        Title        = $p.Title
        Service      = ($p.Services -join ", ")
        Region       = ($p.ImpactedRegions -join ", ")
        Status       = $p.Status
        ScheduledTime = if ($p.StartDateTime) { [datetime]$p.StartDateTime } else { "" }
        LastUpdated  = if ($p.LastModifiedDateTime) { [datetime]$p.LastModifiedDateTime } else { "" }
    }
}

$healthHtml = @()
foreach ($h in $healthAdvisories) {
    $healthHtml += [PSCustomObject]@{
        Title       = $h.Title
        Service     = ($h.Services -join ", ")
        Region      = ($h.ImpactedRegions -join ", ")
        Status      = $h.Status
        Published   = if ($h.StartDateTime) { [datetime]$h.StartDateTime } else { "" }
        LastUpdated = if ($h.LastModifiedDateTime) { [datetime]$h.LastModifiedDateTime } else { "" }
    }
}

$securityHtml = @()
foreach ($s in $securityAdvisories) {
    $securityHtml += [PSCustomObject]@{
        Title       = $s.Title
        Service     = ($s.Services -join ", ")
        Region      = ($s.ImpactedRegions -join ", ")
        Status      = $s.Status
        Published   = if ($s.StartDateTime) { [datetime]$s.StartDateTime } else { "" }
        LastUpdated = if ($s.LastModifiedDateTime) { [datetime]$s.LastModifiedDateTime } else { "" }
    }
}

# Assemble final HTML
$htmlBody = @()
$htmlBody += Build-SectionHtml "📢 Service Health" $alertsHtml "#E81123"
$htmlBody += Build-SectionHtml "🛠 Planned Maintenance" $plannedHtml "#F7630C"
$htmlBody += Build-SectionHtml "⚠ Health Advisories" $healthHtml "#FFB900"
$htmlBody += Build-SectionHtml "🔒 Security Advisories" $securityHtml "#6B4EFF"

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

# ----------------------------
# SAVE HTML REPORT LOCALLY
# ----------------------------
$emailBody | Out-File -FilePath $OutputHtmlFile -Encoding UTF8
Write-Host "✅ HTML report saved to: $OutputHtmlFile"

# ----------------------------
# SEND EMAIL VIA MICROSOFT GRAPH
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
