<#
.SYNOPSIS
  Tenant-wide Azure Service Health & Security Report (HTML email) using Managed Identity
.DESCRIPTION
  - Connects using system-assigned managed identity (ClientId provided)
  - Fetches tenant-level Service Issues, Planned Maintenance, Health Advisories, Security Advisories, Billing Updates from Microsoft Graph (beta)
  - Resolves Scope for each event using Azure Resource Graph (batched)
  - Generates a polished HTML report and sends via Microsoft Graph (email body)
USAGE
  - Run in Cloud Shell or Azure Automation Runbook with system-assigned managed identity enabled.
  - Ensure the managed identity has:
      * Microsoft Graph: ServiceAnnouncement.Read.All (app permission) or appropriate delegated permission
      * Azure Resource Graph: Reader across subscriptions (or Resource Graph Reader)
      * Microsoft Graph Mail.Send (to send mail) or permission to send mail as the user/account used
  - Edit $To and $Subject as needed.
#>

param()

# ----------------------------
# Configuration
# ----------------------------
$ClientId = "9464e54c-6ec0-4b15-8380-6172a2e3114b"    # system-assigned or user-assigned MI client id (provided)
$To = "you@example.com"                              # recipient
$Subject = "Azure Tenant Service Health Report - $(Get-Date -Format 'dd-MMM-yyyy')"

# Output local path (kept optional - script only sends email)
$OutputHtmlFile = "$env:TEMP\AzureTenantServiceHealthReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"

# ----------------------------
# Helper: Ensure Modules
# ----------------------------
function Ensure-Module {
    param(
        [string]$Name
    )
    if (-not (Get-Module -ListAvailable -Name $Name)) {
        try {
            Write-Host "Installing module $Name ..."
            Install-Module -Name $Name -Force -Scope CurrentUser -AllowClobber -ErrorAction Stop
        } catch {
            Write-Warning "Failed to install module $Name: $($_.Exception.Message)"
        }
    }
}

$required = @('Az.Accounts','Az.Resources','Az.ResourceGraph','Microsoft.Graph','Az.Profile')
foreach ($m in $required) { Ensure-Module -Name $m }

# ----------------------------
# CONNECT: Azure & Microsoft Graph (Managed Identity)
# ----------------------------
Write-Host "Connecting to Azure using Managed Identity..."
try {
    Connect-AzAccount -Identity -ErrorAction Stop
    Write-Host "Connected to Azure."
} catch {
    Write-Error "Failed to Connect-AzAccount -Identity: $($_.Exception.Message)"
    throw
}

Write-Host "Connecting to Microsoft Graph (Managed Identity)..."
try {
    # Install/Import Microsoft.Graph module usage
    Import-Module Microsoft.Graph -ErrorAction SilentlyContinue
    Connect-MgGraph -Identity -ClientId $ClientId -ErrorAction Stop
    # Use beta because serviceAnnouncement endpoints are only in beta at times
    Select-MgProfile -Name "beta"
    Write-Host "Connected to Microsoft Graph (beta)."
} catch {
    Write-Error "Failed to connect to Microsoft Graph with managed identity: $($_.Exception.Message)"
    throw
}

# ----------------------------
# FETCH: Service Issues + Messages from Graph
# ----------------------------
Write-Host "Fetching Service Issues and Messages from Microsoft Graph..."
try {
    $issuesResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/admin/serviceAnnouncement/issues" -ErrorAction Stop
    $issues = $issuesResponse.value
} catch {
    Write-Warning "Failed to fetch issues: $($_.Exception.Message)"
    $issues = @()
}

try {
    $messagesResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/beta/admin/serviceAnnouncement/messages" -ErrorAction Stop
    $messages = $messagesResponse.value
} catch {
    Write-Warning "Failed to fetch messages: $($_.Exception.Message)"
    $messages = @()
}

Write-Host "Fetched $($issues.Count) issues and $($messages.Count) messages."

# ----------------------------
# Normalise and Combine Events
# ----------------------------
$allEvents = @()

# Normalize issues
foreach ($i in $issues) {
    $allEvents += [PSCustomObject]@{
        Category     = "Service Issue"
        Title        = $i.Title
        TrackingId   = $i.Id
        EventLevel   = $i.Classification
        Services     = ($i.ImpactedService -join ', ')
        Regions      = ($i.ImpactedRegions -join ', ')
        StartTime    = if ($i.StartDateTime) {[datetime]$i.StartDateTime} else {$null}
        LastUpdated  = if ($i.LastModifiedDateTime) {[datetime]$i.LastModifiedDateTime} else {$null}
        EventTags    = ($i.Tags -join ', ')
        DetailsUrl   = $i.DetailsUrl
    }
}

# Normalize messages (planned maintenance, advisories)
foreach ($m in $messages) {
    # map category values to human names (best-effort)
    $cat = switch -Regex ($m.Category) {
        'plannedMaintenance' { 'Planned Maintenance' }
        'healthAdvisory'     { 'Health Advisory' }
        'securityAdvisory'   { 'Security Advisory' }
        'billingUpdate'      { 'Billing Update' }
        default              { $m.Category }
    }

    $allEvents += [PSCustomObject]@{
        Category     = $cat
        Title        = $m.Title
        TrackingId   = $m.Id
        EventLevel   = $m.Severity
        Services     = ($m.Services -join ', ')
        Regions      = ($m.ImpactedRegions -join ', ')
        StartTime    = if ($m.StartDateTime) {[datetime]$m.StartDateTime} else {$null}
        LastUpdated  = if ($m.LastModifiedDateTime) {[datetime]$m.LastModifiedDateTime} else {$null}
        EventTags    = ($m.Tags -join ', ')
        DetailsUrl   = $m.DetailsUrl
    }
}

if ($allEvents.Count -eq 0) {
    Write-Host "No service health events found. Sending an 'All Healthy' email."
}

# ----------------------------
# Resolve Scope per Event using Azure Resource Graph (batched)
# ----------------------------
# We will collect tracking IDs and query Resource Graph once.
function Resolve-Scopes-Batched {
    param(
        [array]$EventList     # list of PSCustomObject with TrackingId property
    )
    $resultMap = @{}

    $trackingIds = $EventList | ForEach-Object { $_.TrackingId } | Where-Object { $_ } | Select-Object -Unique
    if (-not $trackingIds -or $trackingIds.Count -eq 0) {
        return $resultMap
    }

    # ARG Kusto: resourcehealth events sometimes stored under resourceproviders or under resourcehealth namespace
    # We'll search for resources where properties.trackingId matches any of our tracking IDs.
    $idsCsv = ($trackingIds | ForEach-Object { "'$_'" }) -join ","
    $kql = @"
Resources
| where isnotempty(properties) 
| where tostring(properties.trackingId) in ($idsCsv) or tostring(properties.trackingId) in ($idsCsv)
| project id, type, subscriptionId, name, properties
"@

    try {
        $rgHits = Search-AzGraph -Query $kql -First 10000 -ErrorAction Stop
    } catch {
        Write-Warning "Resource Graph query failed: $($_.Exception.Message)"
        return $resultMap
    }

    # Map trackingId -> list of resourceIds and subscriptionIds
    foreach ($hit in $rgHits) {
        $props = $hit.properties
        if ($props -and $props.trackingId) {
            $tid = [string]$props.trackingId
            if (-not $resultMap.ContainsKey($tid)) { $resultMap[$tid] = [ordered]@{ ResourceIds = @(); SubscriptionIds = @() } }
            $resultMap[$tid].ResourceIds += $hit.id
            if ($hit.subscriptionId) { $resultMap[$tid].SubscriptionIds += $hit.subscriptionId }
            # Also check impactedResources inside properties (if present)
            if ($props.impactedResources) {
                foreach ($ir in $props.impactedResources) {
                    if ($ir.resourceId) { $resultMap[$tid].ResourceIds += $ir.resourceId }
                    if ($ir.subscriptionId) { $resultMap[$tid].SubscriptionIds += $ir.subscriptionId }
                }
            }
        }
    }

    # Deduplicate lists
    foreach ($k in $resultMap.Keys) {
        $resultMap[$k].ResourceIds = ($resultMap[$k].ResourceIds | Select-Object -Unique)
        $resultMap[$k].SubscriptionIds = ($resultMap[$k].SubscriptionIds | Select-Object -Unique)
    }

    return $resultMap
}

Write-Host "Resolving event scopes using Azure Resource Graph..."
$map = Resolve-Scopes-Batched -EventList $allEvents

# Attach Scope to each event object
foreach ($evt in $allEvents) {
    $scopeValue = "Tenant"
    if ($evt.TrackingId -and $map.ContainsKey($evt.TrackingId)) {
        $entry = $map[$evt.TrackingId]
        if ($entry.ResourceIds.Count -gt 0) {
            # present up to 3 resource ids
            $sample = $entry.ResourceIds | Select-Object -First 3
            $more = ""
            if ($entry.ResourceIds.Count -gt 3) { $more = " (+$($entry.ResourceIds.Count - 3) more)" }
            $scopeValue = "Resources: $($sample -join '; ')$more"
        } elseif ($entry.SubscriptionIds.Count -gt 0) {
            if ($entry.SubscriptionIds.Count -eq 1) {
                $subName = $entry.SubscriptionIds[0]
                $scopeValue = "Subscription: $subName"
            } else {
                $scopeValue = "Subscriptions: $($entry.SubscriptionIds.Count)"
            }
        }
    }
    $evt | Add-Member -NotePropertyName Scope -NotePropertyValue $scopeValue
}

# ----------------------------
# Prepare Summary Counts
# ----------------------------
$summary = [ordered]@{
    TotalEvents = $allEvents.Count
    ServiceIssues = ($allEvents | Where-Object { $_.Category -eq 'Service Issue' }).Count
    PlannedMaintenance = ($allEvents | Where-Object { $_.Category -eq 'Planned Maintenance' }).Count
    HealthAdvisories = ($allEvents | Where-Object { $_.Category -eq 'Health Advisory' }).Count
    SecurityAdvisories = ($allEvents | Where-Object { $_.Category -eq 'Security Advisory' }).Count
    BillingUpdates = ($allEvents | Where-Object { $_.Category -eq 'Billing Update' }).Count
}

# ----------------------------
# Build HTML Report
# ----------------------------
function Build-HTML {
    param([array]$Events, [hashtable]$Summary)

    $style = @"
<style>
body { font-family: 'Segoe UI', Arial, sans-serif; background:#f4f6f8; color:#222; padding:16px; }
.header { display:flex; justify-content:space-between; align-items:center; }
.h-title { color:#0078D4; font-size:20px; margin:0; }
.summary { margin-top:8px; font-size:14px; }
.card { display:inline-block; padding:8px 12px; border-radius:6px; margin-right:8px; color:white; font-weight:600; }
.card.red { background:#e81123; } .card.orange { background:#f7630c; } .card.yellow { background:#ffb900; } .card.purple { background:#6b4eff; } .card.blue { background:#0078d4; }
table { width:100%; border-collapse:collapse; margin-top:12px; }
th, td { padding:8px; border:1px solid #e1e4e8; font-size:13px; text-align:left; vertical-align:top; }
th { background:#f0f4f8; color:#222; font-weight:700; }
tr:nth-child(even) td { background:#ffffff; }
a { color:#0078D4; text-decoration:none; }
.small { font-size:12px; color:#666; }
.footer { margin-top:12px; font-size:12px; color:#666; }
</style>
"@

    $header = "<div class='header'><div><h1 class='h-title'>Azure Tenant Service Health Report</h1><div class='small'>Generated: $(Get-Date -Format 'dd-MMM-yyyy HH:mm:ss')</div></div></div>"

    $cards = "<div class='summary'>"
    $cards += "<span class='card red'>Service Issues: $($Summary.ServiceIssues)</span>"
    $cards += "<span class='card orange'>Planned Maintenance: $($Summary.PlannedMaintenance)</span>"
    $cards += "<span class='card yellow'>Health Advisories: $($Summary.HealthAdvisories)</span>"
    $cards += "<span class='card purple'>Security Advisories: $($Summary.SecurityAdvisories)</span>"
    $cards += "<span class='card blue'>Billing Updates: $($Summary.BillingUpdates)</span>"
    $cards += "</div>"

    $bodyHtml = $header + $cards

    # Group by category and render tables
    $groups = $Events | Group-Object -Property Category
    foreach ($g in $groups) {
        $bodyHtml += "<h2 style='margin-top:16px;'>$($g.Name) ($($g.Count))</h2>"
        $bodyHtml += "<table><thead><tr>
                        <th style='width:28%'>Issue name</th>
                        <th>Event level</th>
                        <th>Tracking ID</th>
                        <th>Services</th>
                        <th>Regions</th>
                        <th>Scope</th>
                        <th>Start time</th>
                        <th>Last updated</th>
                        <th>Event tags</th>
                        <th>Details</th>
                      </tr></thead><tbody>"
        foreach ($r in $g.Group) {
            $start = if ($r.StartTime) { $r.StartTime.ToString('dd-MMM-yyyy HH:mm') } else { "" }
            $last = if ($r.LastUpdated) { $r.LastUpdated.ToString('dd-MMM-yyyy HH:mm') } else { "" }
            $level = [string]$r.EventLevel
            $levelCell = if ($level -match 'Critical|Active|Investigating|Error') { "<span style='color:#e81123;font-weight:600;'>$level</span>" } else { "<span style='color:#107C10;'>$level</span>" }
            $detailsLink = if ($r.DetailsUrl) { "<a href='$($r.DetailsUrl)' target='_blank'>View</a>" } else { "" }

            $bodyHtml += "<tr>
                            <td>$([System.Web.HttpUtility]::HtmlEncode($r.Title))</td>
                            <td>$levelCell</td>
                            <td>$([System.Web.HttpUtility]::HtmlEncode($r.TrackingId))</td>
                            <td>$([System.Web.HttpUtility]::HtmlEncode($r.Services))</td>
                            <td>$([System.Web.HttpUtility]::HtmlEncode($r.Regions))</td>
                            <td>$([System.Web.HttpUtility]::HtmlEncode($r.Scope))</td>
                            <td>$start</td>
                            <td>$last</td>
                            <td>$([System.Web.HttpUtility]::HtmlEncode($r.EventTags))</td>
                            <td>$detailsLink</td>
                          </tr>"
        }
        $bodyHtml += "</tbody></table>"
    }

    $bodyHtml += "<div class='footer'>Report generated automatically using Managed Identity + Microsoft Graph & Azure Resource Graph.</div>"

    return "<html><head>$style</head><body>$bodyHtml</body></html>"
}

$html = Build-HTML -Events $allEvents -Summary $summary

# Save locally (optional)
try {
    $html | Out-File -FilePath $OutputHtmlFile -Encoding UTF8
    Write-Host "Local copy saved to $OutputHtmlFile"
} catch {
    Write-Warning "Failed to save local HTML file: $($_.Exception.Message)"
}

# ----------------------------
# SEND EMAIL via Microsoft Graph
# ----------------------------
Write-Host "Sending report email via Microsoft Graph..."
try {
    # Using Send-MgUserMail -UserId 'me' (the identity used to connect). In some managed identity scenarios,
    # sending mail 'as' a user may not be permitted. If Send-MgUserMail fails, consider Send-MgGraphRequest POST to /users/{userId}/sendMail
    $message = @{
        message = @{
            subject = $Subject
            body = @{
                contentType = "html"
                content = $html
            }
            toRecipients = @(@{emailAddress = @{address = $To}})
        }
        saveToSentItems = $false
    }

    # Try Send-MgUserMail (preferred)
    try {
        Send-MgUserMail -UserId "me" -BodyParameter $message -ErrorAction Stop
        Write-Host "Email sent via Send-MgUserMail."
    } catch {
        Write-Warning "Send-MgUserMail failed: $($_.Exception.Message) - attempting REST POST /me/sendMail"

        # Fallback: use REST API
        $uri = "https://graph.microsoft.com/beta/me/sendMail"
        Invoke-MgGraphRequest -Method POST -Uri $uri -Body ($message | ConvertTo-Json -Depth 10) -ErrorAction Stop
        Write-Host "Email sent via Invoke-MgGraphRequest (me/sendMail)."
    }
} catch {
    Write-Error "Failed to send email: $($_.Exception.Message)"
}

Write-Host "Done."
