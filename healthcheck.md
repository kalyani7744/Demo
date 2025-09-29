# Connect to Azure AD
#Import-Module "Microsoft.Graph.Authentication"
#Import-Module "Microsoft.Graph.Mail"
#Import-Module "Microsoft.Graph.Users.Actions"

$alertClientId = "a2e0604e-752c-47b2-b7a8-57ae8f128555"
$alertThumbprint = "31B8B5BBAFD6C1188D6D32AC7D37197227EC800B"
$tenant = "39130803-0a38-4cc1-a114-117be7255ced"
Connect-MgGraph -CertificateThumbprint $alertThumbprint -ClientId $alertClientId -TenantId $tenant


Disable-AzContextAutosave -Scope Process

# Connect to Azure with user-assigned managed identity
$AzureContext = (Connect-AzAccount -Identity -AccountId "a98eec03-5a89-4bc0-a6eb-a5021321f65a").context

#Select-MGProfile -Name v1.0


$user = "vidula.nandasena@acuitykp.com"
$name = "Vidula Nandasena"
$subject = "Azure Resource Health Summary"
$type = "html"

$recipients = @()
$recipients += @{
    emailAddress = @{
        address = "azurealerts@acuitykp.com"
    }
}






# set and store context
#$AzureContext = Set-AzContext -SubscriptionName $AzureContext.Subscription -DefaultProfile $AzureContext


# Get the subscription ID
$subscriptions = (Get-AzSubscription | ? { $_.Name -like "akp-az*"}).Id


$availabilityInfo = @()

#For each subscription
foreach($subscription in $subscriptions)
{
    #Change the context to the subscription
    Set-AZContext -Subscription $subscription

    if(((Get-AzResourceprovider -ProviderNamespace "Microsoft.ResourceHealth").RegistrationState) -contains "NotRegistered")
    {
        Register-AzResourceProvider -ProviderNamespace "Microsoft.ResourceHealth"
        while(((Get-AzResourceprovider -ProviderNamespace "Microsoft.ResourceHealth").RegistrationState) -contains "Registering")
        {
            Start-Sleep -Seconds 30
        }
    }

    #Get an auth token from this context
    $token = Get-AzAccessToken -ResourceUrl "https://management.azure.com/"

    
    # Define the API URL and parameters
    $apiUrl = "https://management.azure.com/subscriptions/$subscription/providers/Microsoft.ResourceHealth/availabilityStatuses?api-version=2015-01-01"
    $header = @{ 'Authorization' = 'Bearer ' + $token.Token }
    
    $response = (Invoke-RestMethod -Uri $apiUrl -Headers $header -Method GET).value
    foreach($rs in $response)
    {
        #Azure ResourceURI Format
        #/subscriptions/{subscriptionId}/resourceGroups/{resourceGroupName}/providers/{resourceProviderNamespace}/{resourceType}/{resourceName}
        $rsComponents = $rs.id.Split('/') #Split the resource URI
        $resourceGroup =  $rsComponents[4]
        $resourceType = $rsComponents[7]
        $resourceName = $rsComponents[8]

        $availabilityInfo += [PSCustomObject]@{
            "subscription" = (Get-AzSubscription -SubscriptionId $subscription).Name
            "resourceLocation" = $rs.location;
            "resourceGroup" = $resourceGroup;
            "resourceType" = $resourceType;
            "resourceName" = $resourceName;
            "resourceAvailabilityState" = $rs.properties.availabilityState
            "resourceAvailabilityTitle" = $rs.properties.title
            "resourceAvailabilityCause" = $rs.properties.reasonType
            "resourceStatusSummary" = $rs.properties.summary
            "resourceStateReportedTime" = $rs.properties.reportedTime
        }
    }
}


#NotAailale
$notavailable = $availabilityInfo | ? { $_.resourceAvailabilityState -ne "Available"} 
$available = $availabilityInfo | ? { $_.resourceAvailabilityState -eq "Available"} 

$html = @"
<style>
*{
    box-sizing: border-box;
    -webkit-box-sizing: border-box;
    -moz-box-sizing: border-box;
}
body{
    font-family: Helvetica;
    -webkit-font-smoothing: antialiased;
    background: rgba(7, 10, 37, 0.97);
}
h1{
    text-align: justified;
    font-size: 18px;
    text-transform: uppercase;
    letter-spacing: 1px;
    color: #999;
    padding: 10px 0;
}

/* Table Styles */

.table-wrapper{
    margin: 10px 70px 70px;
    box-shadow: 0px 35px 50px rgba( 0, 0, 0, 0.2 );
}

.fl-table {
    border-radius: 5px;
    font-size: 12px;
    font-weight: normal;
    border: none;
    border-collapse: collapse;
    width: 100%;
    max-width: 100%;
    white-space: nowrap;
    background-color: white;
}

.fl-table td, .fl-table th {
    text-align: justified;
    padding: 8px;
}

.fl-table td {
    border-right: 1px solid #f8f8f8;
    font-size: 12px;
}

.fl-table thead th {
    color: #ffffff;
    background: #324960;
}

.fl-table thead th:nth-child(odd) {
    color: #ffffff;
    background: #324960;
}

.fl-table tr:nth-child(even) {
    background: #F8F8F8;
}

/* Responsive */

@media (max-width: 767px) {
    .fl-table {
        display: block;
        width: 100%;
    }
    .table-wrapper:before{
        content: "Scroll horizontally >";
        display: block;
        text-align: right;
        font-size: 11px;
        color: white;
        padding: 0 0 10px;
    }
    .fl-table thead, .fl-table tbody, .fl-table thead th {
        display: block;
    }
    .fl-table thead th:last-child{
        border-bottom: none;
    }
    .fl-table thead {
        float: left;
    }
    .fl-table tbody {
        width: auto;
        position: relative;
        overflow-x: auto;
    }
    .fl-table td, .fl-table th {
        padding: 20px .625em .625em .625em;
        height: 60px;
        vertical-align: middle;
        box-sizing: border-box;
        overflow-x: hidden;
        overflow-y: auto;
        width: 120px;
        font-size: 13px;
        text-overflow: ellipsis;
    }
    .fl-table thead th {
        text-align: left;
        border-bottom: 1px solid #f7f7f9;
    }
    .fl-table tbody tr {
        display: table-cell;
    }
    .fl-table tbody tr:nth-child(odd) {
        background: none;
    }
    .fl-table tr:nth-child(even) {
        background: transparent;
    }
    .fl-table tr td:nth-child(odd) {
        background: #F8F8F8;
        border-right: 1px solid #E6E4E4;
    }
    .fl-table tr td:nth-child(even) {
        border-right: 1px solid #E6E4E4;
    }
    .fl-table tbody td {
        display: block;
        text-align: center;
    }
}
</style>
<div class="table-wrapper">
<h1>Resources with error states</h1>
<table class="fl-table">
<thead>
    <tr>
        <th>Resource Name</th>
        <th>State</th>
        <th>Cause</th>
        <th>Summary</th>
        <th>Type</th>
        <th>Subscription</th>
        <th>Resource Group</th>
        <th>Reported Time</th>
    </tr>
</thead>
<tbody>
"@
foreach($rs in $notavailable)
{
    $appendHTML = @"
    <tr>
        <td>$($rs.resourceName)</td>
        <td>$($rs.resourceAvailabilityState)</td>
        <td>$($rs.resourceAvailabilityCause)</td>
        <td>$($rs.resourceStatusSummary)</td>
        <td>$($rs.resourceType)</td>
        <td>$($rs.subscription)</td>
        <td>$($rs.resourceGroup)</td>
        <td>$($rs.resourceStateReportedTime)</td>
    </tr
"@
$html += $appendHTML
}

$html += @"
</tbody></table></div>
"@

$html += @"
<div class="table-wrapper">
<h1>Healthy Resources</h1>
<table class="fl-table">
<thead>
    <tr>
        <th>Resource Name</th>
        <th>State</th>
        <th>Cause</th>
        <th>Summary</th>
        <th>Type</th>
        <th>Subscription</th>
        <th>Resource Group</th>
        <th>Reported Time</th>
    </tr>
</thead>
<tbody>
"@
foreach($rs in $available)
{
    $appendHTML = @"
    <tr>
        <td>$($rs.resourceName)</td>
        <td>$($rs.resourceAvailabilityState)</td>
        <td>$($rs.resourceAvailabilityCause)</td>
        <td>$($rs.resourceStatusSummary)</td>
        <td>$($rs.resourceType)</td>
        <td>$($rs.subscription)</td>
        <td>$($rs.resourceGroup)</td>
        <td>$($rs.resourceStateReportedTime)</td>
    </tr
"@
    $html += $appendHTML
}
$html += @"
</tbody></table></div>
"@


$message = @{
      "Subject"=$subject;
      "From"=@{
         "EmailAddress"=@{
            "Name"="#IT Notifications";
            "Address"="ITNotifications@acuitykp.com"
         }
      };
      "Body"=@{
         "ContentType"="HTML";
         "Content"= $html
      };
      "ToRecipients"= $recipients
}

Send-MgUserMail -UserId $user -Message $message
