# Suppress errors globally
$ErrorActionPreference = "SilentlyContinue"

# Connect to Microsoft Graph with System-Assigned Managed Identity
Connect-MgGraph -Identity -ClientId "9464e54c-6ec0-4b15-8380-6172a2e3114b" -ErrorAction SilentlyContinue

# Define the location filter
$locationFilter = "Acuity | Colombo WTC"

# Retrieve all users with selected properties
$allUsers = Get-MgUser -All -Select "DisplayName,MailNickname,Mail,EmployeeId,JobTitle,Department,MobilePhone,BusinessPhones,OfficeLocation" -ExpandProperty "Manager" -ErrorAction SilentlyContinue

# Filter users based on location and job title
$filteredUsers = $allUsers | Where-Object {
    $_.OfficeLocation -eq $locationFilter -and $_.JobTitle -ne $null
}

# Prepare user data for export
$userInfo = $filteredUsers | ForEach-Object {
    # Attempt to retrieve manager's display name
    $managerName = $null
    try {
        $managerResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/users/$($_.Id)/manager" -ErrorAction Stop
        if ($managerResponse -ne $null) {
            $managerName = $managerResponse.DisplayName
        }
    } catch {
        Write-Output "Failed to retrieve manager for user: $($_.DisplayName)"
    }

    # Create object with required properties
    [PSCustomObject]@{
        Name             = $_.DisplayName
        SamAccountName   = $_.MailNickname
        Email            = $_.Mail
        EmployeeID       = $_.EmployeeId
        Title            = $_.JobTitle
        Department       = $_.Department
        Mobile           = $_.MobilePhone
        TelephoneNumber  = if ($_.BusinessPhones) { $_.BusinessPhones -join ", " } else { $null }
        ManagerName      = $managerName
    }
}

# Export sorted user information to XLSX if data is available
if ($userInfo.Count -gt 0) {
    $xlsxFilePath = "$env:TEMP\CMB_EMPDBlatest.xlsx"
    $userInfo | Sort-Object -Property Name | Export-Excel -Path $xlsxFilePath -WorksheetName "UserData" -AutoSize
    Write-Output "XLSX file created at: $xlsxFilePath"
} else {
    Write-Output "No user data available for export."
}

# Get SharePoint Site ID for upload
$siteName = "IT-CMB"
try {
    $siteNameResponse = Invoke-MgGraphRequest -Method GET -Uri "https://graph.microsoft.com/v1.0/sites/acuitykp.sharepoint.com:/sites/$siteName"
    $siteID = $siteNameResponse.id
    Write-Output "SharePoint Site ID retrieved: $siteID"

    # Upload XLSX file to SharePoint
    $fileContent = [System.IO.File]::ReadAllBytes($xlsxFilePath)
    $uploadUri = "https://graph.microsoft.com/v1.0/sites/$siteID/drive/root:/CMB - Database/CMB_EMPDBlatest.xlsx:/content"
    Invoke-MgGraphRequest -Method PUT -Uri $uploadUri -Body $fileContent
    Write-Output "File uploaded to SharePoint 'CMB-Database' library successfully."
} catch {
    Write-Output "Failed to retrieve site ID or upload the file to SharePoint."
}
