$tenantId = "tenant id"
$clientId = "client id"
$clientSecret = "secret value"

#Prepare the body for the token request
$body = @{
    client_id     = $clientId
    scope         = "https://graph.microsoft.com/.default"
    client_secret = $clientSecret
    grant_type    = "client_credentials"
}

#Retrieve the OAuth 2.0 token
$response = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$tenantId/oauth2/v2.0/token" -Method Post -ContentType "application/x-www-form-urlencoded" -Body $body
$accessToken = $response.access_token

#Set the auth header
$headers = @{
    Authorization = "Bearer $accessToken"
}

#Get all SharePoint sites
$sitesResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites?search=*" -Headers $headers -Method Get
$sites = $sitesResponse.value

#Set path
$excelFilePath = "SharePointFoldersReport.xlsx"

#Iterate over each site
foreach ($site in $sites) {
    Write-Host "Processing site: $($site.webUrl)"

    # Initialize collection as an expandable ArrayList to store folder info
    $folderInfoCollection = New-Object System.Collections.ArrayList

    # Access the default drive (library) of the site
    $driveResponse = Invoke-RestMethod -Uri "https://graph.microsoft.com/v1.0/sites/$($site.id)/drive" -Headers $headers -Method Get
    $driveId = $driveResponse.id

    # Retrieve only the top-level entries (children of root)
    $childrenUri = "https://graph.microsoft.com/v1.0/drives/$driveId/root/children"
    $foldersResponse = Invoke-RestMethod -Uri $childrenUri -Headers $headers -Method Get
    $folders = $foldersResponse.value

    foreach ($folder in $folders) {
        if ($folder.folder) {
            # Convert size from bytes to gigabytes, rounded to two decimal places
            $sizeInGB = [Math]::Round($folder.size / 1GB, 2)

            # Add folder info to collection
            $folderInfoCollection.Add([PSCustomObject]@{
                Name = $folder.name
                Id   = $folder.id
                SizeGB = $sizeInGB
            }) | Out-Null
        }
    }

    #Use only the site display name for the worksheet name, and truncate if needed
    $siteName = ($site.displayName -replace '\W', '_') # Replace illegal characters
    $sheetName = $siteName.Substring(0, [Math]::Min(31, $siteName.Length)) # Limit to 31 characters

    #Export results
    if ($folderInfoCollection.Count -gt 0) {
        try {
            $folderInfoCollection | Export-Excel -Path $excelFilePath -WorksheetName $sheetName -AutoSize
        } catch {
            Write-Host "An error occurred while exporting data to Excel for site: $siteName. Error: $_"
        }
    } else {
        Write-Host "No folder information found for site: $($site.webUrl)"
    }
}

Write-Host "Export complete: $excelFilePath."
