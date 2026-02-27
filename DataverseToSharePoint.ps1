param (
    [Parameter(Mandatory = $true)]
    [string]$DataverseRowId = "3df13404-7a53-f011-877a-002248222778",

    [Parameter(Mandatory = $true)]
    [string]$FileName = "FileName.pdf",

    [Parameter(Mandatory = $false)]
    [string]$DataverseTableName = "annotations",

    [Parameter(Mandatory = $false)]
    [string]$DataverseFileColumnName = "documentbody",

    [Parameter(Mandatory = $false)]
    [string]$DataverseUrl = "https://orgb9d46ed6.crm.dynamics.com/",

    [Parameter(Mandatory = $false)]
    [string]$SharePointSiteUrl = "https://jakegwynndemo.sharepoint.com/sites/AllowList-Automation",

    [Parameter(Mandatory = $false)]
    [string]$SharePointDocLibName = "DocLibTest",

    [Parameter(Mandatory = $false)]
    [string]$TenantId = "04b9e073-f7cf-4c95-9f91-e6d55d5a3797",

    [Parameter(Mandatory = $false)]
    [string]$CertificateThumbprint = "CCBE000B80BE1112C553F635C20DC3C8D9412528"
)

# Provide details about parameters that were passed to the script
Write-Output "Dataverse Table Name:               $DataverseTableName"
Write-Output "Dataverse File Column Name:         $DataverseFileColumnName"
Write-Output "Dataverse Row ID:                   $DataverseRowId"
Write-Output "Dataverse URL:                      $DataverseUrl"
Write-Output "SharePoint Site URL:                $SharePointSiteUrl"
Write-Output "SharePoint Document Library Name:   $SharePointDocLibName"
Write-Output "File Name:                          $FileName"
Write-Output "Tenant ID:                          $TenantId"

$AppRegistrationCredentials = Get-AutomationPSCredential -Name "AppRegistration-Dev"
if (-not $AppRegistrationCredentials) {
    $result = @{
        success = $false
        error = "Failed to retrieve credentials. Ensure the credential 'AppRegistration' exists in Azure Automation."
    }
    Write-Output ($result | ConvertTo-Json -Compress)
    exit
}

# Set ClientId and ClientSecret variables for Dataverse Authentication
$ClientId = $AppRegistrationCredentials.UserName
$ClientSecret = $AppRegistrationCredentials.GetNetworkCredential().Password

# Get OAuth 2.0 Token for Dataverse Authentication
$TokenUrl = "https://login.microsoftonline.com/$TenantId/oauth2/v2.0/token"
$Body = @{
    grant_type    = "client_credentials"
    client_id     = $ClientId
    client_secret = $ClientSecret
    scope         = "$DataverseUrl/.default"
}

try {
    Write-Output "Retrieving access token for Dataverse..." 
    $TokenResponse = Invoke-RestMethod -Method Post -Uri $TokenUrl -Body $Body -ContentType "application/x-www-form-urlencoded"
    $AccessToken = $TokenResponse.access_token
} catch {
    $result = @{
        success = $false
        error = "Failed to retrieve access token for Dataverse: $($_.Exception.Message)"
    }
    Write-Output ($result | ConvertTo-Json -Compress)
    exit
}

# Connect to SharePoint using PnP PowerShell
try {
    Write-Output "Connecting to SharePoint using Connect-PnPOnline ..." 
    Connect-PnPOnline -Url $SharePointSiteUrl -ClientId $ClientId -Thumbprint $CertificateThumbprint -Tenant $TenantId  -ErrorAction Stop
} catch {
    $result = @{
        success = $false
        error = "Failed to connect to SharePoint Online: $($_.Exception.Message)"
    }
    Write-Output ($result | ConvertTo-Json -Compress)
    exit
}

$Headers = @{
    Authorization = "Bearer $AccessToken"
}

$TempFilePath = Join-Path $env:TEMP $FileName

# Safely trim any trailing slash from the base URL to prevent "//api" malformed URLs
$CleanDataverseUrl = $DataverseUrl.TrimEnd('/')

try {
    Write-Output "Initializing Dataverse file download..."

    if ($DataverseTableName -eq "annotations") {
        # Annotations require a POST action to initialize chunking
        $InitUrl = "$CleanDataverseUrl/api/data/v9.2/InitializeAnnotationBlocksDownload"
        $InitBody = @{
            Target = @{
                "@odata.type" = "Microsoft.Dynamics.CRM.annotation"
                annotationid = $DataverseRowId
            }
        } | ConvertTo-Json -Depth 2
        
        $InitResponse = Invoke-RestMethod -Method Post -Uri $InitUrl -Headers $Headers -Body $InitBody -ContentType "application/json"
    } 
    else {
        # Modern File/Image Columns use a GET function bound to the column
        $InitUrl = "$CleanDataverseUrl/api/data/v9.2/$DataverseTableName($DataverseRowId)/$DataverseFileColumnName/Microsoft.Dynamics.CRM.InitializeFileBlocksDownload"
        $InitResponse = Invoke-RestMethod -Method Get -Uri $InitUrl -Headers $Headers
    }
    
    $FileSize = $InitResponse.FileSizeInBytes
    # Do NOT URL-encode the token, as it is now being passed safely inside a JSON body
    $Token = $InitResponse.FileContinuationToken 
    $BlockSize = 4194304 # 4 MB chunks
    $Offset = 0
    
    Write-Output "Downloading $FileName ($FileSize bytes) in 4MB chunks..."
    
    # Create an empty file and open a file stream to write directly to the Temp Disk
    New-Item -Path $TempFilePath -ItemType File -Force | Out-Null
    $FileStream = [System.IO.File]::OpenWrite($TempFilePath)
    
    try {
        while ($Offset -lt $FileSize) {
            # DownloadBlock is an Action, so it requires a POST request
            $DownloadUrl = "$CleanDataverseUrl/api/data/v9.2/DownloadBlock"
            
            $DownloadBody = @{
                FileContinuationToken = $Token
                Offset = $Offset
                BlockLength = $BlockSize
            } | ConvertTo-Json -Compress

            # Make the POST request
            $BlockResponse = Invoke-RestMethod -Method Post -Uri $DownloadUrl -Headers $Headers -Body $DownloadBody -ContentType "application/json"
            
            # The Web API returns 'Data' as a Base64-encoded string. Convert it back to bytes.
            $BlockBytes = [System.Convert]::FromBase64String($BlockResponse.Data)
            
            # Write the bytes to the disk stream
            $FileStream.Write($BlockBytes, 0, $BlockBytes.Length)
            
            $Offset += $BlockSize
        }
    } finally {
        # Always ensure the stream is closed so SharePoint can safely lock and upload it
        $FileStream.Close() 
    }
    
    Write-Output "Dataverse File retrieved and temporarily saved to disk successfully"
} 
catch {
    $result = @{ success = $false; error = "Failed to retrieve Dataverse file: $($_.Exception.Message)" }
    Write-Output ($result | ConvertTo-Json -Compress)
    exit
}

# ---------------------------------------------------------
# Upload to SharePoint
# ---------------------------------------------------------
try {
    Write-Output "Uploading file from temp disk to SharePoint..."
    
    # Add-PnPFile reads directly from the disk path
    $FileUpload = Add-PnPFile -Path $TempFilePath -Folder $SharePointDocLibName
    
    Write-Output "Dataverse file successfully uploaded to SharePoint"
} 
catch {
    $result = @{ success = $false; error = "Failed to upload Dataverse file to SharePoint: $($_.Exception.Message)" }
    Write-Output ($result | ConvertTo-Json -Compress)
    exit
} 
finally {
    # Clean up the sandbox environment
    if (Test-Path $TempFilePath) {
        Remove-Item $TempFilePath -Force
    }
}

Disconnect-PnPOnline
Write-Output "Script completed successfully" 

# Return success result as JSON
$result = @{
    success = $true
    error = ""
}
Write-Output ($result | ConvertTo-Json -Compress)
