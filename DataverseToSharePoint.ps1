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

$AppRegistrationCredentials = Get-AutomationPSCredential -Name "AppRegistration"
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

try {
    Write-Output "Retrieving Dataverse file..."

    if ($DataverseTableName -eq "annotations") {
        # Annotations store files as a single Base64 string (cannot be chunked via API)
        $FileUrl = "$DataverseUrl/api/data/v9.2/$DataverseTableName($DataverseRowId)?`$select=$DataverseFileColumnName"
        $response = Invoke-RestMethod -Method Get -Uri $FileUrl -Headers $Headers
        
        $FileBytes = [System.Convert]::FromBase64String($response.documentbody)
        [System.IO.File]::WriteAllBytes($TempFilePath, $FileBytes)
        
        # Free up memory immediately
        $response = $null
        $FileBytes = $null
        [GC]::Collect()
    } 
    else {
        # Modern File/Image Columns support chunked downloading
        $InitUrl = "$DataverseUrl/api/data/v9.2/$DataverseTableName($DataverseRowId)/$DataverseFileColumnName/Microsoft.Dynamics.CRM.InitializeFileBlocksDownload"
        $InitResponse = Invoke-RestMethod -Method Get -Uri $InitUrl -Headers $Headers
        
        $FileSize = $InitResponse.FileSizeInBytes
        $Token = [uri]::EscapeDataString($InitResponse.FileContinuationToken)
        $BlockSize = 4194304 # 4 MB chunks
        $Offset = 0
        
        # Create an empty file and open a file stream
        New-Item -Path $TempFilePath -ItemType File -Force | Out-Null
        $FileStream = [System.IO.File]::OpenWrite($TempFilePath)
        
        try {
            while ($Offset -lt $FileSize) {
                $DownloadUrl = "$DataverseUrl/api/data/v9.2/DownloadBlock(FileContinuationToken='$Token',Offset=$Offset,BlockLength=$BlockSize)"
                
                # Retrieve the block and write directly to the file stream
                $BlockResponse = Invoke-WebRequest -Method Get -Uri $DownloadUrl -Headers $Headers
                $BlockBytes = $BlockResponse.Content
                $FileStream.Write($BlockBytes, 0, $BlockBytes.Length)
                
                $Offset += $BlockSize
            }
        } finally {
            $FileStream.Close() # Always ensure the stream is closed so SharePoint can lock it
        }
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
    
    # Add-PnPFile reads directly from the disk path, bypassing MemoryStream
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


