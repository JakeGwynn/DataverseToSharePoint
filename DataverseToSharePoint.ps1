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

if($DataverseTableName -eq "annotations"){
    $FileUrl = "$DataverseUrl/api/data/v9.2/$DataverseTableName($DataverseRowId)?`$select=$DataverseFileColumnName"
} else {
    $FileUrl = "$DataverseUrl/api/data/v9.2/$DataverseTableName($DataverseRowId)/$DataverseFileColumnName/`$value"
}

try {
    Write-Output "Retrieving Dataverse file" 
    $response = Invoke-WebRequest -Method Get -Uri $FileUrl -Headers $Headers

    if($DataverseTableName -eq "annotations") {
        $FileContent = [System.Convert]::FromBase64String(($response.Content | ConvertFrom-Json).documentbody)
    } else {
        $FileContent = $response.Content 
    }

    Write-Output "Dataverse File retrieved successfully" 
} catch {
    $result = @{
        success = $false
        error = "Failed to retrieve Dataverse file: $($_.Exception.Message)"
    }
    Write-Output ($result | ConvertTo-Json -Compress)
    exit
}

try {
    $Stream = [IO.MemoryStream]::new([byte[]]$FileContent)
    $FileUpload = Add-PnPFile -FileName $FileName -Folder $SharePointDocLibName -Stream $Stream
    Write-Output "Dataverse file successfully uploaded to SharePoint" 
} catch {
    $result = @{
        success = $false
        error = "Failed to upload Dataverse file to SharePoint: $($_.Exception.Message)"
    }
    Write-Output ($result | ConvertTo-Json -Compress)
    exit
}

Disconnect-PnPOnline
Write-Output "Script completed successfully" 

# Return success result as JSON
$result = @{
    success = $true
    error = ""
}
Write-Output ($result | ConvertTo-Json -Compress)

