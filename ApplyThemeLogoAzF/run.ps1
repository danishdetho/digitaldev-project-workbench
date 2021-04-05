using namespace System.Net

# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

# Interact with query parameters or the body of the request.
$name = $Request.Query.Name
if (-not $name) {
    $name = $Request.Body.Name
}

$body = "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response."

if ($name) {
    $body = "Hello, $name. This HTTP triggered function executed successfully."
}
try
{
   
    $siterequestitemID=65
    $siterequestlist = "Project Workbench Sites"
    $adminSiteURL = "https://johnholland.sharepoint.com/sites/workbench-admin-test"
    $siteurl = $Request.Body["SiteURL"]
    $theme = $Request.Body["Theme"]
    $logoURL = $Request.Body["LogoURL"]
    
    Write-Host $siteurl $theme $logoURL

    $clientId = ls env:APPSETTING_AzureADApp_ClientID
    $thumbprint = ls env:APPSETTING_AzureADApp_Thumbprint
    $tenant = ls env:APPSETTING_Workbench_Tenant
    Write-Host $clientId.value
    Write-Host $thumbprint.value
    Connect-PnPOnline -Url $siteurl -Tenant $tenant.value -ClientId $clientId.value -Thumbprint $thumbprint.value

    $theme = $theme -replace '\s',''
    Invoke-PnPSiteTemplate -Path "D:\home\site\WorkBenchTemplate\theme.xml" -Parameters @{"theme"=$theme }
    Write-Host "Theme Applied Successfully"
    $logoFileName = $logoURL.split('/')[-1]
    $sitelogourl = -join($siteurl,"/SiteAssets/" , $logoFileName)
    Write-Host   $sitelogourl
   
    $copiedFile = Copy-PnPFile -SourceUrl $logoURL -TargetUrl "SiteAssets" -Force
    Write-Host $copiedFile $copiedFile.FileName 
    Set-PnPWeb -SiteLogoUrl $sitelogourl
    Write-host $sitelogourl
    Write-Host "Logo Set for new Site with logo URL"

    Write-Host "Done"

}
catch
{
    $ErrorMessage = $_.Exception.Message
    if ($ErrorMessage)
    {
        Write-Host "Error: " $ErrorMessage  -Fore red
    }
    throw
    Set-PnPTraceLog -Off
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})
