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
   
    $adminSiteURL = $Request.Body["AdminSiteURL"]
    $siteurl = $Request.Body["SiteURL"]
    $theme = $Request.Body["Theme"]
    $logoURL = $Request.Body["SiteLogo"]
    $requestType  = $Request.Body["RequestType"] -split ','
    $projectWebsite = $Request.Body["ProjectWebsite"]
    Write-Host $siteurl $theme $requestType

    $clientId = ls env:APPSETTING_AzureADApp_ClientID
    $thumbprint = ls env:APPSETTING_AzureADApp_Thumbprint
    $tenant = ls env:APPSETTING_Workbench_Tenant
    Write-Host $clientId.value
    Write-Host $thumbprint.value
    Connect-PnPOnline -Url $siteurl -Tenant $tenant.value -ClientId $clientId.value -Thumbprint $thumbprint.value

    if($requestType -contains "Theme")
    {
         Write-Host "Change Theme"
        if($theme -eq "John Holland")
        {
            Get-PnPTenantTheme -Name "John Holland" | Set-PnPWebTheme
        }
        else
        {
            $theme = $theme -replace '\s',''
            Invoke-PnPSiteTemplate -Path "D:\home\site\WorkBenchTemplate\theme.xml" -Parameters @{"theme"=$theme }
            Write-Host "Theme Applied Successfully"
        }
    }
    if($requestType -contains "Logo")
    {
        Write-Host "Change Logo"
        $logoFileName = $logoURL.split('/')[-1]
        $sitelogourl = -join($siteurl,"/SiteAssets/" , $logoFileName)
        $logoURL = $logoURL -replace "https://johnholland.sharepoint.com",""
        Write-Host   $logoURL

        Copy-PnPFile -SourceUrl $logoURL -TargetUrl "SiteAssets" -Force
        Set-PnPWeb -SiteLogoUrl $sitelogourl
        Write-host $sitelogourl
        Write-Host "Logo Set for new Site with logo URL"
    }
    if($requestType -contains "Project Website")
    {
        Write-Host "Change Project Website"
        $projectwebsiteitem = Set-PnPListItem -List "Project Details" -Identity 1 -Values @{"ProjectWebsite" = $projectWebsite}
        $projectWebsiteNavLink = Get-PnPNavigationNode -Location Footer  |  Where-Object {$_.Title -Contains "Project website"} 
        if($projectWebsiteNavLink)
        {
            $id =  $projectWebsiteNavLink.ID
            Remove-PnPNavigationNode -Identity $id -Force
            $newWebsiteLink = Add-PnPNavigationNode -Location Footer -Title "Project website" -Url $projectWebsite -First
            Write-Host "Project Website Link Removed and new one Added. " $projectWebsite 
        }
        else
        {
            $newWebsiteLink = Add-PnPNavigationNode -Location Footer -Title "Project Website" -Url $projectWebsite -First
            Write-Host "Project Website  Link Added. " $projectWebsite
        }
    }
     Write-Host $callbackuri
    $scv = 200
        $body = @{
            status="200"
            result="Success"
            error ="NA"
        } | ConvertTo-Json
        #$logicAppResponse = Invoke-RestMethod -Uri $callbackuri -Method Post -Body $body -UseBasicParsing -SkipHttpErrorCheck -StatusCodeVariable "scv"
    
    Write-Host "Done"
    return $body 

}
catch
{
    $ErrorMessage = $_.Exception.Message
    if ($ErrorMessage)
    {
        $scv = 501
        $body = @{
            status="500"
            resut="Fail"
            error =$_.Exception.Message
    } | ConvertTo-Json
    Write-Host "Error: " $ErrorMessage  -Fore red
    return $body 
            #$logicAppResponse = Invoke-RestMethod -Uri $callbackuri -Method Post -Body $body -UseBasicParsing -SkipHttpErrorCheck -StatusCodeVariable "scv"
        }
    throw
}

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})
