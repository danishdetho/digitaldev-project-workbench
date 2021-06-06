using namespace System.Net
param($name)

try
{
    Write-Host "Apply Site Content Template Invoked"
    $url = $name['newsiteurl']
    $siterequestitemID = $name['siterequestitemid']
    $siterequestlist = $name['siterequestlist']
    $adminSiteURL = $name['adminsiteurl']
 
    $clientId = ls env:APPSETTING_AzureADApp_ClientID
    $thumbprint = ls env:APPSETTING_AzureADApp_Thumbprint
    $tenant = ls env:APPSETTING_Workbench_Tenant
    Connect-PnPOnline -Url $url -Tenant $tenant.value -ClientId $clientId.value -Thumbprint $thumbprint.value

    $site = Get-PnPSite
    if($site -ne  $null)
    {   
        Write-Host "Site Found"
        Invoke-PnPSiteTemplate -Path "D:\home\site\WorkBenchTemplate\mastersite.xml"  -ResourceFolder "D:\home\site\WorkBenchTemplate\"  -ClearNavigation -OverwriteSystemPropertyBagValues -ExcludeHandlers SiteSettings,Features,ApplicationLifecycleManagement,RegionalSettings,SearchSettings,ContentTypes,Fields,TermGroups
        Write-Host "Template Content Applied Successfully"

        Connect-PnPOnline -Url $adminSiteURL -Tenant $tenant.value -ClientId $clientId.value -Thumbprint $thumbprint.value
        $siterequest = Set-PnPListItem -List $siterequestlist -Identity $siterequestitemID -Values @{"ProvisioningStage" = "Content Template Applied"}
    }
       return $name
}
catch
{
    $ErrorMessage = $_.Exception.Message
    if ($ErrorMessage)
    {
        Write-Host "Error: " $ErrorMessage  -Fore red
        throw
    }
}
