using namespace System.Net
param($name)
$global:erroractionpreference = 1

try
{
    Write-Host "Apply Site Template Invoked"
    $url = $name['newsiteurl']
    $prefix = $name['prefix']
    $theme = $name['theme']
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
        Invoke-PnPSiteTemplate -Path "D:\home\site\WorkBenchTemplate\Lists.xml"
        Write-Host "Template Lists Applied Successfully"

        Connect-PnPOnline -Url $adminSiteURL -Tenant $tenant.value -ClientId $clientId.value -Thumbprint $thumbprint.value
        $siterequest = Set-PnPListItem -List $siterequestlist -Identity $siterequestitemID -Values @{"ProvisioningStage" = "Lists Template Applied"}
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