using namespace System.Net
param($name)
$global:erroractionpreference = 1

try
{
    Write-Host "Apply Site Template Invoked"
    $url = $name['newsiteurl']
    
    $clientId = ls env:APPSETTING_AzureADApp_ClientID
    $thumbprint = ls env:APPSETTING_AzureADApp_Thumbprint
    $tenant = ls env:APPSETTING_Workbench_Tenant
    Connect-PnPOnline -Url $url -Tenant $tenant.value -ClientId $clientId.value -Thumbprint $thumbprint.value

    $site = Get-PnPSite
    if($site -ne  $null)
    {   Write-Host "Site Found"
        Invoke-PnPSiteTemplate -Path "D:\home\site\WorkBenchTemplate\Libraries.xml"
        Write-Host "Template Libraries Applied Successfully"
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