using namespace System.Net
param($name)
$global:erroractionpreference = 1
try
{
    # Getting Values of Passed Parameters.
    $sitename = $name['projectname']
    $newsiteurl = $name['newsiteurl']
    $businessUnit = $name['businessunit']
    $description = "Workbench Project Site For " + $name['projectname']
    $tenantadminSiteURL = $name['tenantadminsiteurl']
    $adminSiteURL = $name['adminsiteurl']
    $prefix = $name['prefix']
    $siterequestitemID = $name['siterequestitemid']
    $siterequestlist = $name['siterequestlist']
    $siteowner =   $name['siteowner']

    $clientId = ls env:APPSETTING_AzureADApp_ClientID
    $thumbprint = ls env:APPSETTING_AzureADApp_Thumbprint
    $tenant = ls env:APPSETTING_Workbench_Tenant

    Connect-PnPOnline -Url $tenantadminSiteURL -Tenant $tenant.value -ClientId $clientId.value -Thumbprint $thumbprint.value
    
    $site = Get-PnPTenantSite -Identity $newsiteurl -ErrorAction SilentlyContinue
    if($site -ne  $null)
    {
        Write-Host "Site Exists. Skip Site Creation."
        Set-PnPTenantSite -Identity $newsiteurl -DenyAddAndCustomizePages:$false -SharingCapability ExternalUserSharingOnly
    }  
    else{
        Write-Host "Site Doesnt Exist"
        $site = New-PnPTenantSite -Title  $sitename -Url $newsiteurl -Owner $siteowner -Template 'SITEPAGEPUBLISHING#0' -Lcid 1033 -TimeZone 76 -Wait
        Write-Host  $site.Title 
        Set-PnPTenantSite -Identity $newsiteurl -DenyAddAndCustomizePages:$false -SharingCapability ExternalUserSharingOnly
      
        Connect-PnPOnline -Url $adminSiteURL -Tenant $tenant.value -ClientId $clientId.value -Thumbprint $thumbprint.value
        $siterequest = Set-PnPListItem -List $siterequestlist -Identity $siterequestitemID -Values @{"ProvisioningStage" = "Site Created";"Status" = "Provisioned"}
    }
    return $name
}
catch
{
    $ErrorMessage = $_.Exception.Message
    if ($ErrorMessage)
    {
        Write-Host "Error: " $ErrorMessage  -Fore red
    }
    throw
}
