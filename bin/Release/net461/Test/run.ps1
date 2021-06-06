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
    $userId = "danishdetho@bazzinga2019.onmicrosoft.com"    
    $plainText= "rum.dani2021"  
    $pwd = ConvertTo-SecureString $plainText -AsPlainText -Force    
    $creds = New-Object System.Management.Automation.PSCredential($userId,$pwd)  
    #Write-Host $(Get-Module -ListAvailable | Select-Object Name, Path)

    $siterequestitemID=65
    $siterequestlist = "Project Workbench Sites"
    $adminSiteURL = "https://johnholland-admin.sharepoint.com/"
    $newsiteurl = "https://johnholland.sharepoint.com/sites/sc4jv"
    $templatesiteurl = "https://johnholland.sharepoint.com/sites/Workbench-template-test"
    
    $clientId = ls env:APPSETTING_AzureADApp_ClientID
    $thumbprint = ls env:APPSETTING_AzureADApp_Thumbprint
    $tenant = ls env:APPSETTING_Workbench_Tenant
    Write-Host $clientId.value
    Write-Host $thumbprint.value
    
    Connect-PnPOnline -Url $newsiteurl -Tenant $tenant.value -ClientId $clientId.value -Thumbprint $thumbprint.value
    #$site = Get-PnPTenantSite -Url $newsiteurl
    #Set-PnPTenantSite -Identity $site -DenyAddAndCustomizePages:$false
    $workbenchManagementLink = Get-PnPListItem -List "Featured Site Content" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>Workbench Management</Value></Eq></Where></Query></View>"   
    if($workbenchManagementLink -ne $null){
        foreach($link in $workbenchManagementLink)
        {
            Write-Host $link["Title"]
            Set-PnPListItemPermission -List 'Featured Site Content' -Identity $link -Group 'Accounting' -RemoveRole 'Workbench Read'
            Set-PnPListItemPermission -List 'Featured Site Content' -Identity $link -Group 'Commercial' -RemoveRole 'Workbench Read'
            Set-PnPListItemPermission -List 'Featured Site Content' -Identity $link -Group 'Finance' -RemoveRole 'Workbench Read'
            Set-PnPListItemPermission -List 'Featured Site Content' -Identity $link -Group 'HR' -RemoveRole 'Workbench Read'
            Set-PnPListItemPermission -List 'Featured Site Content' -Identity $link -Group 'Workbench Project Staff' -RemoveRole 'Workbench Read'
        }
    }

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
    Body = $(Get-Module -ListAvailable | Select-Object Name, Path)
})
