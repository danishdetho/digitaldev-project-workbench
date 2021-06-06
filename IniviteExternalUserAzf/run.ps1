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
    $callbackuri  = $Request.Body["CallbackUrl"]
    $externalUserEmail = $Request.Body["Email"]#"danishali.se@Gmail.com"
    $emailBody = $Request.Body["Message"]
    $groupName = $Request.Body["WorkbenchGroup"]#"Workbench External Users"
    Write-Host $externalUserEmail
 
    $clientId = ls env:APPSETTING_AzureADApp_ClientID
    $thumbprint = ls env:APPSETTING_AzureADApp_Thumbprint
    $tenant = ls env:APPSETTING_Workbench_Tenant
    $azureADtenantId = ls env:APPSETTING_AzureAD_TenantID
    $serviceAccountPwd = ls env:APPSETTING_KeyVault_ServiceAccount_Pwd
    $serviceAccountUserName = ls env:APPSETTING_ServiceAccount_UserName
    Write-Host  $serviceAccountUserName.Value $serviceAccountPwd.Value 
        
    Import-Module "D:\home\data\ManagedDependencies\210505063222250.r\AzureAD\2.0.2.130\AzureAD.psd1" -UseWindowsPowerShell
    Write-Host "Importred"
    #PowerShell to Add External User to group
    
    $pwd = ConvertTo-SecureString $serviceAccountPwd.Value -AsPlainText -Force  
    $MyCredential = New-Object -TypeName System.Management.Automation.PSCredential  -ArgumentList $serviceAccountUserName.Value , $pwd
    Connect-AzureAD -credential $MyCredential
    #Connect-AzureAD -TenantDomain $azureADtenantId -CertificateThumbprint $thumbprint.value -ApplicationId $clientId.value
    Connect-PnPOnline -Url $siteurl -Tenant $tenant.value -ClientId $clientId.value -Thumbprint $thumbprint.value
    
    Write-Host "Connected"
    $externalUserEmailList = $externalUserEmail.split(';');
    foreach($exernalEmail in $externalUserEmailList)
    {
        $guestUserEmail = $exernalEmail -replace '@','_'
        $guestUserEmail = $guestUserEmail + "#EXT#@johnholland.onmicrosoft.com"
        Write-Host $guestUserEmail
        
        #Add Guest Account if doesnt exis
        $guestUserADGuestAcc = Get-AzureADUser -ObjectId $guestUserEmail -ErrorAction SilentlyContinue
        if($null -ne $guestUserADGuestAcc){
            Write-Host "User guest account exists in Azure AD"
        }
        else{
            Write-Host "User doesnt Exist. Create Azure AD guest invite"
            New-AzureADMSInvitation -InvitedUserEmailAddress $exernalEmail -SendInvitationMessage $False -InviteRedirectUrl $siteurl #-InvitedUserMessageInfo $invitation 

        }
        #Add User to Site Group
        Add-PnPGroupMember -Identity $groupName -EmailAddress $exernalEmail -SendEmail -EmailBody $emailBody
    }

    
      
        $scv = 200
        $body = @{
            status="200"
            result="Success"
            error ="NA"
        } | ConvertTo-Json
       return $body
    Write-Host "Done"

}
catch
{
    $ErrorMessage = $_.Exception.Message
    if ($ErrorMessage)
    {
        Write-Host "Error: " $ErrorMessage  -Fore red
    }
    $body = @{
            status="500"
            result="Success"
            error =$ErrorMessage
        } | ConvertTo-Json
       return $body
    throw
    Set-PnPTraceLog -Off
}
# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    Body = $body
})
