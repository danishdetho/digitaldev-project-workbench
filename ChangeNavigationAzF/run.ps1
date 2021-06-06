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
    $parent = $Request.Body["Parent"]
    $order = $Request.Body["Order"]
    $adminSiteURL = "https://johnholland.sharepoint.com/sites/workbench-admin-test"
    
    Write-Host $siteurl $order $parent

    $clientId = ls env:APPSETTING_AzureADApp_ClientID
    $thumbprint = ls env:APPSETTING_AzureADApp_Thumbprint
    $tenant = ls env:APPSETTING_Workbench_Tenant
    Write-Host $clientId.value
    Write-Host $thumbprint.value
   
    Write-Host "Change Navigation"
    Connect-PnPOnline -Url $adminSiteURL -Tenant $tenant.value -ClientId $clientId.value -Thumbprint $thumbprint.value
    $navSiteLinkRequests =  Get-PnPListItem -List "Workbench Site Navigation" -Query "<View><Query><Where><Eq><FieldRef Name='SiteURL'/><Value Type='Text'>$siteURL</Value></Eq></Where><OrderBy><FieldRef Name='ParentNavigationMenu' /><FieldRef Name='SubMenuOrder' Ascending='False'/></OrderBy></Query></View>"
        
    $ctx = Get-PnPContext
    $field = Get-PnPField -List "Workbench Site Navigation" -Identity ParentNavigationMenu
    $fieldChoice = [Microsoft.SharePoint.Client.ClientContext].GetMethod("CastTo").MakeGenericMethod([Microsoft.SharePoint.Client.FieldChoice]).Invoke($ctx, $field)    
    Connect-PnPOnline -Url $siteurl -Tenant $tenant.value -ClientId $clientId.value -Thumbprint $thumbprint.value

    foreach($choice in $fieldChoice.Choices){
        Write-Host $choice 
        $rootNodeTitle =  $choice.Split(">")[0].Trim()
        $rootNode = Get-PnPNavigationNode -Location QuickLaunch  |  Where-Object {($_.Title -Contains $rootNodeTitle)}
        Write-Host $rootNode.Id
        $parentRoot = Get-PnPNavigationNode -Id $rootNode.Id
        $parentRootChildren = $parentRoot.Children
        
        foreach($rootChildNode in $parentRootChildren){
            if($rootChildNode.Title -eq $choice.Split(">")[1].Trim()){
                $rootChild = Get-PnPNavigationNode -Id $rootChildNode.Id
                $parentSubChildren = $rootChild.Children
                foreach($parentSubNode in $parentSubChildren){
                    Write-Host $parentSubNode.Title "|" $parentSubNode.Id
                    Remove-PnPNavigationNode -Identity $parentSubNode.Id -Force
                }
                $newNavListItems =  $navSiteLinkRequests | Where-Object {($_["ParentNavigationMenu"] -eq $choice)}
    
                foreach($newNavItem in $newNavListItems){
                    $changeNavNodeTitle = $newNavItem["Title"]
                    $changeNavNodeOrder = $newNavItem["SubMenuOrder"]
                    $changeNavNodeAddress = $newNavItem["Address"]
                    $parentTitle = $newNavItem["ParentNavigationMenu"]
                    Write-Host $changeNavNodeTitle " : " $parentRootTitle "," $parenSubtTitle "," $changeNavNodeOrder "," $changeNavNodeAddress
                    $newnode = Add-PnPNavigationNode -Title $changeNavNodeTitle -Url $changeNavNodeAddress -Location "QuickLaunch" -Parent $rootChild.Id -First
    
                }
            
            }
        }
    }


    Write-Host "Done"

    # Associate values to output bindings by calling 'Push-OutputBinding'.
    Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
        StatusCode = [HttpStatusCode]::OK
        Body = $body
    })
}
catch
{
    $ErrorMessage = $_.Exception.Message
    Write-Host "Error: " $ErrorMessage  -Fore red
    throw
}

