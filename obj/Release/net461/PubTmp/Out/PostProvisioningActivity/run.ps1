using namespace System.Net
param($name)
$global:erroractionpreference = 1
try
{
    # Interact with query parameters or the body of the request.
    $newsiteurl = $name['newsiteurl']
    $adminSiteURL = $name['adminsiteurl']
    $prefix = $name['prefix']
    $siterequestitemID = $name['siterequestitemid']
    $siterequestlist = $name['siterequestlist']
    $siteowner =   $name['siteowner']
    $projectmanager =   $name['projectmanager']
    $officemanager =   $name['officemanager']
    $contentowners =   $name['contentowners']
    $templatesiteurl = $name['templatesiteurl']
   
    $clientId = ls env:APPSETTING_AzureADApp_ClientID
    $thumbprint = ls env:APPSETTING_AzureADApp_Thumbprint
    $tenant = ls env:APPSETTING_Workbench_Tenant
    Connect-PnPOnline -Url $adminSiteURL -Tenant $tenant.value -ClientId $clientId.value -Thumbprint $thumbprint.value

    #Copy logo to new site and set logo on New Site
    $siteAssets = "SiteAssets"
    $listitem = Get-PnPListItem -List $siterequestlist -Id $siterequestitemID
    $attachments = Get-PnPProperty -ClientObject $listitem -Property "AttachmentFiles"
    $title = $listitem["Title"] 
    $projectCode = $listitem["ProjectCode"]
    $projectPrefix = $listitem["ProjectPrefix"]
    $projectPhase = $listitem["ProjectPhase"]
    $businessUnit = $listitem["BusinessUnit"]
    $projectWebsite = $listitem["SiteURL"]
    $projecttype =  $listitem["ProjectType"]
    Write-Host $title $projectcode $projectprefix $projectPhase

    $source =  $attachments[0].ServerRelativeUrl
    $logoFileName = $attachments[0].FileName 
 
    $streamResult = (Get-PnPFile -Url  $source ).OpenBinaryStream()
    Invoke-PnPQuery
    Write-Host "Logo File Found"
    $logofile = Add-PnPFile -FileName $logoFileName -Folder "WorkbenchSiteLogo" -Stream $streamResult.Value
    $logofileUrl =  $logofile.ServerRelativeUrl  
    $logoFileName =  $logofile.Name
    Write-Host "Logo File Found : "  $logoFileName 
    $sitelogourl = -join($newsiteurl,"/SiteAssets/" , $logoFileName)
    Write-Host   $sitelogourl
   
    Connect-PnPOnline -Url $newsiteurl -Tenant $tenant.value -ClientId $clientId.value -Thumbprint $thumbprint.value
    Copy-PnPFile -SourceUrl $logofileUrl -TargetUrl "SiteAssets" -Force
     
    Set-PnPWeb -SiteLogoUrl $sitelogourl
    Write-host $sitelogourl
    Write-Host "Logo Set for new Site with logo URL"

    #Add Item in Project Details if not exist
    $projectDetailsList = Get-PnPList | Where-Object {$_.Title -Contains "Project Details"} | Select-Object ItemCount
    if($projectDetailsList.ItemCount)
    {
            Write-Host "Projec Details already exists. " $projectDetailsList.ItemCount
    }
    else {
        $projectDetails = Add-PnPListItem -List "Project Details" -Values @{"Title" = $title ;"ProjectCode" = $projectCode; "ProjectPrefix" = $projectprefix; "ProjectPhase" = $projectPhase; "BusinessUnit" = $businessUnit; "ProjectWebsite" = $projectWebsite }
        Write-Host "Project Details Added"
    
    }
    
    #Add items in Contact List if not exist
    $contactsList = Get-PnPList | Where-Object {$_.Title -Contains "Contacts"} | Select-Object ItemCount
    if($contactsList.ItemCount)
    {
            Write-Host "Contacts already exists. " $contactsList.ItemCount
    }
    else {
        $contact = Add-PnPListItem -List "Contacts" -Values @{"Title" = "Workbench Manager" ;"Person" = $siteowner; "Order0" = "1";"ShowOnHomePage"="true" }
        $contact = Add-PnPListItem -List "Contacts" -Values @{"Title" = "Project Manager" ;"Person" = $projectmanager; "Order0" = "2";"ShowOnHomePage"="true" }
        $contact = Add-PnPListItem -List "Contacts" -Values @{"Title" = "Office Manager" ;"Person" = $officemanager; "Order0" = "3";"ShowOnHomePage"="true" }
        Write-Host "Key Contacts Added"
    }
    
    #Add Links for External Users
    <#$folder = Add-PnPFolder -Name "Shared-with-all-external-users" -Folder "Lists/ProjectSystems"
    Set-PnPFolderPermission -List 'Project Systems' -Identity $folder -Group 'Workbench External Users' -AddRole 'Workbench Read'
    $externalUserLink = Add-PnPListItem -List "Project Systems" -Folder "Shared-with-all-external-users" -Values @{"Title" = "Aconex" ;"URL" = "https://www.johnholland.com";"NewTab"="true"; "Icon" = "Globe"; "Comments" = "Construction documentation management system";"Disciplines"="Project Operations, Procurement, Design"; "CoreSystem" = "true"; "Order0" = "10" }
    $externalUserLink = Add-PnPListItem -List "Project Systems" -Folder "Shared-with-all-external-users" -Values @{"Title" = "PPW External" ;"URL" = "https://www.johnholland.com";"NewTab"="true"; "Icon" = "ProjectCollection"; "Comments" = "Project Pack Web"; "CoreSystem" = "true"; "Order0" = "10" }

    $folder = Add-PnPFolder -Name "Shared-with-all-external-users" -Folder "Lists/FeaturedSiteContent"
    Set-PnPFolderPermission -List 'Featured Site Content' -Identity $folder -Group 'Workbench External Users' -AddRole 'Workbench Read'
    $externalUserLink = Add-PnPListItem -List "Featured Site Content" -Folder "Shared-with-all-external-users" -Values @{"Title" = "Org Chart" ;"URL" = "https://www.johnholland.com";"NewTab"="true"; "Icon" = "Org"; "Description" = "View our project org structure";"Featured"="true"; "Order0" = "10" }
    $externalUserLink = Add-PnPListItem -List "Featured Site Content" -Folder "Shared-with-all-external-users" -Values @{"Title" = "About the Project" ;"URL" = $newsiteurl + "/SitePages/Shared-with-all-external-users/About-The-Project.aspx";"NewTab"="true"; "Icon" = "Info"; "Description" = ""; "Featured" = "true"; "Order0" = "10" }
    #>
    Connect-PnPOnline -Url $templatesiteurl -Tenant $tenant.value -ClientId $clientId.value -Thumbprint $thumbprint.value
    $listItemsProjectSys = Get-PnPListItem -List "Project Systems" | Where-Object {$_.FieldValues.FileRef -like "*Shared-with-all-external-users*" -And $_.FileSystemObjectType -ne "Folder"}    
    $listItemsFeaturedContent = Get-PnPListItem -List "Featured Site Content" | Where-Object {$_.FieldValues.FileRef -like "*Shared-with-all-external-users*" -And $_.FileSystemObjectType -ne "Folder"}    
    
    Connect-PnPOnline -Url $newsiteurl -Tenant $tenant.value -ClientId $clientId.value -Thumbprint $thumbprint.value
    foreach ($item in $listItemsProjectSys)
    {
        if($item.FieldValues.FileRef -like "*Shared-with-all-external-users-only*" ){
            Write-Host $item.FieldValues.FileRef $item.FileSystemObjectType $item["Title"] $item["URL"].Url
            $externalUserLink = Add-PnPListItem -List "Project Systems" -Folder "Shared-with-all-external-users-only" -Values @{"Title" = $item["Title"] ;"URL" = $item["URL"].Url ;"NewTab"=$item["NewTab"]; "Icon" = $item["Icon"]; "Comments" = $item["Comments"];"Disciplines"=$item["Disciplines"]; "CoreSystem" = $item["CoreSystem"]; "Order0" = $item["Order0"] }
        }
        else{
            Write-Host $item.FieldValues.FileRef $item.FileSystemObjectType $item["Title"] $item["URL"].Url
            $externalUserLink = Add-PnPListItem -List "Project Systems" -Folder "Shared-with-all-external-users" -Values @{"Title" = $item["Title"] ;"URL" = $item["URL"].Url ;"NewTab"=$item["NewTab"]; "Icon" = $item["Icon"]; "Comments" = $item["Comments"];"Disciplines"=$item["Disciplines"]; "CoreSystem" = $item["CoreSystem"]; "Order0" = $item["Order0"] }
        }
    }
    foreach ($item in $listItemsFeaturedContent)
    {
        if($item.FieldValues.FileRef -like "*Shared-with-all-external-users*" ){
            
            Write-Host $item.FieldValues.FileRef $item.FileSystemObjectType $item["Title"] $item["URL"].Url
            $url = $item["URL"].Url -replace 'https://johnholland.sharepoint.com/sites/workbench-template-test',$newsiteurl
            $externalUserLink = Add-PnPListItem -List "Featured Site Content" -Folder "Shared-with-all-external-users" -Values @{"Title" = $item["Title"] ;"URL" = $url  ;"NewTab"=$item["NewTab"]; "Icon" = $item["Icon"]; "Description" = $item["Description"];"Featured" = $item["Featured"]; "Order0" = $item["Order0"] }
        }
    }
   
    #Remove Workbench Site Management Link Item Permissions
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
    
    #Remove Recent Node
    Remove-PnPNavigationNode -Title "Recent" -Location QuickLaunch -Force
    Write-Host "Removed Recent Node"
    
    #Add Submit an Idea Footer Link if doest exist
    Set-PnPFooter -Title "© John Holland Group" -LogoUrl "" -Enabled:$true    
    $submitIdeaNavLink = Get-PnPNavigationNode -Location Footer | Where-Object {$_.Title -Contains "Submit an idea"} | Select Title, Id
    if($submitIdeaNavLink)
    {
        Write-Host "Submit an Idea Link Alredy Exists. " $submitIdeaNavLink 
    }
    else
    {
        $submitideaurl = $newsiteurl + "/SitePages/Submit-an-Idea.aspx"
        $submitIdeaNavLink = Add-PnPNavigationNode -Title "Submit an idea" -Url $submitideaurl  -Location "Footer" -First -ErrorAction SilentlyContinue
        Write-Host "Footer Nodes Added"
    }

    #Search Page Changed to New Site and Disable allow Sharing
    $searchResultPage = -join($newsiteurl, "/SitePages/Shared-with-all-external-users/Search-Results.aspx")
    Write-Host $searchResultPage
    Set-PnPSearchSettings -SearchPageUrl $searchResultPage
    Set-PnPRequestAccessEmails -Disabled:$true
    
    #Check If Joint Venture Site and if it is then remove Intranet from News
    Write-host $projecttype
    if($projecttype -eq "Joint Venture / Alliance")
    {
        Write-Host "Removing Intranet Site from News"
        #Update Page News webpart to only this site
        $page=Get-PnPClientSidePage -Identity "Shared-with-all-external-users/Home.aspx"
        $webPart = $page.Controls  | ? {$_.Title -eq "News"} 
        $newsWPInstanceId=$webpart.InstanceId
        $webpartJson = $webpart.PropertiesJson  
        $webpartobj = ConvertFrom-Json -InputObject $webpartJson  
        [System.Collections.ArrayList]$newsListJV = $webpartobj.newsSiteList
        if( $newsListJV.Length>1)
        {
            $newsListJV.RemoveAt(1)
        }
        [System.Collections.ArrayList]$sitestJV = $webpartobj.sites
        if( $sitestJV.Length>1)
        {
            $sitestJV.RemoveAt(1)
        }
        $webpartobj.newsDataSourceProp = 1;
        $webpartobj.dataProviderId = "news";
        $webpartobj.newsSiteList = $newsListJV
        $webpartobj.sites =  $sitestJV
        
        $webpartJsonJV = ConvertTo-Json $webpartobj
        Set-PnPPageWebPart -Page $page -Identity $newsWPInstanceId -PropertiesJson $webpartJsonJV
        $page.Publish()
    }
   
    #Remove OOTB ShrePoint Associated Groups
    $web = Get-PnPWeb
    $ownerGroup =  $web.AssociatedOwnerGroup
    Write-Host $ownerGroup
    $membersGroup = $web.AssociatedMemberGroup 
    $visitorGroup = $web.AssociatedVisitorGroup
    if($ownerGroup -ne  $null){
        Remove-PnPGroup -Identity $ownerGroup -Force
    }
    if($membersGroup -ne  $null){
        Remove-PnPGroup -Identity $membersGroup -Force
    }
    if($visitorGroup -ne  $null){
        Remove-PnPGroup -Identity $visitorGroup -Force
    }    
    Write-Host "OOTB groups removed"
    
    #Add Project Manager, Office Manager and Content Owners in SharePoint Groups
    $group=Get-PnPGroup -Identity "Workbench Managers"
    Add-PnPGroupMember -EmailAddress $siteowner -Group $group.Title
    Write-Host "Adding siteowner with Email" $siteowner " to " $group.Title

    $group=Get-PnPGroup -Identity "Workbench Project Staff"
    if($projectmanager){
        Add-PnPGroupMember -EmailAddress $projectmanager -Group $group.Title
        Write-Host "Adding projectmanager with Email" $projectmanager " to " $group.Title
    }
    else{Write-Host "Project Manager is not provided"}
    if($contentowners){
        $contentowners.Split(";") | ForEach {
            if($_){
                Write-Host "Adding Content owner with Email" $_ " to " $group.Title
                Add-PnPGroupMember -EmailAddress $_ -Group $group.Title
            } 
        }
    }
    else{Write-Host "No Content Owners Provided"}

    #Set Uniques Permissions for Site Pages
    Set-PnPList -Identity "Site Pages" -BreakRoleInheritance
    Set-PnPListPermission -Identity "Site Pages" -Group 'Workbench Managers' -AddRole 'Workbench Contribute'
    Set-PnPListPermission -Identity "Site Pages" -Group 'Accounting' -AddRole 'Workbench Read'
    Set-PnPListPermission -Identity "Site Pages" -Group 'Finance' -AddRole 'Workbench Read'
    Set-PnPListPermission -Identity "Site Pages" -Group 'HR' -AddRole 'Workbench Read'
    Set-PnPListPermission -Identity "Site Pages" -Group 'Commercial' -AddRole 'Workbench Read'
    Set-PnPListPermission -Identity "Site Pages" -Group 'Workbench Project Staff' -AddRole 'Workbench Read'
    
    $folders = Get-PnPFolder -List 'Site Pages' 
    foreach ($folder in $folders){
        if($folder.Name -eq "Shared-with-all-external-users"){
            Set-PnPFolderPermission -List "Site Pages" -Identity $folder -Group 'Workbench External Users' -AddRole 'Workbench Read'
        }
        if($folder.Name -eq "Workbench-managers-only"){
            Set-PnPFolderPermission -List "Site Pages" -Identity $folder -Group 'Accounting' -RemoveRole 'Workbench Read'
            Set-PnPFolderPermission -List "Site Pages" -Identity $folder -Group 'Commercial' -RemoveRole 'Workbench Read'
            Set-PnPFolderPermission -List "Site Pages" -Identity $folder -Group 'Finance' -RemoveRole 'Workbench Read'
            Set-PnPFolderPermission -List "Site Pages" -Identity $folder -Group 'HR' -RemoveRole 'Workbench Read'
            Set-PnPFolderPermission -List "Site Pages" -Identity $folder -Group 'Workbench Project Staff' -RemoveRole 'Workbench Read'
            Set-PnPFolderPermission -List "Site Pages" -Identity $folder -Group 'Workbench Managers' -RemoveRole 'Workbench Contribute'
             Set-PnPFolderPermission -List "Site Pages" -Identity $folder -Group 'Workbench Managers' -AddRole 'Workbench Read'
        }
    }
    $sitepages= (Get-PnPListItem -List "Site Pages"  -Fields "Title")  
    foreach($listItem in $sitepages){ 
        Write-Host  $listItem["Title"] 
        if($listItem["Title"] -eq "Home" -or $listItem["Title"] -eq "Submit an Idea" -or $listItem["Title"] -eq "Search Results"  -or $listItem["Title"] -eq "Files" ){
            Set-PnPListItemPermission -List 'Site Pages' -Identity $listItem -Group 'Workbench Managers' -RemoveRole 'Workbench Contribute'
            Set-PnPListItemPermission -List 'Site Pages' -Identity $listItem -Group 'Workbench Managers' -AddRole 'Workbench Read'
        }
    }

    Connect-PnPOnline -Url $adminSiteURL -Tenant $tenant.value -ClientId $clientId.value -Thumbprint $thumbprint.value
    $siterequest = Set-PnPListItem -List $siterequestlist -Identity $siterequestitemID -Values @{"ProvisioningStage" = "Post Site Provisioning"}
    return $name
}
catch
{
    $ErrorMessage = $_.Exception.Message
    if ($ErrorMessage)
    {
        Write-Host "Error: " $ErrorMessage  -Fore red
    }
}
