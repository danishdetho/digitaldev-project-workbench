using namespace System.Net
param($name)
try
{
    $templatesiteurl = $name['templatesiteurl']
    $newsiteurl = $name['newsiteurl']
    $siterequestitemID = $name['siterequestitemid']
    $siterequestlist = $name['siterequestlist']
    $adminSiteURL = $name['adminsiteurl']

    $clientId = ls env:APPSETTING_AzureADApp_ClientID
    $thumbprint = ls env:APPSETTING_AzureADApp_Thumbprint
    $tenant = ls env:APPSETTING_Workbench_Tenant
    Connect-PnPOnline -Url $templatesiteurl -Tenant $tenant.value -ClientId $clientId.value -Thumbprint $thumbprint.value

    $site = Get-PnPSite
    if($site -ne  $null)
    {
        Write-Host "Template Master Site Found"
       
        Get-PnPSiteTemplate -Out "D:\home\site\WorkBenchTemplate\mastersite.xml" -PersistBrandingFiles -IncludeAllPages  -IncludeSearchConfiguration -IncludeSiteGroups -ExcludeHandlers TermGroups,CustomActions,Lists -Force
        Write-Host "Master Template Extracted"
        
        #Replace JH Human cause space is failing and Text Allignment 
        ((Get-Content -path "D:\home\site\WorkBenchTemplate\mastersite.xml" -Raw) -replace 'JH Human','sitelogo') | Set-Content -Path "D:\home\site\WorkBenchTemplate\mastersite.xml"
        ((Get-Content -path "D:\home\site\WorkBenchTemplate\mastersite.xml" -Raw) -replace 'TextAlignment="Center"','TextAlignment="Left"') | Set-Content -Path "D:\home\site\WorkBenchTemplate\mastersite.xml"
    
        #Get List Titles and Generate Template
        [System.Collections.Generic.List[string]]$listTitles = @()
            
        $Lists =  Get-PnPList | Where-Object {$_.BaseTemplate -eq 100  -and $_.Hidden -eq $false }
        foreach($list in $Lists){
            $listTitles += $list.Title
        }
        Write-Host $listTitles
        Get-PnPSiteTemplate -Out "D:\home\site\WorkBenchTemplate\Lists.xml" -Handlers Lists -ListsToExtract $listTitles -Force -ErrorAction Stop
        Write-Host "Lists Template Extracted Successfully"

        #Check if links already cretaed by previous runs and if not then add links to List template
        Connect-PnPOnline -Url $newsiteurl -Tenant $tenant.value -ClientId $clientId.value -Thumbprint $thumbprint.value

        $projectSystemsList = Get-PnPList | Where-Object {$_.Title -Contains "Project Systems"} | Select-Object ItemCount
        if($projectSystemsList.ItemCount)
        {
            Write-Host "Applications and Project Systems Links already exists. " $projectSystemsList.ItemCount
        }
        else {
            Write-Host "Applications and Project Systems doesnt exists."
            Connect-PnPOnline -Url $templatesiteurl -Tenant $tenant.value -ClientId $clientId.value -Thumbprint $thumbprint.value
            Add-PnPListFoldersToSiteTemplate -Path "D:\home\site\WorkBenchTemplate\Lists.xml" -List "Project Systems"-Recursive -IncludeSecurity
            Write-Host "Folders added to Project Systems Lists Succssfully"
            Add-PnPListFoldersToSiteTemplate -Path "D:\home\site\WorkBenchTemplate\Lists.xml" -List "Featured Site Content"-Recursive -IncludeSecurity
            Write-Host "Folders added to Featured Contents Lists Succssfully"
            Add-PnPDataRowsToSiteTemplate -Path "D:\home\site\WorkBenchTemplate\Lists.xml"  -Query '<View></View>' -List "Project Systems"
            Write-Host "Data added to Project Systems Lists Succssfully"
            Add-PnPDataRowsToSiteTemplate -Path "D:\home\site\WorkBenchTemplate\Lists.xml"  -Query '<View></View>' -List "Featured Site Content" 
            Write-Host "Data added to Featured Contents Lists Succssfully"
        }
        ((Get-Content -path "D:\home\site\WorkBenchTemplate\Lists.xml" -Raw) -replace 'https://johnholland.sharepoint.com/sites/workbench-template-test',$newsiteurl) | Set-Content -Path "D:\home\site\WorkBenchTemplate\Lists.xml"
        
        #Get Libraries Titles and Generate Templates
        $DocumentLibraries = Get-PnPList | Where-Object {$_.BaseTemplate -eq 101 -and $_.Hidden -eq $false -and $_.Title -ne "Style Library" -and $_.Title -ne "Documents" -and $_.Title -ne "Form Templates"} #Or $_.BaseType -eq "DocumentLibrary"
        [System.Collections.Generic.List[string]]$docLibTitles = @()
        [System.Collections.Generic.List[string]]$docLibTitles1 = @()
        
        $i = 0
        foreach($doclib in $DocumentLibraries){
            if($i -le 13){
                $docLibTitles += $doclib.Title
            }
            else{
                $docLibTitles1 += $doclib.Title
            }
            $i++
        }
        Write-Host $docLibTitles
        Write-Host $docLibTitles1

        Get-PnPSiteTemplate -Out "D:\home\site\WorkBenchTemplate\Libraries.xml" -Handlers Lists -ListsToExtract $docLibTitles -Force
        Get-PnPSiteTemplate -Out "D:\home\site\WorkBenchTemplate\Libraries1.xml" -Handlers Lists -ListsToExtract $docLibTitles1 -Force
        
        Write-Host "Libraries Templates Extracted"
        $i = 0
        foreach($doclib in $DocumentLibraries){
            if($i -le 13){
                    Add-PnPListFoldersToSiteTemplate -Path "D:\home\site\WorkBenchTemplate\Libraries.xml" -List $doclib.Title -Recursive -IncludeSecurity
            }
            else{
                    Add-PnPListFoldersToSiteTemplate -Path "D:\home\site\WorkBenchTemplate\Libraries1.xml" -List $doclib.Title -Recursive -IncludeSecurity
            }
            $i++
        }
        Write-Host "Folders added to Document Libraries Succssfully"

        Connect-PnPOnline -Url $adminSiteURL -Tenant $tenant.value -ClientId $clientId.value -Thumbprint $thumbprint.value
        $siterequest = Set-PnPListItem -List $siterequestlist -Identity $siterequestitemID -Values @{"ProvisioningStage" = "Template Generated"}
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