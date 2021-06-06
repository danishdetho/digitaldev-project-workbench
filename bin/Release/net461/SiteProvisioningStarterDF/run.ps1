using namespace System.Net

param($Request, $TriggerMetadata)

Write-Host $Request.Query
Write-Host $Request.Body.Values
$FunctionName = $Request.Body['FunctionName']

$OrchestratorInput = @{
    'newsiteurl' = $Request.Body['NewSiteURL']
    'templatesiteurl' = $Request.Body['TemplateSiteUrl']
    'projectname' = $Request.Body['ProjectName']
    'businessunit' = $Request.Body['BusinessUnit']
    'prefix' = $Request.Body['Prefix']
    'sitelogo' = $Request.Body['SiteLogo']
    'callbackuri' = $Request.Body['callbacKURL']
    'adminsiteurl'= $Request.Body['AdminSiteURL']
    'tenantadminsiteurl'= $Request.Body['TenantAdminSiteURL']
    'siterequestitemid' = $Request.Body['SiteRequestItemID']
    'siterequestlist' = $Request.Body['SiteRequestList']
    'theme' = $Request.Body['Theme']
    'siteowner' = $Request.Body['SiteOwner']
    'projectmanager' = $Request.Body['ProjectManager']
    'officemanager' = $Request.Body['OfficeManager']
    'contentowners' = $Request.Body['ContentOwners']
}
Write-host $OrchestratorInput
$InstanceId = Start-NewOrchestration -FunctionName $FunctionName -InputObject $OrchestratorInput
Write-Host "Started orchestration with ID = '$InstanceId'"

$Response = New-OrchestrationCheckStatusResponse -Request $Request -InstanceId $InstanceId
Push-OutputBinding -Name Response -Value $Response
Write-host $Response.statusCode