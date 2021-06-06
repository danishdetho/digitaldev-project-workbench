param($Context)
$obj = $Context.Input 
$obj  | Export-Csv -Path ".\OrchestratorInput.csv" -IncludeTypeInformation
$callbackuri = $obj['callbackuri']
$callbackuri
$output = @()
try
    {
        if($null -eq $obj)
        {
            Write-host "Null"
            $obj = Import-Csv -Path ".\OrchestratorInput.csv"
        }
        Write-Host "Create Site Activity"
        $output += Invoke-ActivityFunction -FunctionName 'CreateSiteActivity' -Input $obj
   
        Write-Host "Generate Template Activity"
        $output += Invoke-ActivityFunction -FunctionName 'GetSiteTemplateActivity' -Input $obj

        if($null -eq $obj)
        {
            Write-host "Null"
            $obj = Import-Csv -Path ".\OrchestratorInput.csv"
        }
        Write-Host "Apply Template Activity"
        $output += Invoke-ActivityFunction -FunctionName 'ApplyTemplateActivity' -Input $obj

        Write-Host "Apply Template Lists Activity"
        $output += Invoke-ActivityFunction -FunctionName 'ApplyTemplateListsActivity' -Input $obj

        Write-Host "Apply Template Libraries Activity"
        $output += Invoke-ActivityFunction -FunctionName 'ApplyTemplateLibrariesActivity' -Input $obj

        Write-Host "Apply Template Libraries Activity"
        $output += Invoke-ActivityFunction -FunctionName 'ApplyTemplateLibraries1Activity' -Input $obj

        Write-Host "Apply Template Content Activity"
        $output += Invoke-ActivityFunction -FunctionName 'ApplyTemplateContentActivity' -Input $obj

        if($null -eq $obj)
        {
            $obj = Import-Csv -Path ".\OrchestratorInput.csv"
        }
        Write-Host "Post Provisioning Activity"
        $output += Invoke-ActivityFunction -FunctionName 'PostProvisioningActivity' -Input $obj
        
        Write-Host "Site Provisioning Orchestration successful"
        $scv = 200
        $body = @{
            status="200"
            resut="Success"
            error ="NA"
        } | ConvertTo-Json
        $logicAppResponse = Invoke-RestMethod -Uri $callbackuri -Method Post -Body $body -UseBasicParsing -SkipHttpErrorCheck -StatusCodeVariable "scv"

    }
    catch
    {
        $ErrorMessage = $_.Exception.Message
        if ($ErrorMessage)
        {
            $scv = 501
            $body = @{
                status="500"
                resut="Fail"
                error =$_.Exception.Message
        } | ConvertTo-Json
            Write-Host "Error: " $ErrorMessage  -Fore red
             $logicAppResponse = Invoke-RestMethod -Uri $callbackuri -Method Post -Body $body -UseBasicParsing -SkipHttpErrorCheck -StatusCodeVariable "scv"
        }
        throw
    }