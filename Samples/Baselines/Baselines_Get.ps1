<#

.COPYRIGHT
Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
See LICENSE in the project root for license information.

#>

####################################################

$IntuneModule = Get-Module -Name "Microsoft.Graph.Intune" -ListAvailable

if (!$IntuneModule){

    write-host "Microsoft.Graph.Intune Powershell module not installed..." -f Red
    write-host "Install by running 'Install-Module Microsoft.Graph.Intune' from an elevated PowerShell prompt" -f Yellow
    write-host "Script can't continue..." -f Red
    write-host
    exit

}

####################################################

if(!(Connect-MSGraph)){

    Connect-MSGraph

}

####################################################

Update-MSGraphEnvironment -SchemaVersion beta -Quiet

####################################################

$DirPath = "C:\Temp\Noyce"

####################################################

$uri = "https://graph.microsoft.com/beta/deviceManagement/templates?`$filter=(isof(%27microsoft.graph.securityBaselineTemplate%27))"

$IntuneBaselines = (Invoke-MSGraphRequest -Url $uri -HttpMethod GET).value

Write-Host

foreach($Baseline in $IntuneBaselines){
    
    $BDN = $Baseline.displayName.replace(" ","")

    Add-Content -Value "Category,PolicyName,PolicyType,PolicySetting,DocURL" -Path "$DirPath\$BDN.csv"

    Write-Host $Baseline.displayName -ForegroundColor Cyan
    Write-Host

    $TemplateID = $Baseline.id

    $uri = "https://graph.microsoft.com/beta/deviceManagement/templates/$TemplateID/categories"

    $Categories = (Invoke-MSGraphRequest -Url $uri -HttpMethod GET).value

        foreach($Category in ($Categories| Sort-Object displayName)){
            
            $CategoryID = $Category.id
            $CategoryName = $Category.displayName
            Write-Host $CategoryName -ForegroundColor Magenta

            $uri = "https://graph.microsoft.com/beta/devicemanagement/templates/$TemplateID/categories/$CategoryID/settingDefinitions"

            $settingDefinitions = (Invoke-MSGraphRequest -Url $uri -HttpMethod GET).value

            foreach($Setting in $settingDefinitions){
            
                $DisplayName = $Setting.displayName.replace(":","").replace(",",":")
                $Path =  $Setting.id.replace("--",": ").replace("_"," - ")
                $DocURL = $Setting.documentationUrl

                $Policy = $Setting.id.split("_")

                $PolicyTYpe = $Policy[0].split("--")[2]
                $PolicySetting = $Policy[1]

                Add-Content -Value "$CategoryName,$DisplayName,$PolicyType,$PolicySetting,$DocURL" -Path "$DirPath\$BDN.csv"

                $DisplayName
                $Path
                $DocURL
                Write-Host               

            }

        }

    Write-Host

}
