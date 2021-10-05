<#

COPYRIGHT
Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
See LICENSE in the project root for license information.

#>

####################################################

$IntuneModule = Get-Module -Name "Microsoft.Graph.Intune" -ListAvailable

if (!$IntuneModule){

    write-host "Microsoft.Graph.Intune Powershell module is not installed..." -f Red
    write-host "Install the module by running 'Install-Module Microsoft.Graph.Intune' from an elevated PowerShell prompt" -f Yellow
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

$AndroidApps = Get-IntuneMobileApp -Filter "(isof('microsoft.graph.androidLobApp') or isof('microsoft.graph.androidStoreApp'))" -Expand assignments | Get-MSGraphAllPages

if($AndroidApps){

    $apps = @()

    Write-Host
    Write-Host "Count of Android Lob Apps:" @($AndroidApps | ? { $_.'@odata.type' -eq "#microsoft.graph.androidLobApp" }).count
    Write-Host "Count of Android Store Apps:" @($AndroidApps | ? { $_.'@odata.type' -eq "#microsoft.graph.androidStoreApp" }).count
    Write-Host
    write-host "-------------------------------------------------------------------"
    Write-Host

    $AndroidApps | foreach {

        $AppId = $_.id
        $DN = $_.displayName
        $Type = $_.'@odata.type'

        Write-Host "Android Device Admin App: $DN - $Type" -ForegroundColor Cyan
        
        if($_.assignments -ne $null -and $DeviceStatuses){

            $_.assignments | foreach {

                if($_.target.'@odata.type' -eq "#microsoft.graph.groupAssignmentTarget" -or $_.target.'@odata.type' -eq "#microsoft.graph.exclusionGroupAssignmentTarget"){

                    $AAD_GroupName = (Get-AADGroup -groupId $_.target.groupId).displayName

                    if($_.target.'@odata.type'.contains("exclusion")){ $Mode = "Excluded" }
                    else { $Mode = "Included" }

                }

                else {

                    $AAD_GroupName = $_.target.'@odata.type'.split(".")[2]

                    $AAD_GroupName = ($AAD_GroupName -creplace  '([A-Z\W_]|\d+)(?<![a-z])',' $&').trim()

                }

                $Intent = $_.intent

                $item = New-Object PSObject
                $item | Add-Member -type NoteProperty -Name 'DisplayName' -Value $DN
                $item | Add-Member -type NoteProperty -Name 'Type' -Value ($Type.split(".")[2])
                $item | Add-Member -type NoteProperty -Name 'InstallIntent' -Value $Intent
                $item | Add-Member -type NoteProperty -Name 'Mode' -Value $Mode
                $item | Add-Member -type NoteProperty -Name 'AADGroupName' -Value $AAD_GroupName

                $apps += $item

                Write-Host "Assignment: $Intent Intent - $AAD_GroupName" -ForegroundColor Green

            }

        }

        elseif($_.assignments -ne $null) {

            $_.assignments | foreach {

                if($_.target.'@odata.type' -eq "#microsoft.graph.groupAssignmentTarget" -or $_.target.'@odata.type' -eq "#microsoft.graph.exclusionGroupAssignmentTarget"){

                    $AAD_GroupName = (Get-AADGroup -groupId $_.target.groupId).displayName

                    if($_.target.'@odata.type'.contains("exclusion")){ $Mode = "Excluded" }
                    else { $Mode = "Included" }

                }

                else {

                    $AAD_GroupName = $_.target.'@odata.type'.split(".")[2]

                    $AAD_GroupName = ($AAD_GroupName -creplace  '([A-Z\W_]|\d+)(?<![a-z])',' $&').trim()

                }

                $Intent = $_.intent

                $item = New-Object PSObject
                $item | Add-Member -type NoteProperty -Name 'DisplayName' -Value $DN
                $item | Add-Member -type NoteProperty -Name 'Type' -Value ($Type.split(".")[2])
                $item | Add-Member -type NoteProperty -Name 'InstallIntent' -Value $Intent
                $item | Add-Member -type NoteProperty -Name 'Mode' -Value $Mode
                $item | Add-Member -type NoteProperty -Name 'AADGroupName' -Value $AAD_GroupName

                $apps += $item

                Write-Host "Assignment: $Intent Intent - $AAD_GroupName"

            }

        }

        else {

            $item = New-Object PSObject
            $item | Add-Member -type NoteProperty -Name 'DisplayName' -Value $DN
            $item | Add-Member -type NoteProperty -Name 'Type' -Value ($Type.split(".")[2])
            $item | Add-Member -type NoteProperty -Name 'InstallIntent' -Value "No Assignments found"
            $item | Add-Member -type NoteProperty -Name 'Mode' -Value ""
            $item | Add-Member -type NoteProperty -Name 'AADGroupName' -Value ""

            $apps += $item
            
            Write-Host "No Assignments found..." -ForegroundColor Red

        }

        Write-Host

    }

    $header = "<style>"
    $header = $header + "BODY{background-color:white;font-family:verdana}"
    $header = $header + "TABLE{border-width: 1px;border-style: solid;border-color: black;border-collapse: collapse;font-size:12px}"
    $header = $header + "TH{border-width: 1px;padding: 0px;border-style: solid;border-color: black;background-color:#0078D4;color:white;padding: 5px;}"
    $header = $header + "TD{border-width: 1px;padding: 0px;border-style: solid;border-color: black;padding: 5px;}"
    $header = $header + "</style>"

    $date = (get-date)
    $FileName_HTML = "IntuneAndroidDeviceAdminApps.html"

    $apps | ConvertTo-Html -Title "Intune - Android Device Admin Apps Report" -Head $Header `
    -Body "<H3>Intune - Android Device Admin Apps Report</H3>" `
    -Pre "<P>Generated on the $date</P>" `
    | Out-File -FilePath $env:temp\$FileName_HTML

    write-host "HTML Report created '$env:temp\$FileName_HTML'..." -ForegroundColor Green
    Write-Host

}

else {

    Write-Host "No Android Device Admin Apps found..." -ForegroundColor Red
    Write-Host

}

