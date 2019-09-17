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

$ExportPath = Read-Host -Prompt "Please specify a path to export Managed Devices hardware data to e.g. C:\IntuneOutput"

    # If the directory path doesn't exist prompt user to create the directory

    if(!(Test-Path "$ExportPath")){

    Write-Host
    Write-Host "Path '$ExportPath' doesn't exist, do you want to create this directory? Y or N?" -ForegroundColor Yellow

    $Confirm = read-host

        if($Confirm -eq "y" -or $Confirm -eq "Y"){

        new-item -ItemType Directory -Path "$ExportPath" | Out-Null
        Write-Host

        }

        else {

        Write-Host "Creation of directory path was cancelled..." -ForegroundColor Red
        Write-Host
        break

        }

    }

Write-Host

####################################################

$Devices = Get-IntuneManagedDevice

if($Devices){

    $Results = @()

    foreach($Device in $Devices){

    $DeviceID = $Device.id

    Write-Host "Device found:" $Device.deviceName -ForegroundColor Yellow
    Write-Host

    $uri = "https://graph.microsoft.com/beta/deviceManagement/manageddevices('$DeviceID')?`$select=hardwareinformation,iccid,udid"
    $DeviceInfo = Invoke-MSGraphRequest -Url $uri -HttpMethod Get

    $DeviceNoHardware = $Device | select * -ExcludeProperty hardwareInformation,deviceActionResults,userId,imei,manufacturer,model,isSupervised,isEncrypted,serialNumber,meid,subscriberCarrier,iccid,udid
    $HardwareExcludes = $DeviceInfo.hardwareInformation | select * -ExcludeProperty sharedDeviceCachedUsers,phoneNumber
    $OtherDeviceInfo = $DeviceInfo | select iccid,udid

        $Object = New-Object System.Object

            foreach($Property in $DeviceNoHardware.psobject.Properties){

                $Object | Add-Member -MemberType NoteProperty -Name $Property.Name -Value $Property.Value

            }

            foreach($Property in $HardwareExcludes.psobject.Properties){

                $Object | Add-Member -MemberType NoteProperty -Name $Property.Name -Value $Property.Value

            }

            foreach($Property in $OtherDeviceInfo.psobject.Properties){

                $Object | Add-Member -MemberType NoteProperty -Name $Property.Name -Value $Property.Value

            }

        $Results += $Object

        $Object

    }

    $Date = get-date

    $Output = "ManagedDeviceHardwareInfo_" + $Date.Day + "-" + $Date.Month + "-" + $Date.Year + "_" + $Date.Hour + "-" + $Date.Minute

    # Exporting Data to CSV file in provided directory
    $Results | Export-Csv "$ExportPath\$Output.csv" -NoTypeInformation
    write-host "CSV created in $ExportPath\$Output.csv..." -f cyan

}

else {

write-host "No Intune Managed Devices found..." -f green
Write-Host

}