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

$Devices = Get-IntuneManagedDevice

if($Devices){

Write-Host

    foreach($Device in $Devices){

    $DeviceID = $Device.id

    Write-Host "Device found:" $Device.deviceName -ForegroundColor Yellow

        if($Device.ownerType -eq "personal"){

        Write-Host "Device Ownership:" $Device.ownerType -ForegroundColor Cyan

        }

        elseif($Device.ownerType -eq "company"){

        Write-Host "Device Ownership:" $Device.ownerType -ForegroundColor Magenta

        }
    
    $uri = "https://graph.microsoft.com/beta/deviceManagement/manageddevices('$DeviceID')?`$expand=detectedApps"
    $DetectedApps = (Invoke-MSGraphRequest -Url $uri -HttpMethod Get).detectedApps

        if($DetectedApps){
            
            $DetectedApps | select displayName,version | ft

        }

        else {
        
            Write-Host "No detected apps found..."
            Write-Host
            
        }

    }

}

else {

    write-host "No Intune Managed Devices found..." -f green
    Write-Host

}