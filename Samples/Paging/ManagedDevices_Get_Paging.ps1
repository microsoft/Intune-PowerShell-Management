
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

$ManagedDevices = Get-IntuneManagedDevice | Get-MSGraphAllPages

if($ManagedDevices){

    foreach($MD in $ManagedDevices){

        write-host "Managed Device" $MD.deviceName "found..." -ForegroundColor Yellow
        Write-Host
        $MD

    }

}

else {

    Write-Host
    Write-Host "No Managed Devices found..." -ForegroundColor Red
    Write-Host

}