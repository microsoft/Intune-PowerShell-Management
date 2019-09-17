
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

# Filter for the minimum number of days where the device hasn't checked in
$days = 30

$daysago = "{0:s}" -f (get-date).AddDays(-$days) + "Z"

$CurrentTime = [System.DateTimeOffset]::Now

Write-Host
Write-Host "Checking to see if there are devices that haven't synced in the last $days days..." -f Yellow
Write-Host

####################################################

$Devices = Get-IntuneManagedDevice -Filter "lastSyncDateTime le $daysago" | sort deviceName

$Devices = $Devices | ? { $_.managementAgent -ne "eas" }

# If there are devices not synced in the past 30 days script continues
        
if($Devices){

$DeviceCount = @($Devices).count

Write-Host "There are" $Devices.count "devices that have not synced in the last $days days..." -ForegroundColor Red
Write-Host

    # Looping through all the devices returned
                       
    foreach($Device in $Devices){

        write-host "------------------------------------------------------------------"
        Write-Host

        $DeviceID = $Device.id
        $LSD = $Device.lastSyncDateTime
        $EDT = $Device.enrolledDateTime

        write-host "Device Name:"$Device.deviceName -f Green
        write-host "Management State:"$Device.managementState
        write-host "Operating System:"$Device.operatingSystem
        write-host "Device Type:"$Device.deviceType
        write-host "Last Sync Date Time:"$Device.lastSyncDateTime
        write-host "Enrolled Date Time:"$Device.enrolledDateTime
        write-host "Jail Broken:"$Device.jailBroken
        write-host "Compliance State:"$Device.complianceState
        write-host "AAD Registered:"$Device.aadRegistered
        write-host "Management Agent:"$Device.managementAgent

        $TimeDifference = $CurrentTime - $LSD

        $TotalMinutes = ($TimeDifference.TotalMinutes).tostring().split(".")[0]

        write-host
        write-host "Device last synced"$TimeDifference.days "days ago..." -ForegroundColor Red
        Write-Host

    }

}

else {

    write-host "------------------------------------------------------------------"
    Write-Host
    write-host "No Devices not checked in the last $days days found..." -f green
    Write-Host

}