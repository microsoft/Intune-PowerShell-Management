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

Function Set-ManagedDevice(){

<#
.SYNOPSIS
This function is used to set Managed Device property from the Graph API REST interface
.DESCRIPTION
The function connects to the Graph API Interface and sets a Managed Device property
.EXAMPLE
Set-ManagedDevice -id $id -ownerType company
Returns Managed Devices configured in Intune
.NOTES
NAME: Set-ManagedDevice
#>

[cmdletbinding()]

param
(
    [parameter(Mandatory=$true)]
    $id,
    [parameter(Mandatory=$true)]
    [ValidateSet('personal','company')]
    $ownertype
)


$graphApiVersion = "Beta"
$Resource = "deviceManagement/managedDevices"

    try {

        if($ownerType -eq "company"){

$JSON = @"

{
    ownerType:"company"
}

"@

            write-host
            write-host "Are you sure you want to change the device ownership to 'company' on this device? Y or N?"
            $Confirm = read-host

            if($Confirm -eq "y" -or $Confirm -eq "Y"){
            
            # Send Patch command to Graph to change the ownertype
            $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices('$ID')"
            Invoke-MSGraphRequest -Url $uri -HttpMethod PATCH -Content $Json

            }

            else {

            Write-Host "Change of Device Ownership for the device $ID was cancelled..." -ForegroundColor Yellow
            Write-Host

            }
            
        }

        elseif($ownerType -eq "personal"){

$JSON = @"

{
    ownerType:"personal"
}

"@

            write-host
            write-host "Are you sure you want to change the device ownership to 'personal' on this device? Y or N?"
            $Confirm = read-host

            if($Confirm -eq "y" -or $Confirm -eq "Y"){
            
            # Send Patch command to Graph to change the ownertype
            $uri = "https://graph.microsoft.com/beta/deviceManagement/managedDevices('$ID')"
            Invoke-MSGraphRequest -Url $uri -HttpMethod PATCH -Content $Json

            }

            else {

            Write-Host "Change of Device Ownership for the device $ID was cancelled..." -ForegroundColor Yellow
            Write-Host

            }

        }

    }

    catch {

    $ex = $_.Exception
    $errorResponse = $ex.Response.GetResponseStream()
    $reader = New-Object System.IO.StreamReader($errorResponse)
    $reader.BaseStream.Position = 0
    $reader.DiscardBufferedData()
    $responseBody = $reader.ReadToEnd();
    Write-Host "Response content:`n$responseBody" -f Red
    Write-Error "Request to $Uri failed with HTTP Status $($ex.Response.StatusCode) $($ex.Response.StatusDescription)"
    write-host
    break

    }

}

####################################################

Write-Host

$DeviceName = "IPADMINI4"

# Filter for the Managed Device of your choice
$ManagedDevice = Get-IntuneManagedDevice -Filter "deviceName eq '$DeviceName'"

if($ManagedDevice){

    if(@($ManagedDevice.count) -gt 1){

    Write-Host "More than 1 device was found, script supports single deviceID..." -ForegroundColor Red
    Write-Host
    break

    }

    else {

    write-host "Device Name:"$ManagedDevice.deviceName -ForegroundColor Cyan
    write-host "Management State:"$ManagedDevice.managementState
    write-host "Operating System:"$ManagedDevice.operatingSystem
    write-host "Device Type:"$ManagedDevice.deviceType
    write-host "Last Sync Date Time:"$ManagedDevice.lastSyncDateTime
    write-host "Jail Broken:"$ManagedDevice.jailBroken
    write-host "Compliance State:"$ManagedDevice.complianceState
    write-host "Enrollment Type:"$ManagedDevice.enrollmentType
    write-host "AAD Registered:"$ManagedDevice.aadRegistered
    write-host "Management Agent:"$ManagedDevice.managementAgent
    Write-Host "User Principal Name:"$ManagedDevice.userPrincipalName
    Write-Host "Owner Type:"$ManagedDevice.ownerType -ForegroundColor Yellow

    Set-ManagedDevice -id $ManagedDevice.id -ownertype company
    
    Write-Host

    }

}

else {

    Write-Host "No Managed Device found with name '$DeviceName'..." -ForegroundColor Red
    Write-Host

}