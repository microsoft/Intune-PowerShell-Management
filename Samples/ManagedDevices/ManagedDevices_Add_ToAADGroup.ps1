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

Function Get-AADDevice(){

<#
.SYNOPSIS
This function is used to get an AAD Device from the Graph API REST interface
.DESCRIPTION
The function connects to the Graph API Interface and gets an AAD Device registered with AAD
.EXAMPLE
Get-AADDevice -DeviceID $DeviceID
Returns an AAD Device from Azure AD
.NOTES
NAME: Get-AADDevice
#>

[cmdletbinding()]

param
(
    $DeviceID
)

# Defining Variables
$graphApiVersion = "v1.0"
$Resource = "devices"
    
    try {

    $uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)?`$filter=deviceId eq '$DeviceID'"

    $uri = $uri.Replace(" ","%20")

    (Invoke-MSGraphRequest -Url $uri -HttpMethod Get).value 

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

Function Add-AADGroupMember(){

<#
.SYNOPSIS
This function is used to add an member to an AAD Group from the Graph API REST interface
.DESCRIPTION
The function connects to the Graph API Interface and adds a member to an AAD Group registered with AAD
.EXAMPLE
Add-AADGroupMember -GroupId $GroupId -AADMemberID $AADMemberID
Returns all users registered with Azure AD
.NOTES
NAME: Add-AADGroupMember
#>

[cmdletbinding()]

param
(
    $GroupId,
    $AADMemberId
)

# Defining Variables
$graphApiVersion = "v1.0"
$Resource = "groups"
    
    try {

    $uri = "https://graph.microsoft.com/$graphApiVersion/$Resource/$GroupId/members/`$ref"

$JSON = @"

{
    "@odata.id": "https://graph.microsoft.com/v1.0/directoryObjects/$AADMemberId"
}

"@

    Invoke-MSGraphRequest -Url $uri -HttpMethod Post -Content $Json

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

Update-MSGraphEnvironment -SchemaVersion beta -Quiet

####################################################

#region AAD Group

Write-Host

# Setting application AAD Group to assign application
$AADGroup = Read-Host -Prompt "Enter the Azure AD device group name where devices will be assigned as members" 

$GroupId = (Get-AADGroup -Filter "displayname eq '$AADGroup'").id

    if($GroupId -eq $null -or $GroupId -eq ""){

    Write-Host "AAD Group - '$AADGroup' doesn't exist, please specify a valid AAD Group..." -ForegroundColor Red
    Write-Host
    exit

    }

    else {

    $GroupMembers = Get-AADGroupMember -groupId $GroupId

    }

#endregion

####################################################

#region Variables and Filter

Write-Host

# Variable used for filter on users displayname
# Note: The filter is case sensitive

$FilterName = Read-Host -Prompt "Specify the Azure AD display name search string" 

    if($FilterName -eq "" -or $FilterName -eq $null){

    Write-Host
    Write-Host "A string is required to identify the set of users." -ForegroundColor Red
    Write-Host
    break

    }

# Count used to calculate how many devices were added to the Group

$count = 0

# Count to check if any devices have already been added to the Group

$countAdded = 0

#endregion

####################################################

Write-Host
Write-Host "Checking if any Managed Devices are registered with Intune..." -ForegroundColor Cyan
Write-Host

$Devices = Get-IntuneManagedDevice

if($Devices){

    Write-Host "Intune Managed Devices found..." -ForegroundColor Yellow
    Write-Host

    foreach($Device in $Devices){

    $DeviceID = $Device.id
    $AAD_DeviceID = $Device.azureActiveDirectoryDeviceId
    $LSD = $Device.lastSyncDateTime
    $userId = $Device.userPrincipalName

    # Getting User information from AAD to get the users displayName

    $UserDisplayName = $Device.userDisplayName

        # Filtering on the display Name to add users device to a specific group

        if(($UserDisplayName).contains("$FilterName")){

        Write-Host "----------------------------------------------------"
        Write-Host

        write-host "Device Name:"$Device.deviceName -f Green
        write-host "Management State:"$Device.managementState
        write-host "Operating System:"$Device.operatingSystem
        write-host "Device Type:"$Device.deviceType
        write-host "Last Sync Date Time:"$Device.lastSyncDateTime
        write-host "Jail Broken:"$Device.jailBroken
        write-host "Compliance State:"$Device.complianceState
        write-host "Enrollment Type:"$Device.enrollmentType
        write-host "AAD Registered:"$Device.aadRegistered
        Write-Host "UPN:"$Device.userPrincipalName
        write-host
        write-host "User Details:" -f Green
        write-host "User Display Name:"$Device.userDisplayName

        Write-Host "Adding user device" $Device.deviceName "to AAD Group $AADGroup..." -ForegroundColor Yellow

        # Getting Device information from Azure AD Devices

        $AAD_Device = Get-AADDevice -DeviceID $AAD_DeviceID    

        $AAD_Id = $AAD_Device.id

            if($GroupMembers.id -contains $AAD_Id){

                Write-Host "Device already exists in AAD Group..." -ForegroundColor Red

                $countAdded++

            }

            else {

                Write-Host "Adding Device to AAD Group..." -ForegroundColor Yellow

                Add-AADGroupMember -GroupId $GroupId -AADMemberId $AAD_Id

                $count++

            }

        Write-Host

        }

    }
    
    Write-Host "----------------------------------------------------"
    Write-Host
    Write-Host "$count devices added to AAD Group '$AADGroup' with filter '$filterName'..." -ForegroundColor Green
    Write-Host "$countAdded devices already in AAD Group '$AADGroup' with filter '$filterName'..." -ForegroundColor Yellow
    Write-Host

}

else {

    write-host "No Intune Managed Devices found..." -f green
    Write-Host

}