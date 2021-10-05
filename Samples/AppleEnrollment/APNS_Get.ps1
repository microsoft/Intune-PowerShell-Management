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

$APNS = Get-IntuneApplePushNotificationCertificate

$APNS

$DateTime = $APNS.expirationDateTime.DateTime

Write-Host "Expiration Date Time:" $DateTime
