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

Function Get-RemoteActionAudit(){

<#
.SYNOPSIS
This function is used to get Remote Action Audits from the Graph API REST interface
.DESCRIPTION
The function connects to the Graph API Interface and gets any Remote Action Audits
.EXAMPLE
Get-RemoteActionAudit
Returns any device compliance policies configured in Intune
.NOTES
NAME: Get-RemoteActionAudit
#>

[cmdletbinding()]

$graphApiVersion = "Beta"
$Resource = "deviceManagement/remoteActionAudits"

    try {

    $uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)"
    (Invoke-MSGraphRequest -Url $uri -HttpMethod Get).Value

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

Get-RemoteActionAudit