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

Function Get-DirectoryRoles(){

<#
.SYNOPSIS
This function is used to get Directory Roles from the Graph API REST interface
.DESCRIPTION
The function connects to the Graph API Interface and gets any Directory Role registered
.EXAMPLE
Get-DirectoryRoles
Returns all Directory Roles registered
.NOTES
NAME: Get-DirectoryRoles
#>

[cmdletbinding()]

param
(
    $RoleId,
    [ValidateSet("members")]
    [string]
    $Property
)

# Defining Variables
$graphApiVersion = "v1.0"
$Resource = "directoryRoles"
    
    try {
        
        if($RoleId -eq "" -or $RoleId -eq $null){
        
        $uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)"
        (Invoke-MSGraphRequest -Url $uri -HttpMethod Get).Value
        
        }

        else {
            
            if($Property -eq "" -or $Property -eq $null){

            $uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)/$RoleId"
            Invoke-MSGraphRequest -Url $uri -HttpMethod Get

            }

            else {

            $uri = "https://graph.microsoft.com/$graphApiVersion/$($Resource)/$RoleId/$Property"
            (Invoke-MSGraphRequest -Url $uri -HttpMethod Get).Value

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
Write-Host "Please specify which Directory Role you want to query for User membership:" -ForegroundColor Yellow
Write-Host

$Roles = (Get-DirectoryRoles | Select-Object displayName).displayName | Sort-Object

$menu = @{}

for ($i=1;$i -le $Roles.count; $i++) 
{ Write-Host "$i. $($Roles[$i-1])" 
$menu.Add($i,($Roles[$i-1]))}

Write-Host

[int]$ans = Read-Host 'Enter Directory Role to query (Numerical value)'

$selection = $menu.Item($ans)

    if($selection){

    Write-Host
    Write-Host $selection -f Cyan

    $Directory_Role = (Get-DirectoryRoles | Where-Object { $_.displayName -eq "$Selection" })

    $Members = Get-DirectoryRoles -RoleId $Directory_Role.id -Property members

        if($Members){

            $Members | ForEach-Object { $_.displayName + " - " + $_.userPrincipalName }

        }

        else {

            Write-Host "No Users assigned to '$selection' Directory Role..." -ForegroundColor Red

        }

    }
        
    else {

        Write-Host
        Write-Host "Directory Role specified is invalid..." -ForegroundColor Red
        Write-Host "Please specify a valid Directory Role..." -ForegroundColor Red
        Write-Host
        break

    }

Write-Host