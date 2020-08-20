
<#

.COPYRIGHT
Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
See LICENSE in the project root for license information.

#>

param
(
    [Parameter(Mandatory = $true, HelpMessage = "FileIdToRemove, list files and find the Id")]
    $FileIdToRemove,
    [Parameter(Mandatory = $false, HelpMessage = "Environment")]
    [ValidateSet('OneDF', 'Selfhost', 'CTIP', 'PE')]
    $Environment = "PE"
)

####################################################

function Get-AuthToken {

<#
.SYNOPSIS
This function is used to authenticate with the Graph API REST interface
.DESCRIPTION
The function authenticate with the Graph API Interface with the tenant name
.EXAMPLE
Get-AuthToken
Authenticates you with the Graph API interface
.NOTES
NAME: Get-AuthToken
#>

[cmdletbinding()]

param
(
    [Parameter(Mandatory = $true, HelpMessage = "Environment")]
    [ValidateSet('OneDF', 'Selfhost', 'CTIP', 'Canary', 'PE')]
    $Environment = "",
    
    [Parameter(Mandatory=$true)]
    $User
)

$userUpn = New-Object "System.Net.Mail.MailAddress" -ArgumentList $User

$tenant = $userUpn.Host

Write-Host "Checking for AzureAD module..."

    $AadModule = Get-Module -Name "AzureAD" -ListAvailable

    if ($AadModule -eq $null) {

        Write-Host "AzureAD PowerShell module not found, looking for AzureADPreview"
        $AadModule = Get-Module -Name "AzureADPreview" -ListAvailable

    }

    if ($AadModule -eq $null) {
        write-host
        write-host "AzureAD Powershell module not installed..." -f Red
        write-host "Install by running 'Install-Module AzureAD' or 'Install-Module AzureADPreview' from an elevated PowerShell prompt" -f Yellow
        write-host "Script can't continue..." -f Red
        write-host
        exit
    }

# Getting path to ActiveDirectory Assemblies
# If the module count is greater than 1 find the latest version

    if($AadModule.count -gt 1){

        $Latest_Version = ($AadModule | select version | Sort-Object)[-1]

        $aadModule = $AadModule | ? { $_.version -eq $Latest_Version.version }

            # Checking if there are multiple versions of the same module found

            if($AadModule.count -gt 1){

            $aadModule = $AadModule | select -Unique

            }

        $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"

    }

    else {

        $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
        $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"

    }

[System.Reflection.Assembly]::LoadFrom($adal) | Out-Null

[System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null



$clientId = "d1ddf0e4-d672-4dae-b554-9d5bdfd93547"

$resourceAppIdURI = "https://graph.microsoft.com"

$authority = "https://login.microsoftonline.com/$Tenant"

if ($Environment -eq "OneDF")
{
  $resourceAppIdURI = "https://graph.microsoft-ppe.com"
  $authority = "https://login.windows-ppe.net/$Tenant"
  $clientId="fc20a1af-7a79-4693-afdb-0d402f43f24a"
}

$redirectUri = "urn:ietf:wg:oauth:2.0:oob"

    try {

    $authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $authority

    # https://msdn.microsoft.com/en-us/library/azure/microsoft.identitymodel.clients.activedirectory.promptbehavior.aspx
    # Change the prompt behaviour to force credentials each time: Auto, Always, Never, RefreshSession

    $platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Auto"

    $userId = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier" -ArgumentList ($User, "OptionalDisplayableId")

    $authResult = $authContext.AcquireTokenAsync($resourceAppIdURI,$clientId,$redirectUri,$platformParameters,$userId).Result

        # If the accesstoken is valid then create the authentication header

        if($authResult.AccessToken){

        # Creating header for Authorization token

        $authHeader = @{
            'Content-Type'='application/json'
            'Authorization'="Bearer " + $authResult.AccessToken
            'ExpiresOn'=$authResult.ExpiresOn
            }

        return $authHeader

        }

        else {

        Write-Host
        Write-Host "Authorization Access Token is null, please re-run authentication..." -ForegroundColor Red
        Write-Host
        break

        }

    }

    catch {

    write-host $_.Exception.Message -f Red
    write-host $_.Exception.ItemName -f Red
    write-host
    break

    }

}

####################################################

Function Remove-ADMXFile(){

<#
.SYNOPSIS
This function is used list ADMX file
.DESCRIPTION
.EXAMPLE
.NOTES
#>

[cmdletbinding()]

param
(
    $FileIdToRemove, 
    $serviceEndpoint
)

$DCP_resource = "deviceManagement/groupPolicyUploadedDefinitionFiles('$($FileIdToRemove)')/remove"
Write-Verbose "Resource: $DCP_resource"

    try {
        $uri = "$serviceEndpoint/$($DCP_resource)"
        Invoke-RestMethod -Uri $uri -Headers $authToken -Method Post -ContentType "none"
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

Function Test-JSON(){

<#
.SYNOPSIS
This function is used to test if the JSON passed to a REST Post request is valid
.DESCRIPTION
The function tests if the JSON passed to the REST Post is valid
.EXAMPLE
Test-JSON -JSON $JSON
Test if the JSON is valid before calling the Graph REST interface
.NOTES
NAME: Test-JSON
#>

param (

$JSON

)

    try {

    $TestJSON = ConvertFrom-Json $JSON -ErrorAction Stop
    $validJson = $true

    }

    catch {

    $validJson = $false
    $_.Exception

    }

    if (!$validJson){

    Write-Host "Provided JSON isn't in valid JSON format" -f Red
    break

    }

}

####################################################


#region Authentication

write-host

# Checking if authToken exists before running authentication
if($global:authToken){

    # Setting DateTime to Universal time to work in all timezones
    $DateTime = (Get-Date).ToUniversalTime()

    # If the authToken exists checking when it expires
    $TokenExpires = ($authToken.ExpiresOn.datetime - $DateTime).Minutes

        if($TokenExpires -le 0){

        write-host "Authentication Token expired" $TokenExpires "minutes ago" -ForegroundColor Yellow
        write-host

            # Defining User Principal Name if not present

            if($User -eq $null -or $User -eq ""){

            $User = Read-Host -Prompt "Please specify your user principal name for Azure Authentication"
            Write-Host

            }

        $global:authToken = Get-AuthToken -User $User -Environment $Environment

        }
}

# Authentication doesn't exist, calling Get-AuthToken function

else {

    if($User -eq $null -or $User -eq ""){

    $User = Read-Host -Prompt "Please specify your user principal name for Azure Authentication"
    Write-Host

    }

# Getting the authorization token
$global:authToken = Get-AuthToken -User $User -Environment $Environment

}

#endregion

####################################################

$iOS = @"

{
    "@odata.type": "#microsoft.graph.iosGeneralDeviceConfiguration",
    "description": "iOS Device Restriction Policy",
    "displayName": "iOS Device Restriction Policy",
    "accountBlockModification": false,
    "activationLockAllowWhenSupervised": false,
    "airDropBlocked": false,
    "airDropForceUnmanagedDropTarget": false,
    "airPlayForcePairingPasswordForOutgoingRequests": false,
    "appleWatchBlockPairing": false,
    "appleWatchForceWristDetection": false,
    "appleNewsBlocked": false,
    "appsVisibilityList": [],
    "appsVisibilityListType": "none",
    "appStoreBlockAutomaticDownloads": false,
    "appStoreBlocked": false,
    "appStoreBlockInAppPurchases": false,
    "appStoreBlockUIAppInstallation": false,
    "appStoreRequirePassword": false,
    "bluetoothBlockModification": false,
    "cameraBlocked": false,
    "cellularBlockDataRoaming": false,
    "cellularBlockGlobalBackgroundFetchWhileRoaming": false,
    "cellularBlockPerAppDataModification": false,
    "cellularBlockVoiceRoaming": false,
    "certificatesBlockUntrustedTlsCertificates": false,
    "classroomAppBlockRemoteScreenObservation": false,
    "compliantAppsList": [],
    "compliantAppListType": "none",
    "configurationProfileBlockChanges": false,
    "definitionLookupBlocked": false,
    "deviceBlockEnableRestrictions": false,
    "deviceBlockEraseContentAndSettings": false,
    "deviceBlockNameModification": false,
    "diagnosticDataBlockSubmission": false,
    "diagnosticDataBlockSubmissionModification": false,
    "documentsBlockManagedDocumentsInUnmanagedApps": false,
    "documentsBlockUnmanagedDocumentsInManagedApps": false,
    "emailInDomainSuffixes": [],
    "enterpriseAppBlockTrust": false,
    "enterpriseAppBlockTrustModification": false,
    "faceTimeBlocked": false,
    "findMyFriendsBlocked": false,
    "gamingBlockGameCenterFriends": true,
    "gamingBlockMultiplayer": false,
    "gameCenterBlocked": false,
    "hostPairingBlocked": false,
    "iBooksStoreBlocked": false,
    "iBooksStoreBlockErotica": false,
    "iCloudBlockActivityContinuation": false,
    "iCloudBlockBackup": true,
    "iCloudBlockDocumentSync": true,
    "iCloudBlockManagedAppsSync": false,
    "iCloudBlockPhotoLibrary": false,
    "iCloudBlockPhotoStreamSync": true,
    "iCloudBlockSharedPhotoStream": false,
    "iCloudRequireEncryptedBackup": false,
    "iTunesBlockExplicitContent": false,
    "iTunesBlockMusicService": false,
    "iTunesBlockRadio": false,
    "keyboardBlockAutoCorrect": false,
    "keyboardBlockPredictive": false,
    "keyboardBlockShortcuts": false,
    "keyboardBlockSpellCheck": false,
    "kioskModeAllowAssistiveSpeak": false,
    "kioskModeAllowAssistiveTouchSettings": false,
    "kioskModeAllowAutoLock": false,
    "kioskModeAllowColorInversionSettings": false,
    "kioskModeAllowRingerSwitch": false,
    "kioskModeAllowScreenRotation": false,
    "kioskModeAllowSleepButton": false,
    "kioskModeAllowTouchscreen": false,
    "kioskModeAllowVoiceOverSettings": false,
    "kioskModeAllowVolumeButtons": false,
    "kioskModeAllowZoomSettings": false,
    "kioskModeAppStoreUrl": null,
    "kioskModeRequireAssistiveTouch": false,
    "kioskModeRequireColorInversion": false,
    "kioskModeRequireMonoAudio": false,
    "kioskModeRequireVoiceOver": false,
    "kioskModeRequireZoom": false,
    "kioskModeManagedAppId": null,
    "lockScreenBlockControlCenter": false,
    "lockScreenBlockNotificationView": false,
    "lockScreenBlockPassbook": false,
    "lockScreenBlockTodayView": false,
    "mediaContentRatingAustralia": null,
    "mediaContentRatingCanada": null,
    "mediaContentRatingFrance": null,
    "mediaContentRatingGermany": null,
    "mediaContentRatingIreland": null,
    "mediaContentRatingJapan": null,
    "mediaContentRatingNewZealand": null,
    "mediaContentRatingUnitedKingdom": null,
    "mediaContentRatingUnitedStates": null,
    "mediaContentRatingApps": "allAllowed",
    "messagesBlocked": false,
    "notificationsBlockSettingsModification": false,
    "passcodeBlockFingerprintUnlock": false,
    "passcodeBlockModification": false,
    "passcodeBlockSimple": true,
    "passcodeExpirationDays": null,
    "passcodeMinimumLength": 4,
    "passcodeMinutesOfInactivityBeforeLock": null,
    "passcodeMinutesOfInactivityBeforeScreenTimeout": null,
    "passcodeMinimumCharacterSetCount": null,
    "passcodePreviousPasscodeBlockCount": null,
    "passcodeSignInFailureCountBeforeWipe": null,
    "passcodeRequiredType": "deviceDefault",
    "passcodeRequired": true,
    "podcastsBlocked": false,
    "safariBlockAutofill": false,
    "safariBlockJavaScript": false,
    "safariBlockPopups": false,
    "safariBlocked": false,
    "safariCookieSettings": "browserDefault",
    "safariManagedDomains": [],
    "safariPasswordAutoFillDomains": [],
    "safariRequireFraudWarning": false,
    "screenCaptureBlocked": false,
    "siriBlocked": false,
    "siriBlockedWhenLocked": false,
    "siriBlockUserGeneratedContent": false,
    "siriRequireProfanityFilter": false,
    "spotlightBlockInternetResults": false,
    "voiceDialingBlocked": false,
    "wallpaperBlockModification": false
}

"@

####################################################

$Android = @"

{
    "@odata.type": "#microsoft.graph.androidGeneralDeviceConfiguration",
    "description": "Android Device Restriction Policy",
    "displayName": "Android Device Restriction Policy",
    "appsBlockClipboardSharing": false,
    "appsBlockCopyPaste": false,
    "appsBlockYouTube": false,
    "bluetoothBlocked": false,
    "cameraBlocked": false,
    "cellularBlockDataRoaming": true,
    "cellularBlockMessaging": false,
    "cellularBlockVoiceRoaming": false,
    "cellularBlockWiFiTethering": false,
    "compliantAppsList": [],
    "compliantAppListType": "none",
    "diagnosticDataBlockSubmission": false,
    "locationServicesBlocked": false,
    "googleAccountBlockAutoSync": false,
    "googlePlayStoreBlocked": false,
    "kioskModeBlockSleepButton": false,
    "kioskModeBlockVolumeButtons": false,
    "kioskModeManagedAppId": null,
    "nfcBlocked": false,
    "passwordBlockFingerprintUnlock": true,
    "passwordBlockTrustAgents": false,
    "passwordExpirationDays": null,
    "passwordMinimumLength": 4,
    "passwordMinutesOfInactivityBeforeScreenTimeout": null,
    "passwordPreviousPasswordBlockCount": null,
    "passwordSignInFailureCountBeforeFactoryReset": null,
    "passwordRequiredType": "deviceDefault",
    "passwordRequired": true,
    "powerOffBlocked": false,
    "factoryResetBlocked": false,
    "screenCaptureBlocked": false,
    "deviceSharingBlocked": false,
    "storageBlockGoogleBackup": true,
    "storageBlockRemovableStorage": false,
    "storageRequireDeviceEncryption": true,
    "storageRequireRemovableStorageEncryption": true,
    "voiceAssistantBlocked": false,
    "voiceDialingBlocked": false,
    "webBrowserAllowPopups": false,
    "webBrowserBlockAutofill": false,
    "webBrowserBlockJavaScript": false,
    "webBrowserBlocked": false,
    "webBrowserCookieSettings": "browserDefault",
    "wiFiBlocked": false
}

"@

####################################################
if ($Environment -eq "OneDF")
{
  $serviceEndpoint = "https://graph.microsoft-ppe.com/testppebeta_intune_onedf"
}

if ($Environment -eq "Selfhost")
{
  $serviceEndpoint = "https://canary.graph.microsoft.com/testprodbeta_intune_sh"
}

if ($Environment -eq "CTIP")
{
  $serviceEndpoint = "https://graph.microsoft.com/beta"
}

if ($Environment -eq "PE")
{
  $serviceEndpoint = "https://graph.microsoft.com/beta"
}

Remove-ADMXFile $FileIdToRemove $serviceEndpoint

# SIG # Begin signature block
# MIIjkgYJKoZIhvcNAQcCoIIjgzCCI38CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDmPTm6sFr0JhIm
# QbRU5oic1uRPXAft/W1KoojuFyiUjqCCDYEwggX/MIID56ADAgECAhMzAAABh3IX
# chVZQMcJAAAAAAGHMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjAwMzA0MTgzOTQ3WhcNMjEwMzAzMTgzOTQ3WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQDOt8kLc7P3T7MKIhouYHewMFmnq8Ayu7FOhZCQabVwBp2VS4WyB2Qe4TQBT8aB
# znANDEPjHKNdPT8Xz5cNali6XHefS8i/WXtF0vSsP8NEv6mBHuA2p1fw2wB/F0dH
# sJ3GfZ5c0sPJjklsiYqPw59xJ54kM91IOgiO2OUzjNAljPibjCWfH7UzQ1TPHc4d
# weils8GEIrbBRb7IWwiObL12jWT4Yh71NQgvJ9Fn6+UhD9x2uk3dLj84vwt1NuFQ
# itKJxIV0fVsRNR3abQVOLqpDugbr0SzNL6o8xzOHL5OXiGGwg6ekiXA1/2XXY7yV
# Fc39tledDtZjSjNbex1zzwSXAgMBAAGjggF+MIIBejAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUhov4ZyO96axkJdMjpzu2zVXOJcsw
# UAYDVR0RBEkwR6RFMEMxKTAnBgNVBAsTIE1pY3Jvc29mdCBPcGVyYXRpb25zIFB1
# ZXJ0byBSaWNvMRYwFAYDVQQFEw0yMzAwMTIrNDU4Mzg1MB8GA1UdIwQYMBaAFEhu
# ZOVQBdOCqhc3NyK1bajKdQKVMFQGA1UdHwRNMEswSaBHoEWGQ2h0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY0NvZFNpZ1BDQTIwMTFfMjAxMS0w
# Ny0wOC5jcmwwYQYIKwYBBQUHAQEEVTBTMFEGCCsGAQUFBzAChkVodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY0NvZFNpZ1BDQTIwMTFfMjAx
# MS0wNy0wOC5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAgEAixmy
# S6E6vprWD9KFNIB9G5zyMuIjZAOuUJ1EK/Vlg6Fb3ZHXjjUwATKIcXbFuFC6Wr4K
# NrU4DY/sBVqmab5AC/je3bpUpjtxpEyqUqtPc30wEg/rO9vmKmqKoLPT37svc2NV
# BmGNl+85qO4fV/w7Cx7J0Bbqk19KcRNdjt6eKoTnTPHBHlVHQIHZpMxacbFOAkJr
# qAVkYZdz7ikNXTxV+GRb36tC4ByMNxE2DF7vFdvaiZP0CVZ5ByJ2gAhXMdK9+usx
# zVk913qKde1OAuWdv+rndqkAIm8fUlRnr4saSCg7cIbUwCCf116wUJ7EuJDg0vHe
# yhnCeHnBbyH3RZkHEi2ofmfgnFISJZDdMAeVZGVOh20Jp50XBzqokpPzeZ6zc1/g
# yILNyiVgE+RPkjnUQshd1f1PMgn3tns2Cz7bJiVUaqEO3n9qRFgy5JuLae6UweGf
# AeOo3dgLZxikKzYs3hDMaEtJq8IP71cX7QXe6lnMmXU/Hdfz2p897Zd+kU+vZvKI
# 3cwLfuVQgK2RZ2z+Kc3K3dRPz2rXycK5XCuRZmvGab/WbrZiC7wJQapgBodltMI5
# GMdFrBg9IeF7/rP4EqVQXeKtevTlZXjpuNhhjuR+2DMt/dWufjXpiW91bo3aH6Ea
# jOALXmoxgltCp1K7hrS6gmsvj94cLRf50QQ4U8Qwggd6MIIFYqADAgECAgphDpDS
# AAAAAAADMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0
# ZSBBdXRob3JpdHkgMjAxMTAeFw0xMTA3MDgyMDU5MDlaFw0yNjA3MDgyMTA5MDla
# MH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMT
# H01pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTEwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQCr8PpyEBwurdhuqoIQTTS68rZYIZ9CGypr6VpQqrgG
# OBoESbp/wwwe3TdrxhLYC/A4wpkGsMg51QEUMULTiQ15ZId+lGAkbK+eSZzpaF7S
# 35tTsgosw6/ZqSuuegmv15ZZymAaBelmdugyUiYSL+erCFDPs0S3XdjELgN1q2jz
# y23zOlyhFvRGuuA4ZKxuZDV4pqBjDy3TQJP4494HDdVceaVJKecNvqATd76UPe/7
# 4ytaEB9NViiienLgEjq3SV7Y7e1DkYPZe7J7hhvZPrGMXeiJT4Qa8qEvWeSQOy2u
# M1jFtz7+MtOzAz2xsq+SOH7SnYAs9U5WkSE1JcM5bmR/U7qcD60ZI4TL9LoDho33
# X/DQUr+MlIe8wCF0JV8YKLbMJyg4JZg5SjbPfLGSrhwjp6lm7GEfauEoSZ1fiOIl
# XdMhSz5SxLVXPyQD8NF6Wy/VI+NwXQ9RRnez+ADhvKwCgl/bwBWzvRvUVUvnOaEP
# 6SNJvBi4RHxF5MHDcnrgcuck379GmcXvwhxX24ON7E1JMKerjt/sW5+v/N2wZuLB
# l4F77dbtS+dJKacTKKanfWeA5opieF+yL4TXV5xcv3coKPHtbcMojyyPQDdPweGF
# RInECUzF1KVDL3SV9274eCBYLBNdYJWaPk8zhNqwiBfenk70lrC8RqBsmNLg1oiM
# CwIDAQABo4IB7TCCAekwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFEhuZOVQ
# BdOCqhc3NyK1bajKdQKVMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1Ud
# DwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFHItOgIxkEO5FAVO
# 4eqnxzHRI4k0MFoGA1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwubWljcm9zb2Z0
# LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcmwwXgYIKwYBBQUHAQEEUjBQME4GCCsGAQUFBzAChkJodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcnQwgZ8GA1UdIASBlzCBlDCBkQYJKwYBBAGCNy4DMIGDMD8GCCsGAQUFBwIB
# FjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2RvY3MvcHJpbWFyeWNw
# cy5odG0wQAYIKwYBBQUHAgIwNB4yIB0ATABlAGcAYQBsAF8AcABvAGwAaQBjAHkA
# XwBzAHQAYQB0AGUAbQBlAG4AdAAuIB0wDQYJKoZIhvcNAQELBQADggIBAGfyhqWY
# 4FR5Gi7T2HRnIpsLlhHhY5KZQpZ90nkMkMFlXy4sPvjDctFtg/6+P+gKyju/R6mj
# 82nbY78iNaWXXWWEkH2LRlBV2AySfNIaSxzzPEKLUtCw/WvjPgcuKZvmPRul1LUd
# d5Q54ulkyUQ9eHoj8xN9ppB0g430yyYCRirCihC7pKkFDJvtaPpoLpWgKj8qa1hJ
# Yx8JaW5amJbkg/TAj/NGK978O9C9Ne9uJa7lryft0N3zDq+ZKJeYTQ49C/IIidYf
# wzIY4vDFLc5bnrRJOQrGCsLGra7lstnbFYhRRVg4MnEnGn+x9Cf43iw6IGmYslmJ
# aG5vp7d0w0AFBqYBKig+gj8TTWYLwLNN9eGPfxxvFX1Fp3blQCplo8NdUmKGwx1j
# NpeG39rz+PIWoZon4c2ll9DuXWNB41sHnIc+BncG0QaxdR8UvmFhtfDcxhsEvt9B
# xw4o7t5lL+yX9qFcltgA1qFGvVnzl6UJS0gQmYAf0AApxbGbpT9Fdx41xtKiop96
# eiL6SJUfq/tHI4D1nvi/a7dLl+LrdXga7Oo3mXkYS//WsyNodeav+vyL6wuA6mk7
# r/ww7QRMjt/fdW1jkT3RnVZOT7+AVyKheBEyIXrvQQqxP/uozKRdwaGIm1dxVk5I
# RcBCyZt2WwqASGv9eZ/BvW1taslScxMNelDNMYIVZzCCFWMCAQEwgZUwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMQITMwAAAYdyF3IVWUDHCQAAAAABhzAN
# BglghkgBZQMEAgEFAKCBrjAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgbzqsKFVv
# fdHbQr/5CZewh++HBLeqSDCPQnPWisGvXbowQgYKKwYBBAGCNwIBDDE0MDKgFIAS
# AE0AaQBjAHIAbwBzAG8AZgB0oRqAGGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbTAN
# BgkqhkiG9w0BAQEFAASCAQBdPuJ68PoEpFlXbHXBgNmW9Ti1KVhnwG97Iiy6pC7W
# QRQJUWDpt9EJI9nGIS4SoN7lL4j4lVDYfe170TsFm/6qn1XpYLDEtlxlHpDw/BJ5
# GNH5wUA50hkCJLXZlQBMnSbQFf2T+i4uQchxdvd/KIp7V0vxMM4MKbqq0wFjJHuj
# vbw45s+7fI9BtHXOqqUJTUyLNgAjVlhYZljXcrhvQmqiMY/ncRUHIc3fuoIjuKUF
# bfhZ021VrAdV8JHdrOoGP1ncnrJ3JTKqDIgsnE9oiNkVaq56oA70FPIWtu9uu0Pt
# b0kT+p/ojaf8biQsrfnz1WUjw1NNdr3NcMk3papvQdMEoYIS8TCCEu0GCisGAQQB
# gjcDAwExghLdMIIS2QYJKoZIhvcNAQcCoIISyjCCEsYCAQMxDzANBglghkgBZQME
# AgEFADCCAVUGCyqGSIb3DQEJEAEEoIIBRASCAUAwggE8AgEBBgorBgEEAYRZCgMB
# MDEwDQYJYIZIAWUDBAIBBQAEIB0CBi9KKFw7v2/7FqYCHzqlheuihNnl8vNaytDw
# pS0OAgZfO+NB2pUYEzIwMjAwODIwMjI0NjEyLjU5NVowBIACAfSggdSkgdEwgc4x
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKTAnBgNVBAsTIE1p
# Y3Jvc29mdCBPcGVyYXRpb25zIFB1ZXJ0byBSaWNvMSYwJAYDVQQLEx1UaGFsZXMg
# VFNTIEVTTjo2MEJDLUUzODMtMjYzNTElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUt
# U3RhbXAgU2VydmljZaCCDkQwggT1MIID3aADAgECAhMzAAABJt+6SyK5goIHAAAA
# AAEmMA0GCSqGSIb3DQEBCwUAMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNo
# aW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEw
# MB4XDTE5MTIxOTAxMTQ1OVoXDTIxMDMxNzAxMTQ1OVowgc4xCzAJBgNVBAYTAlVT
# MRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQK
# ExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKTAnBgNVBAsTIE1pY3Jvc29mdCBPcGVy
# YXRpb25zIFB1ZXJ0byBSaWNvMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjo2MEJD
# LUUzODMtMjYzNTElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vydmlj
# ZTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAJ4wvoacTvMNlXQTtfF/
# Cx5Ol3X0fcjUNMvjLgTmO5+WHYJFbp725P3+qvFKDRQHWEI1Sz0gB24urVDIjXjB
# h5NVNJVMQJI2tltv7M4/4IbhZJb3xzQW7LolEoZYUZanBTUuyly9osCg4o5joViT
# 2GtmyxK+Fv5kC20l2opeaeptd/E7ceDAFRM87hiNCsK/KHyC+8+swnlg4gTOey6z
# QqhzgNsG6HrjLBuDtDs9izAMwS2yWT0T52QA9h3Q+B1C9ps2fMKMe+DHpG+0c61D
# 94Yh6cV2XHib4SBCnwIFZAeZE2UJ4qPANSYozI8PH+E5rCT3SVqYvHou97HsXvP2
# I3MCAwEAAaOCARswggEXMB0GA1UdDgQWBBRJq6wfF7B+mEKN0VimX8ajNA5hQTAf
# BgNVHSMEGDAWgBTVYzpcijGQ80N7fEYbxTNoWoVtVTBWBgNVHR8ETzBNMEugSaBH
# hkVodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNU
# aW1TdGFQQ0FfMjAxMC0wNy0wMS5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUF
# BzAChj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1RpbVN0
# YVBDQV8yMDEwLTA3LTAxLmNydDAMBgNVHRMBAf8EAjAAMBMGA1UdJQQMMAoGCCsG
# AQUFBwMIMA0GCSqGSIb3DQEBCwUAA4IBAQBAlvudaOlv9Cfzv56bnX41czF6tLtH
# LB46l6XUch+qNN45ZmOTFwLot3JjwSrn4oycQ9qTET1TFDYd1QND0LiXmKz9OqBX
# ai6S8XdyCQEZvfL82jIAs9pwsAQ6XvV9jNybPStRgF/sOAM/Deyfmej9Tg9FcRwX
# ank2qgzdZZNb8GoEze7f1orcTF0Q89IUXWIlmwEwQFYF1wjn87N4ZxL9Z/xA2m/R
# 1zizFylWP/mpamCnVfZZLkafFLNUNVmcvc+9gM7vceJs37d3ydabk4wR6ObR34sW
# aLppmyPlsI1Qq5Lu6bJCWoXzYuWpkoK6oEep1gML6SRC3HKVS3UscZhtMIIGcTCC
# BFmgAwIBAgIKYQmBKgAAAAAAAjANBgkqhkiG9w0BAQsFADCBiDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJv
# b3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5IDIwMTAwHhcNMTAwNzAxMjEzNjU1WhcN
# MjUwNzAxMjE0NjU1WjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDCCASIw
# DQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAKkdDbx3EYo6IOz8E5f1+n9plGt0
# VBDVpQoAgoX77XxoSyxfxcPlYcJ2tz5mK1vwFVMnBDEfQRsalR3OCROOfGEwWbEw
# RA/xYIiEVEMM1024OAizQt2TrNZzMFcmgqNFDdDq9UeBzb8kYDJYYEbyWEeGMoQe
# dGFnkV+BVLHPk0ySwcSmXdFhE24oxhr5hoC732H8RsEnHSRnEnIaIYqvS2SJUGKx
# Xf13Hz3wV3WsvYpCTUBR0Q+cBj5nf/VmwAOWRH7v0Ev9buWayrGo8noqCjHw2k4G
# kbaICDXoeByw6ZnNPOcvRLqn9NxkvaQBwSAJk3jN/LzAyURdXhacAQVPIk0CAwEA
# AaOCAeYwggHiMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBTVYzpcijGQ80N7
# fEYbxTNoWoVtVTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMC
# AYYwDwYDVR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBTV9lbLj+iiXGJo0T2UkFvX
# zpoYxDBWBgNVHR8ETzBNMEugSaBHhkVodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20v
# cGtpL2NybC9wcm9kdWN0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcmwwWgYI
# KwYBBQUHAQEETjBMMEoGCCsGAQUFBzAChj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNydDCBoAYDVR0g
# AQH/BIGVMIGSMIGPBgkrBgEEAYI3LgMwgYEwPQYIKwYBBQUHAgEWMWh0dHA6Ly93
# d3cubWljcm9zb2Z0LmNvbS9QS0kvZG9jcy9DUFMvZGVmYXVsdC5odG0wQAYIKwYB
# BQUHAgIwNB4yIB0ATABlAGcAYQBsAF8AUABvAGwAaQBjAHkAXwBTAHQAYQB0AGUA
# bQBlAG4AdAAuIB0wDQYJKoZIhvcNAQELBQADggIBAAfmiFEN4sbgmD+BcQM9naOh
# IW+z66bM9TG+zwXiqf76V20ZMLPCxWbJat/15/B4vceoniXj+bzta1RXCCtRgkQS
# +7lTjMz0YBKKdsxAQEGb3FwX/1z5Xhc1mCRWS3TvQhDIr79/xn/yN31aPxzymXlK
# kVIArzgPF/UveYFl2am1a+THzvbKegBvSzBEJCI8z+0DpZaPWSm8tv0E4XCfMkon
# /VWvL/625Y4zu2JfmttXQOnxzplmkIz/amJ/3cVKC5Em4jnsGUpxY517IW3DnKOi
# PPp/fZZqkHimbdLhnPkd/DjYlPTGpQqWhqS9nhquBEKDuLWAmyI4ILUl5WTs9/S/
# fmNZJQ96LjlXdqJxqgaKD4kWumGnEcua2A5HmoDF0M2n0O99g/DhO3EJ3110mCII
# YdqwUB5vvfHhAN/nMQekkzr3ZUd46PioSKv33nJ+YWtvd6mBy6cJrDm77MbL2IK0
# cs0d9LiFAR6A+xuJKlQ5slvayA1VmXqHczsI5pgt6o3gMy4SKfXAL1QnIffIrE7a
# KLixqduWsqdCosnPGUFN4Ib5KpqjEWYw07t0MkvfY3v1mYovG8chr1m1rtxEPJdQ
# cdeh0sVV42neV8HR3jDA/czmTfsNv11P6Z0eGTgvvM9YBS7vDaBQNdrvCScc1bN+
# NR4Iuto229Nfj950iEkSoYIC0jCCAjsCAQEwgfyhgdSkgdEwgc4xCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKTAnBgNVBAsTIE1pY3Jvc29mdCBP
# cGVyYXRpb25zIFB1ZXJ0byBSaWNvMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjo2
# MEJDLUUzODMtMjYzNTElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vy
# dmljZaIjCgEBMAcGBSsOAwIaAxUACmcyOWmZxErpq06B8dy6oMZ6//yggYMwgYCk
# fjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQD
# Ex1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDANBgkqhkiG9w0BAQUFAAIF
# AOLpBKswIhgPMjAyMDA4MjAxODE4MTlaGA8yMDIwMDgyMTE4MTgxOVowdzA9Bgor
# BgEEAYRZCgQBMS8wLTAKAgUA4ukEqwIBADAKAgEAAgIjHwIB/zAHAgEAAgIQ+jAK
# AgUA4upWKwIBADA2BgorBgEEAYRZCgQCMSgwJjAMBgorBgEEAYRZCgMCoAowCAIB
# AAIDB6EgoQowCAIBAAIDAYagMA0GCSqGSIb3DQEBBQUAA4GBAFxKKtpuiooafnQ0
# C3M4H28h/BQ5/iSqvLheSfiZggdq8VbltyEgRGdJ8is/jhmW1+Z3+V0ZFysO8gu2
# 2s54xp32IkHrYhxFyze3+mjHo5h47u9cXHU63spqB9aRy51gWQ5Xuab6uHr4XHnY
# TpAS7/7HJkFvB620lz9HTCOyYLhQMYIDDTCCAwkCAQEwgZMwfDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBIDIwMTACEzMAAAEm37pLIrmCggcAAAAAASYwDQYJYIZIAWUD
# BAIBBQCgggFKMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG9w0B
# CQQxIgQgS65PG61Om02frwgs8eqEcWdPvuGGGkGli6p7tj/oAmAwgfoGCyqGSIb3
# DQEJEAIvMYHqMIHnMIHkMIG9BCA2/c/vnr1ecAzvapOWZ2xGfAkzrkfpGcrvMW07
# CQl1DzCBmDCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
# MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
# b24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAAB
# Jt+6SyK5goIHAAAAAAEmMCIEIK/yd0w/VaMqcx14hqgQtZZ9z+NEVqnzHcZfBZiY
# 15hfMA0GCSqGSIb3DQEBCwUABIIBAANPK7iHf3VWh4iCqC3bVLtwwhj1JXr4v69D
# 3MxPmTlSwmPTGtnZ/7HFj/vRXzY0baiATzJnDCp6nnaVoEyRyHKwk/W6RFSGmpzx
# 4feygdfX0dwWch+uOttt6XBhQ8zUvLY99ie1xt4+WmjjIwOh+t1mYXQAPj+Zka2h
# K6jIrHFRYYlWPgsLaX6dJghN30dOlMFqlc+iwPSjUHnRdodaEt/hqTT3UkvB6p21
# 60q4ey06v6ejp5mo6JboywmmkqT80mY1B0itnw8R4O+d2v22QBHsjU609bfxjzVc
# stw3Q5O4GRPXG9u02kv5b362iJ+tUtZ693sUEGqGyT6/rS+RMag=
# SIG # End signature block
