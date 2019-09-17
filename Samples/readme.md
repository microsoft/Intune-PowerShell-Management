# Intune PowerShell Module

The Intune PowerShell Module enables access to Intune programmatically for your tenant, the Module performs the same Intune operations as those available through the Azure Portal.  

Intune provides data into the Microsoft Graph in the same way as other cloud services do, with rich entity information and relationship navigation.

* Intune PowerShell Module: https://www.powershellgallery.com/packages/Microsoft.Graph.Intune
* Intune PowerShell SDK (GitHub): https://github.com/microsoft/Intune-PowerShell-SDK

### Notable features

Standard PowerShell objects are used for input/output, meaning that all built-in PowerShell functionality work, including:
  - Piping of objects between cmdlets
  - Formatting of output: `Format-Table`, `Out-GridView`, `ConvertTo-Csv`, `ConvertTo-Json`, etc.
  - Getting help on usage: `Get-Help`
  - Visualizing input parameters: `Show-Command`
  - Using the 'tab' key to auto-complete or cycle through available options
  - Auto-complete and validation on Enum parameters as well as some query parmeters (i.e. $select, $expand and $orderBy)

- Utility cmdlets for some common tasks
  - Getting the authentication token: `Connect-MSGraph`
  - Getting service metadata: `Get-MSGraphMetadata`
  - Paging: `Get-MSGraphNextPage` and `Get-MSGraphAllPages`
  - Changing environment settings, e.g. Graph schema version: `Update-MSGraphEnvironment -Schema beta -AppId 00000000-0000-0000-0000-000000000000`
  - Make arbitrary Graph calls using the `Invoke-MSGraphRequest` cmdlet

#### Disclaimer
Some script samples retrieve information from your Intune tenant, and others create, delete or update data in your Intune tenant.  Understand the impact of each sample script prior to running it; samples should be run using a non-production or "test" tenant account.

## Getting started
### One-time setup (PowerShell Gallery)
Install the Microsoft.Graph.Intune module from: https://www.powershellgallery.com/packages/Microsoft.Graph.Intune
```PowerShell
Install-Module -Name Microsoft.Graph.Intune
```
### One-time setup (GitHub)
Download the module from the [Releases](https://github.com/Microsoft/Intune-PowerShell-SDK/releases) tab in the GitHub repository
* The "drop\outputs\build\Release\net471" folder in the zip file contains the module.
* If you are using Windows, extract the "net471" folder.  **You must have .NET 4.7.1 or higher installed**.
* The module manifest is the "Microsoft.Graph.Intune.psd1" file inside this folder.  This is the file you would refer to when importing the module.
* Import the module:
```PowerShell
Import-Module $sdkDir/Microsoft.Graph.Intune.psd1
```

### Each time you use the module
To authenticate with Microsoft Graph (this is not required when using CloudShell):

```PowerShell
Connect-MSGraph
```

To authenticate with Microsoft Graph using [System.Management.Automation.PSCredential]

```PowerShell
$adminUPN=Read-Host -Prompt "Enter UPN"
$adminPwd=Read-Host -AsSecureString -Prompt "Enter password for $adminUPN"
$creds = New-Object System.Management.Automation.PSCredential ($AdminUPN, $adminPwd)
$connection = Connect-MSGraph -PSCredential $creds
```

To authenticate in a non-standard environment:

```PowerShell
# 1. Setup the environment
# For example, in a National Cloud environment, the following is required before logging in
Update-MSGraphEnvironment -AuthUrl 'https://login.microsoftonline.us/common' -GraphBaseUrl 'https://graph.microsoft.us' -GraphResourceId 'https://graph.microsoft.us' -SchemaVersion 'beta'
```

### Discovering available commands

Get the full list of available cmdlets:

```PowerShell
Get-Command -Module Microsoft.Graph.Intune
```

Get documentation on a particular cmdlet:

```PowerShell
Get-Help <cmdlet name>
```

Use a UI to see the parameter sets more easily:

```PowerShell
Show-Command <cmdlet name>
```

## Basics

### List Objects

Get all Intune applications:

```PowerShell
Get-IntuneMobileApp
```

### Filter objects

Use -Select to restrict properties to display:

```PowerShell
Get-IntuneMobileApp -Select displayName, publisher
```

Use -Filter to filter results:

```PowerShell
Get-IntuneMobileApp -Select displayName, publisher -Filter "contains(publisher, 'Microsoft')"
```

### Bulk create objects

Bulk create multiple webApp objects (they should appear in the Azure Portal)

```PowerShell
$createdApps = 'https://www.bing.com', 'https://developer.microsoft.com/graph', 'https://portal.azure.com' `
| ForEach-Object { `
    New-IntuneMobileApp `
        -webApp `
        -displayName $_ `
        -publisher 'IT Professional' `
        -appUrl $_ `
        -useManagedBrowser $false `
}
```

Display using GridView:

```PowerShell
1..15 | ForEach-Object { `
    New-IntuneMobileApp `
        -webApp `
        -displayName "Bing #$_" `
        -publisher 'Microsoft' `
        -appUrl 'https://www.bing.com' `
        -useManagedBrowser ([bool]($_ % 2)) `
} | Out-GridView
```

Remove all webApps:
```PowerShell
# Remove all web apps
$appsToDelete = Get-IntuneMobileApp -Filter "isof('microsoft.graph.webApp')"
$appsToDelete | Remove-IntuneMobileApp
```

### Paging
Show paging of audit events (run this in a different window).
```PowerShell
# Audit events are accessible from the beta schema
Update-MSGraphEnvironment -SchemaVersion 'beta'
Connect-MSGraph

# Make the call to get audit events
$auditEvents = Invoke-MSGraphRequest -HttpMethod GET -Url 'deviceManagement/auditEvents'
$auditEvents # more than 1000 results, so they are wrapped in an object with the nextLink
$auditEvents.value | measure

# We can get the next page
$auditEvents2 = $auditEvents | Get-MSGraphNextPage
$auditEvents.value | measure # have to unwrap the results again

# Get all pages of audit events
$auditEvents = Invoke-MSGraphRequest -HttpMethod GET -Url 'deviceManagement/auditEvents' | Get-MSGraphAllPages

# Switch back to v1.0
Update-MSGraphEnvironment -SchemaVersion 'v1.0'
```
### Getting Extended Debug information
If for some reason, a cmdlet fails. Use Get-MSGraphInfo to get extended information.
A sample failure is listed below:
```PowerShell
# Call that failed
Invoke-IntuneDeviceCompliancePolicyAssign : 500 Internal Server Error
{
  "error": {
    "code": "InternalError",
    "message": "{\r\n  \"_version\": 3,\r\n  \"Message\": \"An internal server error has occurred - Operation ID (for customer support): 00000000-0000-0000-0000-000000000000 - Activity ID: a02e4ad2-efdb-4ae0-8b36-7c990a228f21 -
Url: https://fef.msua06.manage.microsoft.com/StatelessDeviceConfigurationFEService/deviceManagement/deviceCompliancePolicies%28%27bc4c48a9-4120-4531-8870-f57767d43da4%27%29/microsoft.management.services.api.assign?api-version=2018
-06-29\",\r\n  \"CustomApiErrorPhrase\": \"\",\r\n  \"RetryAfter\": null,\r\n  \"ErrorSourceService\": \"\",\r\n  \"HttpHeaders\": \"{}\"\r\n}",
    "innerError": {
      "request-id": "a02e4ad2-efdb-4ae0-8b36-7c990a228f21",
      "date": "2018-11-28T21:44:56"
    }
  }
}
At line:1 char:1
+ Invoke-IntuneDeviceCompliancePolicyAssign   -deviceCompliancePolicyId ...
+ ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    + CategoryInfo          : ConnectionError: (@{Request=; Response=}:PSObject) [Invoke-IntuneDe...ncePolicyAssign], HttpRequestException
    + FullyQualifiedErrorId : PowerShellGraphSDK_HttpRequestError,Microsoft.Intune.PowerShellGraphSDK.PowerShellCmdlets.Invoke_IntuneDeviceCompliancePolicyAssign

# Get Debug information
Get-MSGraphDebugInfo

Request
-------
@{HttpMethod=POST; URL=https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies/bc4c48a9-4120-4531-8870-f57767d43da4/assign; Headers=; Content={...

# Look into the Request
(Get-MSGraphDebugInfo).Request

HttpMethod URL                                                                                                                    Headers
---------- ---                                                                                                                    -------
POST       https://graph.microsoft.com/v1.0/deviceManagement/deviceCompliancePolicies/bc4c48a9-4120-4531-8870-f57767d43da4/assign @{Authorization=Bearer eyJ0eXAiOiJKV1QiLCJub25jZSI6IkFRQUJBQUFBQUFDNXVuYTBFVUZnVElGOEVsYXh0V2pUam...

# Look into the Response
(Get-MSGraphDebugInfo).Response

HttpStatusCode HttpStatusPhrase      Headers
-------------- ----------------      -------
           500 Internal Server Error @{Transfer-Encoding=chunked; request-id=a02e4ad2-efdb-4ae0-8b36-7c990a228f21; client-request-id=a02e4ad2-efdb-4ae0-8b36-7c990a228f21; x-ms-ags-diagnostic={"ServerInfo":{"DataCenter":"West Ce...

# Inspect the Response headers
(Get-MSGraphDebugInfo).Response.Headers

Transfer-Encoding         : chunked
request-id                : a02e4ad2-efdb-4ae0-8b36-7c990a228f21
client-request-id         : a02e4ad2-efdb-4ae0-8b36-7c990a228f21
x-ms-ags-diagnostic       : {"ServerInfo":{"DataCenter":"West Central US","Slice":"SliceC","Ring":"1","ScaleUnit":"001","Host":"AGSFE_IN_4","ADSiteName":"WCU"}}
Duration                  : 496.4757
Strict-Transport-Security : max-age=31536000
Cache-Control             : private
Date                      : Wed, 28 Nov 2018 21:44:55 GMT

```

## Scenario Samples
### Create Compliance Policies and Assign it to an AAD Group
```PowerShell
# Search the AAD Group
$AADGroupId = (Get-Groups -Filter "displayName eq 'Intune POC Users'").id
```

### Create an iOS Compliance Policy
```PowerShell
$iOSCompliancePolicy = New-IntuneDeviceCompliancePolicy `
    -iosCompliancePolicy `
    -displayName "Chicago - iOS Compliance Policy" `
    -passcodeRequired $true `
    -passcodeMinimumLength 6 `
    -passcodeMinutesOfInactivityBeforeLock 15 `
    -securityBlockJailbrokenDevices $true `
    -scheduledActionsForRule `
        (New-DeviceComplianceScheduledActionForRuleObject -ruleName PasswordRequired `
            -scheduledActionConfigurations `
                (New-DeviceComplianceActionItemObject -gracePeriodHours 0 `
                -actionType block `
                -notificationTemplateId "" `
                )`
        )

# Assign the newly created compliance policy to the AAD Group
Invoke-IntuneDeviceCompliancePolicyAssign  -deviceCompliancePolicyId $iOSCompliancePolicy.id `
    -assignments `
        (New-DeviceCompliancePolicyAssignmentObject `
            -target `
                (New-DeviceAndAppManagementAssignmentTargetObject `
                    -groupAssignmentTarget `
                    -groupId "$AADGroupId" `
                ) `
        )
```

### Create Android Compliance Policy
```PowerShell
$androidCompliancePolicy = New-IntuneDeviceCompliancePolicy `
    -androidCompliancePolicy `
    -displayName "Chicago - Android Compliance Policy"  `
    -passwordRequired $true `
    -passwordMinimumLength 6 `
    -securityBlockJailbrokenDevices $true `
    -passwordMinutesOfInactivityBeforeLock 15 `
    -scheduledActionsForRule `
    (New-DeviceComplianceScheduledActionForRuleObject `
        -ruleName PasswordRequired `
        -scheduledActionConfigurations `
        (New-DeviceComplianceActionItemObject `
            -gracePeriodHours 0 `
            -actionType block `
            -notificationTemplateId "" `
        )`
    )

# Assign the newly created compliance policy to the AAD Group
Invoke-IntuneDeviceCompliancePolicyAssign -deviceCompliancePolicyId $androidCompliancePolicy.id `
    -assignments `
    (New-DeviceCompliancePolicyAssignmentObject `
        -target `
        (New-DeviceAndAppManagementAssignmentTargetObject `
            -groupAssignmentTarget `
            -groupId "$AADGroupId" `
        ) `
    )
```

### Create Windows 10 Compliance Policy
```PowerShell
$windows10CompliancePolicy = New-IntuneDeviceCompliancePolicy `
    -windows10CompliancePolicy `
    -displayName "Chicago - Windows 10 Compliance Policy" `
    -osMinimumVersion 10.0.16299 `
    -scheduledActionsForRule `
    (New-DeviceComplianceScheduledActionForRuleObject `
        -ruleName PasswordRequired `
        -scheduledActionConfigurations `
        (New-DeviceComplianceActionItemObject `
            -gracePeriodHours 0 `
            -actionType block `
            -notificationTemplateId "" `
        ) `
    )

# Assign the newly created compliance policy to the AAD Group
Invoke-IntuneDeviceCompliancePolicyAssign -deviceCompliancePolicyId $windows10CompliancePolicy.id `
    -assignments `
        (New-DeviceCompliancePolicyAssignmentObject `
            -target `
            (New-DeviceAndAppManagementAssignmentTargetObject `
                -groupAssignmentTarget `
                -groupId "$AADGroupId" `
            ) `
        )
```

### Create MacOS Compliance Policy
```PowerShell
$macOSCompliancePolicy = New-IntuneDeviceCompliancePolicy `
    -macOSCompliancePolicy `
    -displayName "Chicago - MacOS Compliance Policy" `
    -passwordRequired $true `
    -passwordBlockSimple $false `
    -passwordRequiredType deviceDefault `
    -scheduledActionsForRule `
    (New-DeviceComplianceScheduledActionForRuleObject `
        -ruleName PasswordRequired `
        -scheduledActionConfigurations `
        (New-DeviceComplianceActionItemObject `
            -gracePeriodHours 0 `
            -actionType block `
            -notificationTemplateId "" `
        ) `
    )

# Assign the newly created compliance policy to the AAD Group
Invoke-IntuneDeviceCompliancePolicyAssign -deviceCompliancePolicyId $macOSCompliancePolicy.id `
    -assignments `
    (New-DeviceCompliancePolicyAssignmentObject `
    -target `
        (New-DeviceAndAppManagementAssignmentTargetObject `
            -groupAssignmentTarget `
            -groupId "$AADGroupId" `
        )`
    )
```
## Create Configuration Policies and Assign it to an AAD Group
```PowerShell
# Search the AAD Group
$AADGroupId = (Get-Groups -Filter "displayName eq 'Intune POC Users'").id
```

### Create iOS Restriction Policy
```PowerShell
$iosGeneralDeviceConfiguration = New-IntuneDeviceConfigurationPolicy `
    -iosGeneralDeviceConfiguration `
    -displayName "Chicago - iOS Device Restriction Policy" `
    -iCloudBlockBackup $true `
    -iCloudBlockDocumentSync $true `
    -iCloudBlockPhotoStreamSync $true

# Assign the newly created configuration policy to the AAD Group
Invoke-IntuneDeviceConfigurationPolicyAssign -deviceConfigurationId $iosGeneralDeviceConfiguration.id `
    -assignments `
    (New-DeviceConfigurationAssignmentObject `
        -target `
        (New-DeviceAndAppManagementAssignmentTargetObject `
            -groupAssignmentTarget `
            -groupId "$AADGroupId" `
        ) `
    )
```

### Create Android Restriction Policy
```PowerShell
$androidGeneralDeviceConfiguration = New-IntuneDeviceConfigurationPolicy `
    -androidGeneralDeviceConfiguration `
    -displayName "Chicago - Android Device Restriction Policy" `
    -passwordRequired $true `
    -passwordRequiredType deviceDefault `
    -passwordMinimumLength 4

# Assign the newly created configuration policy to the AAD Group
Invoke-IntuneDeviceConfigurationPolicyAssign -deviceConfigurationId $androidGeneralDeviceConfiguration.id `
    -assignments `
        (New-DeviceConfigurationAssignmentObject `
        -target `
            (New-DeviceAndAppManagementAssignmentTargetObject `
                -groupAssignmentTarget -groupId "$AADGroupId" `
            ) `
        )

```
## Create App Protection Polies and assign it to an AAD Group
### iOS App Protection Policy Creation
```PowerShell
# Get the list of iOS managed mobileapp objects
$appsiOS = @()
$iosManagedAppProtectionApps = Get-IntuneMobileApp | ? { $_.appAvailability -eq "global" -and ($_.'@odata.type').contains("managedIOS") }
foreach($app in $iosManagedAppProtectionApps)
{
    $bundleId = $app.bundleId
    $appsiOS += (New-ManagedMobileAppObject -mobileAppIdentifier (New-MobileAppIdentifierObject -iosMobileAppIdentifier -bundleId "$bundleId"))
}

# Create the ios App Protection Policy
$iosManagedAppProtection = New-IntuneAppProtectionPolicy `
    -iosManagedAppProtection `
    -displayName "iOS MAM / APP Policy" `
    -periodOfflineBeforeAccessCheck (New-TimeSpan -Hours 12) `
    -periodOnlineBeforeAccessCheck (New-TimeSpan -Minutes 30)`
    -allowedInboundDataTransferSources managedApps `
    -allowedOutboundDataTransferDestinations managedApps `
    -allowedOutboundClipboardSharingLevel managedAppsWithPasteIn `
    -organizationalCredentialsRequired $false `
    -dataBackupBlocked $true `
    -managedBrowserToOpenLinksRequired $false `
    -deviceComplianceRequired $false `
    -saveAsBlocked $true `
    -periodOfflineBeforeWipeIsEnforced (New-TimeSpan -Days 30) `
    -pinRequired $true `
    -maximumPinRetries 5 `
    -simplePinBlocked $false `
    -minimumPinLength 4 `
    -pinCharacterSet numeric `
    -periodBeforePinReset (New-TimeSpan -Days 30) `
    -allowedDataStorageLocations @("oneDriveForBusiness","sharePoint") `
    -contactSyncBlocked $false `
    -printBlocked $true `
    -fingerprintBlocked $false `
    -disableAppPinIfDevicePinIsSet $false `
    -apps $appsiOS

# Assign ios App Protection Policy to the AAD Group
Invoke-IntuneAppProtectionPolicyIosAssign -iosManagedAppProtectionId $iosManagedAppProtection.id `
    -assignments `
    (New-TargetedManagedAppPolicyAssignmentObject `
            -target `
            (New-DeviceAndAppManagementAssignmentTargetObject `
            -groupAssignmentTarget -groupId "$AADGroupId" `
            ) `
    )
```
### Android App Protection Policy Creation
```PowerShell

# Get the list of Android managed mobileapp objects
$appsAndroid = @()
$AndroidAPPapps = Get-IntuneMobileApp | ? { $_.appAvailability -eq "global" -and ($_.'@odata.type').contains("managedAndroid") }
foreach($app in $AndroidAPPapps)
{
    $PackageId = $app.packageId
    $appsAndroid += (New-ManagedMobileAppObject -mobileAppIdentifier (New-MobileAppIdentifierObject -androidMobileAppIdentifier -packageId "$PackageId"))
}

# Create the Android App Protection Policy
$androidManagedAppProtectionPolicy = New-IntuneAppProtectionPolicy `
    -androidManagedAppProtection -displayName "Android MAM / APP Policy" `
    -periodOfflineBeforeAccessCheck (New-TimeSpan -Hours 12) `
    -periodOnlineBeforeAccessCheck (New-TimeSpan -Minutes 30)`
    -allowedInboundDataTransferSources managedApps `
    -allowedOutboundDataTransferDestinations managedApps `
    -allowedOutboundClipboardSharingLevel managedAppsWithPasteIn `
    -organizationalCredentialsRequired $false `
    -dataBackupBlocked $true `
    -managedBrowserToOpenLinksRequired $false `
    -deviceComplianceRequired $false `
    -saveAsBlocked $true `
    -periodOfflineBeforeWipeIsEnforced (New-TimeSpan -Days 30) `
    -pinRequired $true `
    -maximumPinRetries 5 `
    -simplePinBlocked $false `
    -minimumPinLength 4 `
    -pinCharacterSet numeric `
    -periodBeforePinReset (New-TimeSpan -Days 30) `
    -allowedDataStorageLocations @("oneDriveForBusiness","sharePoint") `
    -contactSyncBlocked $false `
    -printBlocked $true `
    -disableAppPinIfDevicePinIsSet $false `
    -screenCaptureBlocked $true `
    -apps $appsAndroid

# Assign Android App Protection Policy to the AAD Group
Invoke-IntuneAppProtectionPolicyAndroidAssign -androidManagedAppProtectionId $androidManagedAppProtectionPolicy.id `
    -assignments `
    (New-TargetedManagedAppPolicyAssignmentObject `
            -target `
            (New-DeviceAndAppManagementAssignmentTargetObject `
            -groupAssignmentTarget -groupId "$AADGroupId" `
            ) `
    )
```
