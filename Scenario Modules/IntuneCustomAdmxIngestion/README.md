
# Overview 
Microsoft Endpoint Manager currently has the capability to managed ADMX settings for Edge, Office, and a subset of MDM relevant Windows features. We have built, based on customer feedback, the ability for an admin to upload their own ADMX packages. 

With this feature, not only are the setting configurations deployed to targeted devices, but the ADMX and ADML files as well. 

This allows an IT admin to lay down app configuration before app install as the ADMX is installed beforehand. The uploaded files will also appear in the admin console with easy to use configuration.


# Notes for the preview 
The following should be considered for the preview and are subject to change in future iterations.
1.  There is a maximum of 10 packages that can be uploaded
2.  There is a file limit of 1mb
3.  Uploading the ADMX that already exists in the admin console will result in duplicate settings in the MEM admin console.
4.  The upload is available in graph and will be accessible without UX for the preview timeframe.
5.  Removing an ADMX package will require removing any settings configured in profiles before removal.
6.  An ADMX package may have pre-requisites that must be uploaded first. For example, chrome.admx requires first google.admx and windows.admx.
7.  Only 1 ADML file can be uploaded per ADMX file. This means each package will support 1 language for now.
8.  AMDX files with unsupported setting types will fail at upload. For example, combo box settings.

# Procedure for Uploading custom ADMX templates
This procedure will walk you through uploading an ADMX/ADML package, listing the uploaded packages, and removing a package.
## Download the ADMX Template from the web (download locations below for popular templates)
    * Windows: https://www.microsoft.com/en-us/download/101445 (this is mandatory irrespective of the 3rd party app that needs to be managed)
        
        Download the installer and install to a local folder e.g. C:\temp\Windows-PolicyTemplates

    * Chrome ADM/ADMX Templates 71.1 MB Download: https://chromeenterprise.google/browser/download/ 
    
        Download and unzip it to local folder e.g. C:\temp\Chrome-PolicyTemplates

## Download ADMX Publishing PowerShell scripts from Github 
    * Download the Github repo zip file: https://github.com/microsoft/Intune-PowerShell-Management/archive/ADMXCustomIngestion.zip 
        * extract it to local folder e.g. C:\temp\Github
        * The ADMX Publishing scripts will be here: 
            
            C:\temp\Github\Intune-PowerShell-Management-ADMXCustomIngestion\Scenario Modules\IntuneCustomAdmxIngestion

## Publish Windows ADMX and ADML files

```PowerShell
cd "C:\temp\Github\Intune-PowerShell-Management-ADMXCustomIngestion\Scenario Modules\IntuneCustomAdmxIngestion"

PowerShell .\upload-admxfile.ps1 -ADMXFilePath "C:\temp\Windows-PolicyTemplates\windows.admx" -ADMLFilePath "C:\temp\Windows-PolicyTemplates\en-US\windows.adml"
```

## Publish 3rd party container ADMX and ADML files
For Example, for Google Apps:
```PowerShell
cd "C:\temp\Github\Intune-PowerShell-Management-ADMXCustomIngestion\Scenario Modules\IntuneCustomAdmxIngestion"

PowerShell .\upload-admxfile.ps1 -ADMXFilePath "C:\temp\Chrome-PolicyTemplates\windows\admx\google.admx" -ADMLFilePath "C:\temp\Chrome-PolicyTemplates\windows\admx\en-US\google.adml"
```
    
* Enter the Intune Tenant Administrator's UPN
* Please grant access to "Microsoft Intune PowerShell" if prompted.

     Note: If you see the following error, you need to work with Intune Engineering to flight your tenant:
    ``` Code
    "error": {
    "code": "BadRequest",
    "message": "{\r\n  \"_version\": 3,\r\n  \"Message\": \"Feature disabled. - Operation ID (for customer support): 00000000-0000-0000-0000-000000000000   CustomApiErrorPhrase: FeatureDisabled\",\r\n  \"CustomApiErrorPhrase\": \"FeatureDisabled\",\r\n  \"RetryAfter\": null,\r\n  \"ErrorSourceService\": \"\",\r\n  \"HttpHeaders\": \"{}\"\r\n}"
    }
    ```
* On Success, you should see the following message:
    ``` Code
    @odata.context                   : https://graph.microsoft.com/beta/$metadata#deviceManagement/groupPolicyUploadedDefin
                                   itionFiles/$entity
    displayName                      :
    description                      :
    languageCodes                    : {}
    targetPrefix                     :
    targetNamespace                  : Google.Policies
    policyType                       : admxIngested
    revision                         :
    id                               : a8de1c49-34d2-41e5-b798-da7c6f3709ca
    lastModifiedDateTime             : 2020-08-28T20:50:41.7098404Z
    fileName                         : google.admx
    status                           : uploadInProgress
    content                          :
    uploadDateTime                   : 0001-01-01T00:00:00Z
    defaultLanguageCode              : en-US
    groupPolicyUploadedLanguageFiles : {}
    ```
* Check the status of the uploaded ADMX file
    ``` PowerShell
    powershell .\list-uploadedadmxfiles.ps1 
    ```
    
* The output should list the status as follows for success:
    ``` Code
    displayName                      :
    description                      :
    languageCodes                    : {}
    targetPrefix                     : Google2cde4efa-564a-4d8a-945a-e18bff487ab5
    targetNamespace                  : Google.Policies
    policyType                       : admxIngested
    revision                         : 1.0
    id                               : 2cde4efa-564a-4d8a-945a-e18bff487ab5
    lastModifiedDateTime             : 2020-08-28T21:36:34.6180335Z
    fileName                         : google.admx
    status                           : available
    content                          :
    uploadDateTime                   : 0001-01-01T00:00:00Z
    defaultLanguageCode              : en-US
    groupPolicyUploadedLanguageFiles : {}
    ```        
    
* If an upload fails, use this command to get more information about error.
    ``` PowerShell
    powershell .\list-uploadedadmxfiles.ps1 -FileId "a8de1c49-34d2-41e5-b798-da7c6f3709ca" 
    ```

* The output will be something like the following for errors:
    ``` Code
    displayName                      :
    description                      :
    languageCodes                    : {}
    targetPrefix                     :
    targetNamespace                  :
    policyType                       : admxIngested
    revision                         :
    id                               : a8de1c49-34d2-41e5-b798-da7c6f3709ca
    lastModifiedDateTime             : 2020-08-28T20:51:04.0379614Z
    fileName                         : google.admx
    status                           : uploadFailed
    content                          :
    uploadDateTime                   : 0001-01-01T00:00:00Z
    defaultLanguageCode              : en-US
    groupPolicyUploadedLanguageFiles : {}
    groupPolicyOperations            : {@{operationType=upload; operationStatus=failed; statusDetails=localizedStringValue
                                    cannot be null; id=ab937f1a-4d0a-4320-9adc-55cc7d61ca45;
                                    lastModifiedDateTime=2020-08-28T20:51:04.0692062Z}}
    ```
## Publish app specificADMX and ADML files
* For example: Chrome
    ```PowerShell
    cd "C:\temp\Github\Intune-PowerShell-Management-ADMXCustomIngestion\Scenario Modules\IntuneCustomAdmxIngestion"

    PowerShell .\upload-admxfile.ps1 -ADMXFilePath "C:\temp\Chrome-PolicyTemplates\windows\admx\chrome.admx" -ADMLFilePath "C:\temp\Chrome-PolicyTemplates\windows\admx\en-US\chrome.adml"
    ```

    At this point you have the ADMX template for Chrome in your tenant

## Check status of ADMX file upload:
```PowerShell
powershell .\list-uploadedadmxfiles.ps1
```
Once the status is “Available” all settings will be available in tree view Intune UX, to create policies
 
# Configuring the uploaded settings
Follow the instructions in this document: https://docs.microsoft.com/en-us/mem/intune/configuration/administrative-templates-windows

When selecting settings to configure, you will see your uploaded configurations on the left side pane.

# To remove the uploaded ADMX templates.
```PowerShell
powershell .\remove-uploadedadmxfile.ps1 -Environment OneDF -FileIdToRemove <fileId>
```
All applied settings needs to be unconfigured, before a file can be removed. 

