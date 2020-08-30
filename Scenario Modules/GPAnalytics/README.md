
# Overview 
Microsoft Endpoint Manager Group Policy Analytics allows the:
1. Import Group Policies Objects to Intune and provide insights on support to help customers make informed decisions.
2. Telemetry/insights on MDM gaps and help improve Windows MDM as a platform and
3. Migrate Imported Group Policy Objects into Intune Administrative Templates

# Notes for the preview
1. Update the Migration Readiness Report by using the Update-MigrationReadinessReport if Group Policy Objects were imported in the past.
2. Only ADMX Settings are currently supported in the Migration flow.

# Installation Instructions
* Download the Github repo zip file: https://github.com/microsoft/Intune-PowerShell-Management/archive/GPAnalytics.zip 
* extract it to local folder e.g. C:\temp\Github
* The Group Policy Analytics scripts will be here: 
            
    C:\temp\Github\Intune-PowerShell-Management-GPAnalytics\Scenario Modules\GPAnalytics
* Import the IntunePolicyAnalytics Module
``` Powershell
    CD  C:\temp\Github\Intune-PowerShell-Management-GPAnalytics\Scenario Modules\GPAnalytics
    Import-Module .\IntunePolicyAnalyticsClient.psm1
```

# Generate the Migration Readiness of imported Group Policy Objects
``` Powershell
NAME
    Get-MigrationReadinessReport

SYNOPSIS
    Get-MigrationReadinessReport Gets the Migration Readiness Report for previously uploaded GPOs.


SYNTAX
    Get-MigrationReadinessReport [-TenantAdminUPN] <String> [<CommonParameters>]


DESCRIPTION
    Gets the Migration Readiness Report for previously uploaded GPOs.


PARAMETERS
    -TenantAdminUPN <String>
        AAD User Principal Name of the Intune Tenant Administrator which is required to upload the GPOs.

        Required?                    true
        Position?                    1
        Default value
        Accept pipeline input?       false
        Accept wildcard characters?  false

    <CommonParameters>
        This cmdlet supports the common parameters: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer, PipelineVariable, and OutVariable. For more information, see
        about_CommonParameters (https:/go.microsoft.com/fwlink/?LinkID=113216).

INPUTS

OUTPUTS

    -------------------------- EXAMPLE 1 --------------------------

    PS C:\>Get-MigrationReadinessReport -TenantAdminUPN "admin@IPASHAMSUA01MSIT.onmicrosoft.com"

    Gets the Migration Readiness Report for previously uploaded GPOs.





RELATED LINKS
```

# Update the Migration Readiness of previously imported Group Policy Objects

``` Powershell
NAME
    Update-MigrationReadinessReport

SYNOPSIS
    Update-MigrationReadinessReport Updates the Migration Readiness Report for previously uploaded GPOs.


SYNTAX
    Update-MigrationReadinessReport [-TenantAdminUPN] <String> [<CommonParameters>]


DESCRIPTION
    Updates the Migration Readiness Report for previously uploaded GPOs.


PARAMETERS
    -TenantAdminUPN <String>
        AAD User Principal Name of the Intune Tenant Administrator which is required to upload the GPOs.

        Required?                    true
        Position?                    1
        Default value
        Accept pipeline input?       false
        Accept wildcard characters?  false

    <CommonParameters>
        This cmdlet supports the common parameters: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer, PipelineVariable, and OutVariable. For more information, see
        about_CommonParameters (https:/go.microsoft.com/fwlink/?LinkID=113216).

INPUTS

OUTPUTS

    -------------------------- EXAMPLE 1 --------------------------

    PS C:\>Update-MigrationReadinessReport -TenantAdminUPN "admin@IPASHAMSUA01MSIT.onmicrosoft.com"

    Updates the Migration Readiness Report for previously uploaded GPOs.





RELATED LINKS
```
# Migrate imported Group Policy Object into Intune Administrative Templates
``` Powershell
NAME
    Add-GPToIntuneAdmxMigratedProfile

SYNOPSIS
    Migrates the GPO Migration Reports to Intune Admx Profiles


SYNTAX
    Add-GPToIntuneAdmxMigratedProfile [-TenantAdminUPN] <String> [<CommonParameters>]


DESCRIPTION
    Migrates the GPO Migration Reports to Intune Admx Profiles


PARAMETERS
    -TenantAdminUPN <String>
        The UPN of the Intune Tenant Admin.

        Required?                    true
        Position?                    1
        Default value
        Accept pipeline input?       false
        Accept wildcard characters?  false

    <CommonParameters>
        This cmdlet supports the common parameters: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer, PipelineVariable, and OutVariable. For more information, see
        about_CommonParameters (https:/go.microsoft.com/fwlink/?LinkID=113216).

INPUTS

OUTPUTS

    -------------------------- EXAMPLE 1 --------------------------

    PS C:\>Add-GPToIntuneAdmxMigratedProfile -TenantAdminUPN "admin@IPASHAMSUA01MSIT.onmicrosoft.com"

    Creates Administrative Templates Configuration Profiles to migrate the GPOs
```