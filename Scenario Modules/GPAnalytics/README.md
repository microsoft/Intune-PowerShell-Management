
# Overview 
Microsoft Endpoint Manager Group Policy Analytics allows the:
1. Import Group Policies Objects to Intune and provide insights on support to help customers make informed decisions.
2. Telemetry/insights on MDM gaps and help improve Windows MDM as a platform and
3. Migrate Imported Group Policy Objects into Intune Administrative Templates

# Notes for the preview
1. Update the Migration Readiness Report by using the Update-MigrationReadinessReport if Group Policy Objects were imported in the past.
2. Only ADMX Settings are currently supported in the Migration flow.

# Dependencies
The following dependencies are required for the Group Policy Analytics powershell scripts to work:

    1. Windows Remote Server Administration Tool (RSAT). Instructions on how to turn it on can be found here: https://support.microsoft.com/en-us/help/2693643/remote-server-administration-tools-rsat-for-windows-operating-systems
    2. Microsoft Active Directory PowerShell Module. Additional details are here: https://docs.microsoft.com/en-us/powershell/module/addsadministration/?view=win10-ps
    3. Group Policy PowerShell Module. Additional details are here: https://docs.microsoft.com/en-us/powershell/module/grouppolicy/?view=win10-ps 

Note: The IntunePolicyAnalytics Module will prompt and install the RSAT tool.

# Installation Instructions
* Download the Github repo zip file: https://github.com/microsoft/Intune-PowerShell-Management/archive/GPAnalytics.zip 
* extract it to local folder e.g. C:\temp\Github
* The Group Policy Analytics scripts will be here: 
            
    C:\temp\Github\Intune-PowerShell-Management-GPAnalytics\Scenario Modules\GPAnalytics
* Import the IntunePolicyAnalytics Module
``` Powershell
    CD  C:\temp\Github\Intune-PowerShell-Management-GPAnalytics\Scenario Modules\GPAnalytics
    Import-Module .\IntunePolicyAnalyticsClient.psd1

    PS C:\GitHub\Intune-PowerShell-Management\Scenario Modules\GPAnalytics> Get-Module

    ModuleType Version    Name                                ExportedCommands
    ---------- -------    ----                                ----------------
    Manifest   1.0.0.0    ActiveDirectory                     {Add-ADCentralAccessPolicyMember, Add-ADComputerServiceAccount, Add-ADDomainCo...
    Manifest   1.0.0.0    GroupPolicy                         {Backup-GPO, Copy-GPO, Get-GPInheritance, Get-GPO...}
    Script     6.2107.14  IntunePolicyAnalyticsClient         {Get-GPOMigrationReportCollection, Get-MigrationReadinessReport, Import-GPOCol...
    Manifest   3.1.0.0    Microsoft.PowerShell.Management     {Add-Computer, Add-Content, Checkpoint-Computer, Clear-Content...}
    Manifest   3.1.0.0    Microsoft.PowerShell.Utility        {Add-Member, Add-Type, Clear-Variable, Compare-Object...}
    Script     2.0.0      PSReadline                          {Get-PSReadLineKeyHandler, Get-PSReadLineOption, Remove-PSReadLineKeyHandler, ...
```

# Import Group Policy Object (GPO) Reports to Intune
``` Powershell
NAME
    Import-GPOBackupReports.ps1
    
SYNOPSIS
    Import-GPOBackupReports - Imports all GPO Report Xml backed up on disk into Intune and generate Migration report
    
    
SYNTAX
    .\Import-GPOBackupReports.ps1 [-TenantAdminUPN] <String> 
    [-GpoBackupFolderPath] <String> [[-Environment] <String>] [<CommonParameters>]
    
    
DESCRIPTION
    Imports all GPO Report Xml backed up on disk into Intune and generate Migration report
    

PARAMETERS
    -TenantAdminUPN <String>
        
    -GpoBackupFolderPath <String>
        
    -Environment <String>
        Environment is relevant only for test environments
        
    <CommonParameters>
        This cmdlet supports the common parameters: Verbose, Debug,
        ErrorAction, ErrorVariable, WarningAction, WarningVariable,
        OutBuffer, PipelineVariable, and OutVariable. For more information, see 
        about_CommonParameters (https:/go.microsoft.com/fwlink/?LinkID=113216). 
    
    -------------------------- EXAMPLE 1 --------------------------
    
    PS C:\GitHub\Intune-PowerShell-Management\Scenario Modules\GPAnalytics>.\Import-GPOBackupReports.ps1 -TenantAdminUPN "admin@IPASHAMSUA01MSIT.onmicrosoft.com" -GPOBackupFolderPath "C:\GPOBackup"
    
REMARKS
    To see the examples, type: "get-help 
    .\Import-GPOBackupReports.ps1 -examples".
    For more information, type: "get-help 
    .\Import-GPOBackupReports.ps1 -detailed".
    For technical information, type: "get-help 
    .\Import-GPOBackupReports.ps1 -full".
```

# Update the Migration Readiness of previously imported Group Policy Objects

``` Powershell
NAME
    Update-MigrationReadinessReport.ps1

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

    PS C:\GitHub\Intune-PowerShell-Management\Scenario Modules\GPAnalytics>.\Update-MigrationReadinessReport.ps1 -TenantAdminUPN "admin@IPASHAMSUA01MSIT.onmicrosoft.com"

    Updates the Migration Readiness Report for previously uploaded GPOs.


RELATED LINKS
```
