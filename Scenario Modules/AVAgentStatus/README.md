
# Overview 
This Powershell tool allows admins to export Unhealthy endpoint with Product Status in understandable format instead of Flags enum:


# Dependencies
The tool will install 'Intune Powershell SDK' from Powershell Gallery if it is not already installed.    


# Export Unhealthy endpoint report to CSV format with ProductStatus in understandable format.
``` Powershell
NAME
    Export-AgentStatusOpsReportExpandedProductStatus.ps1
    
SYNOPSIS
       Export-AgentStatusOpsReportWithProductStatusDescription - Exports the AgentStatusOpsReport which ProductStatus in descriptive format.
	   It takes output path for the csv as parameter. 
    
    
SYNTAX
    .\Export-AgentStatusOpsReportExpandedProductStatus.ps1 [-OutputPath] <String>     
    
    
DESCRIPTION
    Exports the AgentStatusOpsReport which ProductStatus Description
    

PARAMETERS
    -TenantAdminUPN <String>
            
    
    -------------------------- EXAMPLE 1 --------------------------
    
    PS C:\GitHub>.\Export-AgentStatusOpsReportExpandedProductStatus.ps1 -OutputPath .\Report.csv   

```
