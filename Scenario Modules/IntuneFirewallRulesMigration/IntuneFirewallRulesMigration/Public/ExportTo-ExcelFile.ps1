. "$PSScriptRoot\..\Private\ConvertTo-IntuneFirewallRule-Helper.ps1"
. "$PSScriptRoot\IntuneFirewallRule.ps1"


class ExcelFormat :  IntuneFirewallRule {
    [String]$errorMessage
    [String]$profileTypes
    [String]$localPortRanges
    [String]$remotePortRanges
    [String] $actualLocalAddressRanges
    [String] $actualRemoteAddressRanges
    [String] $interfaceTypes
}
data SummaryDetails
{
    ConvertFrom-StringData -StringData @'
    NumberofFirewallRules=0
    NumberofSplitRules=0
    NumberofSuccessfulConvertedRules=0
    NumberofSucceededSent=0
    ProfileName=Profile Name
'@
}
function Get-ExcelFormatObject
{
     <#
    .SYNOPSIS
    Converts IntuneFirewall Object to a format that can easily be written to an excel file

    .DESCRIPTION
    This function takes Intunefirewall Object and converts them to a format that is easily writtable to an excel file 
    .EXAMPLE
    Get-ExcelFormatObject intuneFirewallObjects
    Get-ExcelFormatObject intuneFirewallObjects
    Get-ExcelFormatObject intuneFirewallObjects
    .OUTPUTS
    ExcelFormat
    #>
    param(
        [Parameter(Mandatory=$true)]
        $intuneFirewallObjects,
        $errorMessage
    )
    $excelObjects = @()
    foreach($intuneFirewallRuleObject in $intuneFirewallObjects){
        $newExcelObject = New-Object -TypeName ExcelFormat
        $newExcelObject.displayName = $intuneFirewallRuleObject.displayName
        $newExcelObject.description = $intuneFirewallRuleObject.description
        $newExcelObject.packageFamilyName = $intuneFirewallRuleObject.packageFamilyName
        $newExcelObject.serviceName              =$intuneFirewallRuleObject.serviceName
        $newExcelObject.filePath                 =$intuneFirewallRuleObject.filePath
        $newExcelObject.interfaceTypes           =$intuneFirewallRuleObject.interfaceTypes 
        $newExcelObject.actualLocalAddressRanges =(ConvertTo-Json $intuneFirewallRuleObject.actualLocalAddressRanges)
        $newExcelObject.actualRemoteAddressRanges=(ConvertTo-Json $intuneFirewallRuleObject.actualRemoteAddressRanges) 
        $newExcelObject.profileTypes             =(ConvertTo-Json $intuneFirewallRuleObject.profileTypes)  
        $newExcelObject.localPortRanges          =(ConvertTo-Json $intuneFirewallRuleObject.localPortRanges)  
        $newExcelObject.remotePortRanges         =(ConvertTo-Json $intuneFirewallRuleObject.remotePortRanges) 
        $newExcelObject.protocol                 =$intuneFirewallRuleObject.protocol 
        $newExcelObject.action                   =$intuneFirewallRuleObject.action                                       
        $newExcelObject.trafficDirection         =$intuneFirewallRuleObject.trafficDirection                  
        $newExcelObject.useAnyLocalAddressRange  =$intuneFirewallRuleObject.useAnyLocalAddressRange   
        $newExcelObject.useAnyRemoteAddressRange =$intuneFirewallRuleObject.useAnyRemoteAddressRange  

        if($errorMessage){     
            $newExcelObject.errorMessage = $errorMessage
        }

        $excelObjects += $newExcelObject
    }
    return $excelObjects
}

function Export-ExcelFile
{
    <#
    .SYNOPSIS
    Creates an excel file called reports.xlsx 

    .DESCRIPTION
    Creates an excel file called reports.xlsx in a folder called logs in the current directory with 3 worksheets that  holds information of the firewall rules that were successfully imported to intune and those that failed.

    .EXAMPLE
    Export-ExcelFile -fileName [filename] -succededToSend $IntuneFirewallRuleObjects
    Export-ExcelFile -fileName [filename] -failedToSend $IntuneFirewallRuleObjects
    Export-ExcelFile -fileName [filename] -failedToConvert $FirewallRuleObjects
    .OUTPUTS
    .\logs\report.xlsx
    #>
    
    param(
        [Parameter(Mandatory = $true)]
        [string]
        $fileName,
        $succeededToSend,
        $failedToSend,
        $failedToConvert
    )
    $i = 0
    $date = Get-Date -Format "M_dd_yy_HH"
    $path = ".\logs\" + $fileName + "_" + $i + " " + $date + ".xlsx"
    
    while(Test-Path $path){
        $i++
        $path = ".\logs\" + $fileName+ "_" + $i + " " + $date + ".xlsx"
        
    }
    
    $properties = @('displayName', 'description', 'action','trafficDirection','profileTypes','interfaceTypes','localPortRanges','remotePortRanges', 'protocol','actualLocalAddressRanges','actualRemoteAddressRanges', 'errorMessage')
    if($failedToConvert){    
        $excel = $failedToConvert| Select-Object -Property displayName, description, action, trafficDirection, errorMessage | Export-Excel $path -AutoFilter -AutoSize -WorksheetName "Failed FirewallRule Conversion"
    }
    if($failedToSend){
        $excel = $failedToSend |Select-Object -Property $properties | Export-Excel $path -AutoFilter -AutoSize -WorksheetName "Failed to Import to Intune"
    }
    if($succeededToSend){
        $excel = $succeededToSend|Select-Object -Property $properties | Export-Excel $path  -AutoFilter -AutoSize -WorksheetName "Imported To Intune"
    }
    return $excel
}

function Get-SummaryDetail
{
    Write-Host "`rSummary Details`r"
    if([int]$SummaryDetails.NumberofSucceededSent -eq [int]$SummaryDetails.NumberofSplitRules)
    {
        Write-Host "Imported " $SummaryDetails.NumberofSuccessfulConvertedRules "/" $SummaryDetails.NumberofFirewallRules "into the Endpoint Security Firewall Rule Profile '"$SummaryDetails.ProfileName"'"
    }
    else
    {
        Write-Host "Imported " $SummaryDetails.NumberofSucceededSent "/" $SummaryDetails.NumberofSplitRules "into the Endpoint Security Firewall Rule Profile '"$SummaryDetails.ProfileName"'"
    }
    if(Test-Path ".\logs")
    {
        Write-Host "See logs :" (Resolve-Path ".\logs") "for more information`r"
    }
}

function Set-SummaryDetail {
    param(
        
        $successCount,
        $numberOfSplittedRules,
        [String]
        $ProfileName,
        $numberOfFirewallRules,
        $ConvertedRulesNumber
    )
    if ($successCount) {
        $SummaryDetails.NumberofSucceededSent = $successCount
    }
    if ($numberOfSplittedRules) {
        $SummaryDetails.NumberofSplitRules = $numberOfSplittedRules
    }
    if ($ProfileName) {
        $SummaryDetails.ProfileName = $ProfileName
    }
    if ($numberOfFirewallRules) {
        $SummaryDetails.NumberofFirewallRules = $numberOfFirewallRules
    }
    if ($ConvertedRulesNumber) {
        $SummaryDetails.NumberofSuccessfulConvertedRules = $ConvertedRulesNumber
    }
   
    
}
