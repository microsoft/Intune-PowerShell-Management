<#
.Synopsis
   Set-DeviceCategory Sets the Device Category for a given Intune Managed Device.
.DESCRIPTION
   Sets the Device Category for a given Intune Managed Device.
.PARAMETER DeviceName
    Display Name of the Intune Managed Device.
.PARAMETER DeviceCategoryName
	Display Name of the Intune Device Category to applied to the said Intune Managed Device. 
	If the Device Category does not exist in Intune, Create it first in Intune.

.EXAMPLE
     Set-DeviceCategory -DeviceName "DESKTOP-CTGLE11" -DeviceCategoryName "category-03"
       Sets the Device category for "DESKTOP-CTGLE11" to "category-03".
#>

param(
	[Parameter(Mandatory=$true)]
	[string]$DeviceName, 
	
	[Parameter(Mandatory=$true)]
	[string]$DeviceCategoryName
)

#
# Install Microsoft.Graph.Intune if necessary - the Microsoft Intune PowerShell SDK
#
If ($null -eq (Get-module -Name 'Microsoft.Graph.Intune'))
{
	Install-Module 'Microsoft.Graph.Intune'
}

 #
 # Set Graph schema version to 'beta'
 #
(Update-MSGraphEnvironment -SchemaVersion 'beta') | Out-Null

 #
 # Connect to AAD using the Intune ITPro credentials
 #
 (Connect-MSGraph) | Out-Null

 #
 # Get the Device Category Id based on Display Name
 #
Write-Output "Resolving $($DeviceCategoryName) ..."
$deviceCategoryId =  (Get-IntuneDeviceCategory | Where-Object { $_.displayName -eq $DeviceCategoryName }).id

# check if Category name resolved or not
If ($null -eq $deviceCategoryId)
{
	Write-Error "ERROR: Unable to resolve $($DeviceCategoryName) in Intune."
	throw
}
Else
{
	Write-Output "$($DeviceCategoryName) resolved to $($deviceCategoryId)"		
}

#
# Get the Deice Id based on Device Name
#
Write-Output "Resolving $($DeviceName) ..."
$deviceId = (Get-IntuneManagedDevice | Where-Object { $_.deviceName -eq $DeviceName }).id

# check if Device Name resolved or not
If ($null -eq $deviceId)
{
	Write-Error "ERROR: Unable to resolve $($DeviceName) in Intune."
	throw
}
Else
{
	Write-Output "$($DeviceName) resolved to $($deviceId)"
}

 #
 # Set up the Graph call parameters
 #
 $deviceCategoryReqBody = (([PSCustomObject]@{"@odata.id" = "https://graph.microsoft.com/beta/deviceManagement/deviceCategories/$($deviceCategoryId)";})| ConvertTo-Json).ToString()
 $deviceCategoryRefUrl = "https://graph.microsoft.com/beta/deviceManagement/managedDevices/$($deviceId)/deviceCategory/`$ref"

 #
 # Make the MSGraph call
 #
Try
{
	Write-Output "Assigning IntuneDeviceId=$($deviceId) to DeviceCategoryId=$($deviceCategoryId)"
	Invoke-MSGraphRequest -Url $deviceCategoryRefUrl -HttpMethod Put -Content $deviceCategoryReqBody
}
Catch
{
	Write-Error $_
	throw
}

Write-Output "SUCCESS: DeviceName=$($DeviceName) assigned to DeviceCategory=$($DeviceCategoryName) in Intune."