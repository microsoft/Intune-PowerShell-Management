<#
.Synopsis
  IntunePolicyAnalyticsClient - Implementation of Group Policy Object import into Intune.
.DESCRIPTION
   Implementation of Group Policy Object import into Intune.
#>

#region Cmdlets
Function Get-GPOMigrationReportCollection
{
<#
.Synopsis
   Get-GPOMigrationReportCollection Generates Migration reports from previously updated Group Policy Objects
.DESCRIPTION
   Gets Migration report for previously uploaded GPOs from Intune.
.PARAMETER TenantAdminUPN
    AAD User Principal Name of the Intune Tenant Administrator which is required to upload the GPOs.
.PARAMETER ExpandSettings
    Switch if set, expands the SettingsMap
.OUTPUT
    GPO Migration Report based on the previously updated Group Policy Objects
.EXAMPLE
    Get-GPOMigrationReportCollection -TenantAdminUPN "admin@IPASHAMSUA01MSIT.onmicrosoft.com"
        Gets the MigrationReports from Intune
#>
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$TenantAdminUPN,
        [Switch]$ExpandSettings = $false
    )

    Try
    {
        # Initialize Module
        $IPAClientConfiguration = Initialize-IPAClientConfiguration -TenantAdminUPN $TenantAdminUPN

        # Fetch the migration report
        $sw = [Diagnostics.Stopwatch]::StartNew()

        # Initialize Module
        $IPAClientConfiguration = Initialize-IPAClientConfiguration -TenantAdminUPN  $TenantAdminUPN

        # Make the Graph call to fetch the GroupPolicyMigrationReport collection
        $collectionUri = ""
        $nextUrl = $null

        # Iterate through nextlinks to get the complete set of reports
        Do
        {
            $result = Get-IntuneEntityCollection -CollectionPath $collectionUri `
                                     -Uri $nextUrl `
                                     -GraphConfiguration $script:GraphConfiguration
            $migrationReportCollection += $result.Value
            $nextUrl = $result.'@odata.nextLink'            
        }
        While ($nextUrl)

        Write-Log -Message "Get-GPOMigrationReportCollection Found $($migrationReportCollection.Count) MigrationReports.."

        # Instantiate a new collection for the GPO Migration Reports to be fetched from Intune
        $GPOMigrationReportCollection = @{}

        # Populate the groupPolicyMigrationReports collection
        ForEach ($migrationReport in $migrationReportCollection)
        {
            Try
            {
                # Get the groupPolicySettingMappings for each migrationReport
                $groupPolicyObjectId = $migrationReport.groupPolicyObjectId
                $ou = [System.Web.HTTPUtility]::UrlDecode($migrationReport.ouDistinguishedName)

                If ($ExpandSettings)
                {
                    $collectionUri = "('$($groupPolicyObjectId)_$($ou)')?`$expand=groupPolicySettingMappings"
                    Write-Log -Message "Get-GPOMigrationReportCollection: collectionUri=$($collectionUri)"

                    $groupPolicySettingMappingCollection = Get-IntuneEntityCollection -CollectionPath $collectionUri `
                                                                          -Uri $null `
                                                                          -GraphConfiguration $script:GraphConfiguration
                }
                Else
                {
                    $groupPolicySettingMappingCollection = $null
                }

                If ($null -eq $groupPolicySettingMappingCollection)
                {
                    $GPOMigrationReportCollection.Add("$($groupPolicyObjectId)_$($ou)", [PSCustomObject]@{MigrationReport = $migrationReport; `
                                                                                SettingMappings = $null})
                }
                Else
                {
                    $GPOMigrationReportCollection.Add("$($groupPolicyObjectId)_$($ou)", [PSCustomObject]@{MigrationReport = $migrationReport; `
                                                                                SettingMappings = ($groupPolicySettingMappingCollection.groupPolicySettingMappings)})
                }
            }
            Catch
            {
                $exception  = $_
                Write-Log -Message "Get-GPOMigrationReportCollection: Failure: $($exception)" -Level "Warn"
            }
        }
    }
    Catch
    {
        $exception  = $_
        Write-Log -Message "Get-GPOMigrationReportCollection: Failure: $($exception)" -Level "Error"
        Write-Log -Message "Get-GPOMigrationReportCollection: IPAGlobalSettingBag: $($IPAClientConfiguration)"
        Write-Log -Message "Get-GPOMigrationReportCollection: graphConfiguration: $($graphConfiguration)"
        throw
    }
    Finally
    {
        $IPAClientConfiguration | ConvertTo-Json | Out-File -FilePath "$($IPAClientConfiguration.ConfigurationFolderPath)\IPAGlobalSettingBag.json"
        $GPOMigrationReportCollection | ConvertTo-Json | Out-File -FilePath "$($IPAClientConfiguration.StatePath)\GroupPolicyMigrationReportCollection.json"
        $sw.Stop()
        Write-Log -Message "Get-GPOMigrationReportCollection: Elapsed time = $($sw.Elapsed.ToString())"
    }

    Write-Log -Message "Get-GPOMigrationReportCollection: GPOMigrationReports returned=$($GPOMigrationReportCollection.Count)"
    return $GPOMigrationReportCollection
}

Function Get-MigrationReadinessReport
{
<#
.Synopsis
   Get-MigrationReadinessReport Gets the Migration Readiness Report for previously uploaded GPOs.
.DESCRIPTION
   Gets the Migration Readiness Report for previously uploaded GPOs.
.PARAMETER TenantAdminUPN
    AAD User Principal Name of the Intune Tenant Administrator which is required to upload the GPOs.
.EXAMPLE
     Get-MigrationReadinessReport -TenantAdminUPN "admin@IPASHAMSUA01MSIT.onmicrosoft.com"
       Gets the Migration Readiness Report for previously uploaded GPOs.
#>
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$TenantAdminUPN
    )

    Try
    {
        # Initialize Module
        $IPAClientConfiguration = Initialize-IPAClientConfiguration -TenantAdminUPN $TenantAdminUPN

        # Fetch the migration report
        $sw = [Diagnostics.Stopwatch]::StartNew()

        # Headers for the Graph call
        $clonedHeaders = CloneObject (Get-AuthHeader $graphConfiguration);
        $clonedHeaders["api-version"] = "$($script:ApiVersion)";        

        <#
            1. Ask to create the report
            Post https://graph.microsoft-ppe.com/testppebeta_intune_onedf/deviceManagement/reports/cachedReportConfigurations
            Payload: {"id":"GPAnalyticsMigrationReadiness_00000000-0000-0000-0000-000000000001","filter":"","orderBy":[],"select":["SettingName","MigrationReadiness","OSVersion","Scope","ProfileType","SettingCategory"],"metadata":""}
        #>
        $uri = "$($script:GraphConfiguration.GraphBaseAddress)/$($script:GraphConfiguration.SchemaVersion)/deviceManagement/reports/cachedReportConfigurations";        
        $Body = "{            
            `"id`":`"GPAnalyticsMigrationReadiness_00000000-0000-0000-0000-000000000001`",
            `"filter`": `"`",`
            `"select`": [`
            `"SettingName`",`"MigrationReadiness`",`"OSVersion`",`"Scope`",`"ProfileType`",`"SettingCategory`"
            ],`
            `"orderBy`": [`
                `"SettingCategory`"
            ]`
          }"
        $clonedHeaders["content-length"] = $Body.Length;

        Try
        {        
            Write-Log -Message "Get-MigrationReadinessReport: Creating MigrationReadiness Report..."           
            $response = Invoke-RestMethod $uri -Method Post -Headers $clonedHeaders -Body $body;
        }
        Catch
        {
            $exception  = $_
            Write-Log -Message "Get-MigrationReadinessReport: Invoke-RestMethod $uri -Method Post. Size=$($Body.Length). Failure: $($exception)" -Level "Warn"
            throw
        }
            
        <#    
            2. Query, over and over, until the report is complete (you will see 'completed' in the response
            get: https://graph.microsoft-ppe.com/testppebeta_intune_onedf/deviceManagement/reports/cachedReportConfigurations('GPAnalyticsSettingMigrationReadiness_00000000-0000-0000-0000-000000000001')            
        #>
        $uri = "$($script:GraphConfiguration.GraphBaseAddress)/$($script:GraphConfiguration.SchemaVersion)/deviceManagement/reports/cachedReportConfigurations(`'GPAnalyticsSettingMigrationReadiness_00000000-0000-0000-0000-000000000001`')"
        Try
        {
            $Counter = 0
            Write-Log -Message "Get-MigrationReadinessReport: Getting MigrationReadinessReport..."       
            Do
            {                
                $response = (Invoke-RestMethod $uri -Method Get -Headers $clonedHeaders)
                $Counter++
                Write-Log -Message "Get-MigrationReadinessReport: Report creation Status: $($response.Status),  Attempt: $($Counter)"
                Start-Sleep -Seconds 1
            } While (($response.Status -contains "inProgress") -and ($Counter -lt 100))
        }
        Catch
        {
            $exception  = $_
            Write-Log -Message "Get-MigrationReadinessReport: Invoke-RestMethod $uri -Method Get. Failure: $($exception)" -Level "Warn"
            throw
        }

        <#
            3. Get the actual report: (you may want to increase 'top')
            Post: https://graph.microsoft-ppe.com/testppebeta_intune_onedf/deviceManagement/reports/getCachedReport
            Payload: {"Id":"GPAnalyticsMigrationReadiness_00000000-0000-0000-0000-000000000001","Skip":0,"Top":50,"Search":"","OrderBy":[],"Select":["SettingName","MigrationReadiness","OSVersion","Scope","ProfileType","SettingCategory"]}
        #>
        $uri = "$($script:GraphConfiguration.GraphBaseAddress)/$($script:GraphConfiguration.SchemaVersion)/deviceManagement/reports/getCachedReport";        
        $Body = "{            
            `"id`":`"GPAnalyticsMigrationReadiness_00000000-0000-0000-0000-000000000001`",
            `"skip`": `"0`",`
            `"Search`": `"`",`
            `"select`": [`
            `"SettingName`",`"MigrationReadiness`",`"OSVersion`",`"Scope`",`"ProfileType`",`"SettingCategory`"
            ],`
            `"orderBy`": [`
                `"SettingCategory`"`
            ]`
          }"
        $clonedHeaders["content-length"] = $Body.Length;

        Try
        {
            Write-Log -Message "Get-MigrationReadinessReport: Get the created report..."                       
            $response = Invoke-RestMethod $uri -Method Post -Headers $clonedHeaders -Body $body;

            Write-Log -Message "Get-MigrationReadinessReport: $($response.TotalRowCount) records found."
            Write-Log -Message "Get-MigrationReadinessReport: $($response.Values.Count) records downloaded."
        }
        Catch
        {
            $exception  = $_
            Write-Log -Message "Get-MigrationReadinessReport: Invoke-RestMethod $uri -Method Post. Size=$($Body.Length). Failure: $($exception)" -Level "Warn"
            throw
        }     
    }
    Catch
    {
        $exception  = $_
        Write-Log -Message "Get-MigrationReadinessReport: Failure: $($exception)" -Level "Error"
        Write-Log -Message "Get-MigrationReadinessReport: IPAGlobalSettingBag: $($IPAClientConfiguration)"
        Write-Log -Message "Get-MigrationReadinessReport: graphConfiguration: $($graphConfiguration)"
        throw
    }
    Finally
    {
        $IPAClientConfiguration | ConvertTo-Json | Out-File -FilePath "$($IPAClientConfiguration.ConfigurationFolderPath)\IPAGlobalSettingBag.json"
        $GPOMigrationReportCollection | ConvertTo-Json | Out-File -FilePath "$($IPAClientConfiguration.StatePath)\GroupPolicyMigrationReportCollection.json"
        $sw.Stop()
        Write-Log -Message "Get-MigrationReadinessReport Elapsed time = $($sw.Elapsed.ToString())"
    }
    
    return  $response
}

Function Update-MigrationReadinessReport
{
<#
.Synopsis
   Update-MigrationReadinessReport Updates the Migration Readiness Report for previously uploaded GPOs.
.DESCRIPTION
   Updates the Migration Readiness Report for previously uploaded GPOs.
.PARAMETER TenantAdminUPN
    AAD User Principal Name of the Intune Tenant Administrator which is required to upload the GPOs.
.EXAMPLE
     Update-MigrationReadinessReport -TenantAdminUPN "admin@IPASHAMSUA01MSIT.onmicrosoft.com"
       Updates the Migration Readiness Report for previously uploaded GPOs.
#>
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$TenantAdminUPN
    )

    Try
    {
        # Initialize Module
        $IPAClientConfiguration = Initialize-IPAClientConfiguration -TenantAdminUPN $TenantAdminUPN

        # Fetch the migration report
        $sw = [Diagnostics.Stopwatch]::StartNew()

        # Make the Graph call to fetch the GroupPolicyMigrationReport collection        
        $nextUrl = "$($script:GraphConfiguration.GraphBaseAddress)/$($script:GraphConfiguration.SchemaVersion)/deviceManagement/GroupPolicyObjectFiles"

        # Iterate through nextlinks to get the complete set of reports
        Do
        {
            $result = Get-IntuneEntityCollection -Uri $nextUrl `
                                     -GraphConfiguration $script:GraphConfiguration `
                                     -CollectionPath ""
            $GroupPolicyObjectFileCollection += $result.Value       
            $nextUrl = $result.'@odata.nextLink'            
        }
        While ($nextUrl)

        Write-Log -Message "Update-MigrationReadinessReport: Found $($GroupPolicyObjectFileCollection.Count) GPO files.."
        
        # Upload GroupPolicyObjectFile back to Intune
        ForEach ($GroupPolicyObjectFile in $GroupPolicyObjectFileCollection)
        {                     
            $ouDistinguishedName = [System.Web.HTTPUtility]::UrlDecode($GroupPolicyObjectFile.ouDistinguishedName)
            $content = $GroupPolicyObjectFile.content
            $GroupPolicyObjectFileToUpload = [PSCustomObject]@{groupPolicyObjectFile = ([PSCustomObject]@{ouDistinguishedName = $ouDistinguishedName; content = $content})}               
            
            # Upload GroupPolicyObjectFile to Intune
            Try
            {
                Write-Log "Update-MigrationReadinessReport: Updating $($GroupPolicyObjectFile.id)..." 
                $MigrationReportCreated = Add-IntuneEntityCollection "createMigrationReport" ($GroupPolicyObjectFileToUpload |ConvertTo-Json) $script:GraphConfiguration
                Write-Log "Update-MigrationReadinessReport: $($MigrationReportCreated.Value) updated."                                
            }
            Catch
            {
                $exception  = $_
                Write-Log -Message "Update-MigrationReadinessReport: Failure: $($exception)"
            }            
        }
    }
    Catch
    {
        $exception  = $_
        Write-Log -Message "Update-MigrationReadinessReport: Failure: $($exception)" -Level "Error"
        Write-Log -Message "Update-MigrationReadinessReport: IPAGlobalSettingBag: $($IPAClientConfiguration)"
        Write-Log -Message "Update-MigrationReadinessReport: graphConfiguration: $($graphConfiguration)"
        throw
    }
    Finally
    {
        $IPAClientConfiguration | ConvertTo-Json | Out-File -FilePath "$($IPAClientConfiguration.ConfigurationFolderPath)\IPAGlobalSettingBag.json"
        $GPOMigrationReportCollection | ConvertTo-Json | Out-File -FilePath "$($IPAClientConfiguration.StatePath)\GroupPolicyMigrationReportCollection.json"
        $sw.Stop()
        Write-Log -Message "Update-MigrationReadinessReport Elapsed time = $($sw.Elapsed.ToString())"
    }
}

Function Import-GPOCollection
{
<#
.Synopsis
    Import-GPOCollection Gets all the Group Policy Object collection for a given domain, uploads it to Intune and determines what settings are supported.
.DESCRIPTION
    IntunePolicyAnalyticsClient uses the Group Policy cmdlets to get all the Group Policy Objects
    for a given domain, uploads it to Intune and determines what settings are supported.
.PARAMETER Domain
    The local AD Domain for which the GPO collection is fetched.
    Defaults to the local AD Domain for the client on which this script is run on.
.PARAMETER TenantAdminUPN
    AAD User Principal Name of the Intune Tenant Administrator which is required to upload the GPOs.
.PARAMETER OUFilter
    Use OUFilter to constrain the GP Objects to the OU in consideration.
    Specifies a query string that retrieves Active Directory objects. This string uses the PowerShell Expression Language syntax. The PowerShell Expression Language syntax provides rich type-conversion support for value types received by the OUFilter parameter. The syntax uses an in-order representation, which means that the operator is placed between the operand and the value. For more information about the OUFilter parameter, type Get-Help about_ActiveDirectory_Filter.
    Syntax:
        The following syntax uses Backus-Naur form to show how to use the PowerShell Expression Language for this parameter.
        <OUFilter> ::= "{" <FilterComponentList> "}"
        <FilterComponentList> ::= <FilterComponent> | <FilterComponent> <JoinOperator> <FilterComponent> | <NotOperator> <FilterComponent>
        <FilterComponent> ::= <attr> <FilterOperator> <value> | "(" <FilterComponent> ")"
        <FilterOperator> ::= "-eq" | "-le" | "-ge" | "-ne" | "-lt" | "-gt"| "-approx" | "-bor" | "-band" | "-recursivematch" | "-like" | "-notlike"
        <JoinOperator> ::= "-and" | "-or"
        <NotOperator> ::= "-not"
        <attr> ::= <PropertyName> | <LDAPDisplayName of the attribute>
        <value>::= <compare this value with an <attr> by using the specified <FilterOperator>>
        For a list of supported types for <value>, type Get-Help about_ActiveDirectory_ObjectModel.
.OUTPUT
    GPO Collection collected from the local AD domain and sent to Intune
.EXAMPLE
    Import-GPOCollection -Domain "redmond.corp.microsoft.com" -TenantAdminUPN "admin@IPASHAMSUA01MSIT.onmicrosoft.com"
        Gets all the GPOs in the domain "redmond.corp.microsoft.com" and back them up on disk.
.EXAMPLE
    Import-GPOCollection -Domain "redmond.corp.microsoft.com" -OUFilter 'DistinguishedName -like "OU=CoreIdentity,OU=ITServices,DC=redmond,DC=corp,DC=microsoft,DC=com"' -TenantAdminUPN "admin@IPASHAMSUA01MSIT.onmicrosoft.com"
        Gets all the GPOs in a specific OU for the given domain, back them up on disk and upload to Intune.
#>
    [cmdletbinding()]
    param(
        [Alias("Domain")]
        [Parameter(Mandatory=$true)]
        [String]$ADDomain,
        [Parameter(Mandatory=$true)]
        [string]$TenantAdminUPN,
        [String]$OUFilter = 'Name -like "*"'
    )

    Try
    {
        # Start timer
        $sw = [Diagnostics.Stopwatch]::StartNew()

        # Initialize Module
        $IPAClientConfiguration = Initialize-IPAClientConfiguration -TenantAdminUPN $TenantAdminUPN

        # Fetch GPO Xml Reports from local AD Domain
        $gpoReportXmlCollection = @{}
        Try
        {
            Write-Log -Message "Import-GPOCollection: Get GPO backups from ADDomain=$($ADDomain) with OUFilter=$($OUFilter)..."
            $gpoReportXmlCollection = Get-GPOReportXmlCollectionFromAD -ADDomain $($ADDomain) -OUFilter $($OUFilter)
        }
        Catch
        {
            $exception  = $_

            # For non-domain joined loadgens
            Switch ($exception)
            {
                'Unable to contact the server. This may be because this server does not exist, it is currently down, or it does not have the Active Directory Web Services running.'
                {
                    If (($IPAClientConfiguration.Environment -ne "pe") -and (Test-Path -Path "$($IPAClientConfiguration.StatePath)\GPOReportXmlCollection.json"))
                    {
                        (Get-Content -Path "$($IPAClientConfiguration.StatePath)\GPOReportXmlCollection.json" `
                            | ConvertFrom-JSon).psobject.properties | ForEach-Object { $gpoReportXmlCollection[$_.Name] = $_.Value }
                        Write-Log -Message "Import-GPOCollection: Read GPOReportXmlCollection from: $($IPAClientConfiguration.StatePath)\GPOReportXmlCollection.json"
                    }
                    Else
                    {
                        Write-Log -Message "Import-GPOCollection:: Get-GPOReportXmlCollectionFromAD -ADDomain $($ADDomain) -OUFilter $($OUFilter) failed. Failure: $($exception)" -Level "Error"
                        throw
                    }
                }
            }
        }

        # Upload GPOs to Intune
        If ($null -ne $gpoReportXmlCollection)
        {
            Write-Log -Message "Import-GPOCollection: Number of GPOs to upload to Intune=$($gpoReportXmlCollection.Count)"
            $gpoReportXmlCollection.GetEnumerator() | ForEach-Object `
            {
                Try
                {
                    $key = $_.key

                    # Create GroupPolicyObjectFile entity in memory
                    $GroupPolicyObjectFile = [PSCustomObject]@{groupPolicyObjectFile = $_.value}

                    # Upload GroupPolicyObjectFile to Intune
                    $MigrationReportCreated = Add-IntuneEntityCollection "createMigrationReport" ($GroupPolicyObjectFile |ConvertTo-Json) $script:GraphConfiguration
                    Write-Log -Message "Import-GPOCollection: $($MigrationReportCreated.value) uploaded to Intune"
                }
                Catch
                {
                    $exception  = $_
                    Write-Log -Message "Add-IntuneEntityCollection "createMigrationReport" for id: $($key) failed. Failure: $($exception)"
                }
            }
        }
    }
    catch
    {
        # Log error
        $exception  = $_
        Write-Log -Message "Import-GPOCollection: Failure: $($exception)" -Level "Error"
        Write-Log -Message "Import-GPOCollection: IPAGlobalSettingBag: $($IPAClientConfiguration)"
        throw
    }
    Finally
    {
        # Save Configuration Bag
        $IPAClientConfiguration | ConvertTo-Json | Out-File -FilePath "$($IPAClientConfiguration.ConfigurationFolderPath)\IPAGlobalSettingBag.json"

        If (($IPAClientConfiguration.Environment -ne "pe") -and `
            !(Test-Path "$($IPAClientConfiguration.StatePath)\GPOReportXmlCollection.json"))
        {
            # Save GPOReportXmlCollection for non-PE environments
            $gpoReportXmlCollection | ConvertTo-Json -Depth 10 | Set-Content -Path "$($IPAClientConfiguration.StatePath)\GPOReportXmlCollection.json"
        }

        $sw.Stop()
        Write-Log -Message "Import-GPOCollection: Elapsed time: $($sw.Elapsed.ToString())"
    }

    return $gpoReportXmlCollection
}

# Global Configuration settings
$script:IPAClientConfiguration = $null
$script:GraphConfiguration = $null

Function Initialize-IPAClientConfiguration
{
<#
.Synopsis
  Initialize-IPAClientConfiguration - Initializes the Global settings for Intune Policy Analytics Client
.DESCRIPTION
   Initializes the Global settings for Intune Policy Analytics Client
.PARAMETER TenantAdminUPN
    AAD User Principal Name of the Intune Tenant Administrator which is required to upload the GPOs.
.PARAMETER Environment
    Type of Intune environment. Supported values:
        local
        dogfood
        selfhost
        ctip
        pe
.PARAMETER DeltaUpdate
    If set, checks if GPO already uploaded to Intune
#>
    param
    (
        [Parameter(Mandatory=$true)]
        $TenantAdminUPN,
        [Parameter(Mandatory=$false)]
        [String]$Environment = "dogfood",
        [Parameter(Mandatory=$false)]
        [Switch]$DeltaUpdate = $false
    )

    If ($null -eq $script:IPAClientConfiguration)
    {
        $GpoBackupDateTime = Get-Date -UFormat "%Y.%m.%d.%H%M"
        $IPAWorkingFolderPath = "$($env:APPDATA)\IPA"
        $LogFolderPath = "$($IPAWorkingFolderPath)\Logs"
        $script:IPAClientConfiguration = @(
            [PSCustomObject]@{`
                ConfigurationFolderPath = "$($IPAWorkingFolderPath)\Configuration"; `
                Environment = "$($Environment)"; `
                GpoBackupFolderPath = "$($IPAWorkingFolderPath)\GPO\GPOBackup"; `
                LogFilePath = "$($LogFolderPath)\IPAClient.$($TenantAdminUPN).$($Environment).$($GpoBackupDateTime).log"; `
                StatePath = "$($IPAWorkingFolderPath)\State"; `
                TenantAdminUPN = "$($TenantAdminUPN)"; `
                DeltaUpdate = "$($DeltaUpdate)"
            }
        )

        # Initialize logging
        If(!(Test-Path -Path "$($LogFolderPath)" ))
        {
            (New-Item -ItemType directory -Path $LogFolderPath) | Out-Null
        }

        If(!(Test-Path -Path "$($script:IPAClientConfiguration.LogFilePath)"))
        {
            (New-Item -Path $script:IPAClientConfiguration.LogFilePath -Force -ItemType File) | Out-Null
            Write-Log -Message "Initializing IPAClient"
        }

        # Create Configuration Folder path if necessary
        If(!(Test-Path -Path "$($script:IPAClientConfiguration.ConfigurationFolderPath)" ))
        {
            (New-Item -ItemType directory -Path $script:IPAClientConfiguration.ConfigurationFolderPath) | Out-Null
        }

        # Create State Folder path if necessary
        If(!(Test-Path -Path "$($script:IPAClientConfiguration.StatePath)" ))
        {
            (New-Item -ItemType directory -Path $script:IPAClientConfiguration.StatePath) | Out-Null
        }

        # Import pre-requisite modules
        Import-PreRequisiteModuleList

        # Initialize Graph
        $script:GraphConfiguration = Initialize-GraphConfiguration -Environment $Environment -TenantAdminUPN $TenantAdminUPN

        # Connect to Intune
        Connect-Intune $script:GraphConfiguration
    }

    return $script:IPAClientConfiguration
}

Function Remove-GPOMigrationReportCollection
{
<#
.Synopsis
   Remove-GPOMigrationReportCollection Removes Migration Report Collection from Intune
.DESCRIPTION
   Removes Migration reports for previously updated GPOs from Intune.
.PARAMETER TenantAdminUPN
    AAD User Principal Name of the Intune Tenant Administrator which is required to upload the GPOs.
.EXAMPLE
    Remove-GPOMigrationReportCollection -TenantAdminUPN "admin@IPASHAMSUA01MSIT.onmicrosoft.com"
        Removes the MigrationReports from Intune
#>
    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$TenantAdminUPN
    )

    Try
    {
        # Initialize Module
        $IPAClientConfiguration = Initialize-IPAClientConfiguration -TenantAdminUPN $TenantAdminUPN

        # Fetch the migration report
        $sw = [Diagnostics.Stopwatch]::StartNew()

        # Initialize Module
        $IPAClientConfiguration = Initialize-IPAClientConfiguration -TenantAdminUPN  $TenantAdminUPN

        # Make the Graph call to fetch the GroupPolicyMigrationReport collection
        $collectionUri = ""
        $nextUrl = $null

        # Iterate through nextlinks to get the complete set of reports
        Do
        {
            $result = Get-IntuneEntityCollection -CollectionPath $collectionUri `
                                     -Uri $nextUrl `
                                     -GraphConfiguration $script:GraphConfiguration
            $migrationReportCollection += $result.Value
            $nextUrl = $result.'@odata.nextLink'
        }
        While ($nextUrl)

        Write-Log -Message "Remove-GPOMigrationReportCollection: Found $($migrationReportCollection.Count) MigrationReports to delete.."

        # Populate the groupPolicyMigrationReports collection
        ForEach ($migrationReport in $migrationReportCollection)
        {
            Try
            {
                # Get the groupPolicySettingMappings for each migrationReport
                $groupPolicyObjectId = $migrationReport.groupPolicyObjectId
                $ou = [System.Web.HTTPUtility]::UrlDecode($migrationReport.ouDistinguishedName)
                $collectionUri = "('$($groupPolicyObjectId)_$($ou)')"
                Write-Log -Message "Remove-GPOMigrationReportCollection: collectionUri=$($collectionUri)"
                (Remove-IntuneEntityCollection -CollectionPath $collectionUri -Uri $null -GraphConfiguration $script:GraphConfiguration) | Out-Null
            }
            Catch
            {
                $exception  = $_
                Write-Log -Message "Remove-GPOMigrationReportCollection: Failure: $($exception)" -Level "Warn"
            }
        }
    }
    Catch
    {
        $exception  = $_
        Write-Log -Message "Remove-GPOMigrationReportCollection: Failure: $($exception)" -Level "Error"
        Write-Log -Message "Remove-GPOMigrationReportCollection: IPAGlobalSettingBag: $($IPAClientConfiguration)"
        Write-Log -Message "Remove-GPOMigrationReportCollection: graphConfiguration: $($graphConfiguration)"
        throw
    }
    Finally
    {
        $IPAClientConfiguration | ConvertTo-Json | Out-File -FilePath "$($IPAClientConfiguration.ConfigurationFolderPath)\IPAGlobalSettingBag.json"
        $sw.Stop()
        Write-Log -Message "Remove-GPOMigrationReportCollection: Elapsed time = $($sw.Elapsed.ToString())"
    }
}

Function Add-GPToIntuneAdmxMigratedProfile
{
<#
.Synopsis
   Migrates the GPO Migration Reports to Intune Admx Profiles
.DESCRIPTION
   Migrates the GPO Migration Reports to Intune Admx Profiles
.PARAMETER TenantAdminUPN
    The UPN of the Intune Tenant Admin.
.EXAMPLE
    Add-GPToIntuneAdmxMigratedProfile -TenantAdminUPN "admin@IPASHAMSUA01MSIT.onmicrosoft.com"
    Creates Administrative Templates Configuration Profiles to migrate the GPOs
#>

    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$TenantAdminUPN
    )

    Try
    {
        # Initialize Module
        $IPAClientConfiguration = Initialize-IPAClientConfiguration -TenantAdminUPN $TenantAdminUPN

        # Fetch the migration report
        $sw = [Diagnostics.Stopwatch]::StartNew()

        # Get Migration reports (Add filters here -SettingCategory, GPO ... )
        $GPOMigrationReportCollection = Get-GPOMigrationReportCollection -TenantAdminUPN $TenantAdminUPN -ExpandSettings    

        # Dictionary to store Parent Settings
        $AdmxSupportedParentSettings = @{} 
        
        # Dictionary to store Child Settings
        $AdmxSupportedChildSettings = @{}

        # Get all setting mappings
        $SettingMappings = $GPOMigrationReportCollection.GetEnumerator() | ForEach-Object {If(([PSCustomObject]($_.value).SettingMappings)){([PSCustomObject]($_.value).SettingMappings)}}
        
        # Populate dictionary for parent settings
        $SettingMappings | ForEach-Object { If(($null -ne $_) -and ($_.AdmxSettingDefinitionId -ne $null) -and $_.parentId -eq $null){ $AdmxSupportedParentSettings[$_.id]=$_}}

        # Populate dictionary for child settings
        $SettingMappings | ForEach-Object { If(($null -ne $_) -and ($_.AdmxSettingDefinitionId -ne $null) -and $_.parentId -ne $null){ $AdmxSupportedChildSettings[$_.id]=$_}}

        # list of settings for admx profile
        $SettingsForAdmxProfile = @()

         $AdmxSupportedParentSettings.GetEnumerator() | ForEach-Object { 
                
                $AdmxSetting = $_.Value               

                 Switch($AdmxSetting.settingValueType){
                    "Boolean" {
                        $odataType = "#microsoft.graph.omaSettingBoolean";
                        Switch($AdmxSetting.settingValue) {
                            "Enabled" { $value = $true }
                            "Disabled" { $value = $false }
                             Default { $value = $false }
                        }
                    }                    
                }

                $PresentationList = New-Object Collections.Generic.List[PSCustomObject]
                if($value -eq $true -and $AdmxSetting.childIdList.count -gt 0){                   
                   $PresentationList += (ProcessAdmxChildrenForPresentation $AdmxSetting.childIdList $AdmxSetting.AdmxSettingDefinitionId)
                }

                $odataBind = "$($GraphConfiguration.GraphBaseAddress)/$($GraphConfiguration.SchemaVersion)/deviceManagement/groupPolicyDefinitions('$($AdmxSetting.AdmxSettingDefinitionId)')"

                $SettingBody = [PSCustomObject]@{`
                   enabled = $value;`
                    presentationValues = $PresentationList;`
                    'definition@odata.bind' = $($odataBind);`
                }

                $SettingsForAdmxProfile+=($SettingBody)
        }

        try 
        {        
            $admxContainerUri = "$($GraphConfiguration.GraphBaseAddress)/$($GraphConfiguration.SchemaVersion)/deviceManagement/groupPolicyConfigurations"

            $namePrefix = "GPO to Admx Migrated-$(Get-Date -UFormat "%Y/%m/%d-%H%M")"
            Write-Log -Message "Add-GPToIntuneAdmxMigratedProfile: Creating Admx Profile: $namePrefix"

            $admxContainerBody = [PSCustomObject]@{`
                displayName = $namePrefix;`
                description = "Admx Profile created for GPOs imported in IPA";`
                roleScopeTagIds = @("0");`
            }
                
            $admxProfile = (Add-IntuneEntityCollection -Body ($admxContainerBody | ConvertTo-Json) -uri $admxContainerUri -GraphConfiguration $script:GraphConfiguration)
            
            Write-Log -Message "Profile created successfully"
            Write-Log -Message "Id of Admx Profile created. $($admxProfile.Id)"
        }
        catch 
        {
            $exception  = $_
            Write-Host $exception
            Write-Log -Message "Add-GPToIntuneAdmxMigratedProfile: Failure: $($exception)"
            throw
        }
        
        try
        {

            $admxUpdateContainerBody = [PSCustomObject]@{`
                    added = $SettingsForAdmxProfile;`
                    updated = @();`
                    deletedIds = @();`
                    }

            $unescapedContainerBody = ( ConvertTo-Json $admxUpdateContainerBody -depth 10 | % { [regex]::Unescape($_) })

            $admxUpdateContainerUri =   "$($GraphConfiguration.GraphBaseAddress)/$($GraphConfiguration.SchemaVersion)/deviceManagement/groupPolicyConfigurations('$($admxProfile.Id)')/updateDefinitionValues"
            Write-Log -Message "Uri to update profile: . $($admxUpdateContainerUri)"
            $updatedConfiguration = (Add-IntuneEntityCollection -Body $unescapedContainerBody -uri $admxUpdateContainerUri -GraphConfiguration $script:GraphConfiguration)
            Write-Log -Message "Profile updated successfully"
        }
        catch
        {
            $exception  = $_
            Write-Host $exception
            Write-Log -Message "Add-GPToIntuneAdmxMigratedProfile: Failure: $($exception)"  
            throw
        }      

    }
    Catch
    {
        $exception  = $_
        Write-Host $exception
        Write-Log -Message "Add-GPToIntuneAdmxMigratedProfile: Failure: $($exception)" -Level "Error"
        Write-Log -Message "Add-GPToIntuneAdmxMigratedProfile: IPAGlobalSettingBag: $($IPAClientConfiguration)"
        throw
    }
    Finally
    {
        $IPAClientConfiguration | ConvertTo-Json | Out-File -FilePath "$($IPAClientConfiguration.ConfigurationFolderPath)\IPAGlobalSettingBag.json"
        $sw.Stop()
        Write-Log -Message "Add-GPToIntuneAdmxMigratedProfile: Elapsed time = $($sw.Elapsed.ToString())"
    }
   

}

Function ProcessAdmxChildrenForPresentation ($ChildIdList, $ParentAdmxDefinitionId) {
<#.Synopsis
   Processes the child settings and creates List of Presentation objects for ADMX profile
.DESCRIPTION
   Processes the child settings and creates List of Presentation objects for ADMX profile
.PARAMETER $ChildIdList
    List of children to be processed.
.PARAMETER  $ParentAdmxDefinitionId
    Admx id of the parent.
#>

   $listOfSettings = New-Object Collections.Generic.List[PSCustomObject]
   $ChildIdList.GetEnumerator() | ForEach-Object{
        $childSetting = $AdmxSupportedChildSettings[$_]
        Switch($childSetting.settingValueType){
            "Boolean" {

                $odataType = "#microsoft.graph.groupPolicyPresentationValueBoolean";

                Switch($childSetting.settingValue) {
                    "Enabled" { $value = $true }
                    "Disabled" { $value = $false }
                        "true" { $value = $true }
                    "false" { $value = $false }
                        Default { $value = $false }
                }
            }
            "UInt32" {
                $odataType = "#microsoft.graph.groupPolicyPresentationValueDecimal";
                $value = $($childSetting.settingValue) -as [int];
            }
            "UInt64"{
                $odataType = "#microsoft.graph.groupPolicyPresentationValueLongDecimal";
                $value = $($childSetting.settingValue) -as [int64];
            }
            "String" {
                    $odataType = "#microsoft.graph.groupPolicyPresentationValueText";
                    $value = $($childSetting.settingValue);
            }
            "DropDownListEnum" {
                    $odataType = "#microsoft.graph.groupPolicyPresentationValueText";
                    $value = $($childSetting.settingValue);

            }
            "JsonListBox" {
                $odataType = "#microsoft.graph.groupPolicyPresentationValueList";

                $value = New-Object Collections.Generic.List[PSCustomObject]
                $listArray = ConvertFrom-Json $childSetting.settingValue;
                $listArray.GetEnumerator() | ForEach-Object { 

                    if($_.name -eq $null){
                        $listitem = [PSCustomObject]@{`
                                "name" = $_.Data;`                    
                        }

                        $value += $listitem
                    }
                    else
                    {
                        $listitem = [PSCustomObject]@{`
                                "name" = $_.Name;`
                                "value" = $_.Data                    
                        }

                        $value += $listitem
                    }                            
                }
                        
            }
            "JsonDisplayField" {
                    $odataType = "#microsoft.graph.omaSettingString";
                    $value = $($childSetting.settingValue);
            }
            "JsonMultiStrings" {
                    $odataType = "#microsoft.graph.omaSettingString";
                    $value = $($childSetting.settingValue);
            }
                   
                   
            Default {
                    $odataType = $($childSetting.settingValueType);
                    $value = $($childSetting.settingValue);
            }
        }

        $valuePropertyStr = "value"

        if($value.GetType().ToString() -eq "System.Object[]")
        {
            $valuePropertyStr = "values"
        }
                
        $childSettingPresentation = [PSCustomObject]@{`
            "$valuePropertyStr" = $value;`
            '@odata.type' = $odataType;`
            'presentation@odata.bind' = "$($GraphConfiguration.GraphBaseAddress)/$($GraphConfiguration.SchemaVersion)/deviceManagement/groupPolicyDefinitions('$($ParentAdmxDefinitionId)')/presentations('$($childSetting.AdmxSettingDefinitionId)')";`
        }

        $listOfSettings += ($childSettingPresentation)
    
   }
  
   Write-Output $listOfSettings
}

Function Add-GPToIntuneMigratedProfile
{
<#
.Synopsis
   Migrates the GPO Migration Reports for GPOs uploaded to Intune
.DESCRIPTION
   Migrates the GPO Migration Reports for GPOs uploaded to Intune
.PARAMETER TenantAdminUPN
    The UPN of the Intune Tenant Admin.
.EXAMPLE
    Add-GPToIntuneMigratedProfile -TenantAdminUPN "admin@IPASHAMSUA01MSIT.onmicrosoft.com"
    Creates Intune Profiles to migrate the GPOs
#>

    [cmdletbinding()]
    param(
        [Parameter(Mandatory=$true)]
        [string]$TenantAdminUPN
    )

    Try
    {
        # Fetch the migration report
        $sw = [Diagnostics.Stopwatch]::StartNew()

        # Initialize Module 
        $IPAClientConfiguration = Initialize-IPAClientConfiguration -TenantAdminUPN  $TenantAdminUPN -Environment pe

        # Get Migration reports (Add filters here -SettingCategory, GPO ... )
        $GPOMigrationReportCollection = Get-GPOMigrationReportCollection -TenantAdminUPN $TenantAdminUPN -ExpandSettings

        # Create Custom OMA Uri profiles
        $DCProfileCollection = @()
        $DCUri = "$($GraphConfiguration.GraphBaseAddress)/$($GraphConfiguration.SchemaVersion)/deviceManagement/deviceConfigurations"

        $GPOMigrationReportCollection.GetEnumerator() | ForEach-Object {
            $gpoMigrationReport = $_
            $MDMSupportedSettings = @{}
            $SettingMappings = $gpoMigrationReport | ForEach-Object {If(([PSCustomObject]($_.value).SettingMappings)){([PSCustomObject]($_.value).SettingMappings)}}
            $SettingMappings | ForEach-Object {If(($null -ne $_) -and ($_.isMdmSupported -eq $True)){$MDMSupportedSettings[$_.id]=$_}}

            $omaSettings = @()
            $omaSettingCount = 0
            $MDMSupportedSettings.GetEnumerator() | ForEach-Object {
                $MDMSupportedSetting = $_.Value
                Switch($MDMSupportedSetting.settingValueType){
                    "Boolean" {
                        $odataType = "#microsoft.graph.omaSettingBoolean";
                        Switch($MDMSupportedSetting.settingValue) {
                            "Enabled" { $value = "true" }
                            "true" { $value = "true" }
                            Default { $value = "false" }
                        }
                    }
                    "JsonDisplayField" {
                         $odataType = "#microsoft.graph.omaSettingString";
                         $value = $($MDMSupportedSetting.settingValue);
                    }
                    "JsonMultiStrings" {
                         $odataType = "#microsoft.graph.omaSettingString";
                         $value = $($MDMSupportedSetting.settingValue);
                    }
                    "JsonListBox" {
                        $odataType = "#microsoft.graph.omaSettingString";
                        $value = $($MDMSupportedSetting.settingValue);
                   }
                    "String" {
                         $odataType = "#microsoft.graph.omaSettingString";
                         $value = $($MDMSupportedSetting.settingValue);
                    }
                    "UInt32" {
                        $odataType = "#microsoft.graph.omaSettingInteger";
                        $value = $($MDMSupportedSetting.settingValue);
                    }
                    Default {
                         $odataType = $($MDMSupportedSetting.settingValueType);
                         $value = $($MDMSupportedSetting.settingValue);
                    }
                }

                $omaSetting =  [PSCustomObject]@{`
                    displayName = $($MDMSupportedSetting.settingName);`
                    omaUri =  $($MDMSupportedSetting.mdmSettingUri);`
                    '@odata.type' = $($odataType);`
                    value = $($value);`
                }

                #
                # Bugbug: only 1000 settings are allowed
                #
                if ($omaSettingCount -lt 1000)
                {
                    $omaSettings += $omaSetting                    
                }
                else
                {
                    Write-Log -Message "Add-GPToIntuneMigratedProfile: $($MDMSupportedSetting.settingName) skipped..."
                }

                $omaSettingCount++
            }

            $namePrefix = "GPAnalytics-Migrated-$(Get-Date -UFormat "%Y/%m/%d-%H%M")"
            $customOMAProfile = [PSCustomObject]@{`
                displayName = "$($namePrefix)-$($gpoMigrationReport.Name)";`
                description = $($gpoMigrationReport.Value.OU);`
                '@odata.type' = "#microsoft.graph.windows10CustomConfiguration";`
                omaSettings = $omaSettings;`
                }

            try 
            {
                Write-Log -Message "Add-GPToIntuneMigratedProfile: Creating Profile: $($customOMAProfile.displayName)... "
                $DCProfileCollection+= (Add-IntuneEntityCollection -Body ($customOMAProfile | ConvertTo-Json) -uri $DCUri -GraphConfiguration $script:GraphConfiguration)
            }
            catch 
            {
                $exception  = $_
                Write-Log -Message "Add-GPToIntuneMigratedProfile: Failure: $($exception)"
            }            
            
            #
            # Sleep a bit to prevent throttling
            #
            Start-Sleep -Milliseconds 1000
        }
    }
    Catch
    {
        $exception  = $_
        Write-Log -Message "Add-GPToIntuneMigratedProfile: Failure: $($exception)" -Level "Error"
        Write-Log -Message "Add-GPToIntuneMigratedProfile: IPAGlobalSettingBag: $($IPAClientConfiguration)"
        throw
    }
    Finally
    {
        $IPAClientConfiguration | ConvertTo-Json | Out-File -FilePath "$($IPAClientConfiguration.ConfigurationFolderPath)\IPAGlobalSettingBag.json"
        $sw.Stop()
        Write-Log -Message "Add-GPToIntuneMigratedProfile: Elapsed time = $($sw.Elapsed.ToString())"
    }
}


Export-ModuleMember -Function Add-GPToIntuneAdmxMigratedProfile
Export-ModuleMember -Function Add-ChromeToEdgeMigratedProfile
Export-ModuleMember -Function Add-GPToIntuneMigratedProfile
Export-ModuleMember -Function Get-GPOMigrationReportCollection
Export-ModuleMember -Function Get-MigrationReadinessReport
Export-ModuleMember -Function Update-MigrationReadinessReport
Export-ModuleMember -Function Import-GPOCollection
Export-ModuleMember -Function Initialize-IPAClientConfiguration
Export-ModuleMember -Function Remove-GPOMigrationReportCollection
#endregion Cmdlets

#region Configuration Utilities
<#
.Synopsis
  Import-PreRequisiteModuleList - Checks if RSAT is installed or not. If not, prompts user to install.
.DESCRIPTION
   Checks if RSAT is installed or not. If not, prompts user to install.
#>
Function Import-PreRequisiteModuleList
{
    Try
    {
        If (!(Get-module -ListAvailable -Name GroupPolicy))
        {
            $ShouldInstallRSATModule = "Y"
            $ShouldInstallRSATModule = Read-Host -Prompt "RSAT Module not installed. Install it? Y/N Default:[$($ShouldInstallRSATModule)]"

            # Install RSAT only if consented
            If ($ShouldInstallRSATModule -eq "Y")
            {
                # Check for Windows10
                $osVersion =[System.Environment]::OSVersion.Version
                If ($osVersion.Major -ge 10)
                {
                    $RSATx86 = 'https://download.microsoft.com/download/1/D/8/1D8B5022-5477-4B9A-8104-6A71FF9D98AB/WindowsTH-RSAT_WS2016-x86.msu'
                    $RSATx64 = 'https://download.microsoft.com/download/1/D/8/1D8B5022-5477-4B9A-8104-6A71FF9D98AB/WindowsTH-RSAT_WS2016-x64.msu'

                    switch ($env:PROCESSOR_ARCHITECTURE)
                    {
                        'x86' {$RSATDownloadUri = $RSATx86}
                        'AMD64' {$RSATDownloadUri = $RSATx64}
                    }

                    $RSATKBDownloadFileName = $RSATDownloadUri.Split('/')[-1]

                    Write-Log -Message "Downloading RSAT from $($RSATDownloadUri) to $($RSATKBDownloadFileName)"
                    Invoke-WebRequest -Uri $RSATDownloadUri -UseBasicParsing -OutFile "$env:TEMP\$RSATKBDownloadFileName"

                    Write-Log -Message "Start-Process -FilePath wusa.exe -ArgumentList $env:TEMP\$($RSATKBDownloadFileName) /quiet /promptrestart /log"
                    Start-Process -FilePath wusa.exe -ArgumentList "$env:TEMP\$($RSATKBDownloadFileName) /quiet /promptrestart /log" -Wait -Verbose
                }
                Else
                {
                    Write-Log -Message "RSAT install is supported only on Windows 10 and above" -Level "Error"
                    throw
                }
            }
        }

        Import-Module ActiveDirectory
        Import-Module GroupPolicy
    }
    Catch
    {
        $exception  = $_
        Write-Log -Message "Import-PreRequisiteModuleList failed. Failure: $($exception)" -Level "Error"
        throw
    }
}
#endregion

#region Logging Utilities
function Write-Log
{
    Param
    (
        [Parameter(Mandatory=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$Message,
        [Parameter(Mandatory=$false)]
        [string]$Level="Info",
        [Parameter(Mandatory=$false)]
        [string]$LogPath = "$($script:IPAClientConfiguration.LogFilePath)"
    )

    # Format Date for our Log File
    $currentDateTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

    # Write message to error, warning, or verbose pipeline and specify $LevelText
    switch ($Level)
    {
        'Error'
        {
            $logLine = "[$($currentDateTime)] ERROR: $($Message)"
            Write-Error $logLine
        }       
        Default
        {
            $logLine = "[$($currentDateTime)] INFO: $($Message)"
            Write-Progress -Activity "IntunePolicyAnalytics" -PercentComplete -1 -Status $logLine
        }
    }

    # Write log entry to $Path
    $logLine | Out-File -FilePath $LogPath -Append
}
#endregion

#region Graph Utilities
function Initialize-GraphConfiguration
{
<#
.Synopsis
    Put-GraphConfiguration: Initializes the Graph settings
.PARAMETER Environment
    Type of Intune environment. Supported values:
        local
        dogfood
        selfhost
        ctip
        pe
.PARAMETER TenantUPN
    UPN of the Intune Tenant Admin
#>
    param
    (
        [Parameter(Mandatory=$true)]
        [String]$Environment,
        [Parameter(Mandatory=$true)]
        $TenantAdminUPN
    )

    # Graph AuthHeader primitives
    $userUpn = New-Object "System.Net.Mail.MailAddress" -ArgumentList $TenantAdminUPN
    $tenant = $userUpn.Host

    # App IDs to use
    $IntunePowerShellclientId = "d1ddf0e4-d672-4dae-b554-9d5bdfd93547" # Official PowerShell App
    $IntunePolicyAnalyticsClientId = "a1357584-810b-42e6-a9e6-4e7237ccbcea" # PPE IPA PowerShell App

    # RedirectUri to use
    $IntunePowerShellRedirectUri = "urn:ietf:wg:oauth:2.0:oob"
    $IntunePolicyAnalyticsRedirectUri = "https://login.microsoftonline.com/common/oauth2/nativeclient"

    # Graph configuration settings per environment
    $GraphConfiguration = @(       
        [PSCustomObject]@{Environment = "dogfood"; `
            AuthUrl = "https://login.windows-ppe.net/$($tenant)"; `
            ResourceId = "https://graph.microsoft-ppe.com"; `
            GraphBaseAddress = "https://graph.microsoft-ppe.com"; `
            IPARoute = "deviceManagement/groupPolicyMigrationReports"; `
            AppId = "$($IntunePolicyAnalyticsClientId)"; `
            RedirectLink = "$($IntunePolicyAnalyticsRedirectUri)"; `
            SchemaVersion = "testppebeta_intune_onedf"; `
            TenantAdminUPN = "$($TenantAdminUPN)";
            AuthHeader = $null;
            platformParameters = $null;
            userId = $null;
            authContext = $null}
        [PSCustomObject]@{Environment = "selfhost"; `
            AuthUrl = "https://login.microsoftonline.com/$($tenant)"; `
            ResourceId = "https://canary.graph.microsoft.com"; `
            GraphBaseAddress = "https://canary.graph.microsoft.com"; `
            IPARoute = "deviceManagement/groupPolicyMigrationReports"; `
            AppId = "$($IntunePowerShellclientId)"; `
            RedirectLink = "$($IntunePowerShellRedirectUri)"; `
            SchemaVersion = "testprodbeta_intune_sh"; `
            TenantAdminUPN = "$($TenantAdminUPN)";
            AuthHeader = $null;
            platformParameters = $null;
            userId = $null;
            authContext = $null}
        [PSCustomObject]@{Environment = "ctip"; `
            AuthUrl = "https://login.microsoftonline.com/$($tenant)"; `
            ResourceId = "https://canary.graph.microsoft.com"; `
            GraphBaseAddress = "https://canary.graph.microsoft.com"; `
            IPARoute = "deviceManagement/groupPolicyMigrationReports"; `
            AppId = "$($IntunePowerShellclientId)"; `
            RedirectLink = "$($IntunePowerShellRedirectUri)"; `
            SchemaVersion = "testprodbeta_intune_ctip"; `
            TenantAdminUPN = "$($TenantAdminUPN)";
            AuthHeader = $null;
            platformParameters = $null;
            userId = $null;
            authContext = $null}
        [PSCustomObject]@{Environment = "pe_canary"; `
            AuthUrl = "https://login.microsoftonline.com/$($tenant)"; `
            ResourceId = "https://canary.graph.microsoft.com"; `
            GraphBaseAddress = "https://canary.graph.microsoft.com"; `
            IPARoute = "deviceManagement/groupPolicyMigrationReports"; `
            AppId = "$($IntunePowerShellclientId)"; `
            RedirectLink = "$($IntunePowerShellRedirectUri)"; `
            SchemaVersion = "testprodbeta_intune_ctip"; `
            TenantAdminUPN = "$($TenantAdminUPN)";
            AuthHeader = $null;
            platformParameters = $null;
            userId = $null;
            authContext = $null}
        [PSCustomObject]@{Environment = "pe"; `
            AuthUrl = "https://login.microsoftonline.com/$($tenant)"; `
            ResourceId = "https://graph.microsoft.com"; `
            GraphBaseAddress = "https://graph.microsoft.com"; `
            IPARoute = "deviceManagement/groupPolicyMigrationReports"; `
            AppId = "$($IntunePowerShellclientId)"; `
            RedirectLink = "$($IntunePowerShellRedirectUri)"; `
            SchemaVersion = "stagingbeta"; `
            TenantAdminUPN = "$($TenantAdminUPN)";
            AuthHeader = $null;
            platformParameters = $null;
            userId = $null;
            authContext = $null}
        [PSCustomObject]@{Environment = "pe_fxp"; `
            AuthUrl = "https://login.microsoftonline.us/$($tenant)"; `
            ResourceId = "https://graph.microsoft.us"; `
            GraphBaseAddress = "https://graph.microsoft.us"; `
            IPARoute = "deviceManagement/groupPolicyMigrationReports"; `
            AppId = "$($IntunePowerShellclientId)"; `
            RedirectLink = "$($IntunePowerShellRedirectUri)"; `
            SchemaVersion = "beta"; `
            TenantAdminUPN = "$($TenantAdminUPN)";
            AuthHeader = $null;
            platformParameters = $null;
            userId = $null;
            authContext = $null}
        [PSCustomObject]@{Environment = "pe_cnb"; `
            AuthUrl = "https://login.partner.microsoftonline.cn/$($tenant)"; `
            ResourceId = "https://microsoftgraph.chinacloud.cn"; `
            GraphBaseAddress = "https://microsoftgraph.chinacloud.cn"; `
            IPARoute = "deviceManagement/groupPolicyMigrationReports"; `
            AppId = "$($IntunePowerShellclientId)"; `
            RedirectLink = "$($IntunePowerShellRedirectUri)"; `
            SchemaVersion = "beta"; `
            TenantAdminUPN = "$($TenantAdminUPN)";
            AuthHeader = $null;
            platformParameters = $null;
            userId = $null;
            authContext = $null}     
    )

    $graphConfiguration = ($GraphConfiguration | Where-Object {$_.Environment -eq "$($Environment)"})

#region AAD configurations
    $AadModule = Get-Module -Name "AzureAD" -ListAvailable

    If ($null -eq $AadModule)
    {
        Write-Log -Message "AzureAD PowerShell module not found, looking for AzureADPreview"
        $AadModule = Get-Module -Name "AzureADPreview" -ListAvailable
    }

    If ($null -eq $AadModule)
    {
        Install-Module AzureADPreview
    }

    $adal = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.dll"
    $adalforms = Join-Path $AadModule.ModuleBase "Microsoft.IdentityModel.Clients.ActiveDirectory.Platform.dll"

    # Load ADAL types
    [System.Reflection.Assembly]::LoadFrom($adal) | Out-Null
    [System.Reflection.Assembly]::LoadFrom($adalforms) | Out-Null

    $graphConfiguration.platformParameters = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.PlatformParameters" -ArgumentList "Auto"
    $graphConfiguration.userId = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.UserIdentifier" -ArgumentList ($graphConfiguration.TenantAdminUPN, "OptionalDisplayableId")
    $graphConfiguration.authContext = New-Object "Microsoft.IdentityModel.Clients.ActiveDirectory.AuthenticationContext" -ArgumentList $graphConfiguration.AuthUrl
#endregion

    return $graphConfiguration
}

<#
.Synopsis
  Connect-Intune - Connect to Microsoft Intune
.DESCRIPTION
   Get an auth token from AAD.
.PARAMETER GraphConfiguration
   The UPN of the user that shall be used to get an auth token for.
#>
Function Connect-Intune
{
    param
    (
        [Parameter(Mandatory=$true)]
        [PSCustomObject]$graphConfiguration
    )

    Try
    {
        # BUGBUG: We are directly doing auth via ADAL till we figure how to call
        # Connect-MSGraph correctly
        (Get-AuthHeader -graphConfiguration $graphConfiguration) | Out-Null
    }
    Catch
    {
        $exception  = $_
        Write-Log -Message "Connect-Intune Failed. Failure: $($exception)" -Level "Error"
        throw
    }
}

#
# CloneObject: Clones the input object
#
function CloneObject($object)
{
    $stream = New-Object IO.MemoryStream;
    $formatter = New-Object Runtime.Serialization.Formatters.Binary.BinaryFormatter;
    $formatter.Serialize($stream, $object);
    $stream.Position = 0;
    $formatter.Deserialize($stream);
}

<#
.Synopsis
  Get-AuthHeader - Get an auth token from AAD.
.DESCRIPTION
   Get an auth token from AAD.
.PARAMETER GraphConfiguration
   The UPN of the user that shall be used to get an auth token for.
#>
Function Get-AuthHeader
{
    param
    (
        [Parameter(Mandatory=$true)]
        [PSCustomObject]$graphConfiguration
    )

    # Get the AuthToken from AAD
    $currentDateTime = (Get-Date).ToUniversalTime()
    $tokenExpiresOn = ($graphConfiguration.AuthHeader.ExpiresOn.datetime - $currentDateTime).Minutes

    If ($tokenExpiresOn -le 0)
    {        
        $authResult = $graphConfiguration.authContext.AcquireTokenAsync($graphConfiguration.ResourceId,`
                                                        $graphConfiguration.AppId, `
                                                        $graphConfiguration.RedirectLink, `
                                                        $graphConfiguration.platformParameters, `
                                                        $graphConfiguration.userId).Result

        # Creating header for Authorization token
        $graphConfiguration.AuthHeader = @{
            'Content-Type'='application/json'
            'Authorization'="Bearer " + $authResult.AccessToken
            'ExpiresOn'=$authResult.ExpiresOn
            }
    }

    return $graphConfiguration.AuthHeader
}

<#
.Synopsis
  Add-IntuneEntityCollection - Makes a POST request to Intune
.DESCRIPTION
   Make a HTTP POST request to Intune
.PARAMETER CollectionPath
   The Collection path to the Graph Entitie that needs to be fetched.
.PARAMETER Body
   The Json serialized Body of the HTTP POST call
.PARAMETER GraphConfiguration
    Graph Configuration
.PARAMETER Uri
    Graph Base route to use. Defaults to IPA Graph base route.
#>
function Add-IntuneEntityCollection
{
    param
    (
        [Parameter(Mandatory=$false)]
        $CollectionPath,
        [Parameter(Mandatory=$true)]
        $Body,
        [Parameter(Mandatory=$true)]
        $GraphConfiguration,
        [Parameter(Mandatory=$false)]
        $uri = $null
    )

    If ($null -eq $uri)
    {
        $uri = "$($GraphConfiguration.GraphBaseAddress)/$($GraphConfiguration.SchemaVersion)/$($GraphConfiguration.IPARoute)/$($collectionPath)";
    }

    $clonedHeaders = CloneObject (Get-AuthHeader $graphConfiguration);
    $clonedHeaders["content-length"] = $Body.Length;
    $clonedHeaders["api-version"] = "$($script:ApiVersion)";

    Try
    {
        #Write-Log -Message "Add-IntuneEntityCollection: $($uri) -Method Post -Body $($body)"
        $response = Invoke-RestMethod $uri -Method Post -Headers $clonedHeaders -Body $body;
    }
    Catch
    {
        $exception  = $_
        Write-Log -Message "Add-IntuneEntityCollection: Failed. CollectionPath:$($CollectionPath). Failure: $($exception)" -Level "Warn"
        throw
    }

    return $response;
}

<#
.Synopsis
  Get-IntuneEntityCollection - Makes a GET request to Intune
.DESCRIPTION
   Make a HTTP GET request to Intune
.PARAMETER CollectionPath
   The Collection path to the Graph Entitie that needs to be fetched.
#>
Function Get-IntuneEntityCollection
{
    param
    (
        [Parameter(Mandatory=$true)]
        $CollectionPath,
        [Parameter(Mandatory=$false)]
        $Uri,
        [Parameter(Mandatory=$true)]
        $GraphConfiguration
    )

    If ($null -eq $Uri)
    {
        $uri = "$($GraphConfiguration.GraphBaseAddress)/$($GraphConfiguration.SchemaVersion)/$($GraphConfiguration.IPARoute)$($collectionPath)";
    }

    $clonedHeaders = CloneObject (Get-AuthHeader $graphConfiguration);
    $clonedHeaders["api-version"] = "$($script:ApiVersion)";

    Try
    {
        $response = Invoke-RestMethod $uri -Method Get -Headers $clonedHeaders
    }
    Catch
    {
        $exception  = $_
        Switch ($exception)
        {
            'not found'
            {
                $response = $null
                Write-Log -Message "Get-IntuneEntityCollection: GET $($uri) Failed. Failure: $($exception)"
            }
            Default
            {
                $response = $exception
                Write-Log -Message "Get-IntuneEntityCollection: GET $($uri) Failed. Failure: $($exception)" -Level "Warn"
                throw
            }
        }
    }

    return $response;
}

<#
.Synopsis
  Remove-IntuneEntityCollection - Makes a DELETE request to Intune
.DESCRIPTION
   Make a HTTP DELETE request to Intune
.PARAMETER CollectionPath
   The Collection path to the Graph Entitie that needs to be fetched.
#>
Function Remove-IntuneEntityCollection
{
    param
    (
        [Parameter(Mandatory=$true)]
        $CollectionPath,
        [Parameter(Mandatory=$false)]
        $Uri,
        [Parameter(Mandatory=$true)]
        $GraphConfiguration
    )

    If ($null -eq $Uri)
    {
        $uri = "$($GraphConfiguration.GraphBaseAddress)/$($GraphConfiguration.SchemaVersion)/$($GraphConfiguration.IPARoute)$($collectionPath)";
    }

    $clonedHeaders = CloneObject (Get-AuthHeader $graphConfiguration);
    $clonedHeaders["api-version"] = "$($script:ApiVersion)";

    Try
    {
        $response = Invoke-RestMethod $uri -Method Delete -Headers $clonedHeaders
    }
    Catch
    {
        $exception  = $_
        Write-Log -Message "Remove-IntuneEntityCollection: DELETE $($uri) Failed. Failure: $($exception)" -Level "Warn"
        throw
    }

    return $response;
}
#endregion

#region AD Utilities
<#
.Synopsis
 Get-GPOReportXmlCollectionFromAD: Calls Get-GPOReport for all GPO discovered for the given domain.
.DESCRIPTION
   Calls Get-GPOReport for all GPO discovered for the given domain and returns the Xml report collection.
.PARAMETER ADDomain
   The local AD Domain for which the GPO collection is fetched.
   Defaults to the local AD Domain for the client on which this script is run on.
.PARAMETER OUFilter
   Use OUFilter to constrain the GP Objects to the OU in consideration.
   Specifies a query string that retrieves Active Directory objects. This string uses the PowerShell Expression Language syntax.
#>
Function Get-GPOReportXmlCollectionFromAD
{
    param
    (
        [Parameter(Mandatory=$true)]
        [String]$ADDomain,
        [Parameter(Mandatory=$true)]
        [String]$OUFilter
    )

    $gpoReportXmlCollection = @{}
    $gpoMigrationReportCollection = @{}

    If ($script:IPAClientConfiguration.DeltaUpdate -eq $true)
    {
        # If set, check Intune before posting GPOReportXml
        $gpoMigrationReportCollection = Get-GPOMigrationReportCollection -TenantAdminUPN $script:IPAClientConfiguration.TenantAdminUPN
    }

    Try
    {
        # Get the OU collection for the given AD Domain
        $ouCollection = Get-ADOrganizationalUnit -Filter $OUFilter -Server $ADDomain `
            | Select-Object Name,DistinguishedName,LinkedGroupPolicyObjects `
            | Where-Object {$_.LinkedGroupPolicyObjects -ne '{}'}

        # Get GPO backups for each OU
        ForEach ($ou in $ouCollection)
        {
            # Get the GPO collection linked to this ou
            # Each element in $gpoCollection is a fully qualified LDAP name of the Linked GPOs.
            # For example: "cn={A7A7EA17-BF74-4120-ADC3-14FD1DE01B34},cn=policies,cn=system,DC=redmond,DC=corp,DC=microsoft,DC=com"
            $GUIDRegex = "[a-zA-Z0-9]{8}[-][a-zA-Z0-9]{4}[-][a-zA-Z0-9]{4}[-][a-zA-Z0-9]{4}[-][a-zA-Z0-9]{12}"
            $gpoCollection = $ou | Select-Object -ExpandProperty LinkedGroupPolicyObjects
            Write-Log -Message "Get-GPOReportXmlCollectionFromAD: $($ou.DistinguishedName). GPO Count=$($gpoCollection.Count)"

            # Backup GPO from LinkedGroupPolicyObjects
            ForEach ($gpo in $gpoCollection)
            {
                $result = [Regex]::Match($gpo,$GUIDRegex);
                If ($result.Success)
                {
                    # Assign the GPO Guid
                    $gpoGuid = $result.Value

                    Try
                    {
                        # Backup a GPO as Xml in memory if not previously uploaded
                        $gpoReportXmlKey = "$($gpoGuid)_$($ou.DistinguishedName)"

                        If (!$gpoMigrationReportCollection.Contains($gpoReportXmlKey))
                        {
                            [Xml]$gpoReportXml = (Get-GPOReport -Guid $gpoGuid -ReportType Xml -Domain $ADDomain -ErrorAction Stop)
                            $bytes = [System.Text.Encoding]::UNICODE.GetBytes($gpoReportXml.InnerXml)
                            $encodedText = [Convert]::ToBase64String($bytes)
                            $gpoReportXmlCollection.Add($gpoReportXmlKey, [PSCustomObject]@{ouDistinguishedName = $ou.DistinguishedName; content = $encodedText})
                            Write-Log -Message "Get-GPOReportXmlCollectionFromAD:  Backed up GPO Guid=$($gpoGuid), $($ou.DistinguishedName)"
                        }
                        Else
                        {
                            Write-Log -Message "Get-GPOReportXmlCollectionFromAD:  GPO Guid=$($gpoGuid), $($ou.DistinguishedName) previously uploaded"
                        }
                    }
                    Catch
                    {
                        $exception  = $_
                        Write-Log -Message "Get-GPOReportXmlCollectionFromAD:Unable to get $($gpo) xml backup in memory. Failure: $($exception)" -Level "Warn"
                        # We continue to next GPO
                    }
                }
            }
        }

        Write-Log -Message "Get-GPOReportXmlCollectionFromAD: $($gpoReportXmlCollection.Count) GPOs found"
    }
    catch
    {
        # Log error
        $exception  = $_
        Write-Log -Message "Get-GPOReportXmlCollectionFromAD: Failure: ($exception)" -Level "Error"
        throw
    }

    return $gpoReportXmlCollection
}
#endregion AD Utilities
# SIG # Begin signature block
# MIIjewYJKoZIhvcNAQcCoIIjbDCCI2gCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCDPqmL1PNd1w58P
# 87yvL4ST5CuQ/Upe8pqCqTdohQBnn6CCDXYwggX0MIID3KADAgECAhMzAAABhk0h
# daDZB74sAAAAAAGGMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjAwMzA0MTgzOTQ2WhcNMjEwMzAzMTgzOTQ2WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQC49eyyaaieg3Xb7ew+/hA34gqzRuReb9svBF6N3+iLD5A0iMddtunnmbFVQ+lN
# Wphf/xOGef5vXMMMk744txo/kT6CKq0GzV+IhAqDytjH3UgGhLBNZ/UWuQPgrnhw
# afQ3ZclsXo1lto4pyps4+X3RyQfnxCwqtjRxjCQ+AwIzk0vSVFnId6AwbB73w2lJ
# +MC+E6nVmyvikp7DT2swTF05JkfMUtzDosktz/pvvMWY1IUOZ71XqWUXcwfzWDJ+
# 96WxBH6LpDQ1fCQ3POA3jCBu3mMiB1kSsMihH+eq1EzD0Es7iIT1MlKERPQmC+xl
# K+9pPAw6j+rP2guYfKrMFr39AgMBAAGjggFzMIIBbzAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUhTFTFHuCaUCdTgZXja/OAQ9xOm4w
# RQYDVR0RBD4wPKQ6MDgxHjAcBgNVBAsTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEW
# MBQGA1UEBRMNMjMwMDEyKzQ1ODM4NDAfBgNVHSMEGDAWgBRIbmTlUAXTgqoXNzci
# tW2oynUClTBUBgNVHR8ETTBLMEmgR6BFhkNodHRwOi8vd3d3Lm1pY3Jvc29mdC5j
# b20vcGtpb3BzL2NybC9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3JsMGEG
# CCsGAQUFBwEBBFUwUzBRBggrBgEFBQcwAoZFaHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraW9wcy9jZXJ0cy9NaWNDb2RTaWdQQ0EyMDExXzIwMTEtMDctMDguY3J0
# MAwGA1UdEwEB/wQCMAAwDQYJKoZIhvcNAQELBQADggIBAEDkLXWKDtJ8rLh3d7XP
# 1xU1s6Gt0jDqeHoIpTvnsREt9MsKriVGKdVVGSJow1Lz9+9bINmPZo7ZdMhNhWGQ
# QnEF7z/3czh0MLO0z48cxCrjLch0P2sxvtcaT57LBmEy+tbhlUB6iz72KWavxuhP
# 5zxKEChtLp8gHkp5/1YTPlvRYFrZr/iup2jzc/Oo5N4/q+yhOsRT3KJu62ekQUUP
# sPU2bWsaF/hUPW/L2O1Fecf+6OOJLT2bHaAzr+EBAn0KAUiwdM+AUvasG9kHLX+I
# XXlEZvfsXGzzxFlWzNbpM99umWWMQPTGZPpSCTDDs/1Ci0Br2/oXcgayYLaZCWsj
# 1m/a0V8OHZGbppP1RrBeLQKfATjtAl0xrhMr4kgfvJ6ntChg9dxy4DiGWnsj//Qy
# wUs1UxVchRR7eFaP3M8/BV0eeMotXwTNIwzSd3uAzAI+NSrN5pVlQeC0XXTueeDu
# xDch3S5UUdDOvdlOdlRAa+85Si6HmEUgx3j0YYSC1RWBdEhwsAdH6nXtXEshAAxf
# 8PWh2wCsczMe/F4vTg4cmDsBTZwwrHqL5krX++s61sLWA67Yn4Db6rXV9Imcf5UM
# Cq09wJj5H93KH9qc1yCiJzDCtbtgyHYXAkSHQNpoj7tDX6ko9gE8vXqZIGj82mwD
# TAY9ofRH0RSMLJqpgLrBPCKNMIIHejCCBWKgAwIBAgIKYQ6Q0gAAAAAAAzANBgkq
# hkiG9w0BAQsFADCBiDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24x
# EDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlv
# bjEyMDAGA1UEAxMpTWljcm9zb2Z0IFJvb3QgQ2VydGlmaWNhdGUgQXV0aG9yaXR5
# IDIwMTEwHhcNMTEwNzA4MjA1OTA5WhcNMjYwNzA4MjEwOTA5WjB+MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSgwJgYDVQQDEx9NaWNyb3NvZnQg
# Q29kZSBTaWduaW5nIFBDQSAyMDExMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIIC
# CgKCAgEAq/D6chAcLq3YbqqCEE00uvK2WCGfQhsqa+laUKq4BjgaBEm6f8MMHt03
# a8YS2AvwOMKZBrDIOdUBFDFC04kNeWSHfpRgJGyvnkmc6Whe0t+bU7IKLMOv2akr
# rnoJr9eWWcpgGgXpZnboMlImEi/nqwhQz7NEt13YxC4Ddato88tt8zpcoRb0Rrrg
# OGSsbmQ1eKagYw8t00CT+OPeBw3VXHmlSSnnDb6gE3e+lD3v++MrWhAfTVYoonpy
# 4BI6t0le2O3tQ5GD2Xuye4Yb2T6xjF3oiU+EGvKhL1nkkDstrjNYxbc+/jLTswM9
# sbKvkjh+0p2ALPVOVpEhNSXDOW5kf1O6nA+tGSOEy/S6A4aN91/w0FK/jJSHvMAh
# dCVfGCi2zCcoOCWYOUo2z3yxkq4cI6epZuxhH2rhKEmdX4jiJV3TIUs+UsS1Vz8k
# A/DRelsv1SPjcF0PUUZ3s/gA4bysAoJf28AVs70b1FVL5zmhD+kjSbwYuER8ReTB
# w3J64HLnJN+/RpnF78IcV9uDjexNSTCnq47f7Fufr/zdsGbiwZeBe+3W7UvnSSmn
# Eyimp31ngOaKYnhfsi+E11ecXL93KCjx7W3DKI8sj0A3T8HhhUSJxAlMxdSlQy90
# lfdu+HggWCwTXWCVmj5PM4TasIgX3p5O9JawvEagbJjS4NaIjAsCAwEAAaOCAe0w
# ggHpMBAGCSsGAQQBgjcVAQQDAgEAMB0GA1UdDgQWBBRIbmTlUAXTgqoXNzcitW2o
# ynUClTAZBgkrBgEEAYI3FAIEDB4KAFMAdQBiAEMAQTALBgNVHQ8EBAMCAYYwDwYD
# VR0TAQH/BAUwAwEB/zAfBgNVHSMEGDAWgBRyLToCMZBDuRQFTuHqp8cx0SOJNDBa
# BgNVHR8EUzBRME+gTaBLhklodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2Ny
# bC9wcm9kdWN0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3JsMF4GCCsG
# AQUFBwEBBFIwUDBOBggrBgEFBQcwAoZCaHR0cDovL3d3dy5taWNyb3NvZnQuY29t
# L3BraS9jZXJ0cy9NaWNSb29DZXJBdXQyMDExXzIwMTFfMDNfMjIuY3J0MIGfBgNV
# HSAEgZcwgZQwgZEGCSsGAQQBgjcuAzCBgzA/BggrBgEFBQcCARYzaHR0cDovL3d3
# dy5taWNyb3NvZnQuY29tL3BraW9wcy9kb2NzL3ByaW1hcnljcHMuaHRtMEAGCCsG
# AQUFBwICMDQeMiAdAEwAZQBnAGEAbABfAHAAbwBsAGkAYwB5AF8AcwB0AGEAdABl
# AG0AZQBuAHQALiAdMA0GCSqGSIb3DQEBCwUAA4ICAQBn8oalmOBUeRou09h0ZyKb
# C5YR4WOSmUKWfdJ5DJDBZV8uLD74w3LRbYP+vj/oCso7v0epo/Np22O/IjWll11l
# hJB9i0ZQVdgMknzSGksc8zxCi1LQsP1r4z4HLimb5j0bpdS1HXeUOeLpZMlEPXh6
# I/MTfaaQdION9MsmAkYqwooQu6SpBQyb7Wj6aC6VoCo/KmtYSWMfCWluWpiW5IP0
# wI/zRive/DvQvTXvbiWu5a8n7dDd8w6vmSiXmE0OPQvyCInWH8MyGOLwxS3OW560
# STkKxgrCxq2u5bLZ2xWIUUVYODJxJxp/sfQn+N4sOiBpmLJZiWhub6e3dMNABQam
# ASooPoI/E01mC8CzTfXhj38cbxV9Rad25UAqZaPDXVJihsMdYzaXht/a8/jyFqGa
# J+HNpZfQ7l1jQeNbB5yHPgZ3BtEGsXUfFL5hYbXw3MYbBL7fQccOKO7eZS/sl/ah
# XJbYANahRr1Z85elCUtIEJmAH9AAKcWxm6U/RXceNcbSoqKfenoi+kiVH6v7RyOA
# 9Z74v2u3S5fi63V4GuzqN5l5GEv/1rMjaHXmr/r8i+sLgOppO6/8MO0ETI7f33Vt
# Y5E90Z1WTk+/gFcioXgRMiF670EKsT/7qMykXcGhiJtXcVZOSEXAQsmbdlsKgEhr
# /Xmfwb1tbWrJUnMTDXpQzTGCFVswghVXAgEBMIGVMH4xCzAJBgNVBAYTAlVTMRMw
# EQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVN
# aWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNp
# Z25pbmcgUENBIDIwMTECEzMAAAGGTSF1oNkHviwAAAAAAYYwDQYJYIZIAWUDBAIB
# BQCgga4wGQYJKoZIhvcNAQkDMQwGCisGAQQBgjcCAQQwHAYKKwYBBAGCNwIBCzEO
# MAwGCisGAQQBgjcCARUwLwYJKoZIhvcNAQkEMSIEIAkTR7Sm+nCSyyhHdmZhkejJ
# yiZ6VPzHbr5Kvvta+91/MEIGCisGAQQBgjcCAQwxNDAyoBSAEgBNAGkAYwByAG8A
# cwBvAGYAdKEagBhodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20wDQYJKoZIhvcNAQEB
# BQAEggEAhSgqpxXPQAwV4vboYw05Hs6uPJ73h3lYlPKZpl69LYxmF/u76cwdxsEp
# LTC77TDrtSkWuTmt5DPJET8Wg1tDzjfKCZrmabU+FFX7Tf5m7ClnfdmzMMThz/tN
# +lJ2WP88J2Lfe0FKnipX79dpAYHIisKrsmPWVYpVTBNFwNANxvLKS+LDSCe1+35r
# IeJ8sx/nl1RtDhvDA5elNS9Gyv+bKAYRvJwpXgOVxD2j6n1/swD9ruvcab56xlDz
# JuOWBXeTuc97U/uBYK7I//Q497vYh2KYSwGs0bKaG/JephSiIHHJOhDpyY+bbzg/
# 9hVWjJ0cUqgTejx97qDo9qw2ti+ZY6GCEuUwghLhBgorBgEEAYI3AwMBMYIS0TCC
# Es0GCSqGSIb3DQEHAqCCEr4wghK6AgEDMQ8wDQYJYIZIAWUDBAIBBQAwggFRBgsq
# hkiG9w0BCRABBKCCAUAEggE8MIIBOAIBAQYKKwYBBAGEWQoDATAxMA0GCWCGSAFl
# AwQCAQUABCDqiG8l+HXcrlcBnDVASds/2dFFUwBceBCYNEBFPv9kMQIGX1/kh8Zm
# GBMyMDIwMDkyNDAzMzgzOC44OTNaMASAAgH0oIHQpIHNMIHKMQswCQYDVQQGEwJV
# UzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UE
# ChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSUwIwYDVQQLExxNaWNyb3NvZnQgQW1l
# cmljYSBPcGVyYXRpb25zMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjoxMkJDLUUz
# QUUtNzRFQjElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZaCC
# DjwwggTxMIID2aADAgECAhMzAAABIfexgZsjRNcMAAAAAAEhMA0GCSqGSIb3DQEB
# CwUAMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQH
# EwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNV
# BAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMB4XDTE5MTExMzIxNDA0
# MloXDTIxMDIxMTIxNDA0MlowgcoxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNo
# aW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24xJTAjBgNVBAsTHE1pY3Jvc29mdCBBbWVyaWNhIE9wZXJhdGlvbnMx
# JjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNOOjEyQkMtRTNBRS03NEVCMSUwIwYDVQQD
# ExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNlMIIBIjANBgkqhkiG9w0BAQEF
# AAOCAQ8AMIIBCgKCAQEA0Ah4aX8dhsoAEppJN+m/5YWAOKUf1Mtd0UbqWOnc1KVX
# hyfUTkQ7S4RvgJxHlqyXF49itbTdGH3SSKiCBcEcLj5amESLzBlkqGvIPTnzDKGT
# 2YZtOpYpi7ifl+EJVR4MidCodphAfk6eeY03+Th4VmHRhJ8glxjCXI+TJAvVAdLr
# rnR8vobuR8L1taFxnAXEBGs6y7cxtreJZuo8KMWI2gxff9FuAh6mqcQ//KDxHBgo
# 56zZnDHNF7fWh3Z4EyiFf0y/FDrOaWEy/l/TWmAzuhRYAEr31r5Kz+Ns6MRN+qQY
# QFmsfFIU+uypuPtl68/hTfUTLpADrfi4NZSrBqNnxwIDAQABo4IBGzCCARcwHQYD
# VR0OBBYEFIJtosXMbdxf57gMHROp3XbSXvQcMB8GA1UdIwQYMBaAFNVjOlyKMZDz
# Q3t8RhvFM2hahW1VMFYGA1UdHwRPME0wS6BJoEeGRWh0dHA6Ly9jcmwubWljcm9z
# b2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1RpbVN0YVBDQV8yMDEwLTA3LTAx
# LmNybDBaBggrBgEFBQcBAQROMEwwSgYIKwYBBQUHMAKGPmh0dHA6Ly93d3cubWlj
# cm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljVGltU3RhUENBXzIwMTAtMDctMDEuY3J0
# MAwGA1UdEwEB/wQCMAAwEwYDVR0lBAwwCgYIKwYBBQUHAwgwDQYJKoZIhvcNAQEL
# BQADggEBAJnIZ6CfVZSKNf8NDkE6qtxiW+AbTrTW93yFtSZVOcfNPkDj9iQs0fsB
# 8vnoWuDn3km6wWiOPlylb+2Vm1/QwWF9jSY6UzIBMNR/NBYcqoISTzF6f5kG4ZG+
# lKQkOCpXVZbLwNXtfD+8X7emO5ojPmWJvhiGc7TJEIs59IQBlTJ2eoVgyAPBWv9W
# cMRzIh5cwGDwOyWKb1Z36Z2CSH8dJrnvQSONErEFjYk60O7UyKnfTOSJT2fxsuwK
# Vw0yVq8PqmA4y+cpTfr+rhrAvhVznwM0uAdcY9yg6c/w0WqPNluBm+SCwCgVrL24
# vjO5fk4z4LhPvrRFwHBCWMA7FvmwYk8wggZxMIIEWaADAgECAgphCYEqAAAAAAAC
# MA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGlu
# Z3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBv
# cmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRo
# b3JpdHkgMjAxMDAeFw0xMDA3MDEyMTM2NTVaFw0yNTA3MDEyMTQ2NTVaMHwxCzAJ
# BgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25k
# MR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jv
# c29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwMIIBIjANBgkqhkiG9w0BAQEFAAOCAQ8A
# MIIBCgKCAQEAqR0NvHcRijog7PwTl/X6f2mUa3RUENWlCgCChfvtfGhLLF/Fw+Vh
# wna3PmYrW/AVUycEMR9BGxqVHc4JE458YTBZsTBED/FgiIRUQwzXTbg4CLNC3ZOs
# 1nMwVyaCo0UN0Or1R4HNvyRgMlhgRvJYR4YyhB50YWeRX4FUsc+TTJLBxKZd0WET
# bijGGvmGgLvfYfxGwScdJGcSchohiq9LZIlQYrFd/XcfPfBXday9ikJNQFHRD5wG
# Pmd/9WbAA5ZEfu/QS/1u5ZrKsajyeioKMfDaTgaRtogINeh4HLDpmc085y9Euqf0
# 3GS9pAHBIAmTeM38vMDJRF1eFpwBBU8iTQIDAQABo4IB5jCCAeIwEAYJKwYBBAGC
# NxUBBAMCAQAwHQYDVR0OBBYEFNVjOlyKMZDzQ3t8RhvFM2hahW1VMBkGCSsGAQQB
# gjcUAgQMHgoAUwB1AGIAQwBBMAsGA1UdDwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/
# MB8GA1UdIwQYMBaAFNX2VsuP6KJcYmjRPZSQW9fOmhjEMFYGA1UdHwRPME0wS6BJ
# oEeGRWh0dHA6Ly9jcmwubWljcm9zb2Z0LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01p
# Y1Jvb0NlckF1dF8yMDEwLTA2LTIzLmNybDBaBggrBgEFBQcBAQROMEwwSgYIKwYB
# BQUHMAKGPmh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbS9wa2kvY2VydHMvTWljUm9v
# Q2VyQXV0XzIwMTAtMDYtMjMuY3J0MIGgBgNVHSABAf8EgZUwgZIwgY8GCSsGAQQB
# gjcuAzCBgTA9BggrBgEFBQcCARYxaHR0cDovL3d3dy5taWNyb3NvZnQuY29tL1BL
# SS9kb2NzL0NQUy9kZWZhdWx0Lmh0bTBABggrBgEFBQcCAjA0HjIgHQBMAGUAZwBh
# AGwAXwBQAG8AbABpAGMAeQBfAFMAdABhAHQAZQBtAGUAbgB0AC4gHTANBgkqhkiG
# 9w0BAQsFAAOCAgEAB+aIUQ3ixuCYP4FxAz2do6Ehb7Prpsz1Mb7PBeKp/vpXbRkw
# s8LFZslq3/Xn8Hi9x6ieJeP5vO1rVFcIK1GCRBL7uVOMzPRgEop2zEBAQZvcXBf/
# XPleFzWYJFZLdO9CEMivv3/Gf/I3fVo/HPKZeUqRUgCvOA8X9S95gWXZqbVr5MfO
# 9sp6AG9LMEQkIjzP7QOllo9ZKby2/QThcJ8ySif9Va8v/rbljjO7Yl+a21dA6fHO
# mWaQjP9qYn/dxUoLkSbiOewZSnFjnXshbcOco6I8+n99lmqQeKZt0uGc+R38ONiU
# 9MalCpaGpL2eGq4EQoO4tYCbIjggtSXlZOz39L9+Y1klD3ouOVd2onGqBooPiRa6
# YacRy5rYDkeagMXQzafQ732D8OE7cQnfXXSYIghh2rBQHm+98eEA3+cxB6STOvdl
# R3jo+KhIq/fecn5ha293qYHLpwmsObvsxsvYgrRyzR30uIUBHoD7G4kqVDmyW9rI
# DVWZeodzOwjmmC3qjeAzLhIp9cAvVCch98isTtoouLGp25ayp0Kiyc8ZQU3ghvkq
# mqMRZjDTu3QyS99je/WZii8bxyGvWbWu3EQ8l1Bx16HSxVXjad5XwdHeMMD9zOZN
# +w2/XU/pnR4ZOC+8z1gFLu8NoFA12u8JJxzVs341Hgi62jbb01+P3nSISRKhggLO
# MIICNwIBATCB+KGB0KSBzTCByjELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hp
# bmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jw
# b3JhdGlvbjElMCMGA1UECxMcTWljcm9zb2Z0IEFtZXJpY2EgT3BlcmF0aW9uczEm
# MCQGA1UECxMdVGhhbGVzIFRTUyBFU046MTJCQy1FM0FFLTc0RUIxJTAjBgNVBAMT
# HE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WiIwoBATAHBgUrDgMCGgMVAK/G
# ozto6gPrwhnSvhHQV1CqY7tpoIGDMIGApH4wfDELMAkGA1UEBhMCVVMxEzARBgNV
# BAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jv
# c29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAg
# UENBIDIwMTAwDQYJKoZIhvcNAQEFBQACBQDjFkBJMCIYDzIwMjAwOTI0MDU0NDQx
# WhgPMjAyMDA5MjUwNTQ0NDFaMHcwPQYKKwYBBAGEWQoEATEvMC0wCgIFAOMWQEkC
# AQAwCgIBAAICHVMCAf8wBwIBAAICEa4wCgIFAOMXkckCAQAwNgYKKwYBBAGEWQoE
# AjEoMCYwDAYKKwYBBAGEWQoDAqAKMAgCAQACAwehIKEKMAgCAQACAwGGoDANBgkq
# hkiG9w0BAQUFAAOBgQABSKa+8QNC8qmGl4rq9AmDuEZuBy+A27rqb1ZraF0RVVVV
# J58mxghgo5eMLh/wPRlTW6Z4XJyUjyXr35JEJehDwSKmH5I5rTQflUBD65ik5wxh
# DzgQWA5WHLe2Ms3WKXr1o/vuUF6fGJwXQ7fohX0Ili9eBZ/YsxZnJaH5I6EVRzGC
# Aw0wggMJAgEBMIGTMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
# MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
# b24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAAB
# IfexgZsjRNcMAAAAAAEhMA0GCWCGSAFlAwQCAQUAoIIBSjAaBgkqhkiG9w0BCQMx
# DQYLKoZIhvcNAQkQAQQwLwYJKoZIhvcNAQkEMSIEIEZCF7klybAptNW6tugZwa+I
# sdKke14ooeIof6WE2T/3MIH6BgsqhkiG9w0BCRACLzGB6jCB5zCB5DCBvQQg/hFs
# k1fpuwAwAZXwRTU7vmo0LfHeZKeO7Wl4aPq052AwgZgwgYCkfjB8MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQg
# VGltZS1TdGFtcCBQQ0EgMjAxMAITMwAAASH3sYGbI0TXDAAAAAABITAiBCD5HohT
# 6MTDdgGMWhRaS4e7IgU/6+sf3BeAskHzjITeuDANBgkqhkiG9w0BAQsFAASCAQBL
# AwqY+gD9Fhdd9/czocs/6LWM8QRBdM1cJUdj69Ooijjrvdo8S4ldx61A+BcEsKF5
# L0oja/p1B3fk1HsRevRXqr1a/1zz7RVT/lomu2/xSm4VDKGJ8CFoPC3gEKZwxnLq
# IPLfjkBAC2g98HxL6CLlqzAono8+nWAS4K2EfJrxDhckScETap8oQRaTA4bkhJpF
# /DJPgH320DuLLzcQq035GALBNRIpZ5Sn0wio0NzssNCZGB///C8HSaQOZyoR4AmA
# 2DLDyyndX0Z2PHehRfd8/LqFV/5S3AZKPQ9DdtQCJGigqpcOgs0UkD12dhSGH4HV
# E2UzhfkCCj93dFF07vTt
# SIG # End signature block
