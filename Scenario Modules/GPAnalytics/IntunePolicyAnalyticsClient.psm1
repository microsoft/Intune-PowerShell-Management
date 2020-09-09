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

                $odataType = "#microsoft.graph.groupPolicyPresentationValueDecimal";

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
        # Initialize Module
        $IPAClientConfiguration = Initialize-IPAClientConfiguration -TenantAdminUPN $TenantAdminUPN

        # Fetch the migration report
        $sw = [Diagnostics.Stopwatch]::StartNew()

        # Initialize Module
        $IPAClientConfiguration = Initialize-IPAClientConfiguration -TenantAdminUPN  $TenantAdminUPN

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
# MIIjhgYJKoZIhvcNAQcCoIIjdzCCI3MCAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCAeNcwUUUIVCsdV
# YKMJkSlg+epsdPuyK6s5rEBL7DHXuqCCDYEwggX/MIID56ADAgECAhMzAAABh3IX
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
# RcBCyZt2WwqASGv9eZ/BvW1taslScxMNelDNMYIVWzCCFVcCAQEwgZUwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMQITMwAAAYdyF3IVWUDHCQAAAAABhzAN
# BglghkgBZQMEAgEFAKCBrjAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgXMRcTMyx
# 8Wj8TTguF+cu67+18eP+AOnbZwGYI7+kh+8wQgYKKwYBBAGCNwIBDDE0MDKgFIAS
# AE0AaQBjAHIAbwBzAG8AZgB0oRqAGGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbTAN
# BgkqhkiG9w0BAQEFAASCAQAth8TFWwRJL4nFcxX6UAXFx8X8QpnEA4oPiBh6OCWg
# 59o/XuGC15/DP7QOmgNHljctZTB+wfhdTVd9nxgwRaGlmAK8CYI73GTo3/wmTmgn
# L7Yd511twD7HS7r4P8XW5uZPM9Ckgq2fIacSeN81lbdsot2DceUl90YGHV/OUzVm
# gYqMZs8X0HsJiLNj8Sk3ba4iuqxAjlxbha9gwbm9i7roKcCgo2cC5rpki0HY4fLR
# qals3GIcrmOPJJFMTzH40vvYKNqSyYWLZCMkniqb8glmEZbAu5SwgRokDxebUe2C
# p2+/scIv1H7JwO4cQh0eWUGHIvmtzflKUWiruqnpHFLXoYIS5TCCEuEGCisGAQQB
# gjcDAwExghLRMIISzQYJKoZIhvcNAQcCoIISvjCCEroCAQMxDzANBglghkgBZQME
# AgEFADCCAVEGCyqGSIb3DQEJEAEEoIIBQASCATwwggE4AgEBBgorBgEEAYRZCgMB
# MDEwDQYJYIZIAWUDBAIBBQAEIJZ/WToOH59mfj17AxJzZPLKsXuFd5fip7DKFuTj
# ugxdAgZfPS4DobMYEzIwMjAwOTA5MDYyNDUzLjI5MlowBIACAfSggdCkgc0wgcox
# CzAJBgNVBAYTAlVTMQswCQYDVQQIEwJXQTEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQg
# SXJlbGFuZCBPcGVyYXRpb25zIExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1Mg
# RVNOOjNCRDQtNEI4MC02OUMzMSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFt
# cCBTZXJ2aWNloIIOPDCCBPEwggPZoAMCAQICEzMAAAEL5Pm+j29MHdAAAAAAAQsw
# DQYJKoZIhvcNAQELBQAwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0
# b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3Jh
# dGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwHhcN
# MTkxMDIzMjMxOTE1WhcNMjEwMTIxMjMxOTE1WjCByjELMAkGA1UEBhMCVVMxCzAJ
# BgNVBAgTAldBMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlv
# bnMgTGltaXRlZDEmMCQGA1UECxMdVGhhbGVzIFRTUyBFU046M0JENC00QjgwLTY5
# QzMxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2UwggEiMA0G
# CSqGSIb3DQEBAQUAA4IBDwAwggEKAoIBAQCXAtWdRjFBuM+D2nhUKLVuWv9cZcq1
# /8emykQBplDii8DqwwCNnD0zJhz7n94WtWjFsc5KL/dF8gKWTMRH5MVTa5dxCJu6
# VtZobc+sztM+0JPM5Vmcb/7D+AlFERGAkQGGxO/Z4fxHH1/EcZ/iwUimzafXjBOl
# IQ3RSxUAj980liuAyNCrj8JdunGR3nVSRvxJtWpUZvlIUrYY4LDmJJsFsI8gsch3
# LrchmPeBkoxsvy7RpKhcOQtTYacD48vz7fzT2ciciJqAXxZt7fth8sgqKiUURCVu
# SlcUKXBXm/1dcYCKqOoUz2YGu2i0t4K/X17JWZ5jdN1vxqzSQa9P4PHxAgMBAAGj
# ggEbMIIBFzAdBgNVHQ4EFgQUrR/Z6h2KHpzgmA1QRGX/921e3u8wHwYDVR0jBBgw
# FoAU1WM6XIoxkPNDe3xGG8UzaFqFbVUwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDov
# L2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwvcHJvZHVjdHMvTWljVGltU3RhUENB
# XzIwMTAtMDctMDEuY3JsMFoGCCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0
# cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9jZXJ0cy9NaWNUaW1TdGFQQ0FfMjAx
# MC0wNy0wMS5jcnQwDAYDVR0TAQH/BAIwADATBgNVHSUEDDAKBggrBgEFBQcDCDAN
# BgkqhkiG9w0BAQsFAAOCAQEAJuijnanvNrS63e87CK0gwImI8C4JdhxLLPnA6m/p
# USXWel9KCa3t95NRNO36NgemDxhskz7rVHiUigb1pJdm+TB5Shg2DlPi1UhdCTaN
# 5lTWZ+rHAFfDI4i2gdKOwdyug73m5ja2dqfDTl2Di5axwcBgDvGsZLfBm+aGut2v
# UGBBg1QjMKfqQGqMJCYwXPGdHmwRN1UN5MpORBkTmk2DEWWjRm0LKQ1/eV4KYiU5
# cV4GC0/8/q/X71wbrwdyH2Zyvh2mIOE+4T9mZc7H0CzZ8QdqTHd2xbTT1GSNReeY
# YlnTkWlCiELjYkInHUfwumC1pCuZMf4ITNw7KjeOGPyKDTCCBnEwggRZoAMCAQIC
# CmEJgSoAAAAAAAIwDQYJKoZIhvcNAQELBQAwgYgxCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xMjAwBgNVBAMTKU1pY3Jvc29mdCBSb290IENlcnRp
# ZmljYXRlIEF1dGhvcml0eSAyMDEwMB4XDTEwMDcwMTIxMzY1NVoXDTI1MDcwMTIx
# NDY1NVowfDELMAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNV
# BAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQG
# A1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTAwggEiMA0GCSqGSIb3
# DQEBAQUAA4IBDwAwggEKAoIBAQCpHQ28dxGKOiDs/BOX9fp/aZRrdFQQ1aUKAIKF
# ++18aEssX8XD5WHCdrc+Zitb8BVTJwQxH0EbGpUdzgkTjnxhMFmxMEQP8WCIhFRD
# DNdNuDgIs0Ldk6zWczBXJoKjRQ3Q6vVHgc2/JGAyWGBG8lhHhjKEHnRhZ5FfgVSx
# z5NMksHEpl3RYRNuKMYa+YaAu99h/EbBJx0kZxJyGiGKr0tkiVBisV39dx898Fd1
# rL2KQk1AUdEPnAY+Z3/1ZsADlkR+79BL/W7lmsqxqPJ6Kgox8NpOBpG2iAg16Hgc
# sOmZzTznL0S6p/TcZL2kAcEgCZN4zfy8wMlEXV4WnAEFTyJNAgMBAAGjggHmMIIB
# 4jAQBgkrBgEEAYI3FQEEAwIBADAdBgNVHQ4EFgQU1WM6XIoxkPNDe3xGG8UzaFqF
# bVUwGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQDAgGGMA8GA1Ud
# EwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU1fZWy4/oolxiaNE9lJBb186aGMQwVgYD
# VR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3NvZnQuY29tL3BraS9jcmwv
# cHJvZHVjdHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3JsMFoGCCsGAQUFBwEB
# BE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNyb3NvZnQuY29tL3BraS9j
# ZXJ0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcnQwgaAGA1UdIAEB/wSBlTCB
# kjCBjwYJKwYBBAGCNy4DMIGBMD0GCCsGAQUFBwIBFjFodHRwOi8vd3d3Lm1pY3Jv
# c29mdC5jb20vUEtJL2RvY3MvQ1BTL2RlZmF1bHQuaHRtMEAGCCsGAQUFBwICMDQe
# MiAdAEwAZQBnAGEAbABfAFAAbwBsAGkAYwB5AF8AUwB0AGEAdABlAG0AZQBuAHQA
# LiAdMA0GCSqGSIb3DQEBCwUAA4ICAQAH5ohRDeLG4Jg/gXEDPZ2joSFvs+umzPUx
# vs8F4qn++ldtGTCzwsVmyWrf9efweL3HqJ4l4/m87WtUVwgrUYJEEvu5U4zM9GAS
# inbMQEBBm9xcF/9c+V4XNZgkVkt070IQyK+/f8Z/8jd9Wj8c8pl5SpFSAK84Dxf1
# L3mBZdmptWvkx872ynoAb0swRCQiPM/tA6WWj1kpvLb9BOFwnzJKJ/1Vry/+tuWO
# M7tiX5rbV0Dp8c6ZZpCM/2pif93FSguRJuI57BlKcWOdeyFtw5yjojz6f32WapB4
# pm3S4Zz5Hfw42JT0xqUKloakvZ4argRCg7i1gJsiOCC1JeVk7Pf0v35jWSUPei45
# V3aicaoGig+JFrphpxHLmtgOR5qAxdDNp9DvfYPw4TtxCd9ddJgiCGHasFAeb73x
# 4QDf5zEHpJM692VHeOj4qEir995yfmFrb3epgcunCaw5u+zGy9iCtHLNHfS4hQEe
# gPsbiSpUObJb2sgNVZl6h3M7COaYLeqN4DMuEin1wC9UJyH3yKxO2ii4sanblrKn
# QqLJzxlBTeCG+SqaoxFmMNO7dDJL32N79ZmKLxvHIa9Zta7cRDyXUHHXodLFVeNp
# 3lfB0d4wwP3M5k37Db9dT+mdHhk4L7zPWAUu7w2gUDXa7wknHNWzfjUeCLraNtvT
# X4/edIhJEqGCAs4wggI3AgEBMIH4oYHQpIHNMIHKMQswCQYDVQQGEwJVUzELMAkG
# A1UECBMCV0ExEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBD
# b3Jwb3JhdGlvbjEtMCsGA1UECxMkTWljcm9zb2Z0IElyZWxhbmQgT3BlcmF0aW9u
# cyBMaW1pdGVkMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjozQkQ0LTRCODAtNjlD
# MzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2VydmljZaIjCgEBMAcG
# BSsOAwIaAxUA8f35HTFqU9zwihI9ktmsPgpwMFKggYMwgYCkfjB8MQswCQYDVQQG
# EwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwG
# A1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQg
# VGltZS1TdGFtcCBQQ0EgMjAxMDANBgkqhkiG9w0BAQUFAAIFAOMCsnMwIhgPMjAy
# MDA5MDkwOTQ2MjdaGA8yMDIwMDkxMDA5NDYyN1owdzA9BgorBgEEAYRZCgQBMS8w
# LTAKAgUA4wKycwIBADAKAgEAAgIa+wIB/zAHAgEAAgIRrjAKAgUA4wQD8wIBADA2
# BgorBgEEAYRZCgQCMSgwJjAMBgorBgEEAYRZCgMCoAowCAIBAAIDB6EgoQowCAIB
# AAIDAYagMA0GCSqGSIb3DQEBBQUAA4GBABiZHZlS9Q7FBMhU0JxRHC5lpasgTKTh
# LiTL9vgyWR2bcnq1KBDCoFlGNPZKA9eCKFvplu0j4954RMcIwwCxA7dgoUc9aBvE
# VSbkxY+uU0am3WAb9id7+YFBMRKBoBbHApREvp9lcunOyVkoXbGGy6K/8UoQ+PKe
# GjVicgtpgD7EMYIDDTCCAwkCAQEwgZMwfDELMAkGA1UEBhMCVVMxEzARBgNVBAgT
# Cldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNVBAoTFU1pY3Jvc29m
# dCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRpbWUtU3RhbXAgUENB
# IDIwMTACEzMAAAEL5Pm+j29MHdAAAAAAAQswDQYJYIZIAWUDBAIBBQCgggFKMBoG
# CSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG9w0BCQQxIgQgIDQ0ycCf
# VXGdnMpsGdE+rOcsoEn62Uvx8HgVER1RS04wgfoGCyqGSIb3DQEJEAIvMYHqMIHn
# MIHkMIG9BCA0j9DOIFM+OiSX8XAkXAXivRR0LPHA6cVU/ATAE1xziDCBmDCBgKR+
# MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMT
# HU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAABC+T5vo9vTB3QAAAA
# AAELMCIEILDtPGz43RPfx4EvYuObHzbcDJdZJSjHj2n/wHLFJQMYMA0GCSqGSIb3
# DQEBCwUABIIBAEMT9yz7iW+/VKrZwbrtsjEQ2254iz+x/SQLOP+R6ybgf51y13NN
# E/YbvEE+0LlLHHN3dQ56aPCLd1ew5VcRNwMCqQh/5Hsqlz+vpqWtBOeVEoGNuFZG
# vYfnQJwm9DETX5Vu960NLNQBQe8SDOJXAusWEOSGsHYcr0XhNE9gpZY7UbTQ8yqB
# q7vhRNzOocZW0TXGPE5Qzg74V6EPyEwWuncVC36cinhb23PWUSbZru4k2nnrH/+S
# CrryTTw3WDLTUwhKHdB1eV4wgjG4foklre+U4bCQUvRVWmfJW54F9CWMrgKYQELR
# HCd6/tkxGKAPwR/BcKZ4PBasdeewi5Pgsgw=
# SIG # End signature block
