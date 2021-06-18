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
        [Switch]$ExpandSettings = $false,
        [String]$Environment = "pe"
    )

    Try
    {
        # Initialize Module
        $IPAClientConfiguration = Initialize-IPAClientConfiguration -TenantAdminUPN $TenantAdminUPN -Environment $Environment

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
                Write-Log -Message "Get-GPOMigrationReportCollection: $($groupPolicyObjectId)_$($ou)"

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
        [string]$TenantAdminUPN,
        [String]$Environment = "pe"
    )

    Try
    {
        # Initialize Module
        $IPAClientConfiguration = Initialize-IPAClientConfiguration -TenantAdminUPN $TenantAdminUPN -Environment $Environment

        # Fetch the migration report
        $sw = [Diagnostics.Stopwatch]::StartNew()

        # Headers for the Graph call
        $clonedHeaders = CloneObject (Get-AuthHeader $graphConfiguration);
        $clonedHeaders["api-version"] = "$($script:ApiVersion)";        

        <#
            1. Ask to create the report
            Post https://graph.microsoft-ppe.com/testppebeta_intune_onedf/deviceManagement/reports/cachedReportConfigurations
            Payload: {"id":"GPAnalyticsSettingMigrationReadiness_00000000-0000-0000-0000-000000000001","filter":"","orderBy":[],"select":["SettingName","SettingCategory","MigrationReadiness","OSVersion","Scope","ProfileType"],"metadata":""}
        #>
        $uri = "$($script:GraphConfiguration.GraphBaseAddress)/$($script:GraphConfiguration.SchemaVersion)/deviceManagement/reports/cachedReportConfigurations";        
        $Body = "{            
            `"id`":`"GPAnalyticsSettingMigrationReadiness_00000000-0000-0000-0000-000000000001`",
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
            Payload: {"Id":"GPAnalyticsSettingMigrationReadiness_00000000-0000-0000-0000-000000000001","GroupBy":["ProfileType"],"Select":["ProfileType"]}
        #>
        $uri = "$($script:GraphConfiguration.GraphBaseAddress)/$($script:GraphConfiguration.SchemaVersion)/deviceManagement/reports/getCachedReport";        
        $Body = "{            
            `"id`":`"GPAnalyticsSettingMigrationReadiness_00000000-0000-0000-0000-000000000001`",
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
        [string]$TenantAdminUPN,
        [String]$Environment = "pe"
    )

    Try
    {
        # Initialize Module
        $IPAClientConfiguration = Initialize-IPAClientConfiguration -TenantAdminUPN $TenantAdminUPN -Environment $Environment

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
    Additionally - this cmdlet can also be used to upload GPOs from a folder-path where GPOReport.xml have been previously exported to.
.DESCRIPTION
    IntunePolicyAnalyticsClient uses the Group Policy cmdlets to get all the Group Policy Objects
    for a given domain, uploads it to Intune and determines what settings are supported.
.PARAMETER TenantAdminUPN
    AAD User Principal Name of the Intune Tenant Administrator which is required to upload the GPOs.
.PARAMETER Domain
    The local AD Domain for which the GPO collection is fetched.
    Defaults to the local AD Domain for the client on which this script is run on.
    This is an optional parameter. 
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
    This is an optional parameter.
.PARAMETER GpoBackupFolderPath
    The folder on local disk where GPOReport.xml files have been previously exported to from a domain
.OUTPUT
    GPO Collection collected from the local AD domain and sent to Intune
.EXAMPLE
    Import-GPOCollection -Domain "redmond.corp.microsoft.com" -TenantAdminUPN "admin@IPASHAMSUA01MSIT.onmicrosoft.com"
        Gets all the GPOs in the domain "redmond.corp.microsoft.com" and back them up on disk.
.EXAMPLE
    Import-GPOCollection -Domain "redmond.corp.microsoft.com" -OUFilter 'DistinguishedName -like "OU=CoreIdentity,OU=ITServices,DC=redmond,DC=corp,DC=microsoft,DC=com"' -TenantAdminUPN "admin@IPASHAMSUA01MSIT.onmicrosoft.com"
        Gets all the GPOs in a specific OU for the given domain, back them up on disk and upload to Intune.
.EXAMPLE 
    Import-GPOCollection -TenantAdminUPN "admin@IPASHAMSUA01MSIT.onmicrosoft.com" -GPOBackupFolderPath "C:\GPOBackup"
#>
    [cmdletbinding()]
    param(        
        [Parameter(Mandatory=$true)]
        [string]$TenantAdminUPN,
        [Alias("Domain")]
        [String]$ADDomain = $null,
        [String]$OUFilter = $null,
        [String]$GpoBackupFolderPath = $null,
        [String]$Environment = "pe"
    )

    Try
    {
        # Start timer
        $sw = [Diagnostics.Stopwatch]::StartNew()

        # Initialize Module
        $IPAClientConfiguration = Initialize-IPAClientConfiguration -TenantAdminUPN $TenantAdminUPN -Environment $Environment

        # Fetch GPO Xml Reports from local AD Domain
        $gpoReportXmlCollection = @{}
        Try
        {
            If ($null -ne $GpoBackupFolderPath)
            {
                Write-Log -Message "Import-GPOCollection: Read GPO backups from GpoBackupFolderPath $($GpoBackupFolderPath)..."
                $gpoReportXmlCollection = GPOReportXmlCollectionFromDisk -GpoBackupFolderPath $GpoBackupFolderPath
            }
            ElseIf ($null -ne $OUFilter)
            {                        
                Write-Log -Message "Import-GPOCollection: Get GPO backups from ADDomain=$($ADDomain) with OUFilter=$($OUFilter)..."
                $gpoReportXmlCollection = Get-GPOReportXmlCollectionFromAD -ADDomain $($ADDomain) -OUFilter $($OUFilter)
            }
            Else
            {
                Write-Log -Message "Import-GPOCollection:: Please specify either -ADDomain or -GpoBackupFolderPath" -Level "Error"
                throw
            }
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
        [String]$Environment = "pe",
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

        # Create GPO Backup folder path if necessary        
        If(!(Test-Path -Path "$($script:IPAClientConfiguration.GpoBackupFolderPath)" ))
        {
            (New-Item -ItemType directory -Path $script:IPAClientConfiguration.GpoBackupFolderPath) | Out-Null
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
        [string]$TenantAdminUPN,
        [String]$Environment = "pe"
    )

    Try
    {
        # Initialize Module
        $IPAClientConfiguration = Initialize-IPAClientConfiguration -TenantAdminUPN $TenantAdminUPN -Environment $Environment

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
        [string]$TenantAdminUPN,
        [String]$Environment = "pe"
    )

    Try
    {
        # Initialize Module
        $IPAClientConfiguration = Initialize-IPAClientConfiguration -TenantAdminUPN $TenantAdminUPN -Environment $Environment

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

            $unescapedContainerBody = ( ConvertTo-Json $admxUpdateContainerBody -depth 10 | ForEach-Object { [regex]::Unescape($_) })
            $admxUpdateContainerUri =   "$($GraphConfiguration.GraphBaseAddress)/$($GraphConfiguration.SchemaVersion)/deviceManagement/groupPolicyConfigurations('$($admxProfile.Id)')/updateDefinitionValues"

            Write-Log -Message "Add-GPToIntuneAdmxMigratedProfile: Add-IntuneEntityCollection -Body: $($unescapedContainerBody))  -uri: $($admxUpdateContainerUri)"
            $updatedConfiguration = (Add-IntuneEntityCollection -Body ($unescapedContainerBody | ConvertTo-Json -Depth 10) -uri $admxUpdateContainerUri -GraphConfiguration $script:GraphConfiguration)

            Write-Log -Message "Add-GPToIntuneAdmxMigratedProfile: Profile updated successfully"
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
   
    return $updatedConfiguration
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

                    if($null -eq $_.name){
                        $listitem = [PSCustomObject]@{`
                                "name" = $_.Data;`
                        }

                        $value += $listitem
                    }
                    else
                    {
                        $listitem = [PSCustomObject]@{`
                                "name" = $_.name;`
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
        [string]$TenantAdminUPN,
        [String]$Environment = "pe"
    )

    Try
    {
        # Fetch the migration report
        $sw = [Diagnostics.Stopwatch]::StartNew()

        # Initialize Module 
        $IPAClientConfiguration = Initialize-IPAClientConfiguration -TenantAdminUPN  $TenantAdminUPN -Environment $Environment

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

Export-ModuleMember -Function Get-GPOMigrationReportCollection
Export-ModuleMember -Function Get-MigrationReadinessReport
Export-ModuleMember -Function Update-MigrationReadinessReport
Export-ModuleMember -Function Import-GPOCollection
Export-ModuleMember -Function Remove-GPOMigrationReportCollection
Export-ModuleMember -Function Write-Log
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
            SchemaVersion = "beta"; `
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
    Get-GPOReportXmlCollectionFromDisk: Reads previously backed up GPO Report Xml files from disk.
.DESCRIPTION
    Reads previously backed up GPO Report Xml files from disk.
.PARAMETER GpoBackupFolderPath
    Path to the Folder on local disk where the GPOReport.xml files are backed up.
#>
Function Get-GPOReportXmlCollectionFromDisk
{
    param(
        [String]$GpoBackupFolderPath        
    )

    $gpoReportXmlCollection = @{}
    $gpoMigrationReportCollection = @{}

    If ($script:IPAClientConfiguration.DeltaUpdate -eq $true)
    {
        # If set, check Intune before posting GPOReportXml
        $gpoMigrationReportCollection = Get-GPOMigrationReportCollection -TenantAdminUPN $script:IPAClientConfiguration.TenantAdminUPN
    }

    # Get list of backuped GPO Reports
    $groupPolicyObjectXmlFileList = Get-ChildItem $GpoBackupFolderPath -Filter *.xml
    ForEach ($groupPolicyObjectXmlFile in $groupPolicyObjectXmlFileList)
    {
        Try
        {
            [Xml]$gpoReportXml = Get-Content $groupPolicyObjectXmlFile.FullName
            $gpoGuid = $gpoReportXml.GPO.Identifier.Identifier.InnerText.TrimStart("{").TrimEnd("}")            
            $ou = ($gpoReportXml.GPO.Name);
            
            # Backup a GPO as Xml in memory if not previously uploaded
            $gpoReportXmlKey = "$($gpoGuid)_$($ou)"

            If (!$gpoMigrationReportCollection.Contains($gpoReportXmlKey))
            {
                $bytes = [System.Text.Encoding]::UNICODE.GetBytes($gpoReportXml.InnerXml)
                $encodedText = [Convert]::ToBase64String($bytes)
                $gpoReportXmlCollection.Add($gpoReportXmlKey, [PSCustomObject]@{ouDistinguishedName = $ou; content = $encodedText})
                Write-Log -Message "Get-GPOReportXmlCollectionFromDisk:  Backed up GPO Guid=$($gpoGuid), $($ou)"
            }
            Else
            {
                Write-Log -Message "Get-GPOReportXmlCollectionFromDisk:  GPO Guid=$($gpoGuid), $($ou.Name) previously uploaded"
            }
        }
        Catch
        {
            # Log error
            $exception  = $_
            Write-Log -Message "Get-GPOReportXmlCollectionFromAD: Failure: ($exception)" -Level "Warn"
            # We continue to next GPO
        }
    }

    return $gpoReportXmlCollection;
}

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
                            $gpoReportXml.InnerXml | Out-File -FilePath "$($IPAClientConfiguration.GpoBackupFolderPath)\$($gpoReportXmlKey).xml"
                            $bytes = [System.Text.Encoding]::UNICODE.GetBytes($gpoReportXml.InnerXml)
                            $encodedText = [Convert]::ToBase64String($bytes)
                            $gpoReportXmlCollection.Add($gpoReportXmlKey, [PSCustomObject]@{ouDistinguishedName = $ou.Name; content = $encodedText})
                            Write-Log -Message "Get-GPOReportXmlCollectionFromAD:  Backed up GPO Guid=$($gpoGuid), $($ou.Name)"
                        }
                        Else
                        {
                            Write-Log -Message "Get-GPOReportXmlCollectionFromAD:  GPO Guid=$($gpoGuid), $($ou.Name) previously uploaded"
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
# MIIjkgYJKoZIhvcNAQcCoIIjgzCCI38CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCByxJgR4afTjj3H
# RoNUltkpZgk9qgUtvcfZy/hjpDPVk6CCDYEwggX/MIID56ADAgECAhMzAAAB32vw
# LpKnSrTQAAAAAAHfMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjAxMjE1MjEzMTQ1WhcNMjExMjAyMjEzMTQ1WjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQC2uxlZEACjqfHkuFyoCwfL25ofI9DZWKt4wEj3JBQ48GPt1UsDv834CcoUUPMn
# s/6CtPoaQ4Thy/kbOOg/zJAnrJeiMQqRe2Lsdb/NSI2gXXX9lad1/yPUDOXo4GNw
# PjXq1JZi+HZV91bUr6ZjzePj1g+bepsqd/HC1XScj0fT3aAxLRykJSzExEBmU9eS
# yuOwUuq+CriudQtWGMdJU650v/KmzfM46Y6lo/MCnnpvz3zEL7PMdUdwqj/nYhGG
# 3UVILxX7tAdMbz7LN+6WOIpT1A41rwaoOVnv+8Ua94HwhjZmu1S73yeV7RZZNxoh
# EegJi9YYssXa7UZUUkCCA+KnAgMBAAGjggF+MIIBejAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUOPbML8IdkNGtCfMmVPtvI6VZ8+Mw
# UAYDVR0RBEkwR6RFMEMxKTAnBgNVBAsTIE1pY3Jvc29mdCBPcGVyYXRpb25zIFB1
# ZXJ0byBSaWNvMRYwFAYDVQQFEw0yMzAwMTIrNDYzMDA5MB8GA1UdIwQYMBaAFEhu
# ZOVQBdOCqhc3NyK1bajKdQKVMFQGA1UdHwRNMEswSaBHoEWGQ2h0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY0NvZFNpZ1BDQTIwMTFfMjAxMS0w
# Ny0wOC5jcmwwYQYIKwYBBQUHAQEEVTBTMFEGCCsGAQUFBzAChkVodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY0NvZFNpZ1BDQTIwMTFfMjAx
# MS0wNy0wOC5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAgEAnnqH
# tDyYUFaVAkvAK0eqq6nhoL95SZQu3RnpZ7tdQ89QR3++7A+4hrr7V4xxmkB5BObS
# 0YK+MALE02atjwWgPdpYQ68WdLGroJZHkbZdgERG+7tETFl3aKF4KpoSaGOskZXp
# TPnCaMo2PXoAMVMGpsQEQswimZq3IQ3nRQfBlJ0PoMMcN/+Pks8ZTL1BoPYsJpok
# t6cql59q6CypZYIwgyJ892HpttybHKg1ZtQLUlSXccRMlugPgEcNZJagPEgPYni4
# b11snjRAgf0dyQ0zI9aLXqTxWUU5pCIFiPT0b2wsxzRqCtyGqpkGM8P9GazO8eao
# mVItCYBcJSByBx/pS0cSYwBBHAZxJODUqxSXoSGDvmTfqUJXntnWkL4okok1FiCD
# Z4jpyXOQunb6egIXvkgQ7jb2uO26Ow0m8RwleDvhOMrnHsupiOPbozKroSa6paFt
# VSh89abUSooR8QdZciemmoFhcWkEwFg4spzvYNP4nIs193261WyTaRMZoceGun7G
# CT2Rl653uUj+F+g94c63AhzSq4khdL4HlFIP2ePv29smfUnHtGq6yYFDLnT0q/Y+
# Di3jwloF8EWkkHRtSuXlFUbTmwr/lDDgbpZiKhLS7CBTDj32I0L5i532+uHczw82
# oZDmYmYmIUSMbZOgS65h797rj5JJ6OkeEUJoAVwwggd6MIIFYqADAgECAgphDpDS
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
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMQITMwAAAd9r8C6Sp0q00AAAAAAB3zAN
# BglghkgBZQMEAgEFAKCBrjAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQg8BPF6AqA
# d1t3hIsH6hHLDnqiJ0XSO8yNlpUp6j2l22IwQgYKKwYBBAGCNwIBDDE0MDKgFIAS
# AE0AaQBjAHIAbwBzAG8AZgB0oRqAGGh0dHA6Ly93d3cubWljcm9zb2Z0LmNvbTAN
# BgkqhkiG9w0BAQEFAASCAQCk0ykaSz6IQ2uPYNo46HrSFK4yg8Wg053QXDo0aQiw
# evNcVJLuq8IsY3B2DkD4Fp9zY88lMPeMJlx7hlXiAl95nMjSRkVMN/4DJwW7DJ8T
# eIGctyCDeXDgB1Rk0/MZ+Tx2JYRMwqz3jgrSHuPC9ays+YU6JrknkE94/oPXOTD+
# BugJSsVQz2k/XIq/gqhQbbRN7a4DtFH6DtbtkABzi2rSVE5eCOiAwgwuY1AElVH8
# XBtP9fYBNcoYtlm+2w1OqM8owUa1bwe2tE6yYsy1HBMxexGDucJpdFuKJhUe61a9
# +akMc9tDwFz9688K4SvgihIUxw3l5Jr8pN+qxaVhvnDdoYIS8TCCEu0GCisGAQQB
# gjcDAwExghLdMIIS2QYJKoZIhvcNAQcCoIISyjCCEsYCAQMxDzANBglghkgBZQME
# AgEFADCCAVUGCyqGSIb3DQEJEAEEoIIBRASCAUAwggE8AgEBBgorBgEEAYRZCgMB
# MDEwDQYJYIZIAWUDBAIBBQAEINdMRO8i3t3XLRtmfd2x7Oz8dGqqb6jPqj38cZ4x
# 17sOAgZgr7WGZgwYEzIwMjEwNjE4MDY1NTMxLjk3NVowBIACAfSggdSkgdEwgc4x
# CzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRt
# b25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKTAnBgNVBAsTIE1p
# Y3Jvc29mdCBPcGVyYXRpb25zIFB1ZXJ0byBSaWNvMSYwJAYDVQQLEx1UaGFsZXMg
# VFNTIEVTTjpDNEJELUUzN0YtNUZGQzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUt
# U3RhbXAgU2VydmljZaCCDkQwggT1MIID3aADAgECAhMzAAABV0QHYtxv6L4qAAAA
# AAFXMA0GCSqGSIb3DQEBCwUAMHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNo
# aW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29y
# cG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEw
# MB4XDTIxMDExNDE5MDIxM1oXDTIyMDQxMTE5MDIxM1owgc4xCzAJBgNVBAYTAlVT
# MRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQK
# ExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKTAnBgNVBAsTIE1pY3Jvc29mdCBPcGVy
# YXRpb25zIFB1ZXJ0byBSaWNvMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjpDNEJE
# LUUzN0YtNUZGQzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vydmlj
# ZTCCASIwDQYJKoZIhvcNAQEBBQADggEPADCCAQoCggEBAN5tA6dUZvnnwL9qQtXc
# wPANhB4ez+5CQrePp/Z8TH4NBr5vAfGMo0lV/lidBatKTgHErOuKH11xVAfBehHJ
# vH9T/OhOc83CJs9bzDhrld0Jdy3eJyC0yBdxVeucS+2a2ZBd50wBg/5/2YjQ2ylf
# D0dxKK6tQLxdODTuadQMbda05lPGnWGwZ3niSgIKVRgqqCVlhHzwNtRh1AH+Zxbf
# Se7t8z3oEKAdTAy7SsP8ykht3srjdh0BykPFdpaAgqwWCJJJmGk0gArSvHC8+vXt
# Go3MJhWQRe5JtzdD5kdaKH9uc9gnShsXyDEhGZjx3+b8cuqEO8bHv0WPX9MREfrf
# xvkCAwEAAaOCARswggEXMB0GA1UdDgQWBBRdMXu76DghnU/kPTMKdFkR9oCp2TAf
# BgNVHSMEGDAWgBTVYzpcijGQ80N7fEYbxTNoWoVtVTBWBgNVHR8ETzBNMEugSaBH
# hkVodHRwOi8vY3JsLm1pY3Jvc29mdC5jb20vcGtpL2NybC9wcm9kdWN0cy9NaWNU
# aW1TdGFQQ0FfMjAxMC0wNy0wMS5jcmwwWgYIKwYBBQUHAQEETjBMMEoGCCsGAQUF
# BzAChj5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1RpbVN0
# YVBDQV8yMDEwLTA3LTAxLmNydDAMBgNVHRMBAf8EAjAAMBMGA1UdJQQMMAoGCCsG
# AQUFBwMIMA0GCSqGSIb3DQEBCwUAA4IBAQAld3kAgG6XWiZyvdibLRmWr7yb6RSy
# cjVDg8tcCitS01sTVp4T8Ad2QeYfJWfK6DMEk7QRBfKgdN7oE8dXtmQVL+JcxLj0
# pUuy4NB5RchcteD5dRnTfKlRi8vgKUaxDcoFIzNEUz1EHpopeagDb4/uI9Uj5tIu
# wlik/qrv/sHAw7kM4gELLNOgdev9Z/7xo1JIwfe0eoQM3wxcCFLuf8S9OncttaFA
# WHtEER8IvgRAgLJ/WnluFz68+hrDfRyX/qqWSPIE0voE6qFx1z8UvLwKpm65QNyN
# DRMp/VmCpqRZrxB1o0RY7P+n4jSNGvbk2bR70kKt/dogFFRBHVVuUxf+MIIGcTCC
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
# cGVyYXRpb25zIFB1ZXJ0byBSaWNvMSYwJAYDVQQLEx1UaGFsZXMgVFNTIEVTTjpD
# NEJELUUzN0YtNUZGQzElMCMGA1UEAxMcTWljcm9zb2Z0IFRpbWUtU3RhbXAgU2Vy
# dmljZaIjCgEBMAcGBSsOAwIaAxUAES34SWJ7DfbSG/gbIQwTrzgZ8PKggYMwgYCk
# fjB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMH
# UmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQD
# Ex1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDANBgkqhkiG9w0BAQUFAAIF
# AOR2i6MwIhgPMjAyMTA2MTgwNzA0MDNaGA8yMDIxMDYxOTA3MDQwM1owdzA9Bgor
# BgEEAYRZCgQBMS8wLTAKAgUA5HaLowIBADAKAgEAAgIg2gIB/zAHAgEAAgIRlDAK
# AgUA5HfdIwIBADA2BgorBgEEAYRZCgQCMSgwJjAMBgorBgEEAYRZCgMCoAowCAIB
# AAIDB6EgoQowCAIBAAIDAYagMA0GCSqGSIb3DQEBBQUAA4GBAKJ6xThFmJ3V2mql
# isgacPGyukJfggoSad2LkQaxI9Gpi7jPFrBexa+xtdXJfcN+zTNmENwaoTZvTfSR
# Hjsnya4EMwZnse6ayGAZ3haWlnBWvSbhWX4U5yPX2o2E7FCAjCESQxf6fpK14wlW
# BxfPNULTbdf5Ld4kFkkoy4IX3MbIMYIDDTCCAwkCAQEwgZMwfDELMAkGA1UEBhMC
# VVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQxHjAcBgNV
# BAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWljcm9zb2Z0IFRp
# bWUtU3RhbXAgUENBIDIwMTACEzMAAAFXRAdi3G/ovioAAAAAAVcwDQYJYIZIAWUD
# BAIBBQCgggFKMBoGCSqGSIb3DQEJAzENBgsqhkiG9w0BCRABBDAvBgkqhkiG9w0B
# CQQxIgQgALPJu0UFlImBE3z0cAsVMugtE5ql7iE+rBNWNSrLG/8wgfoGCyqGSIb3
# DQEJEAIvMYHqMIHnMIHkMIG9BCAsWo0NQ6vzuuupUsZEMSJ4UsRjtQw2dFxZWkHt
# qRygEzCBmDCBgKR+MHwxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9u
# MRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRp
# b24xJjAkBgNVBAMTHU1pY3Jvc29mdCBUaW1lLVN0YW1wIFBDQSAyMDEwAhMzAAAB
# V0QHYtxv6L4qAAAAAAFXMCIEIMbk8cfabbCSWyYAPeKgD+acumPF5pEk8aEvy+O7
# +F7KMA0GCSqGSIb3DQEBCwUABIIBAJGhAQpcZZH02EKZO+6LLfhQLU7AVuzmCsCc
# hcXYUA3OwJmD+FiaNdCGI04ICI4O4abEN4hjW13tFgWit1h06PwOS41pNh6kTkMP
# PnyHkVXPA5xLjogJwejLvdJTBFkquXNJm6EcCcpb8uV+cscfJ1ETkp1PbgxMyKqZ
# 0Ou7PuyDG/btZDCmEzJICb/D+eGunzwjnP6ejztj6zMJSX8lxCcGn0LyC7BVbQYz
# sGQsKLfnnaZdulCFvjFN0qroWFoVPTh+PJjlUVH6JuXY54twhsiMGY1bL1nXt5U7
# wSB+tZdnde+dZRPRx3Q3ia8drQ7hkHtdKA3EFZcNQb8pDsVERo8=
# SIG # End signature block
