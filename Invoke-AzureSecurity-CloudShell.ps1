<#
.SYNOPSIS
    Exports Microsoft security posture and Defender portal data for the workshop.

.DESCRIPTION
    Collects Defender for Cloud posture through Azure Resource Graph, incidents and
    alerts through Microsoft Graph Security, and optional Defender for Endpoint data.
    Each source is isolated so missing licenses or permissions do not stop the export.

.PARAMETER LookbackDays
    Limits Microsoft Graph incidents and alerts to the most recent number of days.

.PARAMETER SkipDefenderPortal
    Skips Microsoft Graph Security and Defender for Endpoint collection.

.PARAMETER SkipGraphSecurity
    Skips interactive Microsoft Graph Security collection while retaining Defender
    for Cloud and Defender for Endpoint data.

.PARAMETER MaxEndpointRecords
    Maximum records exported from each Defender for Endpoint API dataset.

.EXAMPLE
    ./Invoke-AzureSecurity-CloudShell.ps1
    ./Invoke-AzureSecurity-CloudShell.ps1 -LookbackDays 90
#>

[CmdletBinding()]
param(
    [ValidateRange(1, 365)]
    [int]$LookbackDays = 30,
    [switch]$SkipDefenderPortal,
    [switch]$SkipGraphSecurity,
    [ValidateRange(1, 100000)]
    [int]$MaxEndpointRecords = 5000
)

$ErrorActionPreference = 'Stop'
$outputDirectory = if ($env:AZWORKSHOP_OUTPUT) { $env:AZWORKSHOP_OUTPUT } else { $HOME }
$outputFile = Join-Path $outputDirectory "AzureSecurity_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"
$sourceStatus = [System.Collections.Generic.List[object]]::new()

New-Item -ItemType Directory -Path $outputDirectory -Force | Out-Null

foreach ($moduleName in @('Az.Accounts', 'Az.ResourceGraph', 'ImportExcel')) {
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        Write-Host "  Installing $moduleName..." -ForegroundColor DarkYellow
        Install-Module -Name $moduleName -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck
    }
    Import-Module $moduleName -ErrorAction Stop
}

$subscriptionIds = @($env:AZWORKSHOP_SUBSCRIPTION_IDS -split ',' | Where-Object { $_ })
if ($subscriptionIds.Count -eq 0) {
    $subscriptionIds = @(Get-AzSubscription -ErrorAction Stop | Where-Object { $_.State -eq 'Enabled' } | Select-Object -ExpandProperty Id)
}
if ($subscriptionIds.Count -eq 0) { throw "No enabled subscriptions are accessible in the current tenant." }

function Add-SourceStatus {
    param(
        [string]$Source,
        [string]$Status,
        [int]$Records,
        [string]$Details,
        [string]$RequiredAccess
    )

    $sourceStatus.Add([PSCustomObject]@{
        Source         = $Source
        Status         = $Status
        Records        = $Records
        Details        = $Details
        RequiredAccess = $RequiredAccess
        CollectedAtUtc = (Get-Date).ToUniversalTime().ToString('o')
    })
}

function Get-HttpStatusCode {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)

    try { return [int]$ErrorRecord.Exception.Response.StatusCode } catch { return 0 }
}

function Get-SourceFailureStatus {
    param([System.Management.Automation.ErrorRecord]$ErrorRecord)

    $statusCode = Get-HttpStatusCode -ErrorRecord $ErrorRecord
    if ($statusCode -in @(401, 403)) { return 'Forbidden' }
    if ($statusCode -eq 404) { return 'Unavailable' }
    return 'Error'
}

function ConvertTo-CellValue {
    param($Value)

    if ($null -eq $Value) { return '' }
    if ($Value -is [string] -or $Value -is [ValueType]) { return $Value }
    if ($Value -is [System.Collections.IEnumerable]) {
        return (@($Value) | ForEach-Object { [string]$_ }) -join '; '
    }
    return $Value | ConvertTo-Json -Compress -Depth 8
}

function Write-SecuritySheet {
    param(
        [string]$WorksheetName,
        [object[]]$Rows,
        [string]$EmptyMessage = 'No data returned by this source.'
    )

    $exportRows = @($Rows)
    if ($exportRows.Count -eq 0) {
        $exportRows = @([PSCustomObject]@{
            DataStatus = 'NoData'
            Message    = $EmptyMessage
        })
    }

    $exportRows | Export-Excel -Path $outputFile -WorksheetName $WorksheetName `
        -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Medium6
    Write-Host "  [$WorksheetName] $($Rows.Count) data rows" -ForegroundColor Gray
}

function Invoke-SearchAzGraphWithRetry {
    param([hashtable]$Parameters, [int]$MaxAttempts = 5)
    $attempt = 0
    do {
        try {
            return Search-AzGraph @Parameters -ErrorAction Stop
        } catch {
            $attempt++
            $statusCode = Get-HttpStatusCode -ErrorRecord $_
            if ($attempt -ge $MaxAttempts -or $statusCode -notin @(408, 429, 500, 502, 503, 504)) { throw }
            $delaySeconds = [math]::Min(30, [math]::Pow(2, $attempt))
            Write-Host "  Resource Graph returned HTTP $statusCode. Retrying in $delaySeconds seconds..." -ForegroundColor DarkYellow
            Start-Sleep -Seconds $delaySeconds
        }
    } while ($true)
}

function Invoke-AzureResourceGraphQuery {
    param([string]$Query)

    $allRows = [System.Collections.Generic.List[object]]::new()
    $skipToken = $null
    do {
        $parameters = @{ Query = $Query; Subscription = $subscriptionIds; First = 1000 }
        if ($skipToken) { $parameters.SkipToken = $skipToken }
        $result = Invoke-SearchAzGraphWithRetry -Parameters $parameters
        foreach ($row in @($result.Data)) { $allRows.Add($row) }
        $skipToken = $result.SkipToken
    } while ($skipToken)
    return $allRows.ToArray()
}

function Invoke-PostureCollection {
    $queries = [ordered]@{
        SecureScores = @'
SecurityResources
| where type == 'microsoft.security/securescores'
| extend percentageScore=todouble(properties.score.percentage), currentScore=todouble(properties.score.current), maxScore=todouble(properties.score.max), weight=todouble(properties.weight)
| project tenantId, subscriptionId, percentageScore, currentScore, maxScore, weight
'@
        ScoreControls = @'
SecurityResources
| where type == 'microsoft.security/securescores/securescorecontrols'
| extend controlName=tostring(properties.displayName), controlId=tostring(properties.definition.name), notApplicableResourceCount=toint(properties.notApplicableResourceCount), unhealthyResourceCount=toint(properties.unhealthyResourceCount), healthyResourceCount=toint(properties.healthyResourceCount), percentageScore=todouble(properties.score.percentage), currentScore=todouble(properties.score.current), maxScore=todouble(properties.definition.properties.maxScore), weight=todouble(properties.weight), controlType=tostring(properties.definition.properties.source.sourceType)
| project tenantId, subscriptionId, controlName, controlId, unhealthyResourceCount, healthyResourceCount, notApplicableResourceCount, percentageScore, currentScore, maxScore, weight, controlType
| order by unhealthyResourceCount desc
'@
        Recommendations = @'
SecurityResources
| where type == 'microsoft.security/assessments'
| extend recommendationId=name, recommendationName=tostring(properties.displayName), recommendationState=tostring(properties.status.code), recommendationSeverity=tostring(properties.metadata.severity), description=tostring(properties.metadata.description), remediationDescription=tostring(properties.metadata.remediationDescription), assessmentType=tostring(properties.metadata.assessmentType), policyDefinitionId=tostring(properties.metadata.policyDefinitionId), implementationEffort=tostring(properties.metadata.implementationEffort), userImpact=tostring(properties.metadata.userImpact), category=tostring(properties.metadata.categories), threats=tostring(properties.metadata.threats), source=tostring(properties.resourceDetails.Source), affectedResourceId=tostring(properties.resourceDetails.Id), portalLink=tostring(properties.links.azurePortal)
| project tenantId, subscriptionId, recommendationId, recommendationName, recommendationState, recommendationSeverity, affectedResourceId, description, remediationDescription, assessmentType, policyDefinitionId, implementationEffort, userImpact, category, threats, source, portalLink
| order by recommendationSeverity asc, recommendationState desc
'@
        RegulatoryStandards = @'
SecurityResources
| where type == 'microsoft.security/regulatorycompliancestandards'
| extend complianceStandard=name, state=tostring(properties.state), passedControls=toint(properties.passedControls), failedControls=toint(properties.failedControls), skippedControls=toint(properties.skippedControls), unsupportedControls=toint(properties.unsupportedControls)
| project tenantId, subscriptionId, complianceStandard, state, passedControls, failedControls, skippedControls, unsupportedControls
'@
        MCSBCompliance = @'
SecurityResources
| where type == 'microsoft.security/regulatorycompliancestandards/regulatorycompliancecontrols/regulatorycomplianceassessments'
| extend complianceStandard=extract(@'(?i)/regulatorycompliancestandards/([^/]+)', 1, id), complianceControl=extract(@'(?i)/regulatorycompliancecontrols/([^/]+)', 1, id), assessmentName=tostring(properties.description), state=tostring(properties.state), skippedResources=toint(properties.skippedResources), passedResources=toint(properties.passedResources), failedResources=toint(properties.failedResources)
| where complianceStandard contains 'Azure-Security-Benchmark' or complianceStandard contains 'Microsoft-cloud-security-benchmark' or complianceStandard contains 'MCSB'
| project tenantId, subscriptionId, complianceStandard, complianceControl, assessmentName, state, skippedResources, passedResources, failedResources, id
| order by failedResources desc
'@
        DefenderPlans = @'
SecurityResources
| where type == 'microsoft.security/pricings'
| extend planName=name, pricingTier=tostring(properties.pricingTier), subPlan=tostring(properties.subPlan), extensions=tostring(properties.extensions)
| project tenantId, subscriptionId, planName, pricingTier, subPlan, extensions
| order by subscriptionId asc, planName asc
'@
        CloudAlerts = @'
SecurityResources
| where type =~ 'microsoft.security/locations/alerts'
| extend alertName=tostring(properties.AlertDisplayName), alertType=tostring(properties.AlertType), systemAlertId=tostring(properties.SystemAlertId), status=tostring(properties.Status), severity=tostring(properties.Severity), description=tostring(properties.Description), remediationSteps=tostring(properties.RemediationSteps), detectedTime=todatetime(properties.DetectedTimeUtc), compromisedEntity=tostring(properties.CompromisedEntity), portalLink=tostring(properties.AlertUri)
| project tenantId, subscriptionId, alertName, alertType, systemAlertId, status, severity, description, remediationSteps, detectedTime, compromisedEntity, portalLink
| order by detectedTime desc
'@
    }

    $totalRecords = 0
    $failedQueries = [System.Collections.Generic.List[string]]::new()
    foreach ($entry in $queries.GetEnumerator()) {
        try {
            $rows = @(Invoke-AzureResourceGraphQuery -Query $entry.Value)
            $totalRecords += $rows.Count
            Write-SecuritySheet -WorksheetName $entry.Key -Rows $rows
        } catch {
            $failedQueries.Add($entry.Key)
            Write-SecuritySheet -WorksheetName $entry.Key -Rows @() -EmptyMessage "Collection failed: $($_.Exception.Message)"
            Write-Warning "$($entry.Key) collection failed: $($_.Exception.Message)"
        }
    }

    $status = if ($failedQueries.Count -eq 0) { 'Available' } elseif ($failedQueries.Count -eq $queries.Count) { 'Unavailable' } else { 'Partial' }
    $details = if ($failedQueries.Count -eq 0) {
        'Secure score, controls, assessments, MCSB, plans, and cloud alerts collected.'
    } else {
        "Failed datasets: $($failedQueries -join ', ')."
    }
    Add-SourceStatus -Source 'Defender for Cloud / Azure Resource Graph' -Status $status `
        -Records $totalRecords -Details $details `
        -RequiredAccess 'Azure Reader and Microsoft.Security read access at the assessed scopes.'
}

function Connect-WorkshopGraphSecurity {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Write-Host '  Installing Microsoft.Graph.Authentication...' -ForegroundColor DarkYellow
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    $requiredScopes = @('SecurityIncident.Read.All', 'SecurityAlert.Read.All')
    $graphContext = Get-MgContext
    $missingScopes = @($requiredScopes | Where-Object { -not $graphContext -or $_ -notin $graphContext.Scopes })
    if (-not $graphContext -or $missingScopes.Count -gt 0) {
        $azContext = Get-AzContext
        $connectParameters = @{ Scopes = $requiredScopes; NoWelcome = $true }
        if ($azContext.Tenant.Id) { $connectParameters.TenantId = $azContext.Tenant.Id }
        Connect-MgGraph @connectParameters | Out-Null
    }
}

function Invoke-GraphPagedRequest {
    param([string]$Uri)

    $allRows = [System.Collections.Generic.List[object]]::new()
    $nextLink = $Uri
    while ($nextLink) {
        $attempt = 0
        do {
            try {
                $response = Invoke-MgGraphRequest -Method GET -Uri $nextLink -OutputType PSObject
                break
            } catch {
                $attempt++
                $statusCode = Get-HttpStatusCode -ErrorRecord $_
                if ($attempt -ge 5 -or $statusCode -notin @(408, 429, 500, 502, 503, 504)) { throw }
                $delaySeconds = [math]::Min(30, [math]::Pow(2, $attempt))
                Write-Host "  Graph returned HTTP $statusCode. Retrying in $delaySeconds seconds..." -ForegroundColor DarkYellow
                Start-Sleep -Seconds $delaySeconds
            }
        } while ($true)

        foreach ($row in @($response.value)) { $allRows.Add($row) }
        $nextLink = $response.'@odata.nextLink'
    }
    return $allRows.ToArray()
}

function Invoke-GraphSecurityCollection {
    try {
        Connect-WorkshopGraphSecurity
    } catch {
        $failureStatus = Get-SourceFailureStatus -ErrorRecord $_
        foreach ($sourceName in @('Microsoft Graph Security incidents', 'Microsoft Graph Security alerts')) {
            Add-SourceStatus -Source $sourceName -Status $failureStatus -Records 0 `
                -Details $_.Exception.Message `
                -RequiredAccess 'Delegated SecurityIncident.Read.All or SecurityAlert.Read.All plus a supported Entra security role.'
        }
            Write-SecuritySheet -WorksheetName 'Incidents' -Rows @() -EmptyMessage 'Microsoft Graph Security authentication was unavailable.'
            Write-SecuritySheet -WorksheetName 'Alerts' -Rows @() -EmptyMessage 'Microsoft Graph Security authentication was unavailable.'
        Write-Warning "Microsoft Graph Security authentication failed: $($_.Exception.Message)"
        return
    }

    $since = (Get-Date).ToUniversalTime().AddDays(-$LookbackDays).ToString('yyyy-MM-ddTHH:mm:ssZ')
    $incidentUri = "https://graph.microsoft.com/v1.0/security/incidents?`$top=100&`$filter=lastUpdateDateTime%20ge%20$since"
    try {
        $rawIncidents = @(Invoke-GraphPagedRequest -Uri $incidentUri)
        $incidents = @($rawIncidents | ForEach-Object {
            [PSCustomObject]@{
                IncidentId        = $_.id
                DisplayName       = $_.displayName
                Status            = $_.status
                Severity          = $_.severity
                Classification    = $_.classification
                Determination     = $_.determination
                AssignedTo        = $_.assignedTo
                CreatedDateTime   = $_.createdDateTime
                LastUpdateDateTime = $_.lastUpdateDateTime
                PriorityScore     = $_.priorityScore
                Description       = $_.description
                Summary           = $_.summary
                Tags              = ConvertTo-CellValue $_.customTags
                IncidentWebUrl    = $_.incidentWebUrl
            }
        })
        Write-SecuritySheet -WorksheetName 'Incidents' -Rows $incidents
        Add-SourceStatus -Source 'Microsoft Graph Security incidents' -Status $(if ($incidents.Count) { 'Available' } else { 'NoData' }) `
            -Records $incidents.Count -Details "Incidents updated in the last $LookbackDays days." `
            -RequiredAccess 'SecurityIncident.Read.All and Security Reader, Global Reader, Security Operator, or Security Administrator.'
    } catch {
        Add-SourceStatus -Source 'Microsoft Graph Security incidents' -Status (Get-SourceFailureStatus $_) -Records 0 `
            -Details $_.Exception.Message `
            -RequiredAccess 'SecurityIncident.Read.All and a supported Entra security role.'
        Write-SecuritySheet -WorksheetName 'Incidents' -Rows @() -EmptyMessage "Incident collection failed: $($_.Exception.Message)"
        Write-Warning "Incident collection failed: $($_.Exception.Message)"
    }

    $alertUri = "https://graph.microsoft.com/v1.0/security/alerts_v2?`$top=100&`$filter=lastUpdateDateTime%20ge%20$since"
    try {
        $rawAlerts = @(Invoke-GraphPagedRequest -Uri $alertUri)
        $alerts = @($rawAlerts | ForEach-Object {
            [PSCustomObject]@{
                AlertId           = $_.id
                IncidentId        = $_.incidentId
                Title             = $_.title
                Status            = $_.status
                Severity          = $_.severity
                Classification    = $_.classification
                Determination     = $_.determination
                ServiceSource     = $_.serviceSource
                DetectionSource   = $_.detectionSource
                Category          = $_.category
                AssignedTo        = $_.assignedTo
                CreatedDateTime   = $_.createdDateTime
                LastUpdateDateTime = $_.lastUpdateDateTime
                MitreTechniques   = ConvertTo-CellValue $_.mitreTechniques
                Description       = $_.description
                RecommendedActions = $_.recommendedActions
                AlertWebUrl       = $_.alertWebUrl
                IncidentWebUrl    = $_.incidentWebUrl
            }
        })
        Write-SecuritySheet -WorksheetName 'Alerts' -Rows $alerts
        Add-SourceStatus -Source 'Microsoft Graph Security alerts' -Status $(if ($alerts.Count) { 'Available' } else { 'NoData' }) `
            -Records $alerts.Count -Details "Alerts updated in the last $LookbackDays days." `
            -RequiredAccess 'SecurityAlert.Read.All and Security Reader, Global Reader, Security Operator, or Security Administrator.'
    } catch {
        Add-SourceStatus -Source 'Microsoft Graph Security alerts' -Status (Get-SourceFailureStatus $_) -Records 0 `
            -Details $_.Exception.Message `
            -RequiredAccess 'SecurityAlert.Read.All and a supported Entra security role.'
        Write-SecuritySheet -WorksheetName 'Alerts' -Rows @() -EmptyMessage "Alert collection failed: $($_.Exception.Message)"
        Write-Warning "Alert collection failed: $($_.Exception.Message)"
    }
}

function ConvertFrom-SecureAccessToken {
    param($Token)

    if ($Token -is [System.Security.SecureString]) {
        return [System.Net.NetworkCredential]::new('', $Token).Password
    }
    return [string]$Token
}

function Invoke-DefenderEndpointPagedRequest {
    param(
        [string]$Uri,
        [string]$AccessToken,
        [int]$MaxRecords
    )

    $allRows = [System.Collections.Generic.List[object]]::new()
    $nextLink = $Uri
    $truncated = $false
    $headers = @{ Authorization = "Bearer $AccessToken"; Accept = 'application/json' }
    while ($nextLink) {
        $attempt = 0
        do {
            try {
                $response = Invoke-RestMethod -Method GET -Uri $nextLink -Headers $headers -TimeoutSec 120
                break
            } catch {
                $attempt++
                $statusCode = Get-HttpStatusCode -ErrorRecord $_
                if ($attempt -ge 5 -or $statusCode -notin @(408, 429, 500, 502, 503, 504)) { throw }
                $delaySeconds = [math]::Min(30, [math]::Pow(2, $attempt))
                Write-Host "  Defender for Endpoint returned HTTP $statusCode. Retrying in $delaySeconds seconds..." -ForegroundColor DarkYellow
                Start-Sleep -Seconds $delaySeconds
            }
        } while ($true)

        foreach ($row in @($response.value)) {
            if ($allRows.Count -ge $MaxRecords) {
                $truncated = $true
                break
            }
            $allRows.Add($row)
        }
        $nextLink = if ($response.'@odata.nextLink') { $response.'@odata.nextLink' } else { $null }
        if ($allRows.Count -ge $MaxRecords -and $nextLink) {
            $truncated = $true
            break
        }
    }
    return [PSCustomObject]@{
        Rows      = $allRows.ToArray()
        Truncated = $truncated
    }
}

function Invoke-DefenderEndpointCollection {
    try {
        $tokenResponse = Get-AzAccessToken -ResourceUrl 'https://api.securitycenter.microsoft.com'
        $accessToken = ConvertFrom-SecureAccessToken $tokenResponse.Token
        if (-not $accessToken) { throw 'Azure PowerShell did not return a Defender for Endpoint access token.' }
    } catch {
        $unavailableCollections = @(
            @('Defender for Endpoint machines', 'Machines'),
            @('Defender for Endpoint recommendations', 'EndpointRecommendations'),
            @('Defender for Endpoint vulnerabilities', 'Vulnerabilities')
        )
        foreach ($collection in $unavailableCollections) {
            Add-SourceStatus -Source $collection[0] -Status 'Unavailable' -Records 0 -Details $_.Exception.Message `
                -RequiredAccess 'A token for api.securitycenter.microsoft.com with the corresponding delegated or application permission.'
            Write-SecuritySheet -WorksheetName $collection[1] -Rows @() -EmptyMessage 'Defender for Endpoint authentication was unavailable.'
        }
        Write-Warning "Defender for Endpoint token acquisition failed: $($_.Exception.Message)"
        return
    }

    $collections = @(
        @{
            Source = 'Defender for Endpoint machines'; Sheet = 'Machines'; Uri = 'https://api.securitycenter.microsoft.com/api/machines'; Permission = 'Machine.Read.All'
            Map = {
                param($row)
                [PSCustomObject]@{ MachineId=$row.id; ComputerDnsName=$row.computerDnsName; OsPlatform=$row.osPlatform; OsVersion=$row.osVersion; HealthStatus=$row.healthStatus; RiskScore=$row.riskScore; ExposureLevel=$row.exposureLevel; OnboardingStatus=$row.onboardingStatus; LastSeen=$row.lastSeen; AadDeviceId=$row.aadDeviceId; RbacGroupName=$row.rbacGroupName }
            }
        },
        @{
            Source = 'Defender for Endpoint recommendations'; Sheet = 'EndpointRecommendations'; Uri = 'https://api.securitycenter.microsoft.com/api/recommendations'; Permission = 'SecurityRecommendation.Read.All'
            Map = {
                param($row)
                [PSCustomObject]@{ RecommendationId=$row.id; ProductName=$row.productName; RecommendationName=$row.recommendationName; Weaknesses=ConvertTo-CellValue $row.weaknesses; Vendor=$row.vendor; RecommendedVersion=$row.recommendedVersion; SeverityScore=$row.severityScore; PublicExploit=$row.publicExploit; ActiveAlert=$row.activeAlert; AssociatedThreats=ConvertTo-CellValue $row.associatedThreats; ExposedMachinesCount=$row.exposedMachinesCount; RemediationType=$row.remediationType; Status=$row.status }
            }
        },
        @{
            Source = 'Defender for Endpoint vulnerabilities'; Sheet = 'Vulnerabilities'; Uri = 'https://api.securitycenter.microsoft.com/api/vulnerabilities'; Permission = 'Vulnerability.Read.All'
            Map = {
                param($row)
                [PSCustomObject]@{ VulnerabilityId=$row.id; Name=$row.name; Description=$row.description; Severity=$row.severity; CvssV3=$row.cvssV3; ExposedMachines=$row.exposedMachines; PublishedOn=$row.publishedOn; UpdatedOn=$row.updatedOn; PublicExploit=$row.publicExploit; ExploitVerified=$row.exploitVerified; ExploitInKit=$row.exploitInKit; ExploitTypes=ConvertTo-CellValue $row.exploitTypes }
            }
        }
    )

    foreach ($collection in $collections) {
        try {
            $pageResult = Invoke-DefenderEndpointPagedRequest -Uri $collection.Uri -AccessToken $accessToken -MaxRecords $MaxEndpointRecords
            $rawRows = @($pageResult.Rows)
            $mappedRows = @($rawRows | ForEach-Object { & $collection.Map $_ })
            Write-SecuritySheet -WorksheetName $collection.Sheet -Rows $mappedRows
            $status = if ($pageResult.Truncated) { 'Partial' } elseif ($mappedRows.Count) { 'Available' } else { 'NoData' }
            $details = if ($pageResult.Truncated) { "Collection limited to $MaxEndpointRecords records." } else { 'Collection completed.' }
            Add-SourceStatus -Source $collection.Source -Status $status `
                -Records $mappedRows.Count -Details $details -RequiredAccess $collection.Permission
        } catch {
            Add-SourceStatus -Source $collection.Source -Status (Get-SourceFailureStatus $_) -Records 0 `
                -Details $_.Exception.Message -RequiredAccess $collection.Permission
            Write-SecuritySheet -WorksheetName $collection.Sheet -Rows @() -EmptyMessage "$($collection.Source) collection failed: $($_.Exception.Message)"
            Write-Warning "$($collection.Source) collection failed: $($_.Exception.Message)"
        }
    }
}

function Connect-WorkshopGraphIdentity {
    if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
        Write-Host '  Installing Microsoft.Graph.Authentication...' -ForegroundColor DarkYellow
        Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck
    }
    Import-Module Microsoft.Graph.Authentication -ErrorAction Stop

    $requiredScopes = @('Application.Read.All', 'User.Read.All')
    $graphContext = Get-MgContext
    $missingScopes = @($requiredScopes | Where-Object { -not $graphContext -or $_ -notin $graphContext.Scopes })
    if (-not $graphContext -or $missingScopes.Count -gt 0) {
        $azContext = Get-AzContext
        $connectParameters = @{ Scopes = $requiredScopes; NoWelcome = $true }
        if ($azContext.Tenant.Id) { $connectParameters.TenantId = $azContext.Tenant.Id }
        Connect-MgGraph @connectParameters | Out-Null
    }
}

function Invoke-IdentityRiskCollection {
    try {
        Connect-WorkshopGraphIdentity
    } catch {
        $failureStatus = Get-SourceFailureStatus -ErrorRecord $_
        foreach ($sourceName in @('Microsoft Entra app credential expiry', 'Microsoft Entra guest users')) {
            Add-SourceStatus -Source $sourceName -Status $failureStatus -Records 0 `
                -Details $_.Exception.Message `
                -RequiredAccess 'Delegated Application.Read.All or User.Read.All plus a supported Entra directory role.'
        }
        Write-SecuritySheet -WorksheetName 'AppCredentialExpiry' -Rows @() -EmptyMessage 'Microsoft Graph authentication was unavailable.'
        Write-SecuritySheet -WorksheetName 'GuestUsers' -Rows @() -EmptyMessage 'Microsoft Graph authentication was unavailable.'
        Write-Warning "Microsoft Graph identity authentication failed: $($_.Exception.Message)"
        return
    }

    $now = (Get-Date).ToUniversalTime()
    $appUri = "https://graph.microsoft.com/v1.0/applications?`$select=id,appId,displayName,passwordCredentials,keyCredentials&`$top=999"
    try {
        $rawApps = @(Invoke-GraphPagedRequest -Uri $appUri)
        $credentialRows = [System.Collections.Generic.List[object]]::new()
        foreach ($app in $rawApps) {
            foreach ($cred in @($app.passwordCredentials)) {
                if (-not $cred.endDateTime) { continue }
                $endDate = [datetime]$cred.endDateTime
                $daysLeft = [math]::Round(($endDate - $now).TotalDays)
                if ($daysLeft -gt 90) { continue }
                $credentialRows.Add([PSCustomObject]@{
                    AppDisplayName  = $app.displayName
                    AppId           = $app.appId
                    CredentialType  = 'Secret'
                    StartDateTime   = $cred.startDateTime
                    EndDateTime     = $cred.endDateTime
                    DaysUntilExpiry = $daysLeft
                    Status          = if ($daysLeft -lt 0) { 'Expired' } else { 'ExpiringSoon' }
                })
            }
            foreach ($cred in @($app.keyCredentials)) {
                if (-not $cred.endDateTime) { continue }
                $endDate = [datetime]$cred.endDateTime
                $daysLeft = [math]::Round(($endDate - $now).TotalDays)
                if ($daysLeft -gt 90) { continue }
                $credentialRows.Add([PSCustomObject]@{
                    AppDisplayName  = $app.displayName
                    AppId           = $app.appId
                    CredentialType  = 'Certificate'
                    StartDateTime   = $cred.startDateTime
                    EndDateTime     = $cred.endDateTime
                    DaysUntilExpiry = $daysLeft
                    Status          = if ($daysLeft -lt 0) { 'Expired' } else { 'ExpiringSoon' }
                })
            }
        }
        $credentials = @($credentialRows.ToArray() | Sort-Object DaysUntilExpiry)
        Write-SecuritySheet -WorksheetName 'AppCredentialExpiry' -Rows $credentials
        Add-SourceStatus -Source 'Microsoft Entra app credential expiry' -Status $(if ($credentials.Count) { 'Available' } else { 'NoData' }) `
            -Records $credentials.Count -Details 'App registration secrets/certificates expiring within 90 days or already expired.' `
            -RequiredAccess 'Application.Read.All and Cloud Application Administrator, Application Administrator, or Global Reader.'
    } catch {
        Add-SourceStatus -Source 'Microsoft Entra app credential expiry' -Status (Get-SourceFailureStatus $_) -Records 0 `
            -Details $_.Exception.Message `
            -RequiredAccess 'Application.Read.All and a supported Entra directory role.'
        Write-SecuritySheet -WorksheetName 'AppCredentialExpiry' -Rows @() -EmptyMessage "App credential collection failed: $($_.Exception.Message)"
        Write-Warning "App credential collection failed: $($_.Exception.Message)"
    }

    $guestUri = "https://graph.microsoft.com/v1.0/users?`$filter=userType eq 'Guest'&`$select=id,displayName,mail,createdDateTime,accountEnabled&`$top=999"
    try {
        $rawGuests = @(Invoke-GraphPagedRequest -Uri $guestUri)
        $guests = @($rawGuests | ForEach-Object {
            [PSCustomObject]@{
                DisplayName     = $_.displayName
                Mail            = $_.mail
                CreatedDateTime = $_.createdDateTime
                AccountEnabled  = $_.accountEnabled
            }
        })
        Write-SecuritySheet -WorksheetName 'GuestUsers' -Rows $guests
        Add-SourceStatus -Source 'Microsoft Entra guest users' -Status $(if ($guests.Count) { 'Available' } else { 'NoData' }) `
            -Records $guests.Count -Details 'Guest (B2B) accounts in the tenant.' `
            -RequiredAccess 'User.Read.All and Global Reader or Directory Reader.'
    } catch {
        Add-SourceStatus -Source 'Microsoft Entra guest users' -Status (Get-SourceFailureStatus $_) -Records 0 `
            -Details $_.Exception.Message `
            -RequiredAccess 'User.Read.All and a supported Entra directory role.'
        Write-SecuritySheet -WorksheetName 'GuestUsers' -Rows @() -EmptyMessage "Guest user collection failed: $($_.Exception.Message)"
        Write-Warning "Guest user collection failed: $($_.Exception.Message)"
    }
}

Write-Host "`n=== Azure Security Assessment Export ===" -ForegroundColor Cyan
Write-Host "Output: $outputFile" -ForegroundColor Cyan
Write-Host "Lookback: $LookbackDays days" -ForegroundColor Cyan

$azureContext = Get-AzContext
if (-not $azureContext) {
    Write-Host 'No Azure context found. Starting interactive authentication...' -ForegroundColor Yellow
    Connect-AzAccount | Out-Null
}

Invoke-PostureCollection

if ($SkipDefenderPortal) {
    $skippedCollections = @(
        @('Microsoft Graph Security incidents', 'Incidents'),
        @('Microsoft Graph Security alerts', 'Alerts'),
        @('Microsoft Entra app credential expiry', 'AppCredentialExpiry'),
        @('Microsoft Entra guest users', 'GuestUsers'),
        @('Defender for Endpoint machines', 'Machines'),
        @('Defender for Endpoint recommendations', 'EndpointRecommendations'),
        @('Defender for Endpoint vulnerabilities', 'Vulnerabilities')
    )
    foreach ($collection in $skippedCollections) {
        Add-SourceStatus -Source $collection[0] -Status 'Skipped' -Records 0 -Details 'Skipped by parameter.' -RequiredAccess ''
        Write-SecuritySheet -WorksheetName $collection[1] -Rows @() -EmptyMessage 'Collection skipped by parameter.'
    }
} else {
    if ($SkipGraphSecurity) {
        foreach ($collection in @(
            @('Microsoft Graph Security incidents', 'Incidents'),
            @('Microsoft Graph Security alerts', 'Alerts'),
            @('Microsoft Entra app credential expiry', 'AppCredentialExpiry'),
            @('Microsoft Entra guest users', 'GuestUsers')
        )) {
            Add-SourceStatus -Source $collection[0] -Status 'Skipped' -Records 0 -Details 'Skipped by parameter.' -RequiredAccess ''
            Write-SecuritySheet -WorksheetName $collection[1] -Rows @() -EmptyMessage 'Microsoft Graph Security collection skipped by parameter.'
        }
    } else {
        Invoke-GraphSecurityCollection
        Invoke-IdentityRiskCollection
    }
    Invoke-DefenderEndpointCollection
}

Write-SecuritySheet -WorksheetName 'SourceStatus' -Rows $sourceStatus.ToArray()

Write-Host "`nSecurity export complete." -ForegroundColor Green
Write-Host "File: $outputFile" -ForegroundColor Cyan
Write-Host 'Review SourceStatus for permissions, licensing, or source availability gaps.' -ForegroundColor Yellow