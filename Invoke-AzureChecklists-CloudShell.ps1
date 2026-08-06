<#
  Azure Review Checklists (WAF) - Automated ARG Compliance Scan
  Runs each community checklist item's embedded Azure Resource Graph query against the
  current subscription(s), using the checklist JSON files vendored locally in the
  ./checklists folder (no network/GitHub access needed at run time — only Azure Resource
  Graph, same as every other script in this toolkit).
  Source: https://github.com/Azure/review-checklists (community-maintained, not an
  official Microsoft product; queries are read-only Resource Graph checks). Run
  Update-ReviewChecklists.ps1 to refresh the vendored copies from GitHub.
#>

param(
    # Optional: restrict to specific checklist file stems, e.g. -Services keyvault,aks,storage
    [string[]]$Services
)

$_outDir = if ($env:AZWORKSHOP_OUTPUT) { $env:AZWORKSHOP_OUTPUT } else { $HOME }
$outputFile = Join-Path $_outDir "AzureChecklists_$(Get-Date -Format 'yyyyMMdd_HHmmss').xlsx"

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel

$checklistsDir = Join-Path $PSScriptRoot "checklists"

# Checklist "waf" values (Reliability/Security/Cost/Operations/Performance) don't match
# this toolkit's pillar names 1:1, so translate them here.
$pillarMap = @{
    'reliability' = 'Reliability'
    'security'    = 'Security'
    'cost'        = 'Cost Optimization'
    'operations'  = 'Operational Excellence'
    'performance' = 'Performance Efficiency'
}

function ConvertTo-WafPillar {
    param([string]$Waf)
    $key = if ($Waf) { $Waf.Trim().ToLower() } else { '' }
    if ($pillarMap.ContainsKey($key)) { return $pillarMap[$key] }
    return 'Operational Excellence'
}

function Test-NonCompliant {
    param($Value)
    if ($null -eq $Value) { return $false }
    if ($Value -is [bool]) { return -not $Value }
    $s = $Value.ToString().Trim().ToLower()
    return $s -in @('false', 'non-compliant', 'noncompliant', 'no')
}

# ─── Resource Graph retry (handles 429/5xx throttling across large multi-subscription tenants;
# used by the sequential PS<7 fallback below - the parallel branch inlines its own copy since
# functions defined here aren't visible inside ForEach-Object -Parallel runspaces) ──
function Invoke-SearchAzGraphWithRetry {
    param([hashtable]$Parameters, [int]$MaxAttempts = 6)
    $attempt = 0
    do {
        try {
            return Search-AzGraph @Parameters -ErrorAction Stop
        } catch {
            $attempt++
            $statusCode = try { [int]$_.Exception.Response.StatusCode } catch { 0 }
            $transient = ($statusCode -in 429, 408, 500, 502, 503, 504) -or ($_.Exception.Message -match '429|too many requests|throttl|gateway timeout|server error|service unavailable')
            if ($attempt -ge $MaxAttempts -or -not $transient) { throw }
            $delaySeconds = [math]::Min(60, [math]::Pow(2, $attempt))
            Write-Host "  Resource Graph throttled or unavailable. Retrying in $delaySeconds seconds (attempt $attempt/$MaxAttempts)..." -ForegroundColor DarkYellow
            Start-Sleep -Seconds $delaySeconds
        }
    } while ($true)
}

Write-Host "`n=== Azure Review Checklists (community WAF checks) ===" -ForegroundColor Cyan

if (-not (Test-Path $checklistsDir)) {
    Write-Host "  ! Checklist folder not found: $checklistsDir" -ForegroundColor Yellow
    Write-Host "    Run Update-ReviewChecklists.ps1 once to vendor the checklist files." -ForegroundColor Yellow
    @([PSCustomObject]@{ Result = "Checklist folder not found: $checklistsDir" }) |
        Export-Excel -Path $outputFile -WorksheetName "Findings" -AutoSize
    return
}

$checklistFiles = @(Get-ChildItem -Path $checklistsDir -Filter '*_checklist.en.json')
if ($Services) {
    $checklistFiles = @($checklistFiles | Where-Object {
        $stem = $_.Name -replace '_checklist\.en\.json$', ''
        $Services -contains $stem
    })
}
Write-Host "  Found $($checklistFiles.Count) checklists to scan (local, from ./checklists).`n" -ForegroundColor Gray

$findings = @()
$queriesRun = 0
$queriesFailed = 0

# Flatten every checklist item across all files first, then group by the exact ARG
# query text. Many checklists embed identical/shared queries (e.g. generic tagging or
# naming checks reused across services) - deduplicating avoids re-running the same
# Resource Graph query over the network more than once.
$allItems = @()
foreach ($file in $checklistFiles) {
    try {
        $checklist = Get-Content -Path $file.FullName -Raw | ConvertFrom-Json
    } catch {
        Write-Host "  ! Skipped $($file.Name): $($_.Exception.Message)" -ForegroundColor DarkYellow
        continue
    }

    $graphItems = @($checklist.items | Where-Object { "$($_.graph)".Trim() -ne '' })
    if ($graphItems.Count -eq 0) { continue }

    $serviceName = if ($checklist.metadata.name) { $checklist.metadata.name } else { $file.Name -replace '_checklist\.en\.json$', '' }
    Write-Host "  [$serviceName] $($graphItems.Count) automated checks" -ForegroundColor Gray

    foreach ($item in $graphItems) {
        $allItems += [PSCustomObject]@{ Service = $serviceName; Item = $item }
    }
}

$byQuery = $allItems | Group-Object { $_.Item.graph.Trim() }
Write-Host "  $($allItems.Count) checks map to $($byQuery.Count) unique Resource Graph queries." -ForegroundColor Gray

# Each Search-AzGraph call costs ~1-2s of fixed round-trip/auth overhead regardless of
# subscription size, so with 100+ unique queries that dominates runtime. Run them
# concurrently (PS7+) using -DefaultProfile to pass the auth context per-call instead of
# mutating the shared/global Az context, which keeps this safe across threads.
$throttleLimit = 16
if ($PSVersionTable.PSVersion.Major -ge 7) {
    $azContext = Get-AzContext
    $resultsBag = [System.Collections.Concurrent.ConcurrentBag[object]]::new()
    $failCount = [System.Collections.Concurrent.ConcurrentBag[bool]]::new()

    $byQuery | ForEach-Object -Parallel {
        $group = $_
        $ctx = $using:azContext
        $bag = $using:resultsBag
        $fails = $using:failCount
        $localPillarMap = $using:pillarMap

        $rows = @()
        try {
            $skip = $null
            do {
                $p = @{ Query = $group.Name; First = 1000; DefaultProfile = $ctx }
                if ($skip) { $p['SkipToken'] = $skip }
                $queryAttempt = 0
                do {
                    try {
                        $r = Search-AzGraph @p -ErrorAction Stop
                        break
                    } catch {
                        $queryAttempt++
                        $statusCode = try { [int]$_.Exception.Response.StatusCode } catch { 0 }
                        $transient = ($statusCode -in 429, 408, 500, 502, 503, 504) -or ($_.Exception.Message -match '429|too many requests|throttl|gateway timeout|server error|service unavailable')
                        if ($queryAttempt -ge 6 -or -not $transient) { throw }
                        Start-Sleep -Seconds ([math]::Min(60, [math]::Pow(2, $queryAttempt)))
                    }
                } while ($true)
                $rows += $r.Data
                $skip = $r.SkipToken
            } while ($skip)
        } catch {
            $fails.Add($true)
            return
        }

        $nonCompliantRows = @($rows | Where-Object {
            $propNames = $_.PSObject.Properties.Name
            if ($propNames -notcontains 'compliant') { return $false }
            $v = $_.compliant
            if ($null -eq $v) { return $false }
            if ($v -is [bool]) { return -not $v }
            $s = $v.ToString().Trim().ToLower()
            return $s -in @('false', 'non-compliant', 'noncompliant', 'no')
        })
        if ($nonCompliantRows.Count -eq 0) { return }

        # The same query can back multiple checklist items (different service/guid/text),
        # so tag every non-compliant row with each item that shares this query.
        foreach ($entry in $group.Group) {
            $item = $entry.Item
            $key = if ($item.waf) { $item.waf.Trim().ToLower() } else { '' }
            $pillar = if ($localPillarMap.ContainsKey($key)) { $localPillarMap[$key] } else { 'Operational Excellence' }
            foreach ($row in $nonCompliantRows) {
                $resourceId = if ($row.PSObject.Properties.Name -contains 'id') { $row.id } else { '' }
                $bag.Add([PSCustomObject]@{
                    Service     = $entry.Service
                    WafPillar   = $pillar
                    Category    = $item.category
                    Subcategory = $item.subcategory
                    Severity    = $item.severity
                    Text        = $item.text
                    Link        = $item.link
                    Guid        = $item.guid
                    ResourceId  = $resourceId
                })
            }
        }
    } -ThrottleLimit $throttleLimit

    $queriesRun = $byQuery.Count
    $queriesFailed = $failCount.Count
    $findings = @($resultsBag)
} else {
    foreach ($group in $byQuery) {
        $queriesRun++
        $rows = @()
        try {
            $skip = $null
            do {
                $p = @{ Query = $group.Name; First = 1000 }
                if ($skip) { $p['SkipToken'] = $skip }
                $r = Invoke-SearchAzGraphWithRetry -Parameters $p
                $rows += $r.Data
                $skip = $r.SkipToken
            } while ($skip)
        } catch {
            $queriesFailed++
            continue
        }

        $nonCompliantRows = @($rows | Where-Object {
            $propNames = $_.PSObject.Properties.Name
            $propNames -contains 'compliant' -and (Test-NonCompliant $_.compliant)
        })
        if ($nonCompliantRows.Count -eq 0) { continue }

        foreach ($entry in $group.Group) {
            $item = $entry.Item
            foreach ($row in $nonCompliantRows) {
                $resourceId = if ($row.PSObject.Properties.Name -contains 'id') { $row.id } else { '' }
                $findings += [PSCustomObject]@{
                    Service     = $entry.Service
                    WafPillar   = ConvertTo-WafPillar $item.waf
                    Category    = $item.category
                    Subcategory = $item.subcategory
                    Severity    = $item.severity
                    Text        = $item.text
                    Link        = $item.link
                    Guid        = $item.guid
                    ResourceId  = $resourceId
                }
            }
        }
    }
}

Write-Host "`n  Ran $queriesRun checklist queries ($queriesFailed failed/unsupported)." -ForegroundColor Gray
Write-Host "  $($findings.Count) non-compliant findings." -ForegroundColor Gray

$findingsOut = if ($findings.Count -gt 0) { $findings } else { @([PSCustomObject]@{ Result = "No non-compliant resources found" }) }
$findingsOut | Export-Excel -Path $outputFile -WorksheetName "Findings" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Medium6

$bySeverityPillar = @($findings | Group-Object WafPillar, Severity | ForEach-Object {
    [PSCustomObject]@{
        WafPillar = $_.Values[0]
        Severity  = $_.Values[1]
        Count     = $_.Count
    }
})
if ($bySeverityPillar.Count -eq 0) { $bySeverityPillar = @([PSCustomObject]@{ Result = "No data" }) }
$bySeverityPillar | Export-Excel -Path $outputFile -WorksheetName "SummaryByPillar" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Medium6

$byService = @($findings | Group-Object Service | ForEach-Object {
    [PSCustomObject]@{ Service = $_.Name; Count = $_.Count }
} | Sort-Object Count -Descending)
if ($byService.Count -eq 0) { $byService = @([PSCustomObject]@{ Result = "No data" }) }
$byService | Export-Excel -Path $outputFile -WorksheetName "SummaryByService" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Medium6

Write-Host "`n✅ Checklist scan complete!" -ForegroundColor Green
Write-Host "📁 File: $outputFile" -ForegroundColor Cyan
Write-Host "💡 Source: https://github.com/Azure/review-checklists (vendored locally in ./checklists)`n" -ForegroundColor Yellow
