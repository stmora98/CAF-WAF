<#
.SYNOPSIS
    Azure WAF/CAF Workshop Launcher - Runs all discovery scripts and consolidates output.

.DESCRIPTION
    Master launcher that sequentially executes:
      1. Invoke-AzureDiscovery-CloudShell.ps1       (Resource inventory)
      2. Invoke-AzureAdvisor-CloudShell.ps1         (Advisor recommendations)
      3. Invoke-AzureMetrics-CloudShell.ps1         (Right-sizing & reliability)
      4. Invoke-AzureGovernanceViz-CloudShell.ps1   (Governance data export)
    5. Invoke-AzureSecurity-CloudShell.ps1        (Defender posture + incidents)
    6. Invoke-AzureChecklists-CloudShell.ps1      (Azure/review-checklists ARG compliance)
    7. generate-dashboard.py                      (Consolidated dashboard + action items)
    
    All outputs are consolidated under AzureWorkshop at the repository root.

.PARAMETER OutputDir
    Base output directory. Defaults to "AzureWorkshop" at the repository root. If that
    folder already exists from a previous run, it is wiped and recreated (see -KeepPrevious).

.PARAMETER SkipMetrics
    Skip the metrics script (it is slower, queries per-resource). Use if short on time.

.PARAMETER SkipFinOps
    Skip the FinOps extended export (actual cost, reservation utilization, extra
    optimization recommendations). Writes into 02_Advisor alongside AzureAdvisor.

.PARAMETER KeepPrevious
    Archive an existing output folder (rename with a timestamp suffix) instead of
    deleting it before the run.

.PARAMETER SubscriptionIds
    Comma-separated subscription GUIDs to scan, instead of every enabled subscription in
    the tenant. If omitted in an interactive session, a numbered picker is shown (press
    Enter to scan everything, the previous/default behavior).

.EXAMPLE
    ./powershell/Launch-AzureWorkshop.ps1
    ./powershell/Launch-AzureWorkshop.ps1 -SkipMetrics
    ./powershell/Launch-AzureWorkshop.ps1 -KeepPrevious

.NOTES
    Run in Azure Cloud Shell (PowerShell) or locally with Az modules installed.
#>

[CmdletBinding()]
param(
    [string]$OutputDir,
    [switch]$SkipMetrics,
    [switch]$SkipFinOps,
    # By default each run wipes the previous output folder. Pass -KeepPrevious to
    # archive it (renamed with a timestamp suffix) instead of deleting it.
    [switch]$KeepPrevious,
    # Skip the interactive Microsoft Graph sign-in for Defender XDR incidents/alerts
    # during the Security phase (keeps Defender for Cloud CSPM + Endpoint data).
    [switch]$SkipGraphSecurity,
    # Every run always shows the account picker (WAM SSO, so cached accounts are one click).
    # Pass this to additionally disable WAM for the run and force a full interactive sign-in
    # page - use it only if the normal picker itself seems stuck on a cached token.
    [switch]$ForceAccountSelection,
    # Comma-separated subscription GUIDs to scope the run to. Omit to get an interactive
    # picker (or, when non-interactive, every enabled subscription - unchanged default).
    [string]$SubscriptionIds
)

$ErrorActionPreference = 'Continue'
$scriptDir = $PSScriptRoot
if (-not $scriptDir) { $scriptDir = (Get-Location).Path }
$repoRoot = if (Test-Path (Join-Path $scriptDir "generate-dashboard.py")) {
    $scriptDir
} else {
    Split-Path $scriptDir -Parent
}

if (-not $OutputDir) {
    $OutputDir = Join-Path $repoRoot "AzureWorkshop"
}
if ($OutputDir -match '(?i)[\\/]OneDrive(?: - [^\\/]+)?[\\/]') {
    Write-Warning "OneDrive may automatically encrypt generated workbooks while they are being written. If dashboard inputs come back unreadable, pause OneDrive sync for this folder during the run or re-run with -OutputDir pointing outside OneDrive."
}

if (Test-Path $OutputDir) {
    if ($KeepPrevious) {
        $archived = "$OutputDir" + "_" + (Get-Date -Format "yyyyMMdd_HHmmss")
        Write-Host "Archiving previous run to $archived" -ForegroundColor Yellow
        Rename-Item -Path $OutputDir -NewName (Split-Path $archived -Leaf)
    } else {
        Write-Host "Removing previous run at $OutputDir ..." -ForegroundColor Yellow
        Remove-Item -Path $OutputDir -Recurse -Force
    }
}

New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
New-Item -ItemType Directory -Path "$OutputDir/01_Discovery" -Force | Out-Null
New-Item -ItemType Directory -Path "$OutputDir/02_Advisor" -Force | Out-Null
New-Item -ItemType Directory -Path "$OutputDir/03_Metrics" -Force | Out-Null
New-Item -ItemType Directory -Path "$OutputDir/04_Governance" -Force | Out-Null
New-Item -ItemType Directory -Path "$OutputDir/05_Security" -Force | Out-Null
New-Item -ItemType Directory -Path "$OutputDir/06_Checklists" -Force | Out-Null
New-Item -ItemType Directory -Path "$OutputDir/07_Dashboard" -Force | Out-Null

# --- Ensure modules ---
Write-Host "Checking prerequisites..." -ForegroundColor Yellow

$requiredModules = @(
    'Az.Accounts',
    'Az.Resources',
    'Az.ResourceGraph',
    'Az.Monitor',
    'ImportExcel'
)

foreach ($mod in $requiredModules) {
    if (-not (Get-Module -ListAvailable -Name $mod)) {
        Write-Host "  Installing $mod..." -ForegroundColor DarkYellow
        Install-Module -Name $mod -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck
    }
}

Import-Module Az.Accounts -ErrorAction Stop
Import-Module Az.Resources -ErrorAction Stop
Import-Module Az.ResourceGraph -ErrorAction Stop
Import-Module Az.Monitor -ErrorAction SilentlyContinue
Import-Module ImportExcel -ErrorAction Stop

# Ensure Python + openpyxl for the dashboard agent
# `Get-Command` alone isn't enough: on Windows, `python`/`python3` can resolve to the
# Microsoft Store app-execution-alias stub, which "exists" in PATH but exits 9009 without
# running anything. Validate each candidate actually executes before trusting it.
function Resolve-PythonCommand {
    foreach ($candidate in @(
        @{ Cmd = "py"; Args = @("-3") },
        @{ Cmd = "python3"; Args = @() },
        @{ Cmd = "python"; Args = @() }
    )) {
        if (-not (Get-Command $candidate.Cmd -ErrorAction SilentlyContinue)) { continue }
        try {
            $out = & $candidate.Cmd @($candidate.Args) --version 2>&1
            if ($LASTEXITCODE -eq 0 -and ($out -join ' ') -notmatch 'was not found') {
                return $candidate
            }
        } catch { }
    }
    return $null
}

$python = Resolve-PythonCommand
$pythonCmd = if ($python) { $python.Cmd } else { $null }
[string[]]$pythonArgs = if ($python) { $python.Args } else { @() }

if ($pythonCmd) {
    Write-Host "  Checking Python openpyxl..." -ForegroundColor DarkGray
    & $pythonCmd @pythonArgs -c "import openpyxl" 2>$null
    if ($LASTEXITCODE -ne 0) {
        Write-Host "  Installing openpyxl..." -ForegroundColor DarkYellow
        & $pythonCmd @pythonArgs -m pip install openpyxl --quiet --user 2>$null
    }
} else {
    Write-Host "  Python not found - dashboard generation will be skipped" -ForegroundColor DarkYellow
}

Write-Host "  All prerequisites satisfied.`n" -ForegroundColor Green

# Prefers the WAM broker (Windows Hello/Conditional Access support); falls back to device-code login if the broker fails.
# -ForceAccountSelection disables WAM (process scope only, doesn't touch the saved user preference) so
# WAM's silent single-cached-account SSO is bypassed and the full account-picker sign-in page is shown.
function Connect-AzAccountWithWamFallback {
    param([switch]$ForceAccountSelection)

    if ($ForceAccountSelection) {
        # Hard reset: also disable WAM so even the picker's cached tokens are bypassed and a
        # full interactive sign-in page is shown. Use this if the account picker itself seems stuck.
        Disconnect-AzAccount -ErrorAction SilentlyContinue | Out-Null
        Write-Host "  Forcing a full interactive sign-in (WAM disabled for this run)..." -ForegroundColor DarkGray
        if (Get-Command Update-AzConfig -ErrorAction SilentlyContinue) {
            try { Update-AzConfig -EnableLoginByWam $false -Scope Process -ErrorAction Stop | Out-Null } catch { }
        }
        Connect-AzAccount -SkipContextPopulation -ErrorAction Stop
        return
    }

    if (Get-Command Update-AzConfig -ErrorAction SilentlyContinue) {
        try {
            Update-AzConfig -EnableLoginByWam $true -ErrorAction Stop | Out-Null
        } catch {
            Write-Host "  Could not enable WAM broker login ($($_.Exception.Message)) - continuing with default login method." -ForegroundColor DarkYellow
        }
    }
    try {
        Connect-AzAccount -SkipContextPopulation -ErrorAction Stop
    } catch {
        Write-Host "  WAM sign-in failed ($($_.Exception.Message)). Retrying with device code authentication..." -ForegroundColor DarkYellow
        if (Get-Command Update-AzConfig -ErrorAction SilentlyContinue) {
            try { Update-AzConfig -EnableLoginByWam $false -ErrorAction Stop | Out-Null } catch { }
        }
        Connect-AzAccount -UseDeviceAuthentication -SkipContextPopulation
    }
}

# Always force a brand-new interactive sign-in - never silently reuse a cached Az context or
# a WAM SSO session, so every run genuinely prompts for authentication. -ForceAccountSelection
# is kept as a no-op parameter for backward compatibility (it's effectively always on now).
Write-Host "Signing in..." -ForegroundColor Yellow
Connect-AzAccountWithWamFallback -ForceAccountSelection | Out-Null
$context = Get-AzContext
if (-not $context) { throw "Azure authentication did not create a usable tenant context." }

$allEnabledSubscriptions = @(Get-AzSubscription -TenantId $context.Tenant.Id -ErrorAction Stop |
    Where-Object { $_.State -eq 'Enabled' })
if ($allEnabledSubscriptions.Count -eq 0) {
    throw "No enabled subscriptions are accessible in tenant $($context.Tenant.Id)."
}

# Lets the operator scope the run to a subset of subscriptions instead of the whole tenant.
# -SubscriptionIds skips the prompt entirely (scripted/unattended runs); otherwise a graphical
# WinForms picker (checkboxes, live filter, selection counter) is tried first, falling back to
# a text numbered picker if WinForms isn't available (e.g. Cloud Shell has no display/System.
# Windows.Forms), both defaulting to ALL when nothing is picked.
function Show-SubscriptionPickerGui {
    param([object[]]$AllSubscriptions)

    # Throws on Cloud Shell/non-Windows (no System.Windows.Forms) - caller catches and falls
    # back to the text picker.
    Add-Type -AssemblyName System.Windows.Forms, System.Drawing -ErrorAction Stop

    $checkedState = @{}
    foreach ($s in $AllSubscriptions) { $checkedState[$s.Id] = $false }
    $displayToId = @{}

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Azure WAF/CAF Workshop - Select Subscriptions"
    # ClientSize (not Size) fixes the usable interior area precisely - Size also counts the
    # title bar/borders, which previously left the bottom buttons cramped/clipped depending
    # on Windows theme and DPI scaling.
    $form.ClientSize = New-Object System.Drawing.Size(560, 520)
    $form.MinimumSize = New-Object System.Drawing.Size(480, 460)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

    $header = New-Object System.Windows.Forms.Label
    $header.Text = "Select subscriptions to scan"
    $header.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
    $header.AutoSize = $true
    $header.Location = New-Object System.Drawing.Point(16, 14)
    $form.Controls.Add($header)

    $subtitle = New-Object System.Windows.Forms.Label
    $subtitle.Text = "Choose which subscriptions this run should assess. Leave everything unchecked and click OK to scan ALL."
    $subtitle.ForeColor = [System.Drawing.Color]::FromArgb(92, 92, 92)
    $subtitle.Location = New-Object System.Drawing.Point(16, 42)
    $subtitle.Size = New-Object System.Drawing.Size(520, 32)
    $form.Controls.Add($subtitle)

    $filterLabel = New-Object System.Windows.Forms.Label
    $filterLabel.Text = "Filter:"
    $filterLabel.AutoSize = $true
    $filterLabel.Location = New-Object System.Drawing.Point(16, 82)
    $form.Controls.Add($filterLabel)

    $filterBox = New-Object System.Windows.Forms.TextBox
    $filterBox.Location = New-Object System.Drawing.Point(58, 79)
    $filterBox.Width = 220
    $form.Controls.Add($filterBox)

    $countLabel = New-Object System.Windows.Forms.Label
    $countLabel.Text = "0 of $($AllSubscriptions.Count) selected"
    $countLabel.ForeColor = [System.Drawing.Color]::FromArgb(15, 108, 189)
    $countLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $countLabel.AutoSize = $true
    $countLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
    $countLabel.Location = New-Object System.Drawing.Point(370, 82)
    $form.Controls.Add($countLabel)

    $listBox = New-Object System.Windows.Forms.CheckedListBox
    $listBox.Location = New-Object System.Drawing.Point(16, 112)
    $listBox.Size = New-Object System.Drawing.Size(520, 336)
    $listBox.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    $listBox.CheckOnClick = $true
    $form.Controls.Add($listBox)

    function Update-SelectionCount {
        $n = @($checkedState.Values | Where-Object { $_ }).Count
        $countLabel.Text = "$n of $($AllSubscriptions.Count) selected"
    }

    function Update-SubscriptionList {
        param([string]$Query)
        for ($i = 0; $i -lt $listBox.Items.Count; $i++) {
            $id = $displayToId[$listBox.Items[$i].ToString()]
            if ($id) { $checkedState[$id] = $listBox.GetItemChecked($i) }
        }
        $listBox.Items.Clear()
        $displayToId.Clear()
        foreach ($s in $AllSubscriptions) {
            if ($Query -and $s.Name -notlike "*$Query*" -and $s.Id -notlike "*$Query*") { continue }
            $text = "$($s.Name)   ($($s.Id))"
            $displayToId[$text] = $s.Id
            $idx = $listBox.Items.Add($text)
            $listBox.SetItemChecked($idx, [bool]$checkedState[$s.Id])
        }
        Update-SelectionCount
    }

    $listBox.add_ItemCheck({
        param($evtSender, $e)
        $id = $displayToId[$listBox.Items[$e.Index].ToString()]
        if ($id) { $checkedState[$id] = ($e.NewValue -eq [System.Windows.Forms.CheckState]::Checked) }
        $form.BeginInvoke([Action] { Update-SelectionCount }) | Out-Null
    })
    $filterBox.add_TextChanged({ Update-SubscriptionList -Query $filterBox.Text })
    Update-SubscriptionList -Query ""

    $btnSelectAll = New-Object System.Windows.Forms.Button
    $btnSelectAll.Text = "Select All"
    $btnSelectAll.Location = New-Object System.Drawing.Point(16, 468)
    $btnSelectAll.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
    $btnSelectAll.add_Click({
        foreach ($k in @($checkedState.Keys)) { $checkedState[$k] = $true }
        for ($i = 0; $i -lt $listBox.Items.Count; $i++) { $listBox.SetItemChecked($i, $true) }
        Update-SelectionCount
    })
    $form.Controls.Add($btnSelectAll)

    $btnSelectNone = New-Object System.Windows.Forms.Button
    $btnSelectNone.Text = "Select None"
    $btnSelectNone.Location = New-Object System.Drawing.Point(104, 468)
    $btnSelectNone.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left
    $btnSelectNone.add_Click({
        foreach ($k in @($checkedState.Keys)) { $checkedState[$k] = $false }
        for ($i = 0; $i -lt $listBox.Items.Count; $i++) { $listBox.SetItemChecked($i, $false) }
        Update-SelectionCount
    })
    $form.Controls.Add($btnSelectNone)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel (scan ALL)"
    $btnCancel.Location = New-Object System.Drawing.Point(276, 468)
    $btnCancel.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $form.Controls.Add($btnCancel)

    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "OK"
    $btnOk.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnOk.Location = New-Object System.Drawing.Point(456, 468)
    $btnOk.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Right
    $btnOk.add_Click({
        # Clicking OK with nothing checked would silently scan ALL - always confirm that
        # explicitly instead of assuming it, so an accidental click never scans more than intended.
        $anyChecked = @($checkedState.Values | Where-Object { $_ }).Count -gt 0
        if (-not $anyChecked) {
            $confirm = [System.Windows.Forms.MessageBox]::Show(
                $form,
                "No subscriptions are checked. Scan ALL enabled subscriptions instead?",
                "Confirm scan scope",
                [System.Windows.Forms.MessageBoxButtons]::YesNo,
                [System.Windows.Forms.MessageBoxIcon]::Question)
            if ($confirm -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        }
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })
    $form.Controls.Add($btnOk)

    $form.AcceptButton = $btnOk
    $form.CancelButton = $btnCancel

    $dialogResult = $form.ShowDialog()
    $form.Dispose()

    if ($dialogResult -ne [System.Windows.Forms.DialogResult]::OK) { return , @() }
    return , @($checkedState.GetEnumerator() | Where-Object { $_.Value } | ForEach-Object { $_.Key })
}

function Select-WorkshopSubscriptions {
    param([object[]]$AllSubscriptions, [string]$RequestedIds)

    if ($RequestedIds) {
        $requested = @($RequestedIds -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ })
        $matched = @($AllSubscriptions | Where-Object { $_.Id -in $requested })
        $missing = @($requested | Where-Object { $_ -notin $matched.Id })
        if ($missing.Count -gt 0) {
            Write-Warning "These -SubscriptionIds were not found among your enabled subscriptions and will be skipped: $($missing -join ', ')"
        }
        if ($matched.Count -eq 0) { throw "None of the requested -SubscriptionIds matched an enabled subscription." }
        return $matched
    }

    if (-not [Environment]::UserInteractive) { return $AllSubscriptions }

    try {
        $pickedIds = Show-SubscriptionPickerGui -AllSubscriptions $AllSubscriptions
        if ($pickedIds.Count -eq 0) {
            Write-Host "No subscriptions selected in the picker - scanning ALL enabled subscriptions." -ForegroundColor DarkYellow
            return $AllSubscriptions
        }
        return @($AllSubscriptions | Where-Object { $_.Id -in $pickedIds })
    } catch {
        Write-Host "  Graphical subscription picker unavailable ($($_.Exception.Message)) - falling back to a text picker." -ForegroundColor DarkYellow
    }

    Write-Host "`nEnabled subscriptions in this tenant:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $AllSubscriptions.Count; $i++) {
        Write-Host ("  [{0}] {1} ({2})" -f ($i + 1), $AllSubscriptions[$i].Name, $AllSubscriptions[$i].Id)
    }
    $selection = Read-Host "`nEnter subscription numbers to scan (e.g. 1,3,5), or press Enter to scan ALL"
    if (-not $selection) { return $AllSubscriptions }

    $indexes = @($selection -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | ForEach-Object { [int]$_ - 1 })
    $picked = @($indexes | Where-Object { $_ -ge 0 -and $_ -lt $AllSubscriptions.Count } | ForEach-Object { $AllSubscriptions[$_] })
    if ($picked.Count -eq 0) {
        Write-Warning "No valid selection recognized - scanning ALL enabled subscriptions instead."
        return $AllSubscriptions
    }
    return $picked
}

$assessmentSubscriptions = @(Select-WorkshopSubscriptions -AllSubscriptions $allEnabledSubscriptions -RequestedIds $SubscriptionIds)
$env:AZWORKSHOP_SUBSCRIPTION_IDS = ($assessmentSubscriptions.Id -join ',')

Write-Host "Azure WAF/CAF Workshop - Discovery Launcher" -ForegroundColor Cyan
Write-Host "Tenant:  $($context.Tenant.Id)" -ForegroundColor Cyan
Write-Host "Account: $($context.Account.Id)" -ForegroundColor Cyan
Write-Host "Scope:   $($assessmentSubscriptions.Count) enabled subscription(s)" -ForegroundColor Cyan
Write-Host "Output:  $OutputDir`n" -ForegroundColor Cyan

$startTime = Get-Date

# Re-runs a phase script and validates its newest .xlsx isn't OneDrive-encrypted (OLE header),
# retrying a few times since the encryption/decryption race is timing-dependent.
function Invoke-PhaseWithValidation {
    param(
        [string]$Name,
        [string]$ScriptPath,
        [string]$OutputFolder,
        [int]$MaxAttempts = 3,
        [switch]$RequireWorkbook,
        [hashtable]$ScriptArgs = @{}
    )
    for ($attempt = 1; $attempt -le $MaxAttempts; $attempt++) {
        try {
            & $ScriptPath @ScriptArgs *>&1 | ForEach-Object { Write-Host $_ }
        } catch {
            Write-Host "  $Name error: $($_.Exception.Message)" -ForegroundColor Yellow
        }

        $xlsx = Get-ChildItem -Path $OutputFolder -Filter *.xlsx -ErrorAction SilentlyContinue |
            Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if (-not $xlsx) {
            if (-not $RequireWorkbook) { return }
            if ($attempt -lt $MaxAttempts) {
                Write-Host "  ! $Name did not produce an Excel workbook (attempt $attempt/$MaxAttempts). Retrying..." -ForegroundColor DarkYellow
                continue
            }
            throw "$Name failed to produce a required Excel workbook in $OutputFolder."
        }

        $header = [byte[]]::new(4)
        $stream = [System.IO.File]::OpenRead($xlsx.FullName)
        try { [void]$stream.Read($header, 0, 4) } finally { $stream.Close() }
        $isValidZip = ($header[0] -eq 0x50 -and $header[1] -eq 0x4B)

        if ($isValidZip) { return }

        if ($attempt -lt $MaxAttempts) {
            Write-Host "  ! $Name workbook looks OneDrive-encrypted (attempt $attempt/$MaxAttempts). Retrying..." -ForegroundColor DarkYellow
            Start-Sleep -Seconds 5
        } else {
            Write-Host "  ! $Name workbook still invalid after $MaxAttempts attempts (OneDrive sync). Dashboard will treat this phase as unavailable." -ForegroundColor Yellow
        }
    }
}

# === PHASE 1: Resource Discovery ===
Write-Host " PHASE 1/7: Resource Discovery" -ForegroundColor Cyan
$env:AZWORKSHOP_OUTPUT = "$OutputDir/01_Discovery"
Invoke-PhaseWithValidation -Name "Discovery" -ScriptPath "$scriptDir/Invoke-AzureDiscovery-CloudShell.ps1" -OutputFolder "$OutputDir/01_Discovery"

# === PHASE 2: Advisor Recommendations ===
Write-Host "`n PHASE 2/7: Advisor Recommendations" -ForegroundColor Cyan
$env:AZWORKSHOP_OUTPUT = "$OutputDir/02_Advisor"
Invoke-PhaseWithValidation -Name "Advisor" -ScriptPath "$scriptDir/Invoke-AzureAdvisor-CloudShell.ps1" -OutputFolder "$OutputDir/02_Advisor"

# === PHASE 2b: FinOps Extended (actual cost, reservation utilization, extra recs) ===
if (-not $SkipFinOps) {
    Write-Host "`n PHASE 2b: FinOps Extended Export" -ForegroundColor Cyan
    $env:AZWORKSHOP_OUTPUT = "$OutputDir/02_Advisor"
    Invoke-PhaseWithValidation -Name "FinOps" -ScriptPath "$scriptDir/Invoke-AzureFinOps-CloudShell.ps1" -OutputFolder "$OutputDir/02_Advisor"
} else {
    Write-Host "`n PHASE 2b: FinOps Extended Export - SKIPPED" -ForegroundColor DarkYellow
}

# === PHASE 3: Metrics (Right-Sizing) ===
if (-not $SkipMetrics) {
    Write-Host "`n PHASE 3/7: Metrics & Right-Sizing (this takes longer)" -ForegroundColor Cyan
    $env:AZWORKSHOP_OUTPUT = "$OutputDir/03_Metrics"
    Invoke-PhaseWithValidation -Name "Metrics" -ScriptPath "$scriptDir/Invoke-AzureMetrics-CloudShell.ps1" -OutputFolder "$OutputDir/03_Metrics"
} else {
    Write-Host "`n PHASE 3/7: Metrics - SKIPPED" -ForegroundColor DarkYellow
}

# === PHASE 4: Governance Visualization ===
Write-Host "`n PHASE 4/7: Governance Visualization" -ForegroundColor Cyan
$env:AZWORKSHOP_OUTPUT = "$OutputDir/04_Governance"
Invoke-PhaseWithValidation -Name "Governance" -ScriptPath "$scriptDir/Invoke-AzureGovernanceViz-CloudShell.ps1" -OutputFolder "$OutputDir/04_Governance"

# === PHASE 5: Security Assessment ===
Write-Host "`n PHASE 5/7: Security Assessment" -ForegroundColor Cyan
$env:AZWORKSHOP_OUTPUT = "$OutputDir/05_Security"
$securityArgs = if ($SkipGraphSecurity) { @{ SkipGraphSecurity = $true } } else { @{} }
Invoke-PhaseWithValidation -Name "Security" -ScriptPath "$scriptDir/Invoke-AzureSecurity-CloudShell.ps1" -OutputFolder "$OutputDir/05_Security" -RequireWorkbook -ScriptArgs $securityArgs

# === PHASE 6: Review Checklists (Azure/review-checklists ARG compliance) ===
Write-Host "`n PHASE 6/7: Review Checklists (community WAF checks)" -ForegroundColor Cyan
$env:AZWORKSHOP_OUTPUT = "$OutputDir/06_Checklists"
Invoke-PhaseWithValidation -Name "Checklists" -ScriptPath "$scriptDir/Invoke-AzureChecklists-CloudShell.ps1" -OutputFolder "$OutputDir/06_Checklists"

# === PHASE 7: Consolidated Dashboard ===
Write-Host "`n PHASE 7/7: Generating Consolidated Dashboard" -ForegroundColor Cyan
if ($pythonCmd) {
    try {
        & $pythonCmd @pythonArgs (Join-Path $repoRoot "generate-dashboard.py") $OutputDir *>&1 | ForEach-Object { Write-Host $_ }
        if ($LASTEXITCODE -ne 0) {
            throw "Dashboard generator exited with code $LASTEXITCODE"
        }
        Write-Host "  Dashboard generated in $OutputDir/07_Dashboard/" -ForegroundColor Green
    } catch {
        Write-Host "  Dashboard generation failed: $($_.Exception.Message)" -ForegroundColor Yellow
    }
} else {
    Write-Host "  Python not available. Run manually: python3 generate-dashboard.py `"$OutputDir`"" -ForegroundColor Yellow
}

# Clean up env var
Remove-Item Env:\AZWORKSHOP_OUTPUT -ErrorAction SilentlyContinue
Remove-Item Env:\AZWORKSHOP_SUBSCRIPTION_IDS -ErrorAction SilentlyContinue

# === SUMMARY ===
$elapsed = (Get-Date) - $startTime
Write-Host "`nALL PHASES COMPLETE" -ForegroundColor Green
Write-Host "Duration: $([math]::Round($elapsed.TotalMinutes, 1)) minutes" -ForegroundColor Green
Write-Host "Output:   $OutputDir" -ForegroundColor Green
Write-Host "  01_Discovery/    - Resource inventory (Excel)" -ForegroundColor White
Write-Host "  02_Advisor/      - Advisor recommendations (Excel)" -ForegroundColor White
Write-Host "  03_Metrics/      - Right-sizing analysis (Excel)" -ForegroundColor White
Write-Host "  04_Governance/   - Governance data export (Excel)" -ForegroundColor White
Write-Host "  05_Security/     - Security posture and operations (Excel)" -ForegroundColor White
Write-Host "  06_Checklists/   - Azure/review-checklists ARG compliance (Excel)" -ForegroundColor White
Write-Host "  07_Dashboard/    - Consolidated dashboard (HTML)" -ForegroundColor White
