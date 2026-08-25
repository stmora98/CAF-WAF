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
    the tenant. If omitted in an interactive session, a picker is shown; you must select
    at least one subscription to continue - closing the picker or selecting nothing stops
    the run.

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
# Force UTF-8 for both this console's own output (the emoji/box-drawing characters used by
# Write-Host below) and the bytes decoded from the Python dashboard subprocess later - without
# this, the console's default (often CP1252/OEM) codepage misdecodes UTF-8 into garbled
# sequences like "ΓòÉΓòÉΓòÉ" instead of "═══" or "📊"/"✅".
[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
$OutputEncoding = [System.Text.Encoding]::UTF8
try { chcp 65001 | Out-Null } catch { }
# Ensures the Python subprocess (generate-dashboard.py) also always writes UTF-8, regardless
# of how its stdout ends up being connected when it's spawned from here.
$env:PYTHONUTF8 = "1"
$env:PYTHONIOENCODING = "utf-8"
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
# Windows.Forms). Both require at least one subscription to be selected - the run stops if
# none are chosen (closing/cancelling the picker, or an empty text response).
function Show-SubscriptionPickerGui {
    param([object[]]$AllSubscriptions)

    # Throws on Cloud Shell/non-Windows (no System.Windows.Forms) - caller catches and falls
    # back to the text picker.
    Add-Type -AssemblyName System.Windows.Forms, System.Drawing -ErrorAction Stop
    # Must run before the first WinForms control is created in the process; ignore if a
    # previous call (or host) already applied it.
    try { [System.Windows.Forms.Application]::EnableVisualStyles() } catch { }

    $checkedState = @{}
    foreach ($s in $AllSubscriptions) { $checkedState[$s.Id] = $false }
    $displayToId = @{}

    $navy = [System.Drawing.Color]::FromArgb(6, 52, 102)
    $accentBlue = [System.Drawing.Color]::FromArgb(15, 108, 189)
    $infoBg = [System.Drawing.Color]::FromArgb(234, 243, 252)
    $disabledGray = [System.Drawing.Color]::FromArgb(200, 200, 200)
    $stopRed = [System.Drawing.Color]::FromArgb(196, 43, 28)
    $borderGray = [System.Drawing.Color]::FromArgb(210, 210, 214)

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Azure WAF/CAF Workshop - Select Subscriptions"
    $form.ClientSize = New-Object System.Drawing.Size(700, 600)
    $form.MinimumSize = New-Object System.Drawing.Size(620, 520)
    $form.StartPosition = "CenterScreen"
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $form.BackColor = [System.Drawing.Color]::White
    $form.Padding = New-Object System.Windows.Forms.Padding(0)

    # Everything below is Dock/Anchor-driven (a TableLayoutPanel skeleton with AutoSize rows
    # for the header/info/filter/buttons and one Percent(100) row for the list) instead of
    # hand-placed pixel coordinates - this keeps the layout correct even if the window ends
    # up smaller than requested (small/remote screens, unusual DPI, etc.) instead of the
    # previous fixed-coordinate design, which could overlap in those cases.
    $root = New-Object System.Windows.Forms.TableLayoutPanel
    $root.Dock = [System.Windows.Forms.DockStyle]::Fill
    $root.ColumnCount = 1
    $root.RowCount = 5
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$root.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
    $form.Controls.Add($root)

    # --- Row 0: navy header banner ---
    $headerPanel = New-Object System.Windows.Forms.Panel
    $headerPanel.BackColor = $navy
    $headerPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $headerPanel.Padding = New-Object System.Windows.Forms.Padding(24, 16, 24, 16)
    $headerPanel.AutoSize = $true
    $headerPanel.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink

    $header = New-Object System.Windows.Forms.Label
    $header.Text = "Select subscriptions to scan"
    $header.Font = New-Object System.Drawing.Font("Segoe UI", 15, [System.Drawing.FontStyle]::Bold)
    $header.ForeColor = [System.Drawing.Color]::White
    $header.AutoSize = $true
    $header.Dock = [System.Windows.Forms.DockStyle]::Top
    $header.Margin = New-Object System.Windows.Forms.Padding(0)

    $subtitle = New-Object System.Windows.Forms.Label
    $subtitle.Text = "Choose which subscriptions this run should assess."
    $subtitle.ForeColor = [System.Drawing.Color]::FromArgb(200, 220, 240)
    $subtitle.AutoSize = $true
    $subtitle.Dock = [System.Windows.Forms.DockStyle]::Top
    $subtitle.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)

    # Dock=Top controls stack in REVERSE Controls-collection order (the LAST control added
    # ends up docked at the very top), so the subtitle must be added before the title.
    $headerPanel.Controls.Add($subtitle)
    $headerPanel.Controls.Add($header)
    $root.Controls.Add($headerPanel, 0, 0)

    # --- Row 1: light-blue info bar ---
    $infoPanel = New-Object System.Windows.Forms.Panel
    $infoPanel.BackColor = $infoBg
    $infoPanel.Dock = [System.Windows.Forms.DockStyle]::Fill
    $infoPanel.Padding = New-Object System.Windows.Forms.Padding(24, 10, 24, 10)
    $infoPanel.AutoSize = $true
    $infoPanel.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink

    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Text = "At least 1 subscription is required to continue. Closing this window or clicking Cancel stops the run."
    $infoLabel.ForeColor = $accentBlue
    $infoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $infoLabel.AutoSize = $true
    $infoLabel.MaximumSize = New-Object System.Drawing.Size(900, 0)
    $infoLabel.Dock = [System.Windows.Forms.DockStyle]::Top
    $infoPanel.Controls.Add($infoLabel)
    $root.Controls.Add($infoPanel, 0, 1)

    # --- Row 2: filter row (label | textbox that stretches | live count) ---
    $filterRow = New-Object System.Windows.Forms.TableLayoutPanel
    $filterRow.Dock = [System.Windows.Forms.DockStyle]::Fill
    $filterRow.ColumnCount = 3
    $filterRow.RowCount = 1
    $filterRow.AutoSize = $true
    $filterRow.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $filterRow.Margin = New-Object System.Windows.Forms.Padding(24, 12, 24, 8)
    [void]$filterRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))
    [void]$filterRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$filterRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))

    $filterLabel = New-Object System.Windows.Forms.Label
    $filterLabel.Text = "Filter:"
    $filterLabel.AutoSize = $true
    $filterLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    $filterLabel.Margin = New-Object System.Windows.Forms.Padding(0, 6, 8, 0)

    $filterBox = New-Object System.Windows.Forms.TextBox
    $filterBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $filterBox.Margin = New-Object System.Windows.Forms.Padding(0, 2, 12, 0)

    $countLabel = New-Object System.Windows.Forms.Label
    $countLabel.Text = "0 of $($AllSubscriptions.Count) selected"
    $countLabel.ForeColor = $accentBlue
    $countLabel.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $countLabel.AutoSize = $true
    $countLabel.Anchor = [System.Windows.Forms.AnchorStyles]::Right
    $countLabel.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 0)

    $filterRow.Controls.Add($filterLabel, 0, 0)
    $filterRow.Controls.Add($filterBox, 1, 0)
    $filterRow.Controls.Add($countLabel, 2, 0)
    $root.Controls.Add($filterRow, 0, 2)

    # --- Row 3: the subscription list (fills all remaining space) ---
    $listBox = New-Object System.Windows.Forms.CheckedListBox
    $listBox.Dock = [System.Windows.Forms.DockStyle]::Fill
    $listBox.Margin = New-Object System.Windows.Forms.Padding(24, 4, 24, 4)
    $listBox.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $listBox.Font = New-Object System.Drawing.Font("Segoe UI", 10)
    $listBox.CheckOnClick = $true
    $root.Controls.Add($listBox, 0, 3)

    function Update-SelectionCount {
        $n = @($checkedState.Values | Where-Object { $_ }).Count
        $countLabel.Text = "$n of $($AllSubscriptions.Count) selected"
        if (-not $btnOk) { return }
        $btnOk.Enabled = $n -gt 0
        if ($btnOk.Enabled) {
            $btnOk.BackColor = $accentBlue
            $btnOk.ForeColor = [System.Drawing.Color]::White
            $btnOk.Text = "OK"
        } else {
            $btnOk.BackColor = $disabledGray
            $btnOk.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
            $btnOk.Text = "Select at least 1"
        }
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

    # --- Row 4: buttons (Select All/None on the left, Cancel/OK on the right) ---
    # Two-column TableLayoutPanel (Percent(100) then AutoSize) instead of anchored absolute
    # positions, so the right-hand group always ends up flush with the true right edge
    # regardless of the window's actual width.
    $buttonRow = New-Object System.Windows.Forms.TableLayoutPanel
    $buttonRow.Dock = [System.Windows.Forms.DockStyle]::Fill
    $buttonRow.ColumnCount = 2
    $buttonRow.RowCount = 1
    $buttonRow.AutoSize = $true
    $buttonRow.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $buttonRow.Margin = New-Object System.Windows.Forms.Padding(24, 14, 24, 18)
    [void]$buttonRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 100)))
    [void]$buttonRow.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::AutoSize)))

    $leftFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    $leftFlow.FlowDirection = [System.Windows.Forms.FlowDirection]::LeftToRight
    $leftFlow.AutoSize = $true
    $leftFlow.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $leftFlow.Anchor = [System.Windows.Forms.AnchorStyles]::Left
    $leftFlow.Margin = New-Object System.Windows.Forms.Padding(0)

    $rightFlow = New-Object System.Windows.Forms.FlowLayoutPanel
    $rightFlow.FlowDirection = [System.Windows.Forms.FlowDirection]::RightToLeft
    $rightFlow.AutoSize = $true
    $rightFlow.AutoSizeMode = [System.Windows.Forms.AutoSizeMode]::GrowAndShrink
    $rightFlow.Anchor = [System.Windows.Forms.AnchorStyles]::Right
    $rightFlow.Margin = New-Object System.Windows.Forms.Padding(0)

    $btnSelectAll = New-Object System.Windows.Forms.Button
    $btnSelectAll.Text = "Select All"
    $btnSelectAll.Size = New-Object System.Drawing.Size(108, 36)
    $btnSelectAll.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnSelectAll.FlatAppearance.BorderColor = $borderGray
    $btnSelectAll.Margin = New-Object System.Windows.Forms.Padding(0, 0, 8, 0)
    $btnSelectAll.add_Click({
        foreach ($k in @($checkedState.Keys)) { $checkedState[$k] = $true }
        for ($i = 0; $i -lt $listBox.Items.Count; $i++) { $listBox.SetItemChecked($i, $true) }
        Update-SelectionCount
    })
    $leftFlow.Controls.Add($btnSelectAll)

    $btnSelectNone = New-Object System.Windows.Forms.Button
    $btnSelectNone.Text = "Select None"
    $btnSelectNone.Size = New-Object System.Drawing.Size(108, 36)
    $btnSelectNone.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnSelectNone.FlatAppearance.BorderColor = $borderGray
    $btnSelectNone.add_Click({
        foreach ($k in @($checkedState.Keys)) { $checkedState[$k] = $false }
        for ($i = 0; $i -lt $listBox.Items.Count; $i++) { $listBox.SetItemChecked($i, $false) }
        Update-SelectionCount
    })
    $leftFlow.Controls.Add($btnSelectNone)

    # FlowDirection=RightToLeft places the FIRST-added control at the rightmost position, so
    # add OK first (rightmost/primary) then Cancel (to its left).
    $btnOk = New-Object System.Windows.Forms.Button
    $btnOk.Text = "Select at least 1"
    $btnOk.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnOk.Size = New-Object System.Drawing.Size(140, 36)
    $btnOk.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnOk.FlatAppearance.BorderSize = 0
    $btnOk.Enabled = $false
    $btnOk.BackColor = $disabledGray
    $btnOk.ForeColor = [System.Drawing.Color]::FromArgb(90, 90, 90)
    $btnOk.add_Click({
        # OK is only enabled once at least one subscription is checked (see Update-SelectionCount),
        # so reaching here always means a non-empty selection - no scan-ALL fallback anymore.
        $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $form.Close()
    })
    $rightFlow.Controls.Add($btnOk)

    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text = "Cancel"
    $btnCancel.Size = New-Object System.Drawing.Size(110, 36)
    $btnCancel.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $btnCancel.ForeColor = $stopRed
    $btnCancel.FlatAppearance.BorderColor = $stopRed
    $btnCancel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 10, 0)
    $btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel
    $rightFlow.Controls.Add($btnCancel)

    $buttonRow.Controls.Add($leftFlow, 0, 0)
    $buttonRow.Controls.Add($rightFlow, 1, 0)
    $root.Controls.Add($buttonRow, 0, 4)

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

    $guiAvailable = $true
    try {
        $pickedIds = Show-SubscriptionPickerGui -AllSubscriptions $AllSubscriptions
    } catch {
        $guiAvailable = $false
        Write-Host "  Graphical subscription picker unavailable ($($_.Exception.Message)) - falling back to a text picker." -ForegroundColor DarkYellow
    }

    if ($guiAvailable) {
        # OK in the picker is only enabled once >=1 subscription is checked, so an empty
        # result here always means Cancel/closed - stop the run rather than defaulting to ALL.
        if ($pickedIds.Count -eq 0) {
            throw "No subscriptions were selected - stopping the run. Re-run the launcher and select at least one subscription (or pass -SubscriptionIds)."
        }
        return @($AllSubscriptions | Where-Object { $_.Id -in $pickedIds })
    }

    Write-Host "`nEnabled subscriptions in this tenant:" -ForegroundColor Cyan
    for ($i = 0; $i -lt $AllSubscriptions.Count; $i++) {
        Write-Host ("  [{0}] {1} ({2})" -f ($i + 1), $AllSubscriptions[$i].Name, $AllSubscriptions[$i].Id)
    }
    $selection = Read-Host "`nEnter subscription numbers to scan (e.g. 1,3,5) - at least one is required"
    if (-not $selection) { throw "No subscriptions were selected - stopping the run." }

    $indexes = @($selection -split '[,;]' | ForEach-Object { $_.Trim() } | Where-Object { $_ } | ForEach-Object { [int]$_ - 1 })
    $picked = @($indexes | Where-Object { $_ -ge 0 -and $_ -lt $AllSubscriptions.Count } | ForEach-Object { $AllSubscriptions[$_] })
    if ($picked.Count -eq 0) {
        throw "No valid subscription selection recognized - stopping the run."
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
Write-Host "Subscriptions analyzed ($($assessmentSubscriptions.Count)):" -ForegroundColor Green
foreach ($sub in $assessmentSubscriptions) {
    Write-Host "  - $($sub.Name) ($($sub.Id))" -ForegroundColor White
}
Write-Host "Output:   $OutputDir" -ForegroundColor Green
Write-Host "  01_Discovery/    - Resource inventory (Excel)" -ForegroundColor White
Write-Host "  02_Advisor/      - Advisor recommendations (Excel)" -ForegroundColor White
Write-Host "  03_Metrics/      - Right-sizing analysis (Excel)" -ForegroundColor White
Write-Host "  04_Governance/   - Governance data export (Excel)" -ForegroundColor White
Write-Host "  05_Security/     - Security posture and operations (Excel)" -ForegroundColor White
Write-Host "  06_Checklists/   - Azure/review-checklists ARG compliance (Excel)" -ForegroundColor White
Write-Host "  07_Dashboard/    - Consolidated dashboard (HTML)" -ForegroundColor White
