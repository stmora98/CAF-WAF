<#
  Refreshes the vendored Azure/review-checklists JSON files under ./checklists.
  This is a maintenance script, NOT part of the workshop launcher pipeline —
  Invoke-AzureChecklists-CloudShell.ps1 reads the local ./checklists folder only
  and never calls the GitHub API at run time. Run this script manually, on a
  machine with internet access, whenever you want to pick up upstream updates.
#>

param(
    [string]$DestinationFolder = (Join-Path $PSScriptRoot "checklists")
)

$repoApi = "https://api.github.com/repos/Azure/review-checklists/contents/checklists"
$webHeaders = @{ 'User-Agent' = 'AzureWafWorkshop' }

New-Item -ItemType Directory -Path $DestinationFolder -Force | Out-Null

Write-Host "Fetching checklist catalog from Azure/review-checklists..." -ForegroundColor Cyan
$catalog = Invoke-RestMethod -Uri $repoApi -Headers $webHeaders
$files = @($catalog | Where-Object { $_.name -match '_checklist\.en\.json$' })

Write-Host "Downloading $($files.Count) checklist files to $DestinationFolder ..." -ForegroundColor Gray
foreach ($file in $files) {
    Invoke-RestMethod -Uri $file.download_url -Headers $webHeaders -OutFile (Join-Path $DestinationFolder $file.name)
    Write-Host "  ✓ $($file.name)" -ForegroundColor DarkGray
}

Write-Host "`n✅ Vendored $($files.Count) checklist files." -ForegroundColor Green
Write-Host "💡 Commit the ./checklists folder so Invoke-AzureChecklists-CloudShell.ps1 stays offline." -ForegroundColor Yellow
