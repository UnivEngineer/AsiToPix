[CmdletBinding()]
[System.Diagnostics.CodeAnalysis.SuppressMessageAttribute(
    "PSAvoidUsingWriteHost",
    "",
    Justification = "Interactive mode intentionally uses host colors; TSV mode remains pipeline-safe."
)]
param(
    [string[]]$ImportPath = @(),

    [switch]$Tsv
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$reportModule = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.ImportReport.psm1"
Import-Module $reportModule -Force

if ($ImportPath.Count -eq 0) {
    $ImportPath = @(Get-AsiToPixImportRoot)
}

if ($ImportPath.Count -eq 0) {
    throw "No Import folder found by pattern *:\AstroPhoto\Import. Pass its path with -ImportPath."
}

$resolvedImportPaths = foreach ($path in $ImportPath) {
    if (-not (Test-Path -LiteralPath $path -PathType Container)) {
        throw "Import folder not found: $path"
    }

    (Resolve-Path -LiteralPath $path -ErrorAction Stop).ProviderPath
}
$resolvedImportPaths = @($resolvedImportPaths | Sort-Object -Unique)

if ($Tsv) {
    $tsvLines = @(Get-AsiToPixImportReportLine -ImportPath $resolvedImportPaths)
    $tsvText = $tsvLines -join "`r`n"
    Write-Output $tsvText
    return
}

Write-Host "--- ASIAir IMPORT REPORT ---" -ForegroundColor Cyan
foreach ($path in $resolvedImportPaths) {
    Write-Host "[INFO] Import folder found: " -ForegroundColor Cyan -NoNewline
    Write-Host $path -ForegroundColor Green
}
Write-Host "[INFO] Scanning supported image files..." -ForegroundColor Yellow

$report = @(Get-AsiToPixImportReport -ImportPath $resolvedImportPaths -PromptForMissingData)
if ($report.Count -eq 0) {
    throw "No supported light image files found under Import folder(s): $($resolvedImportPaths -join ', ')"
}

$totalFrames = ($report | Measure-Object -Property FrameCount -Sum).Sum
Write-Host "[INFO] Found $totalFrames sub(s) in $($report.Count) object(s).`n" -ForegroundColor Green

$expectSetup = $true
foreach ($line in Get-AsiToPixImportReportPrettyLine -Report $report) {
    if ([string]::IsNullOrEmpty($line)) {
        Write-Host ""
        $expectSetup = $true
    }
    elseif ($expectSetup) {
        Write-Host $line -ForegroundColor Yellow
        $expectSetup = $false
    }
    elseif ($line -match '^Object\s') {
        Write-Host $line -ForegroundColor Cyan
    }
    else {
        Write-Host $line -ForegroundColor Gray
    }
}

Write-Host "`n--- INTEGRATION BY OBJECT ---" -ForegroundColor Cyan
$expectSetup = $true
foreach ($line in Get-AsiToPixIntegrationSummaryPrettyLine -Report $report) {
    if ([string]::IsNullOrEmpty($line)) {
        Write-Host ""
        $expectSetup = $true
    }
    elseif ($expectSetup) {
        Write-Host $line -ForegroundColor Yellow
        $expectSetup = $false
    }
    elseif ($line -match '^Object\s') {
        Write-Host $line -ForegroundColor Cyan
    }
    else {
        Write-Host $line -ForegroundColor Gray
    }
}

Write-Host "`n--- INTEGRATION SUMMARY ---" -ForegroundColor Cyan
foreach ($setupGroup in ($report | Group-Object ImportRoot, Setup)) {
    $setupRows = @($setupGroup.Group)
    $setupSeconds = [decimal](($setupRows | Measure-Object -Property IntegrationSeconds -Sum).Sum)
    $setupFrames = ($setupRows | Measure-Object -Property FrameCount -Sum).Sum
    $setupDuration = Format-AsiToPixIntegrationTime -Seconds $setupSeconds
    Write-Host "  $($setupRows[0].Setup): " -ForegroundColor Gray -NoNewline
    Write-Host $setupDuration -ForegroundColor Yellow -NoNewline
    Write-Host " ($setupFrames subs)" -ForegroundColor DarkGray
}

$totalSeconds = [decimal](($report | Measure-Object -Property IntegrationSeconds -Sum).Sum)
$totalDuration = Format-AsiToPixIntegrationTime -Seconds $totalSeconds
Write-Host "Total integration: " -ForegroundColor Green -NoNewline
Write-Host $totalDuration -ForegroundColor Yellow -NoNewline
Write-Host " ($totalFrames subs)" -ForegroundColor Gray

$clipboardAnswer = (Read-Host "Copy TSV table to clipboard? y/N").Trim()
if ($clipboardAnswer -match '^[yYдД]') {
    $tsvText = (@(Get-AsiToPixImportReportLine -Report $report) -join "`r`n")
    Set-Clipboard -Value $tsvText
    Write-Host "[INFO] TSV table copied to clipboard." -ForegroundColor Green
}
