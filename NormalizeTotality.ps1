[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$AstroPhotoRoot,

    [string]$InputPath,

    [string]$OutputPath,

    [string]$MetricsPath,

    [string]$PixInsightPath,

    [string]$PixInsightScriptPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$pathsModulePath = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.Paths.psm1"
$normalizationModulePath = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.TotalityNormalization.psm1"
$integrationModulePath = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.TotalityIntegration.psm1"
Import-Module $pathsModulePath -Force
Import-Module $normalizationModulePath -Force
Import-Module $integrationModulePath -Force

if ([string]::IsNullOrWhiteSpace($InputPath) -or [string]::IsNullOrWhiteSpace($OutputPath)) {
    if ([string]::IsNullOrWhiteSpace($AstroPhotoRoot)) {
        $AstroPhotoRoot = Resolve-AstroPhotoRoot `
            -Purpose "Totality brightness and color normalization" `
            -SelectionPrompt "Select *:\AstroPhoto or *:\Astro root for Totality normalization"
    }
    else {
        if (-not (Test-Path -LiteralPath $AstroPhotoRoot -PathType Container)) {
            throw "Astrophotography root not found: '$AstroPhotoRoot'."
        }
        $AstroPhotoRoot = (Resolve-Path -LiteralPath $AstroPhotoRoot -ErrorAction Stop).ProviderPath
    }

    $relativeHdrPath = "SharpCap\Totality-2026\HDR"
    if ([string]::IsNullOrWhiteSpace($InputPath)) {
        $InputPath = Join-Path -Path $AstroPhotoRoot -ChildPath (
            Join-Path -Path $relativeHdrPath -ChildPath "Composed"
        )
    }
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $OutputPath = Join-Path -Path $AstroPhotoRoot -ChildPath (
            Join-Path -Path $relativeHdrPath -ChildPath "Calibrated"
        )
    }
}

if (-not (Test-Path -LiteralPath $InputPath -PathType Container)) {
    throw "Totality normalization input directory not found: '$InputPath'."
}
$InputPath = (Resolve-Path -LiteralPath $InputPath -ErrorAction Stop).ProviderPath

if ([System.IO.Path]::IsPathRooted($OutputPath)) {
    $OutputPath = [System.IO.Path]::GetFullPath($OutputPath)
}
else {
    $OutputPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PWD.Path -ChildPath $OutputPath))
}
if (Test-Path -LiteralPath $OutputPath) {
    if (-not (Test-Path -LiteralPath $OutputPath -PathType Container)) {
        throw "Totality normalization output path is not a directory: '$OutputPath'."
    }
    $OutputPath = (Resolve-Path -LiteralPath $OutputPath -ErrorAction Stop).ProviderPath
}
else {
    $outputParent = Split-Path -Path $OutputPath -Parent
    if (-not (Test-Path -LiteralPath $outputParent -PathType Container)) {
        throw "Totality normalization output parent directory not found: '$outputParent'."
    }
}

if ([string]::IsNullOrWhiteSpace($MetricsPath)) {
    $MetricsPath = Join-Path -Path $OutputPath -ChildPath "normalization_metrics.csv"
}
elseif ([System.IO.Path]::IsPathRooted($MetricsPath)) {
    $MetricsPath = [System.IO.Path]::GetFullPath($MetricsPath)
}
else {
    $MetricsPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PWD.Path -ChildPath $MetricsPath))
}

if ([string]::IsNullOrWhiteSpace($PixInsightScriptPath)) {
    $PixInsightScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "pixinsight\NormalizeTotality.js"
}
if (-not (Test-Path -LiteralPath $PixInsightScriptPath -PathType Leaf)) {
    throw "PixInsight Totality normalization script not found: '$PixInsightScriptPath'."
}
$PixInsightScriptPath = (Resolve-Path -LiteralPath $PixInsightScriptPath -ErrorAction Stop).ProviderPath
$PixInsightPath = Resolve-AsiToPixPixInsightPath -Path $PixInsightPath

$frames = @(Get-AsiToPixTotalityNormalizationFrame -InputPath $InputPath)
$plan = @(Get-AsiToPixTotalityNormalizationPlan `
    -Frames $frames `
    -OutputDirectory $OutputPath)

Write-Host "--- Totality normalization plan ---" -ForegroundColor Cyan
Write-Host "Input       : $InputPath"
Write-Host "Output      : $OutputPath"
Write-Host "Metrics CSV : $MetricsPath"
Write-Host "Frames      : $($frames.Count)"
Write-Host "PixInsight  : reuse one running instance; start one if none is running"

$plan | Select-Object `
    @{ Name = "Block"; Expression = { "{0:D3}" -f $_.BlockNumber } }, `
    InputPath, `
    NormalizedPath, `
    ColorCorrectedPath | Format-Table -AutoSize -Wrap

$invokeParameters = @{
    Plan                 = $plan
    MetricsPath          = $MetricsPath
    PixInsightPath       = $PixInsightPath
    PixInsightScriptPath = $PixInsightScriptPath
}

if ($WhatIfPreference) {
    $invokeParameters.WhatIf = $true
    if ($PSBoundParameters.ContainsKey("Confirm")) {
        $invokeParameters.Confirm = [bool]$PSBoundParameters["Confirm"]
    }
    $result = Invoke-AsiToPixTotalityNormalizationPlan @invokeParameters
    Write-Host "--- Totality normalization summary ---" -ForegroundColor Cyan
    Write-Host "Normalized frames : $($result.NormalizedCount)"
    Write-Host "Not approved/WhatIf: $($result.NotApprovedCount)"
    return
}

if (-not (Read-AsiToPixTotalityPlanConfirmation -Prompt "Execute Totality normalization plan? [Y/n]")) {
    Write-Host "Totality normalization plan was not executed." -ForegroundColor Yellow
    return
}

$runningPixInsightProcesses = @(Get-Process -Name "PixInsight" -ErrorAction SilentlyContinue)
if ($runningPixInsightProcesses.Count -gt 1) {
    throw "Totality normalization requires exactly one PixInsight instance. Close extra instances and run the script again."
}
if ($runningPixInsightProcesses.Count -eq 0) {
    Write-Host "Starting PixInsight for interactive reference preview selection..." -ForegroundColor Cyan
    $startedPixInsight = Start-Process `
        -FilePath $PixInsightPath `
        -PassThru `
        -ErrorAction Stop
    $startupDeadline = [DateTime]::UtcNow.AddMinutes(2)
    while (@(Get-Process -Name "PixInsight" -ErrorAction SilentlyContinue).Count -eq 0) {
        $startedPixInsight.Refresh()
        if ($startedPixInsight.HasExited) {
            throw "PixInsight closed before reference frame selection could begin."
        }
        if ([DateTime]::UtcNow -ge $startupDeadline) {
            throw "Timed out after two minutes while starting PixInsight: '$PixInsightPath'."
        }
        Start-Sleep -Milliseconds 500
    }
}

Write-Host ""
Write-Host "Prepare the reference ROI in PixInsight:" -ForegroundColor Cyan
Write-Host "  1. Open one Block-*-HDR.xisf frame from: $InputPath"
Write-Host "  2. Create a preview over the region used for brightness and color calibration."
Write-Host "  3. Activate that preview. If it is the only preview, the main view may remain active."
$null = Read-Host "Press Enter here when the reference frame and preview are ready"

$runningPixInsightProcesses = @(Get-Process -Name "PixInsight" -ErrorAction SilentlyContinue)
if ($runningPixInsightProcesses.Count -ne 1) {
    throw "Expected exactly one running PixInsight instance after reference preparation; found $($runningPixInsightProcesses.Count)."
}

$invokeParameters.Confirm = $false
$result = Invoke-AsiToPixTotalityNormalizationPlan @invokeParameters

Write-Host "--- Totality normalization summary ---" -ForegroundColor Cyan
Write-Host "Normalized frames : $($result.NormalizedCount)" -ForegroundColor Green
Write-Host "Metrics CSV       : $MetricsPath"
Write-Host "Not approved      : $($result.NotApprovedCount)"
