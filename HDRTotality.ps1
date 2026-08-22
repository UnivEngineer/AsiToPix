[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$AstroPhotoRoot,

    [string]$InputPath,

    [string]$OutputPath,

    [string[]]$Exposure,

    [string]$PixInsightPath,

    [ValidateSet("Reuse", "Dedicated")]
    [string]$PixInsightMode,

    [string]$PixInsightScriptPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$pathsModulePath = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.Paths.psm1"
$hdrModulePath = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.TotalityHDR.psm1"
$integrationModulePath = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.TotalityIntegration.psm1"
Import-Module $pathsModulePath -Force
Import-Module $hdrModulePath -Force
Import-Module $integrationModulePath -Force

if ([string]::IsNullOrWhiteSpace($InputPath) -or [string]::IsNullOrWhiteSpace($OutputPath)) {
    if ([string]::IsNullOrWhiteSpace($AstroPhotoRoot)) {
        $AstroPhotoRoot = Resolve-AstroPhotoRoot `
            -Purpose "Totality HDR input and output" `
            -SelectionPrompt "Select *:\AstroPhoto or *:\Astro root for Totality HDR"
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
            Join-Path -Path $relativeHdrPath -ChildPath "Aligned"
        )
    }
    if ([string]::IsNullOrWhiteSpace($OutputPath)) {
        $OutputPath = Join-Path -Path $AstroPhotoRoot -ChildPath (
            Join-Path -Path $relativeHdrPath -ChildPath "Composed"
        )
    }
}

if (-not (Test-Path -LiteralPath $InputPath -PathType Container)) {
    throw "Totality HDR input directory not found: '$InputPath'."
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
        throw "Totality HDR output path is not a directory: '$OutputPath'."
    }
    $OutputPath = (Resolve-Path -LiteralPath $OutputPath -ErrorAction Stop).ProviderPath
}
else {
    $outputParent = Split-Path -Path $OutputPath -Parent
    if (-not (Test-Path -LiteralPath $outputParent -PathType Container)) {
        throw "Totality HDR output parent directory not found: '$outputParent'."
    }
}

if ([string]::IsNullOrWhiteSpace($PixInsightScriptPath)) {
    $PixInsightScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "pixinsight\TotalityHDR.js"
}
if (-not (Test-Path -LiteralPath $PixInsightScriptPath -PathType Leaf)) {
    throw "PixInsight HDR composition script not found: '$PixInsightScriptPath'."
}
$PixInsightScriptPath = (Resolve-Path -LiteralPath $PixInsightScriptPath -ErrorAction Stop).ProviderPath
$PixInsightPath = Resolve-AsiToPixPixInsightPath -Path $PixInsightPath

$frames = @(Get-AsiToPixTotalityHdrFrame -InputPath $InputPath)
$ladderParameters = @{
    Frames = $frames
}
if ($PSBoundParameters.ContainsKey("Exposure")) {
    $ladderParameters.Exposure = $Exposure
}
$ladder = Select-AsiToPixTotalityHdrLadder @ladderParameters
$plan = @(Get-AsiToPixTotalityHdrPlan `
    -Frames $frames `
    -Ladder $ladder `
    -OutputDirectory $OutputPath)

$runningPixInsightProcesses = @(Get-Process -Name "PixInsight" -ErrorAction SilentlyContinue)
if ([string]::IsNullOrWhiteSpace($PixInsightMode)) {
    $PixInsightMode = Read-AsiToPixPixInsightMode `
        -RunningInstanceCount $runningPixInsightProcesses.Count
}
elseif ($PixInsightMode -eq "Reuse" -and $runningPixInsightProcesses.Count -eq 0) {
    throw "Cannot reuse PixInsight because no running PixInsight instance was found. Start PixInsight or use -PixInsightMode Dedicated."
}

Write-Host "--- Totality HDR composition plan ---" -ForegroundColor Cyan
Write-Host "Input          : $InputPath"
Write-Host "Output         : $OutputPath"
Write-Host "Exposure ladder: $($ladder.Label)"
Write-Host "PixInsight mode: $PixInsightMode"
Write-Host "Discovered     : $($frames.Count) frame(s)"

$plan | Select-Object `
    @{ Name = "Block"; Expression = { "{0:D3}" -f $_.BlockNumber } }, `
    @{ Name = "HDR inputs (long to short)"; Expression = { @($_.ExposureLabels) -join " + " } }, `
    @{ Name = "Ready"; Expression = { if ($_.CanCompose) { "Yes" } else { "No" } } }, `
    @{ Name = "Action"; Expression = {
        if (-not $_.CanCompose) { "Skip: $($_.SkipReason)" }
        elseif (Test-Path -LiteralPath $_.OutputPath) { "Skip: output exists" }
        else { "Compose HDR" }
    } }, `
    OutputPath | Format-Table -AutoSize -Wrap

$actionableBlocks = @($plan | Where-Object {
    $_.CanCompose -and -not (Test-Path -LiteralPath $_.OutputPath)
})
if (-not $WhatIfPreference -and $actionableBlocks.Count -gt 0) {
    if (-not (Read-AsiToPixTotalityPlanConfirmation -Prompt "Execute HDR plan? [Y/n]")) {
        Write-Host "HDR composition plan was not executed." -ForegroundColor Yellow
        return
    }
}

$invokeParameters = @{
    Plan                 = $plan
    PixInsightPath       = $PixInsightPath
    PixInsightScriptPath = $PixInsightScriptPath
    PixInsightMode       = $PixInsightMode
    Operation            = "HDR"
}
if ($WhatIfPreference) {
    $invokeParameters.WhatIf = $true
}
if ($PSBoundParameters.ContainsKey("Confirm")) {
    $invokeParameters.Confirm = [bool]$PSBoundParameters["Confirm"]
}

$result = Invoke-AsiToPixTotalityIntegrationPlan @invokeParameters

Write-Host "--- Totality HDR summary ---" -ForegroundColor Cyan
Write-Host "Composed blocks    : $($result.IntegratedCount)" -ForegroundColor Green
Write-Host "Existing outputs  : $($result.ExistingCount)"
Write-Host "Incomplete blocks : $($result.RejectedCount)"
Write-Host "Not approved/WhatIf: $($result.NotApprovedCount)"
