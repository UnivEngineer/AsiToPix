[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$AstroPhotoRoot,

    [string]$TotalityPath,

    [string]$OutputPath,

    [string]$SettingsPath,

    [ValidateSet(0, 90, 180, 270)]
    [int]$Rotation = 0,

    [ValidateSet("None", "Horizontal", "Vertical")]
    [string]$Flip = "None",

    [ValidateSet("Ask", "Saved", "Capture")]
    [string]$ChannelMatchMode = "Ask",

    [string]$PixInsightPath,

    [string]$PixInsightScriptPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$pathsModulePath = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.Paths.psm1"
$alignmentModulePath = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.TotalityAlignment.psm1"
$integrationModulePath = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.TotalityIntegration.psm1"
Import-Module $pathsModulePath -Force
Import-Module $alignmentModulePath -Force
Import-Module $integrationModulePath -Force

if ([string]::IsNullOrWhiteSpace($TotalityPath)) {
    if ([string]::IsNullOrWhiteSpace($AstroPhotoRoot)) {
        $AstroPhotoRoot = Resolve-AstroPhotoRoot `
            -Purpose "Totality channel alignment" `
            -SelectionPrompt "Select *:\AstroPhoto or *:\Astro root for Totality alignment"
    }
    else {
        if (-not (Test-Path -LiteralPath $AstroPhotoRoot -PathType Container)) {
            throw "Astrophotography root not found: '$AstroPhotoRoot'."
        }
        $AstroPhotoRoot = (Resolve-Path -LiteralPath $AstroPhotoRoot -ErrorAction Stop).ProviderPath
    }
    $TotalityPath = Join-Path -Path $AstroPhotoRoot -ChildPath "SharpCap\Totality-2026"
}
if (-not (Test-Path -LiteralPath $TotalityPath -PathType Container)) {
    throw "Totality source directory not found: '$TotalityPath'."
}
$TotalityPath = (Resolve-Path -LiteralPath $TotalityPath -ErrorAction Stop).ProviderPath

if ([string]::IsNullOrWhiteSpace($OutputPath)) {
    $OutputPath = Join-Path -Path $TotalityPath -ChildPath "HDR\Aligned"
}
elseif (-not [System.IO.Path]::IsPathRooted($OutputPath)) {
    $OutputPath = Join-Path -Path $PWD.Path -ChildPath $OutputPath
}
$OutputPath = [System.IO.Path]::GetFullPath($OutputPath)
if ((Test-Path -LiteralPath $OutputPath) -and
    -not (Test-Path -LiteralPath $OutputPath -PathType Container)) {
    throw "Totality alignment output path exists but is not a directory: '$OutputPath'."
}

if ([string]::IsNullOrWhiteSpace($SettingsPath)) {
    $SettingsPath = Join-Path -Path $TotalityPath -ChildPath "HDR\channel_match_offsets.json"
}
elseif (-not [System.IO.Path]::IsPathRooted($SettingsPath)) {
    $SettingsPath = Join-Path -Path $PWD.Path -ChildPath $SettingsPath
}
$SettingsPath = [System.IO.Path]::GetFullPath($SettingsPath)
if ((Test-Path -LiteralPath $SettingsPath) -and
    -not (Test-Path -LiteralPath $SettingsPath -PathType Leaf)) {
    throw "Totality ChannelMatch settings path exists but is not a file: '$SettingsPath'."
}

if (-not $PSBoundParameters.ContainsKey("Rotation")) {
    $Rotation = Read-AsiToPixTotalityAlignmentRotation
}
if (-not $PSBoundParameters.ContainsKey("Flip")) {
    $Flip = Read-AsiToPixTotalityAlignmentFlip
}

if ([string]::IsNullOrWhiteSpace($PixInsightScriptPath)) {
    $PixInsightScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "pixinsight\AlignTotality.js"
}
if (-not (Test-Path -LiteralPath $PixInsightScriptPath -PathType Leaf)) {
    throw "PixInsight Totality alignment script not found: '$PixInsightScriptPath'."
}
$PixInsightScriptPath = (Resolve-Path -LiteralPath $PixInsightScriptPath -ErrorAction Stop).ProviderPath
$PixInsightPath = Resolve-AsiToPixPixInsightPath -Path $PixInsightPath

$frames = @(Get-AsiToPixTotalityAlignmentFrame -TotalityPath $TotalityPath)
$plan = @(Get-AsiToPixTotalityAlignmentPlan `
    -Frames $frames `
    -OutputDirectory $OutputPath `
    -Rotation $Rotation `
    -Flip $Flip)
$referenceFrame = Get-AsiToPixTotalityAlignmentReferenceFrame -Frames $frames
$exposureLabels = @($frames |
    Sort-Object -Property ExposureMilliseconds, ExposureLabel |
    Select-Object -ExpandProperty ExposureLabel -Unique)

$savedSettings = $null
$settingsExist = Test-Path -LiteralPath $SettingsPath -PathType Leaf
if ($ChannelMatchMode -eq "Ask") {
    if ($settingsExist -and
        (Read-AsiToPixTotalityPlanConfirmation `
            -Prompt "Use saved ChannelMatch offsets from '$SettingsPath'? [Y/n]")) {
        $ChannelMatchMode = "Saved"
    }
    else {
        $ChannelMatchMode = "Capture"
    }
}
if ($ChannelMatchMode -eq "Saved") {
    if (-not $settingsExist) {
        throw "Cannot use saved ChannelMatch offsets because the settings file was not found: '$SettingsPath'."
    }
    $savedSettings = Read-AsiToPixTotalityChannelMatchSetting -Path $SettingsPath
}

$rotationText = if ($Rotation -eq 0) { "none" } else { "$Rotation degrees clockwise" }
$flipText = switch ($Flip) {
    "Horizontal" { "horizontal mirror" }
    "Vertical" { "vertical mirror" }
    default { "none" }
}
Write-Host "--- Totality channel alignment plan ---" -ForegroundColor Cyan
Write-Host "Source root       : $TotalityPath"
Write-Host "Output            : $OutputPath"
Write-Host "Settings JSON     : $SettingsPath"
Write-Host "ChannelMatch      : $ChannelMatchMode"
Write-Host "Exposures         : $($exposureLabels -join ', ')"
Write-Host "Integrated frames : $($frames.Count)"
Write-Host "Rotation          : $rotationText"
Write-Host "Mirror            : $flipText"
Write-Host "Suggested reference: $($referenceFrame.Path)"

$plan | Select-Object `
    @{ Name = "Block"; Expression = { "{0:D3}" -f $_.BlockNumber } }, `
    ExposureLabel, `
    Phase, `
    @{ Name = "Action"; Expression = {
        if (Test-Path -LiteralPath $_.OutputPath) { "Skip: output exists" }
        else { "ChannelMatch + FastRotation" }
    } }, `
    OutputPath | Format-Table -AutoSize -Wrap

$invokeParameters = @{
    Plan                 = $plan
    PixInsightPath       = $PixInsightPath
    PixInsightScriptPath = $PixInsightScriptPath
    ChannelMatchMode     = $ChannelMatchMode
    SettingsPath         = $SettingsPath
}
if ($null -ne $savedSettings) {
    $invokeParameters.Channels = $savedSettings.Channels
}
if ($WhatIfPreference) {
    $invokeParameters.WhatIf = $true
    if ($PSBoundParameters.ContainsKey("Confirm")) {
        $invokeParameters.Confirm = [bool]$PSBoundParameters["Confirm"]
    }
    $result = Invoke-AsiToPixTotalityAlignmentPlan @invokeParameters
    Write-Host "--- Totality alignment summary ---" -ForegroundColor Cyan
    Write-Host "Aligned frames     : $($result.AlignedCount)"
    Write-Host "Existing outputs   : $($result.ExistingCount)"
    Write-Host "Not approved/WhatIf: $($result.NotApprovedCount)"
    return
}

$actionableFrames = @($plan | Where-Object { -not (Test-Path -LiteralPath $_.OutputPath) })
if ($actionableFrames.Count -eq 0) {
    Write-Host "All aligned outputs already exist; nothing to do." -ForegroundColor Yellow
    return
}
if (-not (Read-AsiToPixTotalityPlanConfirmation -Prompt "Execute Totality alignment plan? [Y/n]")) {
    Write-Host "Totality alignment plan was not executed." -ForegroundColor Yellow
    return
}

$runningPixInsightProcesses = @(Get-Process -Name "PixInsight" -ErrorAction SilentlyContinue)
if ($runningPixInsightProcesses.Count -gt 1) {
    throw "Totality alignment requires exactly one PixInsight instance. Close extra instances and run the script again."
}
if ($runningPixInsightProcesses.Count -eq 0) {
    Write-Host "Starting PixInsight for ChannelMatch preparation..." -ForegroundColor Cyan
    $startedPixInsight = Start-Process `
        -FilePath $PixInsightPath `
        -PassThru `
        -ErrorAction Stop
    $startupDeadline = [DateTime]::UtcNow.AddMinutes(2)
    while (@(Get-Process -Name "PixInsight" -ErrorAction SilentlyContinue).Count -eq 0) {
        $startedPixInsight.Refresh()
        if ($startedPixInsight.HasExited) {
            throw "PixInsight closed before Totality ChannelMatch preparation could begin."
        }
        if ([DateTime]::UtcNow -ge $startupDeadline) {
            throw "Timed out after two minutes while starting PixInsight: '$PixInsightPath'."
        }
        Start-Sleep -Milliseconds 500
    }
}

if ($ChannelMatchMode -eq "Capture") {
    $baselineIconIds = @(Get-AsiToPixChannelMatchIconSnapshot `
        -PixInsightPath $PixInsightPath `
        -PixInsightScriptPath $PixInsightScriptPath)

    Write-Host ""
    Write-Host "Prepare ChannelMatch in PixInsight:" -ForegroundColor Cyan
    Write-Host "  1. Open this suggested central frame with a middle exposure:"
    Write-Host "     $($referenceFrame.Path)" -ForegroundColor White
    Write-Host "  2. Open ChannelMatch and adjust the R/G/B X/Y offsets."
    Write-Host "  3. Drag the blue New Instance triangle from ChannelMatch to an EMPTY place"
    Write-Host "     in the PixInsight workspace. This creates a new process icon that PJSR can read."
    Write-Host "  4. Return to this console."
    $null = Read-Host "Press Enter after the new ChannelMatch process icon has been created"
    $invokeParameters.BaselineIconIds = $baselineIconIds
}

$invokeParameters.Confirm = $false
$result = Invoke-AsiToPixTotalityAlignmentPlan @invokeParameters

Write-Host "--- Totality alignment summary ---" -ForegroundColor Cyan
Write-Host "Aligned frames   : $($result.AlignedCount)" -ForegroundColor Green
Write-Host "Existing outputs : $($result.ExistingCount)"
Write-Host "Not approved     : $($result.NotApprovedCount)"
if ($null -ne $result.ChannelMatch) {
    Write-Host "ChannelMatch offsets:" -ForegroundColor Cyan
    foreach ($channel in @($result.ChannelMatch.channels)) {
        Write-Host ("  {0}: enabled={1}, dx={2}, dy={3}" -f
            $channel.name, $channel.enabled, $channel.dx, $channel.dy)
    }
}
if ($ChannelMatchMode -eq "Capture") {
    Write-Host "Saved settings   : $SettingsPath"
}
