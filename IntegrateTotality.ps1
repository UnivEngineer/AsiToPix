[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$Exposure,

    [ValidateSet("Registered", "Debayered")]
    [string]$InputStage,

    [string]$ProcessingRoot = "C:\AstroPhoto\Processing\Totality",

    [string]$ManifestPath,

    [string]$PixInsightPath,

    [ValidateSet("Reuse", "Dedicated")]
    [string]$PixInsightMode,

    [string]$PixInsightScriptPath
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$modulePath = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.TotalityIntegration.psm1"
Import-Module $modulePath -Force

function Read-TotalityExposure {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [System.IO.DirectoryInfo[]]$Directories
    )

    Write-Host "Available exposure folders:" -ForegroundColor Cyan
    for ($index = 0; $index -lt $Directories.Count; $index++) {
        Write-Host ("  [{0}] {1}" -f ($index + 1), $Directories[$index].Name)
    }

    while ($true) {
        $answer = ([string](Read-Host "Exposure folder")).Trim()
        $number = 0
        if ([int]::TryParse($answer, [ref]$number) -and $number -ge 1 -and $number -le $Directories.Count) {
            return $Directories[$number - 1]
        }

        $match = @($Directories | Where-Object { $_.Name -ieq $answer })
        if ($match.Count -eq 1) {
            return $match[0]
        }
        Write-Host "Enter a displayed number or exact folder name." -ForegroundColor Yellow
    }
}

if (-not (Test-Path -LiteralPath $ProcessingRoot -PathType Container)) {
    throw "Totality processing root not found: '$ProcessingRoot'."
}
$ProcessingRoot = (Resolve-Path -LiteralPath $ProcessingRoot -ErrorAction Stop).ProviderPath

$exposureDirectories = @(Get-AsiToPixTotalityExposureDirectory -ProcessingRoot $ProcessingRoot)
if ($exposureDirectories.Count -eq 0) {
    throw "No Total-* exposure folders found in '$ProcessingRoot'."
}

if ([string]::IsNullOrWhiteSpace($Exposure)) {
    $exposureDirectory = Read-TotalityExposure -Directories $exposureDirectories
}
else {
    $matchingDirectories = @($exposureDirectories | Where-Object { $_.Name -ieq $Exposure })
    if ($matchingDirectories.Count -ne 1) {
        $availableNames = @($exposureDirectories.Name) -join ", "
        throw "Exposure folder '$Exposure' was not found in '$ProcessingRoot'. Available folders: $availableNames."
    }
    $exposureDirectory = $matchingDirectories[0]
}

$framesPath = Join-Path -Path $exposureDirectory.FullName -ChildPath "Frames"
$exposureLabel = $exposureDirectory.Name.Substring("Total-".Length)
$rawPath = Join-Path -Path $framesPath -ChildPath "raw"
if ([string]::IsNullOrWhiteSpace($InputStage)) {
    $InputStage = Read-AsiToPixTotalityInputStage
}
$inputPath = Join-Path -Path $framesPath -ChildPath $InputStage.ToLowerInvariant()
$integratedPath = Join-Path -Path $framesPath -ChildPath "integrated"
if (-not (Test-Path -LiteralPath $framesPath -PathType Container)) {
    throw "Frames folder not found for exposure '$($exposureDirectory.Name)': '$framesPath'."
}
if (-not (Test-Path -LiteralPath $inputPath -PathType Container)) {
    throw "$InputStage frame folder not found: '$inputPath'."
}
if ([string]::IsNullOrWhiteSpace($ManifestPath)) {
    $ManifestPath = Join-Path -Path $framesPath -ChildPath "manifest.json"
}
elseif ([System.IO.Path]::IsPathRooted($ManifestPath)) {
    $ManifestPath = [System.IO.Path]::GetFullPath($ManifestPath)
}
else {
    $ManifestPath = [System.IO.Path]::GetFullPath((Join-Path -Path $PWD.Path -ChildPath $ManifestPath))
}

$rawFrames = @(Get-AsiToPixTotalityRawFrame -RawPath $rawPath)
$inputFrames = @(Get-AsiToPixTotalityFrame -FramePath $inputPath -AllowEmpty)
$manifestResult = Resolve-AsiToPixTotalityFrameManifest `
    -ManifestPath $ManifestPath `
    -MinimumIndex $rawFrames[0].Index `
    -MaximumIndex $rawFrames[-1].Index `
    -WhatIf:$WhatIfPreference `
    -Confirm:$false
$frameManifest = $manifestResult.Manifest
$plan = @(Get-AsiToPixTotalityFrameManifestPlan `
    -RawFrames $rawFrames `
    -InputFrames $inputFrames `
    -Manifest $frameManifest `
    -OutputDirectory $integratedPath `
    -ExposureLabel $exposureLabel `
    -InputStage $InputStage `
    -MinimumFrameCount 3)

if ([string]::IsNullOrWhiteSpace($PixInsightScriptPath)) {
    $PixInsightScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "pixinsight\IntegrateTotality.js"
}
if (-not (Test-Path -LiteralPath $PixInsightScriptPath -PathType Leaf)) {
    throw "PixInsight integration script not found: '$PixInsightScriptPath'."
}
$PixInsightScriptPath = (Resolve-Path -LiteralPath $PixInsightScriptPath -ErrorAction Stop).ProviderPath
$PixInsightPath = Resolve-AsiToPixPixInsightPath -Path $PixInsightPath
$runningPixInsightProcesses = @(Get-Process -Name "PixInsight" -ErrorAction SilentlyContinue)
if ([string]::IsNullOrWhiteSpace($PixInsightMode)) {
    $PixInsightMode = Read-AsiToPixPixInsightMode -RunningInstanceCount $runningPixInsightProcesses.Count
}
elseif ($PixInsightMode -eq "Reuse" -and $runningPixInsightProcesses.Count -eq 0) {
    throw "Cannot reuse PixInsight because no running PixInsight instance was found. Start PixInsight or use -PixInsightMode Dedicated."
}

Write-Host "--- Totality integration plan ---" -ForegroundColor Cyan
Write-Host "Exposure folder: $($exposureDirectory.FullName)"
Write-Host "Exposure label : $exposureLabel"
Write-Host "Frames         : $framesPath"
Write-Host "Manifest       : $ManifestPath"
Write-Host "Raw            : $rawPath"
Write-Host "Input stage    : $InputStage"
Write-Host "Input folder   : $inputPath"
Write-Host "Integrated     : $integratedPath"
Write-Host "PixInsight mode: $PixInsightMode"
Write-Host "Block length   : $($frameManifest.BlockLength)"
Write-Host "Block start    : $($frameManifest.BlockStartIndex)"
Write-Host "C2 / C3        : $($frameManifest.C2Index) / $($frameManifest.C3Index)"
Write-Host "Skip indexes   : $(Format-AsiToPixTotalityIndexList -Indexes @($frameManifest.SkipIndices))"
Write-Host "Raw frames     : $($rawFrames.Count)"
Write-Host "Input frames   : $($inputFrames.Count)"

$plan | Select-Object `
    @{ Name = "Block"; Expression = { "{0:D3}" -f $_.BlockNumber } }, `
    Phase, `
    @{ Name = "Nominal"; Expression = { "$($_.NominalFirstIndex)-$($_.NominalLastIndex)" } }, `
    @{ Name = "Integrate"; Expression = { Format-AsiToPixTotalityIndexList -Indexes $_.IntegrateIndexes } }, `
    @{ Name = "Skip"; Expression = { Format-AsiToPixTotalityIndexList -Indexes $_.SkippedIndexes } }, `
    @{ Name = "Ready"; Expression = { "$($_.AvailableFrameCount)/$($_.FrameCount)" } }, `
    @{ Name = "Action"; Expression = {
        if (-not $_.CanIntegrate) { "Skip: $($_.SkipReason)" }
        elseif (Test-Path -LiteralPath $_.OutputPath) { "Skip: output exists" }
        else { "Integrate" }
    } } | Format-Table -AutoSize

Write-Host "--- Planned outputs ---" -ForegroundColor Cyan
foreach ($block in $plan) {
    Write-Host ("[{0:D3} {1}] integrate indexes: {2}" -f `
        $block.BlockNumber,
        $block.Phase,
        (Format-AsiToPixTotalityIndexList -Indexes $block.IntegrateIndexes))
    if ($block.SkippedIndexes.Count -gt 0) {
        Write-Host ("    excluded by manifest: {0}" -f `
            (Format-AsiToPixTotalityIndexList -Indexes $block.SkippedIndexes))
    }
    Write-Host "    output: $($block.OutputPath)"
    if (-not $block.CanIntegrate) {
        Write-Host "    not ready: $($block.SkipReason)" -ForegroundColor Yellow
    }
}

$actionableBlocks = @($plan | Where-Object {
    $_.CanIntegrate -and -not (Test-Path -LiteralPath $_.OutputPath)
})
if (-not $WhatIfPreference -and $actionableBlocks.Count -gt 0) {
    if (-not (Read-AsiToPixTotalityPlanConfirmation -Prompt "Execute plan? [Y/n]")) {
        Write-Host "Integration plan was not executed." -ForegroundColor Yellow
        return
    }
}

$invokeParameters = @{
    Plan                     = $plan
    PixInsightPath           = $PixInsightPath
    PixInsightScriptPath     = $PixInsightScriptPath
    PixInsightMode           = $PixInsightMode
}
if ($WhatIfPreference) {
    $invokeParameters.WhatIf = $true
}
if ($PSBoundParameters.ContainsKey("Confirm")) {
    $invokeParameters.Confirm = [bool]$PSBoundParameters["Confirm"]
}

$result = Invoke-AsiToPixTotalityIntegrationPlan @invokeParameters

Write-Host "--- Integration summary ---" -ForegroundColor Cyan
Write-Host "Integrated blocks : $($result.IntegratedCount)" -ForegroundColor Green
Write-Host "Existing outputs  : $($result.ExistingCount)"
Write-Host "Rejected blocks   : $($result.RejectedCount)"
Write-Host "Not approved/WhatIf: $($result.NotApprovedCount)"
