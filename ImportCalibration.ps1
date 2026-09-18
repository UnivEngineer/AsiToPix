[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$SourcePath = "",

    [string]$AstroPhotoRoot = "",

    [Alias("SourceRoot")]
    [string]$SourceAstroPhotoRoot = "",

    [Alias("DestinationRoot")]
    [string]$DestinationAstroPhotoRoot = "",

    [string]$CameraName = "",

    [string]$Gain = "",

    [string]$TemperatureC = "",

    [string]$DarkExposureSeconds = "",

    [string]$FilterName = "",

    [string]$AngleDegrees = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$pathsModule = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.Paths.psm1"
$environmentModule = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.Environment.psm1"
$importModule = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.ImportCalibration.psm1"

Import-Module $environmentModule -Force
Import-Module $importModule -Force
Import-Module $pathsModule -Force

Write-Host "--- CALIBRATION FRAME IMPORT ---" -ForegroundColor Cyan

if (-not [string]::IsNullOrWhiteSpace($AstroPhotoRoot) -and
    -not [string]::IsNullOrWhiteSpace($DestinationAstroPhotoRoot)) {
    $legacyDestinationRoot = (Resolve-Path -LiteralPath $AstroPhotoRoot -ErrorAction Stop).ProviderPath
    $namedDestinationRoot = (Resolve-Path -LiteralPath $DestinationAstroPhotoRoot -ErrorAction Stop).ProviderPath
    if (-not $legacyDestinationRoot.Equals($namedDestinationRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Specify only one destination root: -AstroPhotoRoot (legacy) or -DestinationAstroPhotoRoot."
    }
}

$sourceRootForDiscovery = if ([string]::IsNullOrWhiteSpace($SourceAstroPhotoRoot)) {
    if ([string]::IsNullOrWhiteSpace($SourcePath)) {
        Write-Host "[INFO] Select the calibration source root (read from <selected root>\Import):" -ForegroundColor Cyan
        Resolve-AstroPhotoRoot `
            -Purpose "calibration source" `
            -SelectionPrompt "Select the calibration source root" `
            -AlwaysPrompt
    } else {
        ""
    }
} else {
    if (-not (Test-Path -LiteralPath $SourceAstroPhotoRoot -PathType Container)) {
        throw "AstroPhoto source root not found: $SourceAstroPhotoRoot"
    }
    (Resolve-Path -LiteralPath $SourceAstroPhotoRoot -ErrorAction Stop).ProviderPath
}

$destinationRootInput = if ([string]::IsNullOrWhiteSpace($DestinationAstroPhotoRoot)) {
    $AstroPhotoRoot
} else {
    $DestinationAstroPhotoRoot
}
if ([string]::IsNullOrWhiteSpace($destinationRootInput)) {
    Write-Host "`n[INFO] Select the calibration destination root:" -ForegroundColor Cyan
    $AstroPhotoRoot = Resolve-AstroPhotoRoot -Purpose "calibration destination"
} else {
    if (-not (Test-Path -LiteralPath $destinationRootInput -PathType Container)) {
        throw "AstroPhoto destination root not found: $destinationRootInput"
    }
    $AstroPhotoRoot = (Resolve-Path -LiteralPath $destinationRootInput -ErrorAction Stop).ProviderPath
}
Write-AsiToPixCyrillicPathWarning -Path $AstroPhotoRoot -Context "AstroPhoto destination root"

if (-not (Test-Path -LiteralPath $AstroPhotoRoot -PathType Container)) {
    Write-Host "[!] AstroPhoto destination root not found: $AstroPhotoRoot" -ForegroundColor Red
    exit 1
}

if ([string]::IsNullOrWhiteSpace($sourceRootForDiscovery)) {
    $sourceRootForDiscovery = $AstroPhotoRoot
}

$sourcePaths = @()
$useDiscoveryConfirmation = $false
if ([string]::IsNullOrWhiteSpace($SourcePath)) {
    $importRoot = Join-Path -Path $sourceRootForDiscovery -ChildPath "Import"
    Write-Host "[INFO] Calibration source discovery (read): $importRoot" -ForegroundColor DarkGray
    $discoveredFolders = @(Find-AsiToPixCalibrationImportFolder -ImportRoot $importRoot)
    if ($discoveredFolders.Count -gt 0) {
        Write-Host "`nDetected calibration folders:" -ForegroundColor Cyan
        for ($index = 0; $index -lt $discoveredFolders.Count; $index++) {
            Write-Host " [$($index + 1)] [$($discoveredFolders[$index].Category)] $($discoveredFolders[$index].SourcePath)" `
                -ForegroundColor White
        }

        if (-not $WhatIfPreference -and
            -not (Read-AsiToPixCalibrationConfirmation `
                -Prompt "Import all $($discoveredFolders.Count) detected calibration folder(s)?")) {
            Write-Host "[INFO] Calibration import cancelled." -ForegroundColor Yellow
            return
        }

        $sourcePaths = @($discoveredFolders | Select-Object -ExpandProperty SourcePath)
        $useDiscoveryConfirmation = $true
    } else {
        $SourcePath = (Read-Host "Enter import root containing flat(s), dark(s), and bias(es) folders").Trim('"')
    }
}
Write-AsiToPixCyrillicPathWarning -Path $sourceRootForDiscovery -Context "AstroPhoto source root"
if ($sourcePaths.Count -eq 0) {
    $sourcePaths = @($SourcePath)
}

$calibrationRoot = Join-Path -Path $AstroPhotoRoot -ChildPath "Calibration"
Write-Host "[INFO] Calibration library destination (write): $calibrationRoot" -ForegroundColor DarkGray
if (-not (Test-Path -LiteralPath $calibrationRoot -PathType Container)) {
    Write-Host "[!] Calibration root not found: $calibrationRoot" -ForegroundColor Red
    exit 1
}

foreach ($currentSourcePath in $sourcePaths) {
    Write-AsiToPixCyrillicPathWarning -Path $currentSourcePath -Context "calibration import source path"
    if (-not (Test-Path -LiteralPath $currentSourcePath -PathType Container)) {
        throw "Calibration import source folder not found: $currentSourcePath"
    }

    Import-AsiToPixCalibration `
        -SourcePath $currentSourcePath `
        -CalibrationRoot $calibrationRoot `
        -CameraName $CameraName `
        -Gain $Gain `
        -TemperatureC $TemperatureC `
        -DarkExposureSeconds $DarkExposureSeconds `
        -FilterName $FilterName `
        -AngleDegrees $AngleDegrees `
        -SkipConfirmation:$useDiscoveryConfirmation `
        -WhatIf:$WhatIfPreference `
        -Confirm:$false
}
