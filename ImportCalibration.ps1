[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$SourcePath = "",

    [string]$AstroPhotoRoot = "",

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

Write-Host "--- CALIBRATION FRAME IMPORT ---" -ForegroundColor Cyan

if ([string]::IsNullOrWhiteSpace($AstroPhotoRoot)) {
    Import-Module $pathsModule -Force
    Write-Host "`n[INFO] Select the root used by this calibration import:" -ForegroundColor Cyan
    if ([string]::IsNullOrWhiteSpace($SourcePath)) {
        Write-Host "  Read from : <selected root>\Import" -ForegroundColor White
    } else {
        Write-Host "  Read from : $SourcePath" -ForegroundColor White
    }
    Write-Host "  Write to  : <selected root>\Calibration" -ForegroundColor White
    $AstroPhotoRoot = Resolve-AstroPhotoRoot
}
Write-AsiToPixCyrillicPathWarning -Path $AstroPhotoRoot -Context "AstroPhoto root"

if (-not (Test-Path -LiteralPath $AstroPhotoRoot -PathType Container)) {
    Write-Host "[!] AstroPhoto root not found: $AstroPhotoRoot" -ForegroundColor Red
    exit 1
}

$sourcePaths = @()
$useDiscoveryConfirmation = $false
if ([string]::IsNullOrWhiteSpace($SourcePath)) {
    $importRoot = Join-Path -Path $AstroPhotoRoot -ChildPath "Import"
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
