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

if ([string]::IsNullOrWhiteSpace($SourcePath)) {
    $SourcePath = (Read-Host "Enter import root containing flat(s), dark(s), and bias(es) folders").Trim('"')
}
Write-AsiToPixCyrillicPathWarning -Path $SourcePath -Context "calibration import source path"

if (-not (Test-Path -LiteralPath $SourcePath -PathType Container)) {
    Write-Host "[!] Import root not found: $SourcePath" -ForegroundColor Red
    exit 1
}

if ([string]::IsNullOrWhiteSpace($AstroPhotoRoot)) {
    Import-Module $pathsModule -Force
    $AstroPhotoRoot = Resolve-AstroPhotoRoot
}
Write-AsiToPixCyrillicPathWarning -Path $AstroPhotoRoot -Context "AstroPhoto root"

if (-not (Test-Path -LiteralPath $AstroPhotoRoot -PathType Container)) {
    Write-Host "[!] AstroPhoto root not found: $AstroPhotoRoot" -ForegroundColor Red
    exit 1
}

$calibrationRoot = Join-Path -Path $AstroPhotoRoot -ChildPath "Calibration"
if (-not (Test-Path -LiteralPath $calibrationRoot -PathType Container)) {
    Write-Host "[!] Calibration root not found: $calibrationRoot" -ForegroundColor Red
    exit 1
}

Import-AsiToPixCalibration `
    -SourcePath $SourcePath `
    -CalibrationRoot $calibrationRoot `
    -CameraName $CameraName `
    -Gain $Gain `
    -TemperatureC $TemperatureC `
    -DarkExposureSeconds $DarkExposureSeconds `
    -FilterName $FilterName `
    -AngleDegrees $AngleDegrees `
    -WhatIf:$WhatIfPreference `
    -Confirm:$false
