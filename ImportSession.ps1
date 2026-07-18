[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$SourcePath = "",

    [string]$AstroPhotoRoot = "",

    [string]$ObjectName = "",

    [string]$SeasonName = "",

    [string]$TelescopeName = "",

    [string]$CameraName = "",

    [ValidateSet("", "Copy", "Symlink")]
    [string]$ImportMode = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$pathsModule = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.Paths.psm1"
$environmentModule = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.Environment.psm1"
$importModule = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.ImportSession.psm1"

Import-Module $environmentModule -Force
Import-Module $importModule -Force

Write-Host "--- ASIAir SESSION IMPORT ---" -ForegroundColor Cyan

if ([string]::IsNullOrWhiteSpace($SourcePath)) {
    $SourcePath = (Read-Host "Enter light folder path, FITS/RAW file, or import object name").Trim('"')
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

$sourceResolution = Resolve-AsiToPixImportSourcePath -SourcePath $SourcePath -AstroPhotoRoot $AstroPhotoRoot
$SourcePath = $sourceResolution.SourcePath
if (-not [string]::IsNullOrWhiteSpace($sourceResolution.AstroPhotoRoot)) {
    $AstroPhotoRoot = $sourceResolution.AstroPhotoRoot
}
Write-AsiToPixCyrillicPathWarning -Path $SourcePath -Context "source path"

if (-not (Test-Path -LiteralPath $SourcePath -PathType Container)) {
    Write-Host "[!] Source folder not found: $SourcePath" -ForegroundColor Red
    exit 1
}

Import-AsiToPixSession `
    -SourcePath $SourcePath `
    -AstroPhotoRoot $AstroPhotoRoot `
    -ObjectName $ObjectName `
    -SeasonName $SeasonName `
    -TelescopeName $TelescopeName `
    -CameraName $CameraName `
    -ImportMode $ImportMode `
    -WhatIf:$WhatIfPreference
