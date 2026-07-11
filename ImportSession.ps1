[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$SourcePath = "",

    [string]$AstroPhotoRoot = "",

    [string]$ObjectName = "",

    [string]$SeasonName = "",

    [string]$TelescopeName = "",

    [string]$CameraName = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$pathsModule = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.Paths.psm1"
$importModule = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.ImportSession.psm1"

Import-Module $importModule -Force

Write-Host "--- ASIAir SESSION IMPORT ---" -ForegroundColor Cyan

if ([string]::IsNullOrWhiteSpace($SourcePath)) {
    $SourcePath = (Read-Host "Enter ASIAir Light folder path, copied import folder, or a FITS file inside it").Trim('"')
}

if ($SourcePath -match '\.fits?(\.gz)?$') {
    $SourcePath = Split-Path -Path $SourcePath -Parent
}

if (-not (Test-Path -LiteralPath $SourcePath -PathType Container)) {
    Write-Host "[!] Source folder not found: $SourcePath" -ForegroundColor Red
    exit 1
}

if ([string]::IsNullOrWhiteSpace($AstroPhotoRoot)) {
    Import-Module $pathsModule -Force
    $AstroPhotoRoot = Resolve-AstroPhotoRoot
}

if (-not (Test-Path -LiteralPath $AstroPhotoRoot -PathType Container)) {
    Write-Host "[!] AstroPhoto root not found: $AstroPhotoRoot" -ForegroundColor Red
    exit 1
}

Import-AsiToPixSession `
    -SourcePath $SourcePath `
    -AstroPhotoRoot $AstroPhotoRoot `
    -ObjectName $ObjectName `
    -SeasonName $SeasonName `
    -TelescopeName $TelescopeName `
    -CameraName $CameraName `
    -WhatIf:$WhatIfPreference
