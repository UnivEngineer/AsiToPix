[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$MetaPath,

    [string]$AstroPhotoRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

Write-Host "--- WBPP Master Exporter ---" -ForegroundColor Cyan

$environmentModule = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.Environment.psm1"
Import-Module $environmentModule -Force

$pathsModule = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.Paths.psm1"
Import-Module $pathsModule -Force

$exportModule = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.ExportMasters.psm1"
Import-Module $exportModule -Force

if ([string]::IsNullOrWhiteSpace($MetaPath)) {
    $MetaPath = (Read-Host "Enter project_meta.json path, project folder, or object name").Trim('"')
}

if (-not [string]::IsNullOrWhiteSpace($AstroPhotoRoot)) {
    if (-not (Test-Path -LiteralPath $AstroPhotoRoot -PathType Container)) {
        throw "AstroPhoto root folder not found: '$AstroPhotoRoot'."
    }
    $AstroPhotoRoot = (Resolve-Path -LiteralPath $AstroPhotoRoot -ErrorAction Stop).ProviderPath
}

$trimmedMetadataInput = $MetaPath.Trim().Trim('"')
if ([string]::IsNullOrWhiteSpace($trimmedMetadataInput)) {
    throw "The project metadata path or object name cannot be empty."
}
$metadataInputExists = Test-Path -LiteralPath $trimmedMetadataInput
$metadataInputLooksLikePath = [System.IO.Path]::IsPathRooted($trimmedMetadataInput) -or
    $trimmedMetadataInput.Contains([System.IO.Path]::DirectorySeparatorChar) -or
    $trimmedMetadataInput.Contains([System.IO.Path]::AltDirectorySeparatorChar) -or
    [System.IO.Path]::GetExtension($trimmedMetadataInput) -ieq ".json"
$metadataPathParameters = @{
    InputValue = $trimmedMetadataInput
}
if (-not $metadataInputExists -and -not $metadataInputLooksLikePath) {
    if ([string]::IsNullOrWhiteSpace($AstroPhotoRoot)) {
        $AstroPhotoRoot = Resolve-AstroPhotoRoot
    }

    $metadataPathParameters.ProcessingRoot = Join-Path -Path $AstroPhotoRoot -ChildPath "Processing"
}

$MetaPath = Resolve-AsiToPixProjectMetadataPath @metadataPathParameters
Write-AsiToPixPathConventionWarning -Path $MetaPath -Context "project metadata path"

$metadata = Read-AsiToPixProjectMetadata -Path $MetaPath

if ([string]::IsNullOrWhiteSpace($AstroPhotoRoot) -and
    (Test-AsiToPixProjectMetadataNeedsAstroPhotoRoot -Metadata $metadata)) {
    $AstroPhotoRoot = Resolve-AstroPhotoRoot
}

if (-not [string]::IsNullOrWhiteSpace($AstroPhotoRoot)) {
    Write-AsiToPixPathConventionWarning -Path $AstroPhotoRoot -Context "AstroPhoto root"
}

$planParameters = @{
    Metadata = $metadata
}
if (-not [string]::IsNullOrWhiteSpace($AstroPhotoRoot)) {
    $planParameters.AstroPhotoRoot = $AstroPhotoRoot
}

$plan = Get-AsiToPixMasterExportPlan @planParameters
Write-AsiToPixMasterExportPlan -Plan $plan

$applyParameters = @{
    Plan    = $plan
    Confirm = $false
}
if ($WhatIfPreference) {
    $applyParameters.WhatIf = $true
}

$null = Invoke-AsiToPixMasterExportPlan @applyParameters
