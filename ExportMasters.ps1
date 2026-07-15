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
    $MetaPath = (Read-Host "Paste path to project_meta.json").Trim('"')
}
Write-AsiToPixPathConventionWarning -Path $MetaPath -Context "project metadata path"

$metadata = Read-AsiToPixProjectMetadata -Path $MetaPath

if (-not [string]::IsNullOrWhiteSpace($AstroPhotoRoot)) {
    if (-not (Test-Path -LiteralPath $AstroPhotoRoot -PathType Container)) {
        throw "AstroPhoto root folder not found: '$AstroPhotoRoot'."
    }
    $AstroPhotoRoot = (Resolve-Path -LiteralPath $AstroPhotoRoot -ErrorAction Stop).ProviderPath
} elseif (Test-AsiToPixProjectMetadataNeedsAstroPhotoRoot -Metadata $metadata) {
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
