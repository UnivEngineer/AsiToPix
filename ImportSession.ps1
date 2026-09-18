[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$SourcePath = "",

    [string]$AstroPhotoRoot = "",

    [Alias("SourceRoot")]
    [string]$SourceAstroPhotoRoot = "",

    [Alias("DestinationRoot")]
    [string]$DestinationAstroPhotoRoot = "",

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
Import-Module $pathsModule -Force

Write-Host "--- ASIAir SESSION IMPORT ---" -ForegroundColor Cyan

$sourcePathWasProvided = -not [string]::IsNullOrWhiteSpace($SourcePath)

if (-not [string]::IsNullOrWhiteSpace($AstroPhotoRoot) -and
    -not [string]::IsNullOrWhiteSpace($DestinationAstroPhotoRoot)) {
    $legacyDestinationRoot = (Resolve-Path -LiteralPath $AstroPhotoRoot -ErrorAction Stop).ProviderPath
    $namedDestinationRoot = (Resolve-Path -LiteralPath $DestinationAstroPhotoRoot -ErrorAction Stop).ProviderPath
    if (-not $legacyDestinationRoot.Equals($namedDestinationRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Specify only one destination root: -AstroPhotoRoot (legacy) or -DestinationAstroPhotoRoot."
    }
}

$sourceRootForLookup = if ([string]::IsNullOrWhiteSpace($SourceAstroPhotoRoot)) {
    if (-not $sourcePathWasProvided) {
        Write-Host "[INFO] Select the light import source root (search <selected root>\Import):" -ForegroundColor Cyan
        Resolve-AstroPhotoRoot `
            -Purpose "light import source" `
            -SelectionPrompt "Select the light import source root" `
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
    $AstroPhotoRoot = Resolve-AstroPhotoRoot -Purpose "light archive destination"
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

if ([string]::IsNullOrWhiteSpace($sourceRootForLookup)) {
    $sourceRootForLookup = $AstroPhotoRoot
}
Write-AsiToPixCyrillicPathWarning -Path $sourceRootForLookup -Context "AstroPhoto source root"

if (-not $sourcePathWasProvided) {
    $SourcePath = (Read-Host "Enter light folder path, supported image file, or import object name").Trim('"')
}

$sourceResolution = if ($sourcePathWasProvided) {
    Resolve-AsiToPixImportSourcePath -SourcePath $SourcePath -AstroPhotoRoot $sourceRootForLookup
} else {
    Read-AsiToPixImportSource -InitialValue $SourcePath -AstroPhotoRoot $sourceRootForLookup
}
$SourcePath = $sourceResolution.SourcePath
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
