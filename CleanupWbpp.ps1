[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$AstroPhotoRoot,

    [string]$ProcessingRoot
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$pathsModule = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.Paths.psm1"
Import-Module $pathsModule -Force

$cleanupModule = Join-Path -Path $PSScriptRoot -ChildPath "src\AsiToPix.CleanupWbpp.psm1"
Import-Module $cleanupModule -Force

if ([string]::IsNullOrWhiteSpace($ProcessingRoot)) {
    if ([string]::IsNullOrWhiteSpace($AstroPhotoRoot)) {
        $AstroPhotoRoot = Resolve-AstroPhotoRoot
    }
    elseif (-not (Test-Path -LiteralPath $AstroPhotoRoot -PathType Container)) {
        throw "AstroPhoto root folder not found: '$AstroPhotoRoot'."
    }

    $ProcessingRoot = Join-Path -Path $AstroPhotoRoot -ChildPath "Processing"
}

if (-not (Test-Path -LiteralPath $ProcessingRoot -PathType Container)) {
    throw "Processing folder not found: '$ProcessingRoot'."
}
$ProcessingRoot = (Resolve-Path -LiteralPath $ProcessingRoot -ErrorAction Stop).ProviderPath

Write-Host "--- WBPP temporary file cleanup ---" -ForegroundColor Cyan
Write-Host "Scanning: $ProcessingRoot"

$plan = @(Get-AsiToPixWbppCleanupPlan -ProcessingRoot $ProcessingRoot)
if ($plan.Count -eq 0) {
    Write-Host "No processing projects with removable WBPP files were found." -ForegroundColor Green
    return
}

Write-Host "Found $($plan.Count) project(s) that can be cleaned." -ForegroundColor Yellow
$invokeParameters = @{ Plan = $plan }
if ($WhatIfPreference) {
    $invokeParameters.WhatIf = $true
}
$result = Invoke-AsiToPixWbppCleanupPlan @invokeParameters

Write-Host "`n--- Cleanup summary ---" -ForegroundColor Cyan
Write-Host "Removed files       : $($result.RemovedFileCount)" -ForegroundColor Green
Write-Host "Removed directories : $($result.RemovedDirectoryCount)" -ForegroundColor Green
Write-Host "Declined projects   : $($result.DeclinedProjectCount)" -ForegroundColor DarkYellow
if ($WhatIfPreference) {
    Write-Host ("Would reclaim       : {0:N2} GB" -f ($result.WhatIfBytes / 1GB)) -ForegroundColor Cyan
}
else {
    Write-Host ("Space reclaimed     : {0:N2} GB" -f ($result.ReclaimedBytes / 1GB)) -ForegroundColor Cyan
}
