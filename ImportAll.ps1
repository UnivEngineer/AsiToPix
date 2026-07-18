[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$ImportRoot = "",

    [string]$AstroPhotoRoot = "",

    [string]$SeasonName = "",

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

Write-Host "--- ASIAir BATCH IMPORT ---" -ForegroundColor Cyan

if ([string]::IsNullOrWhiteSpace($AstroPhotoRoot)) {
    Import-Module $pathsModule -Force
    $AstroPhotoRoot = Resolve-AstroPhotoRoot
}
Write-AsiToPixCyrillicPathWarning -Path $AstroPhotoRoot -Context "AstroPhoto root"

if (-not (Test-Path -LiteralPath $AstroPhotoRoot -PathType Container)) {
    Write-Host "[!] AstroPhoto root not found: $AstroPhotoRoot" -ForegroundColor Red
    exit 1
}

if ([string]::IsNullOrWhiteSpace($ImportRoot)) {
    $ImportRoot = Join-Path -Path $AstroPhotoRoot -ChildPath "Import"
}
Write-AsiToPixCyrillicPathWarning -Path $ImportRoot -Context "Import root"

if (-not (Test-Path -LiteralPath $ImportRoot -PathType Container)) {
    Write-Host "[!] Import root not found: $ImportRoot" -ForegroundColor Red
    exit 1
}

$sessions = @(Find-AsiToPixImportSession -ImportRoot $ImportRoot)
if ($sessions.Count -eq 0) {
    Write-Host "[INFO] No importable light sessions found under: $ImportRoot" -ForegroundColor Yellow
    exit 0
}

Write-Host "`nFound import sessions:" -ForegroundColor Cyan
for ($i = 0; $i -lt $sessions.Count; $i++) {
    $session = $sessions[$i]
    Write-Host (" [{0}] {1} / {2}: {3} file(s)" -f ($i + 1), $session.DetectedSetupName, $session.DetectedObject, $session.FileCount) -ForegroundColor White
}

if ([string]::IsNullOrWhiteSpace($SeasonName)) {
    $defaultSeason = (Get-Date).Year.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $answer = (Read-Host "Enter destination season/group name for all imports [$defaultSeason]").Trim()
    $SeasonName = if ([string]::IsNullOrWhiteSpace($answer)) { $defaultSeason } else { $answer }
}
$SeasonName = ConvertTo-AsiToPixPathSegment -Value $SeasonName -ValueName "season/group name"

$resolvedImportMode = Read-AsiToPixImportMode -ImportMode $ImportMode
$objectNameByDetected = @{}
$setupNameBySource = @{}
$plans = @()

foreach ($session in $sessions) {
    $objectName = ""
    if ($objectNameByDetected.ContainsKey($session.DetectedObject)) {
        $objectName = $objectNameByDetected[$session.DetectedObject]
    }

    $setupName = ""
    if ($setupNameBySource.ContainsKey($session.SetupSourcePath)) {
        $setupName = $setupNameBySource[$session.SetupSourcePath]
    }

    $plan = Get-AsiToPixImportPlan `
        -SourcePath $session.SourcePath `
        -AstroPhotoRoot $AstroPhotoRoot `
        -ObjectName $objectName `
        -SeasonName $SeasonName `
        -SetupName $setupName `
        -ImportMode $resolvedImportMode

    if (-not $objectNameByDetected.ContainsKey($session.DetectedObject)) {
        $objectNameByDetected[$session.DetectedObject] = $plan.ObjectName
    }
    if (-not $setupNameBySource.ContainsKey($session.SetupSourcePath)) {
        $setupNameBySource[$session.SetupSourcePath] = $plan.SetupName
    }

    $plans += $plan
}

Write-Host "`nBatch import plan:" -ForegroundColor Cyan
for ($i = 0; $i -lt $plans.Count; $i++) {
    $plan = $plans[$i]
    $fileCount = @($plan.ParsedFiles).Count
    $groups = @($plan.ParsedFiles |
        Group-Object -Property FilterName, DestinationNightFolder |
        ForEach-Object {
            $first = $_.Group[0]
            "$($first.FilterName)/$($first.DestinationNightFolder):$($_.Count)"
        })
    Write-Host (" [{0}] {1} / {2} / {3}: {4} file(s)" -f ($i + 1), $plan.ObjectName, $plan.SeasonName, $plan.SetupName, $fileCount) -ForegroundColor White
    Write-Host "     $($groups -join ', ')" -ForegroundColor DarkGray
}

$confirm = (Read-Host "Apply all import plans? (Y/n)").Trim()
if (-not ([string]::IsNullOrWhiteSpace($confirm) -or $confirm[0] -in @([char]'y', [char]'Y', [char]0x0434, [char]0x0414))) {
    Write-Host "[INFO] Batch import cancelled." -ForegroundColor Yellow
    exit 0
}

$totalImported = 0
$totalExisting = 0
$totalTrash = 0

foreach ($plan in $plans) {
    Write-Host "`n-- $($plan.ObjectName) / $($plan.SetupName) --" -ForegroundColor Cyan
    $result = Invoke-AsiToPixImportPlan -Plan $plan -WhatIf:$WhatIfPreference
    $totalImported += $result.Imported
    $totalExisting += $result.AlreadyInGood
    $totalTrash += $result.PreservedTrash
}

Write-Host "`n[DONE] Batch import finished." -ForegroundColor Cyan
if ($resolvedImportMode -eq "Symlink") {
    Write-Host "  Linked:           $totalImported" -ForegroundColor White
} else {
    Write-Host "  Copied:           $totalImported" -ForegroundColor White
}
Write-Host "  Already in Good:  $totalExisting" -ForegroundColor White
Write-Host "  Preserved Trash:  $totalTrash" -ForegroundColor White
