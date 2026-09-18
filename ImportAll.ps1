[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [string]$ImportRoot = "",

    [string]$AstroPhotoRoot = "",

    [Alias("SourceRoot")]
    [string]$SourceAstroPhotoRoot = "",

    [Alias("DestinationRoot")]
    [string]$DestinationAstroPhotoRoot = "",

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
Import-Module $pathsModule -Force

Write-Host "--- ASIAir BATCH IMPORT ---" -ForegroundColor Cyan

if (-not [string]::IsNullOrWhiteSpace($AstroPhotoRoot) -and
    -not [string]::IsNullOrWhiteSpace($DestinationAstroPhotoRoot)) {
    $legacyDestinationRoot = (Resolve-Path -LiteralPath $AstroPhotoRoot -ErrorAction Stop).ProviderPath
    $namedDestinationRoot = (Resolve-Path -LiteralPath $DestinationAstroPhotoRoot -ErrorAction Stop).ProviderPath
    if (-not $legacyDestinationRoot.Equals($namedDestinationRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Specify only one destination root: -AstroPhotoRoot (legacy) or -DestinationAstroPhotoRoot."
    }
}

$sourceRootForImport = if ([string]::IsNullOrWhiteSpace($SourceAstroPhotoRoot)) {
    if ([string]::IsNullOrWhiteSpace($ImportRoot)) {
        Write-Host "[INFO] Select the light import source root (read from <selected root>\Import):" -ForegroundColor Cyan
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

if ([string]::IsNullOrWhiteSpace($sourceRootForImport)) {
    $sourceRootForImport = $AstroPhotoRoot
}

if ([string]::IsNullOrWhiteSpace($ImportRoot)) {
    $ImportRoot = Join-Path -Path $sourceRootForImport -ChildPath "Import"
}
Write-AsiToPixCyrillicPathWarning -Path $ImportRoot -Context "Import root"
Write-AsiToPixCyrillicPathWarning -Path $sourceRootForImport -Context "AstroPhoto source root"

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
