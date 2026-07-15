# Combine Seasons Script
# Run as Admin for Symlinks
$environmentModule = Join-Path $PSScriptRoot "src\AsiToPix.Environment.psm1"
Import-Module $environmentModule -Force
Initialize-AsiToPixEnvironment

$projectMetadataModule = Join-Path $PSScriptRoot "src\AsiToPix.ProjectMetadata.psm1"
Import-Module $projectMetadataModule -Force

# Enable long path support
try {
    $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem"
    Set-ItemProperty -Path $regPath -Name "LongPathsEnabled" -Value 1 -ErrorAction SilentlyContinue
} catch {
    Write-Host "[INFO] Could not enable long path support in registry" -ForegroundColor DarkYellow
}

Write-Host "--- COMBINE SEASONS ---" -ForegroundColor Cyan

# Ask for the root folder containing season directories
$rootPath = (Read-Host "Enter path to root folder containing season directories").Trim('"')
Write-AsiToPixCyrillicPathWarning -Path $rootPath -Context "root path"

if (!(Test-Path $rootPath)) {
    Write-Host "[!] Root folder not found: $rootPath" -ForegroundColor Red
    exit
}

Write-Host "`nScanning for season folders in: $rootPath" -ForegroundColor Yellow

# Auto-detect season folders containing Source, excluding Combined
$availableSeasons = @()
$seasonFolders = @()

Get-ChildItem $rootPath -Directory | ForEach-Object {
    if ($_.Name -eq "Combined") {
        Write-Host "  Skipping Combined folder" -ForegroundColor DarkGray
        return
    }
    $sourcePath = Join-Path $_.FullName "Source"
    if (Test-Path $sourcePath) {
        $availableSeasons += @{
            Name = $_.Name
            Path = $sourcePath
        }
        Write-Host "  Found season: $($_.Name)" -ForegroundColor Green
    }
}

if ($availableSeasons.Count -eq 0) {
    Write-Host "[!] No season folders with Source subdirectories found!" -ForegroundColor Red
    Write-Host "    Expected structure: <RootPath>\<SeasonName>\Source\" -ForegroundColor Gray
    exit
}

# Let the user choose which seasons to combine
Write-Host "`n--- SELECT SEASONS TO COMBINE ---" -ForegroundColor Cyan
foreach ($season in $availableSeasons) {
    $choice = Read-Host "Add season '$($season.Name)' to combine? (Y/n)"
    if ($choice -eq '' -or $choice -match '^[yYдД]') {
        $seasonFolders += $season.Path
        Write-Host "  [v] Added: $($season.Name)" -ForegroundColor Green
    } else {
        Write-Host "  [x] Skipped: $($season.Name)" -ForegroundColor DarkGray
    }
}

if ($seasonFolders.Count -eq 0) {
    Write-Host "[!] No seasons selected for combining!" -ForegroundColor Red
    Write-Host "    At least one season must be selected." -ForegroundColor Gray
    exit
}

# Shared output folder for the combined project
$combineRoot = Join-Path $rootPath "Combined"
$combineSource = Join-Path $combineRoot "Source"
$combinePix = Join-Path $combineRoot "Pix"
Write-AsiToPixCyrillicPathWarning -Path $combineRoot -Context "combined project path"

# If Combined already exists, ask whether it should be cleaned
if (Test-Path $combineRoot) {
    $cleanChoice = Read-Host "`nCombined folder already exists. Clean it? (Y/n)"
    if ($cleanChoice -eq '' -or $cleanChoice -match '^[yYдД]') { 
        Write-Host "Cleaning Combined folder..." -ForegroundColor Yellow
        Remove-Item $combineRoot -Recurse -Force 
    }
}

# Create folders when they do not exist
if (!(Test-Path $combineSource)) { 
    New-Item -ItemType Directory -Path $combineSource -Force | Out-Null 
}
if (!(Test-Path $combinePix)) { 
    New-Item -ItemType Directory -Path $combinePix -Force | Out-Null 
}

Write-Host "`nCombining $($seasonFolders.Count) seasons into: $combineSource" -ForegroundColor Yellow

foreach ($seasonPath in $seasonFolders) {
    Write-AsiToPixCyrillicPathWarning -Path $seasonPath -Context "season source path"
    if (!(Test-Path $seasonPath)) {
        Write-Host "[!] Season folder not found: $seasonPath" -ForegroundColor Red
        continue
    }
    
    $seasonName = Split-Path (Split-Path $seasonPath -Parent) -Leaf
    Write-Host "`nProcessing season: $seasonName" -ForegroundColor Green
    
    # Iterate over frame type folders and normalize the legacy FlatDarks name.
    foreach ($typeFolder in (Get-ChildItem $seasonPath -Directory)) {
        $typeName = Get-AsiToPixProjectSourceFolderName -Type $typeFolder.Name
        $combineTypeFolder = Join-Path $combineSource $typeName
        
        if (!(Test-Path $combineTypeFolder)) {
            New-Item -ItemType Directory -Path $combineTypeFolder -Force | Out-Null
        }
        
        # Iterate over all symlinks in the frame type folder
        foreach ($symlink in (Get-ChildItem $typeFolder.FullName -Directory)) {
            $originalTarget = $null
            try {
                # Read the original symlink target
                $item = Get-Item $symlink.FullName
                if ($item.LinkType -eq 'SymbolicLink') {
                    # Target may be an array; use the first element
                    $targetArray = $item.Target
                    if ($targetArray -is [array]) {
                        $originalTarget = $targetArray[0]
                    } else {
                        $originalTarget = $targetArray
                    }
                } else {
                    # For ordinary folders, use their own path
                    $originalTarget = $symlink.FullName
                }
            } catch {
                Write-Host "[!] Error reading symlink: $($symlink.Name) - $($_.Exception.Message)" -ForegroundColor Red
                continue
            }
            
            if ($originalTarget -and (Test-Path $originalTarget)) {
                Write-AsiToPixCyrillicPathWarning -Path $originalTarget -Context "symlink target path"
                # Keep the original name; paths should already be unique by date/setup
                $newSymlinkPath = Join-Path $combineTypeFolder $symlink.Name
                
                if (!(Test-Path $newSymlinkPath)) {
                    try {
                        New-Item -ItemType SymbolicLink -Path $newSymlinkPath -Value $originalTarget -Force | Out-Null
                        Write-Host "  + $typeName\$($symlink.Name)" -ForegroundColor Gray
                    } catch {
                        Write-Host "[!] Failed to create symlink '$($symlink.Name)': $($_.Exception.Message)" -ForegroundColor Red
                        Write-Host "    Target was: $originalTarget" -ForegroundColor Yellow
                    }
                } else {
                    # This should not happen because paths are expected to be unique by date/setup
                    Write-Host "[!] CONFLICT: $typeName\$($symlink.Name) already exists! This shouldn't happen." -ForegroundColor Red
                    Write-Host "    Existing: $(Get-Item $newSymlinkPath | Select-Object -ExpandProperty Target)" -ForegroundColor Yellow
                    Write-Host "    New:      $originalTarget" -ForegroundColor Yellow
                }
            } else {
                Write-Host "[!] Target not found or invalid: $originalTarget" -ForegroundColor Red
            }
        }
    }
}

Write-Host "`n[DONE] Combined project root: $combineRoot" -ForegroundColor Yellow
Write-Host "Created folders:" -ForegroundColor Cyan
Write-Host "  - Source: $combineSource (for WBPP input)" -ForegroundColor Gray
Write-Host "  - Pix: $combinePix (for WBPP output)" -ForegroundColor Gray
Write-Host "`nYou can now use this folder for WBPP processing" -ForegroundColor Cyan
