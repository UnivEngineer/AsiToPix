# Combine Seasons Script
# Run as Admin for Symlinks
fsutil behavior set SymlinkEvaluation L2R:1 | Out-Null

# Enable long path support
try {
    $regPath = "HKLM:\SYSTEM\CurrentControlSet\Control\FileSystem"
    Set-ItemProperty -Path $regPath -Name "LongPathsEnabled" -Value 1 -ErrorAction SilentlyContinue
} catch {
    Write-Host "[INFO] Could not enable long path support in registry" -ForegroundColor DarkYellow
}

Write-Host "--- COMBINE SEASONS ---" -ForegroundColor Cyan

# Запрашиваем корневую папку с сезонами
$rootPath = (Read-Host "Enter path to root folder containing season directories").Trim('"')

if (!(Test-Path $rootPath)) {
    Write-Host "[!] Root folder not found: $rootPath" -ForegroundColor Red
    exit
}

Write-Host "`nScanning for season folders in: $rootPath" -ForegroundColor Yellow

# Автоматически находим папки сезонов (содержащие папку Source), исключаем Combined
$seasonFolders = @()
Get-ChildItem $rootPath -Directory | ForEach-Object {
    if ($_.Name -eq "Combined") {
        # Write-Host "  Skipping Combined folder" -ForegroundColor DarkGray
        return
    }
    $sourcePath = Join-Path $_.FullName "Source"
    if (Test-Path $sourcePath) {
        $seasonFolders += $sourcePath
        Write-Host "  Found season: $($_.Name)" -ForegroundColor Green
    }
}

if ($seasonFolders.Count -eq 0) {
    Write-Host "[!] No season folders with Source subdirectories found!" -ForegroundColor Red
    Write-Host "    Expected structure: <RootPath>\<SeasonName>\Source\" -ForegroundColor Gray
    exit
}

# Общая папка для объединения
$combineRoot = Join-Path $rootPath "Combined"
$combineSource = Join-Path $combineRoot "Source"
$combinePix = Join-Path $combineRoot "Pix"

# Проверяем существование папки Combined и запрашиваем очистку если нужно
if (Test-Path $combineRoot) {
    $cleanChoice = Read-Host "`nCombined folder already exists. Clean it? (Y/n)"
    if ($cleanChoice -eq '' -or $cleanChoice -match '^[yYдД]') { 
        Write-Host "Cleaning Combined folder..." -ForegroundColor Yellow
        Remove-Item $combineRoot -Recurse -Force 
    }
}

# Создаем папки если не существуют
if (!(Test-Path $combineSource)) { 
    New-Item -ItemType Directory -Path $combineSource -Force | Out-Null 
}
if (!(Test-Path $combinePix)) { 
    New-Item -ItemType Directory -Path $combinePix -Force | Out-Null 
}

Write-Host "`nCombining $($seasonFolders.Count) seasons into: $combineSource" -ForegroundColor Yellow

foreach ($seasonPath in $seasonFolders) {
    if (!(Test-Path $seasonPath)) {
        Write-Host "[!] Season folder not found: $seasonPath" -ForegroundColor Red
        continue
    }
    
    $seasonName = Split-Path (Split-Path $seasonPath -Parent) -Leaf
    Write-Host "`nProcessing season: $seasonName" -ForegroundColor Green
    
    # Проходим по типам калибровки (Lights, Darks, Biases, Flats, FlatDarks)
    foreach ($typeFolder in (Get-ChildItem $seasonPath -Directory)) {
        $typeName = $typeFolder.Name
        $combineTypeFolder = Join-Path $combineSource $typeName
        
        if (!(Test-Path $combineTypeFolder)) {
            New-Item -ItemType Directory -Path $combineTypeFolder -Force | Out-Null
        }
        
        # Проходим по всем симлинкам в папке типа
        foreach ($symlink in (Get-ChildItem $typeFolder.FullName -Directory)) {
            $originalTarget = $null
            try {
                # Получаем оригинальный путь симлинка
                $item = Get-Item $symlink.FullName
                if ($item.LinkType -eq 'SymbolicLink') {
                    # Исправляем проблему с массивом - берем первый элемент
                    $targetArray = $item.Target
                    if ($targetArray -is [array]) {
                        $originalTarget = $targetArray[0]
                    } else {
                        $originalTarget = $targetArray
                    }
                } else {
                    # Если это обычная папка, используем её путь
                    $originalTarget = $symlink.FullName
                }
            } catch {
                Write-Host "[!] Error reading symlink: $($symlink.Name) - $($_.Exception.Message)" -ForegroundColor Red
                continue
            }
            
            if ($originalTarget -and (Test-Path $originalTarget)) {
                # Используем оригинальное имя - пути уже уникальные по дате/настройкам
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
                    # Это НЕ должно происходить, так как пути уникальны по дате/настройкам
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