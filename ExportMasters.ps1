Write-Host "--- WBPP Master Exporter v6 (JSON Project Meta) ---" -ForegroundColor Cyan

# Запросить путь к project_meta.json
$metaPath = (Read-Host "Paste path to project_meta.json (e.g. D:\Astro\M_101\Season-Scope\project_meta.json)").Trim('"')
if (!(Test-Path $metaPath)) {
    Write-Host "[!] Error: project_meta.json not found!" -ForegroundColor Red; exit
}

# Прочитать и распарсить JSON
$meta = Get-Content $metaPath -Raw | ConvertFrom-Json
$baseZ = "Z:\AstroPhoto\Calibration"

$pixPath = $meta.PixPath
$masterPath = "$pixPath\master"
if (!(Test-Path $masterPath)) {
    Write-Host "[!] Error: Master folder not found! ($masterPath)" -ForegroundColor Red; exit
}

$res = "6248x4176" # Можно добавить в meta при желании

$masters = Get-ChildItem $masterPath -Filter "*.xisf"

foreach ($cam in $meta.Cameras) {
    $camFull = $cam.Name
    $telSetup = $meta.Scope
    Write-Host "`n--- Exporting for camera: $camFull ---" -ForegroundColor Cyan
    $mastersForCam = $masters | Where-Object { $_.Name -like "*$camFull*" }
    foreach ($m in $mastersForCam) {
        $name = $m.Name
        $targetSub = ""
        $newName = ""

        # Парсинг метаданных из длинного имени WBPP
        $gain = if ($name -match "GAIN-(\d+)") { $Matches[1] } else { "120" }
        $filt = if ($name -match "FILTER-([^_]+)") { $Matches[1] } else { "L" }
        $expStr = if ($name -match "EXP-([\d\.]+s?)") { $Matches[1] } else { "300s" }
        $sess = if ($name -match "SESSION-([\d\.]+)") { $Matches[1] } else { "Unknown" }
        $temp = if ($name -match "TEMP-([\-\d\.]+C)") { $Matches[1] } else { "-20C" }

        # Числовое значение экспозиции для логики Flat-Darks
        $expNum = [double]($expStr -replace 's', '')

        if ($name -match "masterFlat") {
            $targetSub = "Master\flats\$telSetup\$sess $filt 0deg"
            $newName = "masterFlat_BIN-1_${res}_FILTER-${filt}.xisf"
        }
        elseif ($name -match "masterDark") {
            if ($expNum -lt 10) {
                # Это Flat-Dark (экспозиция < 10 сек)
                $targetSub = "Master\flat-darks\Gain$gain\$temp\$($expNum)s"
                $newName = "masterDark_BIN-1_${res}_EXPOSURE-$($expNum)s.xisf"
            } else {
                # Это обычный Dark
                $targetSub = "Master\darks\Gain$gain\$temp\$($expNum)s"
                $newName = "masterDark_BIN-1_${res}_EXPOSURE-$($expNum).00s.xisf"
            }
        }
        elseif ($name -match "masterBias") {
            $targetSub = "Master\biases\Gain$gain\$temp"
            $newName = "masterBias_BIN-1_${res}.xisf"
        }

        if ($targetSub) {
            $destFolder = Join-Path "$baseZ\$camFull" $targetSub
            $destFile = Join-Path $destFolder $newName

            if (Test-Path $destFile) {
                $leftPart = "[EXISTS] $newName".PadRight(50)
                $fullTarget = Join-Path $camFull $targetSub
                Write-Host "$leftPart --> $fullTarget" -ForegroundColor Yellow
            } else {
                if (!(Test-Path $destFolder)) { New-Item -ItemType Directory -Path $destFolder -Force | Out-Null }
                Copy-Item -Path $m.FullName -Destination $destFile -Force
                
                $leftPart = "[SAVED]  $newName".PadRight(50)
                $fullTarget = Join-Path $camFull $targetSub
                Write-Host "$leftPart --> $fullTarget" -ForegroundColor Green
            }
        }
    }
}

Write-Host "`nExport process finished!"
