Write-Host "--- WBPP Master Exporter v3 (Safety & Flat-Darks) ---" -ForegroundColor Cyan
$localProject = (Read-Host "Paste local project path (e.g. D:\Astro\M_101)").Trim('"')
$baseZ = "Z:\AstroPhoto\Calibration"

$masterPath = "$localProject\pix\master"
if (!(Test-Path $masterPath)) { 
    Write-Host "[!] Error: Master folder not found!" -ForegroundColor Red; exit 
}

$camFull = Read-Host "Full Camera Name (e.g. ASI2600MM)"
$telSetup = Read-Host "Telescope @ Reducer for Flats (e.g. LX200 @ 0.63x)"
$temp = "-20C" # Default temp
$res = "6248x4176" # Resolution for ASI2600

$masters = Get-ChildItem $masterPath -Filter "*.xisf"

foreach ($m in $masters) {
    $name = $m.Name
    $targetSub = ""
    $newName = ""

    # Парсинг метаданных из длинного имени WBPP
    $gain = if ($name -match "GAIN-(\d+)") { $Matches[1] } else { "120" }
    $filt = if ($name -match "FILTER-([^_]+)") { $Matches[1] } else { "L" }
    $expStr = if ($name -match "EXP-([\d\.]+s?)") { $Matches[1] } else { "300s" }
    $sess = if ($name -match "SESSION-([\d\.]+)") { $Matches[1] } else { "Unknown" }

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

        # ПРОВЕРКА НА ПЕРЕЗАПИСЬ
        $doCopy = $true
        if (Test-Path $destFile) {
            Write-Host "`n[!] FILE EXISTS: $newName" -ForegroundColor Yellow
            $choice = Read-Host "Overwrite on Z:? (y/n)"
            if ($choice -ne "y") { $doCopy = $false; Write-Host "   Skipped." -ForegroundColor Gray }
        }

        if ($doCopy) {
            if (!(Test-Path $destFolder)) { New-Item -ItemType Directory -Path $destFolder -Force | Out-Null }
            Copy-Item -Path $m.FullName -Destination $destFile -Force
            Write-Host " [SAVED] $newName" -ForegroundColor Green
            Write-Host "         To: $targetSub" -ForegroundColor Gray
        }
    }
}

Write-Host "`nExport process finished!" -ForegroundColor Yellow
