# Run as Admin for Symlinks
fsutil behavior set SymlinkEvaluation L2R:1 | Out-Null
$baseZ = "Z:\AstroPhoto"; $localBase = "D:\Astro"

Write-Host "--- ASIAir SMART SCANNER v28 ---" -ForegroundColor Cyan

# 1. AUTO-DETECT
$inputPath = (Read-Host "Paste path to lights folder (or any .fit file inside)").Trim('"')

# If a file path was given, use its parent folder
if ($inputPath -match '\.\w+$') { $inputPath = Split-Path $inputPath -Parent }

# Parse the path for project metadata
if ($inputPath -match 'ASIAir\\(?<obj>[^\\]+)\\(?<season>[^\\]+)\\(?<tel>.*?) @ (?<cam>ASI\d+[^\\]*)') {
    $parsedObj      = $Matches['obj']
    $parsedSeason   = $Matches['season']
    $parsedTel      = $Matches['tel']
    $parsedCamShort = $Matches['cam']
} else { Write-Host "Path parse error! Expected: ...\ASIAir\<Object>\<Season>\<Tel> @ <Cam>\..." -ForegroundColor Red; exit }

# ── Interactive parameter confirmation ───────────────────────────────────────
function Confirm-Param {
    param([string]$Label, [string]$Detected)
    # Print label, arrow and detected value on one line, then prompt on same line
    $labelPad = $Label.PadRight(10)
    $detPad = $Detected.PadRight(35)
    Write-Host "  $labelPad -> " -ForegroundColor DarkGray -NoNewline
    Write-Host $detPad -ForegroundColor Yellow -NoNewline
    Write-Host "  >:" -ForegroundColor DarkGray -NoNewline
    $ans = Read-Host
    if ($ans -eq $null -or $ans.Trim() -eq "") { return $Detected } else { return $ans.Trim() }
}

Write-Host "`n-- Detected parameters -- (Enter to accept, or type override)" -ForegroundColor Cyan
# Print each parameter label and detected value; Confirm-Param will prompt on same line
Write-Host "" -NoNewline
$obj      = Confirm-Param "Object"  $parsedObj
$season   = Confirm-Param "Season"  $parsedSeason
$telSetup = Confirm-Param "Scope"   $parsedTel
Write-Host ""

# camShort = path segment from folder name (e.g. ASI2600, may lack MM/MC suffix)
$camShort = $parsedCamShort

# Structure
$safeObj = $obj.Replace(" ", "_");
$safeSetup = "$($season)_$($telSetup)_$($camShort)".Replace(" ", "_").Replace("@","").Replace("__","_")
$setupRoot = Join-Path (Join-Path $localBase $safeObj) $safeSetup
$sourcePath = Join-Path $setupRoot "Source"; $pixPath = Join-Path $setupRoot "Pix"

if (Test-Path $sourcePath) {
    if ((Read-Host "Clean [Source] folder? (y/n)") -match "^[yYdD]") { Remove-Item $sourcePath -Recurse -Force }
}
foreach ($p in @($sourcePath, $pixPath)) { if (!(Test-Path $p)) { New-Item -ItemType Directory -Path $p -Force | Out-Null } }

$lightsRoot = "$baseZ\ASIAir\$obj\$season\$telSetup @ $camShort\Good"
$pendingLinks = @()
 # Camera mapping summary: collect observed raw filters, sanitized filters and targets per camera
 $camMap = @{}

# Filter mapping: raw filter name from filename -> normalized output filter + WBPP target group
# OSC (*MC): IRC/L/Triband -> L/RGB | None -> None/RGB | HO/UHC -> HO/HO | SO -> SO/SO
# Mono (*MM): L -> L/RGB | None -> None/RGB | R/G/B -> R,G,B/RGB | H/Ha -> H/H | S/SII -> S/S | O/OIII -> O/O
function Get-FilterMap($rawFilt, $camType) {
    $r = $rawFilt.Trim()
    if ($camType -eq "OSC") {
        switch -Regex ($r) {
            '^(IRC|Triband)$'      { return @{ Filt="L";    Target="RGB" } }
            '^L$'                  { return @{ Filt="L";    Target="RGB" } }
            '^None$'               { return @{ Filt="None"; Target="RGB" } }
            '^(HO|UHC)$'           { return @{ Filt="HO";   Target="HO"  } }
            '^SO$'                 { return @{ Filt="SO";   Target="SO"  } }
            default                { return @{ Filt=$r;     Target=$r    } }
        }
    } else {
        # Mono
        switch -Regex ($r) {
            '^L$'                  { return @{ Filt="L";    Target="L" } }
            '^None$'               { return @{ Filt="None"; Target="None" } }
            '^R$'                  { return @{ Filt="R";    Target="R" } }
            '^G$'                  { return @{ Filt="G";    Target="G" } }
            '^B$'                  { return @{ Filt="B";    Target="B" } }
            '^(H|Ha)$'             { return @{ Filt="H";    Target="H"   } }
            '^(S|SII)$'            { return @{ Filt="S";    Target="S"   } }
            '^(O|OIII|OII)$'       { return @{ Filt="O";    Target="O"   } }
            default                { return @{ Filt=$r;     Target=$r    } }
        }
    }
}

# Calibration Path Helper - $cb = calibration base for this session's camera
function Get-CalibPath($type, $gain, $rawTemp, $expValue, $cb) {
    $sub = switch($type) { "Darks" {"darks"} "Biases" {"biases"} "FlatDarks" {"flat-darks"} }
    $numTemp = [double]($rawTemp -replace '[^-\d\.]', '')
    $targetTempFolder = "$([Math]::Round($numTemp / 5.0) * 5)C"
    
    foreach ($m in @("Master", "Source")) {
        $base = "$cb\$m\$sub\Gain$gain\$targetTempFolder"
        if (Test-Path $base) {
            $finalPath = $base
            if ($expValue) {
                $cleanExp = ($expValue -replace '[^0-9\.]', '').Replace(".0","")
                $expDir = Get-ChildItem $base -Directory | Where-Object { $_.Name -like "$cleanExp*" } | Select-Object -First 1
                if ($expDir) { $finalPath = $expDir.FullName } else { continue }
            }
            $display = $finalPath.Replace($cb, "").TrimStart("\")
            return @{Path=$finalPath; Mode=$m; Display="$display"}
        }
    }
    return $null
}

# SCANNER - iterates filter-group folders (L, RGB, Ha...) then date sub-folders
foreach ($fFolder in (Get-ChildItem $lightsRoot -Directory)) {
    foreach ($dFolder in (Get-ChildItem $fFolder.FullName -Directory)) {
        $date = ($dFolder.Name -split ' ')[0]
        $sample = Get-ChildItem $dFolder.FullName -Filter "*.fit*" | Select-Object -First 1
        if (!$sample) { continue }

        $fName = $sample.Name

        # 1. DETECT CAMERA FROM THIS SESSION'S FITS FILENAME
        # Pattern: ..._2600MM_... or ..._2600MC_... or ..._585MC_... etc.
        $sessionCamFull = $camShort   # fallback to path-level short name
        if ($fName -match '_(?<camf>(?:ASI)?\d{3,4}M[MCPmcp]+)_') {
            $sessionCamFull = ("ASI" + $Matches['camf']) -replace '^ASIASI','ASI'
        }
        $sessionCamType  = if ($sessionCamFull -match "MM") { "Mono" } else { "OSC" }
        $sessionCalibBase = "$baseZ\Calibration\$sessionCamFull"
        $camInFile = $sessionCamFull -replace '^ASI',''   # e.g. "2600MM", "2600MC", "585MC"

        # Ensure camera entry exists in map
        if (-not $camMap.ContainsKey($sessionCamFull)) {
            $camMap[$sessionCamFull] = @{ Observed = @(); Sanitized = @(); Targets = @(); Mappings = @() }
        }

        # 2. EXTRACT RAW FILTER FROM FILENAME
        # Pattern: ..._<CamModel>_<Filter>_gain...  e.g. _2600MC_None_gain120_
        $rawFilt = "Unknown"
        if ($fName -match "_$( [regex]::Escape($camInFile) )_(?<realFilt>[^_]+)_gain") {
            $rawFilt = $Matches['realFilt']
        }

        # 3. MAP TO NORMALIZED FILTER + WBPP TARGET
        $fm        = Get-FilterMap $rawFilt $sessionCamType
        $filt      = $fm.Filt      # normalized filter used for folder name and flat matching
        $targetGrp = $fm.Target    # WBPP integration target keyword

        # Record mapping info for summary
        $camMap[$sessionCamFull].Observed += $rawFilt
        $camMap[$sessionCamFull].Sanitized += $filt
        $camMap[$sessionCamFull].Targets += $targetGrp
        $camMap[$sessionCamFull].Mappings += ("$rawFilt -> $filt -> $targetGrp")

        # 3. METADATA PARSING (Angle fix)
        $valExp  = if ($fName -match "_(\d+\.?\d*)s_") { $Matches[1] } else { "300" }
        $curExp  = $valExp.Replace(".0","") + "s"
        $curGain = if ($fName -match "_gain(\d+)_") { $Matches[1] } else { "120" }
        $curTemp = if ($fName -match "_(-?\d+\.?\d*)C_") { $Matches[1] } else { "-20" }
        
        # Strip 'deg' before converting to INT
        $curAngRaw = if ($fName -match "_(\d+)deg_") { $Matches[1] } else { $null }
        $numCurAng = if ($curAngRaw) { [int]$curAngRaw } else { -999 }
        $angDisp = if ($numCurAng -eq -999) { "Unknown" } else { "${numCurAng}deg" }

        # --- Add newline before session header for readability ---
        Write-Host ""
        Write-Host "  [$sessionCamFull/$sessionCamType] $date | Filter: $rawFilt -> $filt | Target: $targetGrp | Angle: $angDisp" -ForegroundColor Green

        $roundT = "$([Math]::Round([double]$curTemp / 5.0) * 5)C"
        $sTag = "Session_${date}_Filter_${filt}_Target_${targetGrp}_Gain_${curGain}_Exp_${curExp}_Cam_${sessionCamFull}"
        $cTag = "Gain_${curGain}_Temp_${roundT}_Exp_${curExp}_Cam_${sessionCamFull}"

        # Always add CAM to tag for all types
        $pendingLinks += [PSCustomObject]@{ Type="Lights"; Tag=$sTag; Src=$dFolder.FullName; Display="Good\$($fFolder.Name)\$date"; Cam=$sessionCamFull }

        # --- FLATS SEARCH ---
        # Build list of all raw filter names that map to the same normalized filter,
        # so a flat shot as "IRC" is accepted for lights shot as "L" (and vice versa),
        # but "None" and "L" are kept separate (different flat corrections).
        $flatAliases = @()
        $flatAliases += $rawFilt
        $flatAliases += $filt
        # Add known synonyms for the normalized filter
        switch ($filt) {
            "L"    { $flatAliases += "IRC"; $flatAliases += "Triband" }
            "H"    { $flatAliases += "Ha" }
            "S"    { $flatAliases += "SII" }
            "O"    { $flatAliases += "OIII" }
            "HO"   { $flatAliases += "UHC" }
        }
        $flatAliases = $flatAliases | Select-Object -Unique

        # --- Robust filter match: look for filter as separate word or after date ---
        function FlatFolderMatchesFilter($folderName, $aliases) {
            $datePart = $folderName.Substring(0,8)
            $rest = $folderName.Substring(8).TrimStart()
            $filtPart = $null
            if ($rest -match '^(filt)?([^ _]+)') { $filtPart = $Matches[2] }
            <#
            Write-Host "[DEBUG] FlatFolder: $folderName | Parsed date: $datePart | Parsed filter: $filtPart | Aliases: $($aliases -join ', ')" -ForegroundColor DarkGray
            #>
            foreach ($alias in $aliases) {
                if ($filtPart -eq $alias) { <#Write-Host "[DEBUG]   MATCH: $folderName <-> $alias" -ForegroundColor Green;#> return $true }
            }
            <#Write-Host "[DEBUG]   SKIP: $folderName (no filter match)" -ForegroundColor Yellow#>
            return $false
        }

        # Known filters for validation
        $knownFilters = @('L','None','R','G','B','H','Ha','S','SII','O','OIII','OII','IRC','Triband','HO','UHC','SO')

        # Helper: try to parse date as yy.MM.dd or dd.MM.yy
        function ParseDate($dateStr) {
            $formats = @('yy.MM.dd','dd.MM.yy')
            foreach ($fmt in $formats) {
                try { return [datetime]::ParseExact($dateStr, $fmt, $null) } catch {} 
            }
            return $null
        }

        # Collect all candidate flats from Master and Source and always show menu for user selection
        $avail = @()
        # Write-Host "[DEBUG] Scanning flat folders in: $sessionCalibBase\*\flats\$telSetup" -ForegroundColor Cyan
        foreach($m in @("Master","Source")) {
            $p = "$sessionCalibBase\$m\flats\$telSetup"
            if (Test-Path $p) {
                Get-ChildItem $p -Directory | ForEach-Object {
                    $item = $_ | Select-Object -Property Name, FullName
                    $folderName = $item.Name
                    $date = $folderName.Substring(0,8)
                    $rest = $folderName.Substring(8).TrimStart()
                    $filt = $null; $ang = $null; $warn = $false
                    if ($rest -match '^(filt)?([^ _]+)') { $filt = $Matches[2] }
                    else { $warn = $true; Write-Host "[!] Flat folder '$folderName' does not match expected pattern (date filt<Filter> or date <Filter>)" -ForegroundColor Red }
                    # Validate date
                    $dateOk = $false
                    $d = ParseDate $date
                    if ($d) { $dateOk = $true } else { $dateOk = $false }
                    if (-not $dateOk) { $warn = $true; Write-Host "[!] Invalid date in flat folder: $folderName" -ForegroundColor Red }
                    # Validate filter
                    if ($filt -and ($knownFilters -notcontains $filt)) { $warn = $true; Write-Host "[!] Unknown filter '$filt' in flat folder: $folderName" -ForegroundColor Red }
                    # Validate angle if present (optional)
                    if ($folderName -match '(\d+)deg') { $ang = $Matches[1] } else { $ang = $null }
                    if ($ang -and ($ang -notmatch '^\d+$')) { $warn = $true; Write-Host "[!] Invalid angle in flat folder: $folderName" -ForegroundColor Red }
                    # Write-Host "[DEBUG] Found flat folder: $folderName | Filter: $filt | Date: $date | Angle: $ang" -ForegroundColor Magenta
                    $obj = [PSCustomObject]@{ Name = $item.Name; FullName = $item.FullName; Origin = $m }
                    $avail += $obj
                }
            }
        }

        if ($avail.Count -eq 0) {
            $lightsDisplay = "Good\$($fFolder.Name)\$date"
            Write-Host "`n[!] Missing Flats for $lightsDisplay (CalibRoot: $sessionCalibBase) (Original: $rawFilt, Target: $targetGrp, Angle: $angDisp)" -ForegroundColor Yellow
        } else {
            # Determine scores and default selection
            $defaultIdx = -1
            for ($i=0; $i -lt $avail.Count; $i++) {
                $fDirName = $avail[$i].Name
                $notBad = ($fDirName -notmatch "\(")
                # date parse from folder name (dd.MM.yy or dd.MM.yyyy)
                $candDateStr = if ($fDirName -match '(?<!\d)(\d{2}\.\d{2}\.\d{2,4})(?!\d)') { $Matches[1] } else { $null }
                $candDate = $null
                if ($candDateStr) {
                    try {
                        if ($candDateStr -match '\.\d{2}$') {
                            $parts = $candDateStr -split '\.'
                            $y = [int]$parts[2]
                            if ($y -lt 50) { $y = 2000 + $y } else { $y = 1900 + $y }
                            $candDate = Get-Date "$($parts[0])/$($parts[1])/$y"
                        } else {
                            $candDate = Get-Date $candDateStr
                        }
                    } catch { $candDate = $null }
                }
                # session date parse
                $sessDate = $null
                if ($date -match '(?<!\d)(\d{2}\.\d{2}\.\d{2,4})(?!\d)') {
                    $sd = $Matches[1]
                    try {
                        if ($sd -match '\.\d{2}$') {
                            $parts = $sd -split '\.'
                            $y = [int]$parts[2]
                            if ($y -lt 50) { $y = 2000 + $y } else { $y = 1900 + $y }
                            $sessDate = Get-Date "$($parts[0])/$($parts[1])/$y"
                        } else { $sessDate = Get-Date $sd }
                    } catch { $sessDate = $null }
                }
                $dateDiff = 9999
                if ($candDate -and $sessDate) { $dateDiff = [Math]::Abs((New-TimeSpan -Start $candDate -End $sessDate).Days) }

                $fAng = if ($fDirName -match "(\d+)deg") { [int]$Matches[1] } else { -888 }
                $angOk = ($numCurAng -ne -999 -and $fAng -ne -888 -and [Math]::Abs($numCurAng - $fAng) -le 2)
                $filtOk = FlatFolderMatchesFilter $fDirName $flatAliases

                $score = 0
                if ($filtOk) { $score += 10 }
                if ($angOk) { $score += 5 }
                if ($dateDiff -le 2) { $score += 3 }

                $avail[$i] | Add-Member -NotePropertyName Score -NotePropertyValue $score -Force
                $avail[$i] | Add-Member -NotePropertyName DateDiff -NotePropertyValue $dateDiff -Force
                $avail[$i] | Add-Member -NotePropertyName Ang -NotePropertyValue $fAng -Force
                $avail[$i] | Add-Member -NotePropertyName FiltOk -NotePropertyValue $filtOk -Force

                if ($defaultIdx -eq -1 -and $notBad -and $filtOk -and $dateDiff -le 0) { $defaultIdx = $i }
            }

            Write-Host "Select flats for session (Good\$($fFolder.Name)\$date) from calib root: $sessionCalibBase" -ForegroundColor Cyan
            $shownAny = $false
            $fullGreenIdx = -1
            for ($i=0; $i -lt $avail.Count; $i++) {
                $ent = $avail[$i]
                if (-not $ent.FiltOk) { continue }
                $shownAny = $true
                # Цвета для даты и угла
                $dateColor = 'White'; $angColor = 'White'; $rowColor = 'DarkGray'
                $dateMatch = $ent.DateDiff -eq 0
                $dateNear = ($ent.DateDiff -le 2)
                $angMatch = ($numCurAng -ne -999 -and $ent.Ang -ne -888 -and [Math]::Abs($numCurAng - $ent.Ang) -le 0)
                $angNear = ($numCurAng -ne -999 -and $ent.Ang -ne -888 -and [Math]::Abs($numCurAng - $ent.Ang) -le 2)
                if ($dateMatch) { $dateColor = 'Green' } elseif ($dateNear) { $dateColor = 'Yellow' }
                if ($angMatch) { $angColor = 'Green' } elseif ($angNear) { $angColor = 'Yellow' }
                # Определяем цвет всей строки
                if ($dateMatch -and $angMatch) { $rowColor = 'Green' }
                elseif ($dateNear -or $angNear) { $rowColor = 'Yellow' }
                else { $rowColor = 'DarkGray' }
                # Дефолтный индекс — первый full match
                if ($fullGreenIdx -eq -1 -and $rowColor -eq 'Green') { $fullGreenIdx = $i }
                $relPath = $ent.FullName.Replace($sessionCalibBase, '').TrimStart('\')
                # Формируем строку с цветными датой и углом
                $idxStr = "[$i]".PadRight(5)
                $origStr = "[$($ent.Origin)]".PadRight(10)
                $nameParts = $ent.Name -split ' '
                $dateStr = $nameParts[0]
                $restStr = $ent.Name.Substring($dateStr.Length).TrimStart()
                if ($ent.Ang -ne -888) { $angStr = "$($ent.Ang)deg" } else { $angStr = "" }
                # Удалить угол из restStr, если он есть
                if ($angStr -ne "") {
                    $restStr = $restStr -replace "\s*\b$($ent.Ang)deg\b", ""
                }
                # Динамическое выравнивание стрелки
                $leftPart = " $idxStr $origStr $dateStr $restStr $angStr"
                $arrowCol = 45
                $arrowPad = ""
                if ($leftPart.Length -lt $arrowCol) { $arrowPad = ' ' * ($arrowCol - $leftPart.Length) }
                $arrowPad += "<--"
                # Выводим строку с цветными датой и углом
                if ($rowColor -eq 'Green') {
                    Write-Host (" $idxStr $origStr ") -NoNewline -ForegroundColor White
                    Write-Host $dateStr -NoNewline -ForegroundColor Green
                    Write-Host (" $restStr ") -NoNewline -ForegroundColor White
                    if ($angStr -ne "") { Write-Host $angStr -NoNewline -ForegroundColor Green }
                    Write-Host ("$arrowPad $relPath") -ForegroundColor Green
                } elseif ($rowColor -eq 'Yellow') {
                    Write-Host (" $idxStr $origStr ") -NoNewline -ForegroundColor White
                    Write-Host $dateStr -NoNewline -ForegroundColor $dateColor
                    Write-Host (" $restStr ") -NoNewline -ForegroundColor White
                    if ($angStr -ne "") { Write-Host $angStr -NoNewline -ForegroundColor $angColor }
                    Write-Host ("$arrowPad $relPath") -ForegroundColor Yellow
                } else {
                    Write-Host (" $idxStr $origStr $($ent.Name)$arrowPad $relPath") -ForegroundColor DarkGray
                }
            }
            # Дефолтный индекс — первый full match (зелёная строка)
            if ($fullGreenIdx -ge 0) { $defaultIdx = $fullGreenIdx }
            if ($defaultIdx -ge 0) { $dtext = " $defaultIdx" } else { $dtext = " - skip" }
            $prompt = "Select Index (Enter to accept default$dtext)"
            $idx = Read-Host -Prompt $prompt
            Write-Host ""
            if ($idx -match '^\d+$') { $fFound = $avail[[int]$idx]; $foundIn = $fFound.Origin }
            elseif ($idx.Trim() -eq "" -and $defaultIdx -ge 0) { $fFound = $avail[$defaultIdx]; $foundIn = $fFound.Origin }
        }

        if ($fFound) {
            # Clean leading slash for flats display too
            $fDisp = "$foundIn\flats\$telSetup\$($fFound.Name)"
            $pendingLinks += [PSCustomObject]@{ Type="Flats"; Tag=$sTag; Src=$fFound.FullName; Display=$fDisp; Cam=$sessionCamFull }

            # Flat-Darks
            $fSample = Get-ChildItem $fFound.FullName -Filter "*.fit*" | Select-Object -First 1
            if ($fSample -and ($fSample.Name -match "_(\d+\.?\d*)m?s_")) {
                $fdExp = $Matches[1] + "s"
                $fd = Get-CalibPath "FlatDarks" $curGain $curTemp $fdExp $sessionCalibBase
                if ($fd) {
                    $fdTag = "Exp_${fdExp}_Gain_${curGain}_Target_${targetGrp}_Filter_${filt}_Cam_${sessionCamFull}"
                    $pendingLinks += [PSCustomObject]@{ Type="FlatDarks"; Tag=$fdTag; Src=$fd.Path; Display=$fd.Display; Cam=$sessionCamFull }
                }
            }
        }

        # --- DARKS & BIAS ---
        # Добавляем Target и Filter в симлинки для калибровочных кадров
        $dTag = "Gain_${curGain}_Temp_${roundT}_Exp_${curExp}_Target_${targetGrp}_Filter_${filt}_Cam_${sessionCamFull}"
        $bTag = "Gain_${curGain}_Temp_${roundT}_Target_${targetGrp}_Filter_${filt}_Cam_${sessionCamFull}"

        $d = Get-CalibPath "Darks" $curGain $curTemp $curExp $sessionCalibBase
        if ($d) { $pendingLinks += [PSCustomObject]@{ Type="Darks"; Tag=$dTag; Src=$d.Path; Display=$d.Display; Cam=$sessionCamFull } }
        
        $b = Get-CalibPath "Biases" $curGain $curTemp $null $sessionCalibBase
        if ($b) { $pendingLinks += [PSCustomObject]@{ Type="Biases"; Tag=$bTag; Src=$b.Path; Display=$b.Display; Cam=$sessionCamFull } }
    }
}

# Print camera mapping summary
Write-Host "`n--- CAMERA MAPPING SUMMARY ---" -ForegroundColor Cyan
foreach ($cam in $camMap.Keys) {
    $entry = $camMap[$cam]
    $obs = ($entry.Observed | Select-Object -Unique) -join ' | '
    $san = ($entry.Sanitized | Select-Object -Unique) -join ' | '
    $tgt = ($entry.Targets | Select-Object -Unique) -join ' | '
    Write-Host "$cam" -ForegroundColor Yellow
    Write-Host "  |-- Observations filters  | $obs"
    Write-Host "  |-- Sanitized filters     | $san"
    Write-Host "  |-- Target combine        | $tgt"
    Write-Host "  |-- Mappings:" -ForegroundColor DarkGray
    ($entry.Mappings | Select-Object -Unique) | ForEach-Object { Write-Host "      $_" -ForegroundColor Gray }
}

# --- REVIEW (v33 - Compact & Clean) ---
Write-Host "`n--- PROJECT TREE: $setupRoot ---" -ForegroundColor Cyan
$pendingLinks | Group-Object Type | ForEach-Object {
    Write-Host "[$($_.Name)]"  -ForegroundColor Yellow
    $_.Group | Sort-Object Tag, Cam | ForEach-Object {
        # Using a slightly smaller PadRight for better window fit
        $camPart = if ($_.Cam) { " [$($_.Cam)]" } else { "" }
        Write-Host " |- $($_.Tag.PadRight(65)) <-- $($_.Display)$camPart" -ForegroundColor Gray
    }
}

# --- FINAL DEBUG TABLE (commented out for clarity) ---
<#
Write-Host "`n--- DEBUG: ALL FLAT FOLDERS ---" -ForegroundColor Magenta
foreach($m in @("Master","Source")) {
    $p = "$sessionCalibBase\$m\flats\$telSetup"
    if (Test-Path $p) {
        Get-ChildItem $p -Directory | ForEach-Object {
            Write-Host $_.Name
        }
    }
}
#>

if ($pendingLinks.Count -gt 0) {
    # Write-Host "`nPlanned symlink targets:" -ForegroundColor Cyan
    foreach ($l in $pendingLinks) {
        $target = Join-Path $sourcePath "$($l.Type)\$($l.Tag)"
        if (Test-Path $target) {
            # If symlink exists, check if it points to the same source
            $existing = Get-Item $target -ErrorAction SilentlyContinue
            if ($existing -and $existing.LinkType -eq 'SymbolicLink' -and $existing.Target -ne $l.Src) {
                Write-Host "[!] Symlink $target already exists and points elsewhere!" -ForegroundColor Red
            }
        }
        # Write-Host "  -> $target  <= $($l.Src)"
    }
    if ((Read-Host "Create symlinks in Source? (y/n)") -match "^[yYdD]") {
        foreach ($l in $pendingLinks) {
            $target = Join-Path $sourcePath "$($l.Type)\$($l.Tag)"
            if (!(Test-Path $target)) { 
                New-Item -ItemType Directory -Path (Split-Path $target) -Force | Out-Null
                New-Item -ItemType SymbolicLink -Path $target -Value $l.Src | Out-Null 
            }
        }
        Write-Host "`n[DONE] Project Root: $setupRoot" -ForegroundColor Yellow
    }
}

# Reminder: which keywords and modes to add
Write-Host "`n--- Keywords to add (for WBPP) ---" -ForegroundColor Cyan
Write-Host "Preferred keywords and suggested mode:" -ForegroundColor Gray
Write-Host "  CAM      : camera identifier            (add to pre and post)" -ForegroundColor Yellow
Write-Host "  FILTER   : normalized filter name       (add to pre)" -ForegroundColor Yellow
Write-Host "  TARGET   : WBPP target group            (add to pre)" -ForegroundColor Yellow
Write-Host "  GAIN     : gain                         (add to pre)" -ForegroundColor Yellow
Write-Host "  TEMP     : sensor temperature           (add to pre)" -ForegroundColor Yellow
Write-Host "  EXP      : exposure time                (add to pre)" -ForegroundColor Yellow
Write-Host "If unsure, add CAM both pre and post. FILTER/TARGET should be present before integration selection (pre)." -ForegroundColor Gray
